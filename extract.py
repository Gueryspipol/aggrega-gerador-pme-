"""Camada de extração: lê a Matriz (Excel) sem depender de posições fixas de colunas.

Princípios:
- as abas são localizadas por nome normalizado (CC / SC); SC pode não existir;
- cada operadora é um bloco que começa numa linha cuja coluna B é "OPERADORA";
- os planos de um bloco são pares (VIDAS, PER CAPITA) a partir da coluna C, e a
  quantidade de pares é detectada pelo cabeçalho (varia por cliente);
- as faixas etárias vão da linha abaixo do cabeçalho até a linha "Total".
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Optional

import openpyxl

from .models import Bloco, CenarioMatriz, Demografia, Matriz, Plano

FAIXA_INICIOS = [0, 19, 24, 29, 34, 39, 44, 49, 54, 59]
FAIXA_ROTULOS = ["00 - 18", "19 - 23", "24 - 28", "29 - 33", "34 - 38",
                 "39 - 43", "44 - 48", "49 - 53", "54 - 58", "59+"]

_PLACEHOLDER = re.compile(r"^plano( \d+)?$")
_PCT = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")


def norm(valor: Any) -> str:
    """minúsculas, sem acento, espaços colapsados."""
    if valor is None:
        return ""
    s = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(valor: Any) -> Optional[float]:
    if isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    return None  # '#DIV/0!', textos e vazios


def achar_aba(nomes: list[str], tipo: str) -> Optional[str]:
    """Localiza a aba do cenário 'CC' ou 'SC'. Devolve None se não existir."""
    padrao = (r"\b(c|com)\b.*copart" if tipo == "CC" else r"\b(s|sem)\b.*copart")
    for nome in nomes:
        if re.search(padrao, norm(nome)):
            return nome
    return None


class _Grade:
    """Acesso 1-indexado seguro a uma planilha lida em memória."""

    def __init__(self, ws):
        self.rows = list(ws.iter_rows(values_only=True))

    @property
    def n(self) -> int:
        return len(self.rows)

    def __call__(self, r: int, c: int) -> Any:
        try:
            return self.rows[r - 1][c - 1]
        except IndexError:
            return None


def _ler_bloco(g: _Grade, indice: int, r: int) -> Bloco:
    nome = str(g(r, 3) or "").strip()
    linha_planos, linha_cab = r + 1, r + 2

    colunas, c = [], 3
    while norm(g(linha_cab, c)) == "vidas":
        colunas.append(c)
        c += 2

    primeira = linha_cab + 1
    faixas, rr = [], primeira
    while not norm(g(rr, 2)).startswith("total"):
        faixas.append(str(g(rr, 2) or "").strip())
        rr += 1
        if rr > g.n or len(faixas) > 25:
            raise ValueError(f"Bloco '{nome}' (linha {r}): linha 'Total' não encontrada")
    linha_total = rr

    planos = []
    for c in colunas:
        rotulo = g(linha_planos, c)
        nome_plano = None if (rotulo is None or _PLACEHOLDER.match(norm(rotulo))) else str(rotulo).strip()
        planos.append(Plano(
            nome=nome_plano,
            vidas=int(_num(g(linha_total, c)) or 0),
            total=_num(g(linha_total, c + 1)) or 0.0,
            por_faixa=[_num(g(primeira + k, c + 1)) for k in range(len(faixas))],
        ))
    return Bloco(indice=indice, nome=nome, linha=r, faixas=faixas, planos=planos)


def _achar_rotulo(g: _Grade, rotulo: str) -> Optional[tuple[int, int]]:
    alvo = norm(rotulo)
    for r in range(1, g.n + 1):
        for c in range(1, len(g.rows[r - 1]) + 1):
            if norm(g(r, c)) == alvo:
                return r, c
    return None


def ler_cenario(ws, tipo: str) -> CenarioMatriz:
    g = _Grade(ws)
    linhas_op = [r for r in range(1, g.n + 1) if norm(g(r, 2)) == "operadora"]
    if not linhas_op:
        raise ValueError(f"Aba {ws.title!r}: nenhum bloco 'OPERADORA' encontrado")
    blocos = [_ler_bloco(g, i, r) for i, r in enumerate(linhas_op)]

    reajuste = None
    for c in range(1, len(g.rows[linhas_op[0] - 1]) + 1):
        v = g(linhas_op[0], c)
        if "reajuste" in norm(v):
            m = _PCT.search(str(v))
            if m:
                reajuste = float(m.group(1).replace(",", ".")) / 100
                break

    def valor_ao_lado(rotulo: str) -> Optional[float]:
        pos = _achar_rotulo(g, rotulo)
        return _num(g(pos[0], pos[1] + 1)) if pos else None

    return CenarioMatriz(
        tipo=tipo, blocos=blocos, reajuste_pct=reajuste,
        fatura_atual=valor_ao_lado("Fatura Atual"),
        fatura_reajustada=valor_ao_lado("Fatura Reajustada"),
    )


def ler_demografia(wb) -> Demografia:
    nome_base = next((n for n in wb.sheetnames if norm(n) in ("base", "base de dados")), None)
    if nome_base is None:
        raise ValueError("Aba 'Base' (ou 'Base de dados') não encontrada na Matriz")
    g = _Grade(wb[nome_base])
    cab = {norm(g(1, c)): c for c in range(1, len(g.rows[0]) + 1) if g(1, c) is not None}
    # 'Idade' repete; só precisamos destes três
    try:
        c_eleg, c_sexo, c_faixa = cab["elegibilidade"], cab["sexo"], cab["faixa etaria"]
    except KeyError as e:
        raise ValueError(f"Aba {nome_base!r}: coluna {e} não encontrada no cabeçalho") from None

    titular = {"Funcionários": 0, "Dependentes": 0}
    genero = {"M": 0, "F": 0}
    faixas = {ini: {"T": 0, "D": 0} for ini in FAIXA_INICIOS}
    for r in range(2, g.n + 1):
        eleg = norm(g(r, c_eleg)).upper()[:1]
        if eleg not in ("T", "D"):
            continue
        titular["Funcionários" if eleg == "T" else "Dependentes"] += 1
        sexo = norm(g(r, c_sexo)).upper()[:1]
        if sexo in genero:
            genero[sexo] += 1
        m = re.search(r"\d+", str(g(r, c_faixa) or ""))
        if m:
            ini = max(i for i in FAIXA_INICIOS if i <= int(m.group()))
            faixas[ini][eleg] += 1
    return Demografia(titularidade=titular, genero=genero, faixa_etaria=faixas)


def _cliente_do_arquivo(caminho: Path) -> str:
    base = re.split(r"_matriz", caminho.stem, flags=re.I)[0]
    return base.replace("_", " ").strip()


def carregar_matriz(caminho: str | Path, cliente: Optional[str] = None) -> Matriz:
    caminho = Path(caminho)
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=False)
    cenarios = {}
    for tipo in ("CC", "SC"):
        aba = achar_aba(wb.sheetnames, tipo)
        if aba:
            cenarios[tipo] = ler_cenario(wb[aba], tipo)
    if not cenarios:
        raise ValueError("Nenhuma aba de cenário (com/sem coparticipação) encontrada")
    return Matriz(
        caminho=str(caminho), cenarios=cenarios, demografia=ler_demografia(wb),
        cliente=cliente or _cliente_do_arquivo(caminho),
    )
