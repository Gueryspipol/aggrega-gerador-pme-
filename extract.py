"""Camada de extração: lê a Matriz (Excel) reconhecendo CONCEITOS, não posições.

O que o leitor procura (em qualquer coluna/linha, sem diferenciar maiúsculas/acentos):
- abas de cenário: as que têm blocos "OPERADORA" + colunas "VIDAS"; CC/SC identificados pelo nome da
  aba, por textos "com/sem coparticipação" dentro da aba ou, no limite, pela ordem das abas;
- bloco de operadora: começa numa célula "OPERADORA"; o nome é a 1ª célula preenchida à direita;
- planos: cada coluna "VIDAS" do cabeçalho é um plano (valor por pessoa na coluna ao lado);
  o nome do plano é a célula preenchida logo acima do cabeçalho;
- faixas etárias: linhas abaixo do cabeçalho até a linha que começa com "Total";
- situação atual do cliente: bloco cujo nome tem "atual" / "sem prestador" (senão, o primeiro);
- demografia: aba que tenha cabeçalhos de elegibilidade, sexo e faixa etária (ou idade).
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
_RE_OPERADORA = re.compile(r"^operadoras?\s*:?$")
_RE_VIDAS = re.compile(r"^(qtd\.?|qtde\.?|quant\.?|n[ºo°]?\.?|numero de)?\s*vidas$")
_RE_ATUAL = re.compile(r"atual|sem prestador|vigente")

SIN_ELEG = {"elegibilidade", "titularidade", "tipo", "tipo beneficiario", "tipo de beneficiario",
            "titular/dependente", "t/d"}
SIN_SEXO = {"sexo", "genero"}
SIN_FAIXA = {"faixa etaria", "faixa", "faixa etaria beneficiario"}
SIN_IDADE = {"idade", "idade atual"}


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


class _Grade:
    """Acesso 1-indexado seguro a uma planilha lida em memória."""

    def __init__(self, ws, max_linhas: int = 2000, max_colunas: int = 200):
        self.rows = [tuple(r) for r in ws.iter_rows(max_row=max_linhas, max_col=max_colunas,
                                                    values_only=True)]
        self.n = len(self.rows)
        self.ncols = max((len(r) for r in self.rows), default=0)

    def __call__(self, r: int, c: int) -> Any:
        try:
            return self.rows[r - 1][c - 1]
        except IndexError:
            return None

    def celulas(self, r1=1, r2=None, c1=1, c2=None):
        for r in range(max(1, r1), min(r2 or self.n, self.n) + 1):
            for c in range(max(1, c1), min(c2 or self.ncols, self.ncols) + 1):
                v = self(r, c)
                if v is not None and str(v).strip() != "":
                    yield r, c, v


# --------------------------------------------------------------------------- #
# Abas de cenário (CC / SC)
# --------------------------------------------------------------------------- #
def achar_aba(nomes: list[str], tipo: str) -> Optional[str]:
    """Aba CC/SC só pelo nome (usado como dica). Devolve None se nenhum nome casar."""
    padrao = (r"\b(c|com)\b.*copart" if tipo == "CC" else r"\b(s|sem)\b.*copart")
    for nome in nomes:
        if re.search(padrao, norm(nome)):
            return nome
    return None


def _ancoras_operadora(g: _Grade) -> list[tuple[int, int]]:
    return [(r, c) for r, c, v in g.celulas() if _RE_OPERADORA.match(norm(v))]


def _tem_blocos_operadora(ws) -> bool:
    g = _Grade(ws, max_linhas=500)
    if not _ancoras_operadora(g):
        return False
    return any(_RE_VIDAS.match(norm(v)) for _, _, v in g.celulas())


def _dica_conteudo(ws) -> Optional[str]:
    g = _Grade(ws, max_linhas=80)
    sem = com = False
    for _, _, v in g.celulas():
        t = norm(v)
        if re.search(r"\bsem co-?(part|participacao)", t):
            sem = True
        elif re.search(r"\bcom co-?(part|participacao)", t):
            com = True
    if sem and not com:
        return "SC"
    if com and not sem:
        return "CC"
    return None


def _dica_nome(nome: str) -> Optional[str]:
    n = norm(nome)
    if re.search(r"\bsem\b|\bsc\b|\bs\b|\bs/|scop", n):
        return "SC"
    if re.search(r"\bcom\b|\bcc\b|\bc\b|\bc/|ccop", n):
        return "CC"
    return None


def achar_abas(wb) -> dict[str, str]:
    """{'CC': nome, 'SC': nome} a partir do conteúdo das abas."""
    candidatas = [n for n in wb.sheetnames if _tem_blocos_operadora(wb[n])]
    rotulo: dict[str, Optional[str]] = {}
    for n in candidatas:
        rotulo[n] = _dica_nome(n) or _dica_conteudo(wb[n])
    achadas: dict[str, str] = {}
    for n in candidatas:                               # rótulos explícitos primeiro
        t = rotulo[n]
        if t and t not in achadas:
            achadas[t] = n
    for n in candidatas:                               # sem dica nenhuma: ordem das abas
        if n in achadas.values():
            continue
        for t in ("CC", "SC"):
            if t not in achadas:
                achadas[t] = n
                break
    return achadas


# --------------------------------------------------------------------------- #
# Leitura de um cenário
# --------------------------------------------------------------------------- #
def _ler_bloco(g: _Grade, indice: int, r: int, c0: int, limite: int) -> Bloco:
    nome = ""
    for c in range(c0 + 1, g.ncols + 1):
        v = g(r, c)
        if v is not None and str(v).strip():
            nome = str(v).strip()
            break

    linha_cab = next((rr for rr in range(r + 1, min(r + 7, limite))
                      if any(_RE_VIDAS.match(norm(g(rr, c))) for c in range(c0 + 1, g.ncols + 1))), None)
    if linha_cab is None:
        raise ValueError(f"Bloco '{nome}' (linha {r}): cabeçalho com 'VIDAS' não encontrado")
    colunas = [c for c in range(c0 + 1, g.ncols + 1) if _RE_VIDAS.match(norm(g(linha_cab, c)))]

    linha_planos = None
    for rr in range(linha_cab - 1, r, -1):
        if any(g(rr, c) not in (None, "") or g(rr, c + 1) not in (None, "") for c in colunas):
            linha_planos = rr
            break

    primeira, faixas, rr = linha_cab + 1, [], linha_cab + 1
    while rr < limite and not norm(g(rr, c0)).startswith("total"):
        faixas.append(str(g(rr, c0) or "").strip())
        rr += 1
    if rr >= limite:
        raise ValueError(f"Bloco '{nome}' (linha {r}): linha 'Total' não encontrada")
    linha_total = rr

    planos = []
    for c in colunas:
        rotulo = g(linha_planos, c) if linha_planos else None
        if rotulo in (None, "") and linha_planos:
            rotulo = g(linha_planos, c + 1)
        nome_plano = None if (rotulo in (None, "") or _PLACEHOLDER.match(norm(rotulo))) \
            else str(rotulo).strip()
        planos.append(Plano(
            nome=nome_plano,
            vidas=int(_num(g(linha_total, c)) or 0),
            total=_num(g(linha_total, c + 1)) or 0.0,
            por_faixa=[_num(g(primeira + k, c + 1)) for k in range(len(faixas))],
        ))
    return Bloco(indice=indice, nome=nome, linha=r, faixas=faixas, planos=planos)


def ler_cenario(ws, tipo: str) -> CenarioMatriz:
    g = _Grade(ws)
    ancoras = _ancoras_operadora(g)
    if not ancoras:
        raise ValueError(f"Aba {ws.title!r}: nenhum bloco 'OPERADORA' encontrado")
    colunas = [c for _, c in ancoras]
    # coluna dos rótulos = a que mais repete "OPERADORA"; em empate, a mais à esquerda
    # (a Matriz costuma ter um espelho "reajustado" mais à direita, que não deve ser lido)
    mais = max(colunas.count(c) for c in set(colunas))
    c0 = min(c for c in set(colunas) if colunas.count(c) == mais)
    ancoras = sorted(a for a in ancoras if a[1] == c0)

    blocos = []
    for i, (r, _) in enumerate(ancoras):
        limite = ancoras[i + 1][0] if i + 1 < len(ancoras) else g.n + 1
        blocos.append(_ler_bloco(g, i, r, c0, limite))

    pos_atual = next((i for i, b in enumerate(blocos) if _RE_ATUAL.search(norm(b.nome))), 0)
    if pos_atual:
        blocos.insert(0, blocos.pop(pos_atual))
    for i, b in enumerate(blocos):
        b.indice = i

    reajuste = None
    for _, _, v in g.celulas(r2=60):
        if "reajuste" in norm(v):
            m = _PCT.search(str(v))
            if m:
                reajuste = float(m.group(1).replace(",", ".")) / 100
                break

    def valor_ao_lado(rotulo: str) -> Optional[float]:
        alvo = norm(rotulo)
        for r, c, v in g.celulas():
            if norm(v) == alvo:
                for cc in range(c + 1, c + 5):
                    x = _num(g(r, cc))
                    if x is not None:
                        return x
        return None

    return CenarioMatriz(tipo=tipo, blocos=blocos, reajuste_pct=reajuste,
                         fatura_atual=valor_ao_lado("Fatura Atual"),
                         fatura_reajustada=valor_ao_lado("Fatura Reajustada"))


def _alinhar_cenarios(cc: CenarioMatriz, sc: CenarioMatriz) -> None:
    """Se os dois cenários trazem as mesmas operadoras em ordens diferentes, alinha o SC ao CC."""
    nc = [norm(b.nome) for b in cc.blocos]
    ns = [norm(b.nome) for b in sc.blocos]
    if (len(nc) == len(ns) and all(nc) and all(ns) and len(set(nc)) == len(nc)
            and sorted(nc) == sorted(ns) and nc != ns):
        por_nome = {norm(b.nome): b for b in sc.blocos}
        sc.blocos = [por_nome[n] for n in nc]
        for i, b in enumerate(sc.blocos):
            b.indice = i


# --------------------------------------------------------------------------- #
# Demografia
# --------------------------------------------------------------------------- #
def _achar_base(wb) -> Optional[tuple[str, int, dict[str, int]]]:
    """(aba, linha_cabeçalho, {campo: coluna}) da primeira aba com elegibilidade + sexo + faixa/idade."""
    for nome in wb.sheetnames:
        g = _Grade(wb[nome], max_linhas=40)
        for r in range(1, min(g.n, 30) + 1):
            cab: dict[str, int] = {}
            for c in range(1, g.ncols + 1):
                n = norm(g(r, c))
                if n in SIN_ELEG:
                    cab.setdefault("eleg", c)
                elif n in SIN_SEXO:
                    cab.setdefault("sexo", c)
                elif n in SIN_FAIXA:
                    cab.setdefault("faixa", c)
                elif n in SIN_IDADE:
                    cab.setdefault("idade", c)
            if "eleg" in cab and "sexo" in cab and ("faixa" in cab or "idade" in cab):
                return nome, r, cab
    return None


def ler_demografia(wb) -> Demografia:
    achou = _achar_base(wb)
    if achou is None:
        raise ValueError("Não encontrei a base de beneficiários (aba com colunas de elegibilidade/"
                         "titularidade, sexo e faixa etária ou idade)")
    nome, r_cab, cab = achou
    g = _Grade(wb[nome])
    titular = {"Funcionários": 0, "Dependentes": 0}
    genero = {"M": 0, "F": 0}
    faixas = {ini: {"T": 0, "D": 0} for ini in FAIXA_INICIOS}
    for r in range(r_cab + 1, g.n + 1):
        e = norm(g(r, cab["eleg"]))[:1]
        eleg = "T" if e in ("t", "f") else ("D" if e in ("d", "a") else None)
        if eleg is None:
            continue
        titular["Funcionários" if eleg == "T" else "Dependentes"] += 1
        sexo = norm(g(r, cab["sexo"]))[:1].upper()
        if sexo in genero:
            genero[sexo] += 1
        idade_ini = None
        if "faixa" in cab:
            m = re.search(r"\d+", str(g(r, cab["faixa"]) or ""))
            idade_ini = int(m.group()) if m else None
        if idade_ini is None and "idade" in cab:
            x = _num(g(r, cab["idade"]))
            idade_ini = int(x) if x is not None else None
        if idade_ini is not None:
            faixas[max(i for i in FAIXA_INICIOS if i <= idade_ini)][eleg] += 1
    return Demografia(titularidade=titular, genero=genero, faixa_etaria=faixas)


# --------------------------------------------------------------------------- #
def diagnosticar(caminho: str | Path) -> str:
    """Relatório legível do que o leitor reconheceu na Matriz (ou por que não reconheceu)."""
    out: list[str] = []
    try:
        wb = openpyxl.load_workbook(caminho, data_only=True)
    except Exception as e:  # noqa: BLE001
        return f"Não consegui abrir o arquivo: {e}"
    out.append("Abas do arquivo: " + ", ".join(repr(n) for n in wb.sheetnames))
    abas = achar_abas(wb)
    if not abas:
        out.append("Abas de cenário: NENHUMA (procuro células 'OPERADORA' e colunas 'VIDAS')")
    for tipo, nome in abas.items():
        out.append(f"\nCenário {tipo} <- aba {nome!r}")
        try:
            cen = ler_cenario(wb[nome], tipo)
        except Exception as e:  # noqa: BLE001
            out.append(f"  ERRO: {e}")
            continue
        out.append(f"  reajuste: {cen.reajuste_pct if cen.reajuste_pct is not None else 'não encontrado'}"
                   f" | fatura atual: {cen.fatura_atual} | fatura reajustada: {cen.fatura_reajustada}")
        for b in cen.blocos:
            planos = ", ".join(f"{p.nome or '—'}({p.vidas}v)" for p in b.planos if p.vidas or p.nome)
            out.append(f"  [{b.indice}] {b.nome or '(sem nome)'} | {len(b.faixas)} faixas | "
                       f"total {b.total_mensal:,.2f} | planos: {planos or '—'}")
    base = _achar_base(wb)
    if base:
        out.append(f"\nBase de beneficiários <- aba {base[0]!r}, cabeçalho na linha {base[1]}, "
                   f"colunas {base[2]}")
    else:
        out.append("\nBase de beneficiários: NÃO encontrada (preciso de colunas elegibilidade, sexo e "
                   "faixa etária/idade no mesmo cabeçalho)")
    return "\n".join(out)


def _cliente_do_arquivo(caminho: Path) -> str:
    base = re.split(r"_matriz", caminho.stem, flags=re.I)[0]
    return base.replace("_", " ").strip()


def carregar_matriz(caminho: str | Path, cliente: Optional[str] = None) -> Matriz:
    caminho = Path(caminho)
    wb = openpyxl.load_workbook(caminho, data_only=True)
    cenarios = {tipo: ler_cenario(wb[aba], tipo) for tipo, aba in achar_abas(wb).items()}
    if not cenarios:
        raise ValueError("Nenhuma aba de cenário encontrada.\n\n" + diagnosticar(caminho))
    if "CC" in cenarios and "SC" in cenarios:
        _alinhar_cenarios(cenarios["CC"], cenarios["SC"])
    try:
        demografia = ler_demografia(wb)
    except ValueError as e:
        raise ValueError(f"{e}\n\n{diagnosticar(caminho)}") from None
    return Matriz(caminho=str(caminho), cenarios=cenarios, demografia=demografia,
                  cliente=cliente or _cliente_do_arquivo(caminho))
