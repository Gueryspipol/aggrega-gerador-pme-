"""Orquestração: Matriz -> cálculo -> PowerPoint."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Sequence

from pptx import Presentation

from .calc import calcular
from .extract import carregar_matriz
from .models import Estudo
from .pptx_writer import remover_slide
from .slides import posicao_menor_custo, preencher_estudo

MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]


def parse_mes(texto: Optional[str]) -> Optional[tuple[str, int]]:
    """'JUN/2026', 'jun 2026', '06/2026' -> ('JUN', 2026)."""
    if not texto:
        return None
    m = re.fullmatch(r"\s*([A-Za-zçÇ]{3}|\d{1,2})\s*[/\- ]\s*(\d{4})\s*", texto)
    if not m:
        raise ValueError(f"Mês/ano inválido: {texto!r} (use JUN/2026)")
    mes = m.group(1)
    nome = MESES[int(mes) - 1] if mes.isdigit() else mes.upper()
    if nome not in MESES:
        raise ValueError(f"Mês inválido: {mes!r}")
    return nome, int(m.group(2))


def mes_do_arquivo(caminho: str | Path) -> Optional[tuple[str, int]]:
    """Tenta ler o mês de nomes como 'Matriz_2026_06_v1' ou 'Apr_X_v1_062026'."""
    nome = Path(caminho).stem
    m = re.search(r"(20\d{2})_(\d{2})(?!\d)", nome)
    if m:
        return MESES[int(m.group(2)) - 1], int(m.group(1))
    m = re.search(r"(?<!\d)(0[1-9]|1[0-2])(20\d{2})(?!\d)", nome)
    if m:
        return MESES[int(m.group(1)) - 1], int(m.group(2))
    return None


@dataclass
class Resultado:
    saida: Path
    estudo: Estudo
    avisos: list[str]


def gerar(matriz: str | Path, template: str | Path, saida: str | Path, *,
          operadoras: Optional[Sequence[int | str]] = None,
          cliente: Optional[str] = None,
          mes_atual: Optional[str] = None,
          mes_reajuste: Optional[str] = None) -> Resultado:
    m = carregar_matriz(matriz, cliente=cliente)
    prs = Presentation(str(template))               # o template original nunca é alterado
    estudo = calcular(m, operadoras, cliente=cliente,
                      mes_atual=parse_mes(mes_atual) or mes_do_arquivo(matriz),
                      mes_reajuste=parse_mes(mes_reajuste),
                      posicao_menor_custo=posicao_menor_custo(prs))
    avisos = list(estudo.avisos)

    remover = preencher_estudo(prs, estudo, avisos)
    for slide in remover:
        remover_slide(prs, slide)

    saida = Path(saida)
    saida.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(saida))
    return Resultado(saida=saida, estudo=estudo, avisos=avisos)
