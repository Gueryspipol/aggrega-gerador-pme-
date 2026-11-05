"""Utilitários seguros de python-pptx: trocam o TEXTO preservando a formatação do template.

Regra de ouro: nunca usar `shape.text = ...` / `text_frame.text = ...` (recria os runs e
perde fonte/tamanho/cor). Aqui sempre se reaproveita o primeiro run do parágrafo; quando o
parágrafo está vazio, o estilo é herdado de `endParaRPr` ou de uma célula-modelo vizinha.
"""
from __future__ import annotations

import copy
import re
from typing import Iterable, Optional, Sequence

from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn


def _rpr_modelo(p_modelo) -> Optional[object]:
    if p_modelo is None or not p_modelo.runs:
        return None
    return p_modelo.runs[0]._r.find(qn("a:rPr"))


def definir_paragrafo(p, texto: str, modelo=None) -> None:
    """Define o texto de um parágrafo mantendo a formatação do primeiro run."""
    runs = p.runs
    if runs:
        runs[0].text = texto
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
        return
    if texto == "":
        return
    r = p.add_run()
    r.text = texto
    fonte = None
    fim = p._p.find(qn("a:endParaRPr"))
    if fim is not None and fim.get("sz"):
        fonte = copy.deepcopy(fim)
        fonte.tag = qn("a:rPr")
    elif _rpr_modelo(modelo) is not None:
        fonte = copy.deepcopy(_rpr_modelo(modelo))
    if fonte is not None:
        antigo = r._r.find(qn("a:rPr"))
        if antigo is not None:
            r._r.remove(antigo)
        r._r.insert(0, fonte)
    if modelo is not None and p._p.find(qn("a:pPr")) is None:
        ppr = modelo._p.find(qn("a:pPr"))
        if ppr is not None:
            p._p.insert(0, copy.deepcopy(ppr))


def definir_linhas(shape, linhas: Sequence[str]) -> None:
    """Parágrafo i recebe linhas[i]; parágrafos além da lista ficam vazios."""
    for i, p in enumerate(shape.text_frame.paragraphs):
        definir_paragrafo(p, linhas[i] if i < len(linhas) else "")


def limpar(shape) -> None:
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.text = ""


def substituir_regex(shape, padrao: str, novo: str) -> bool:
    """Substitui dentro dos runs (preserva o restante do texto e o estilo)."""
    achou = False
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if re.search(padrao, r.text):
                r.text = re.sub(padrao, novo, r.text)
                achou = True
    return achou


def definir_celula(cell, texto: str, modelo=None) -> None:
    definir_paragrafo(cell.text_frame.paragraphs[0], texto,
                      modelo.text_frame.paragraphs[0] if modelo is not None else None)
    for extra in cell.text_frame.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


def formato_serie(chart, k: int) -> str:
    try:
        ser = list(chart.plots[0].series)[k]
        achados = ser._element.xpath("./c:val//c:formatCode")
        if achados and achados[0].text:
            return achados[0].text
    except IndexError:
        pass
    return "General"


def _desvincular_excel_externo(chart) -> None:
    """Gráficos colados do Excel com vínculo apontam para um arquivo externo (caminho de rede),
    que o python-pptx não consegue atualizar. Remove o vínculo; os dados passam a ficar
    embutidos no próprio gráfico. Aparência e formatação não mudam."""
    for ext in chart._chartSpace.xpath("c:externalData"):
        r_id = ext.get(qn("r:id"))
        rel = chart.part.rels.get(r_id) if r_id else None
        if rel is not None and rel.is_external:
            chart._chartSpace.remove(ext)
            chart.part.drop_rel(r_id)


def substituir_dados_grafico(chart, categorias: Sequence[str],
                             series: Sequence[Sequence[Optional[float]]]) -> None:
    """Troca os dados de um gráfico nativo mantendo estilo, cores e rótulos das séries."""
    _desvincular_excel_externo(chart)
    nomes = [s.name for s in chart.plots[0].series]
    cd = CategoryChartData()
    cd.categories = list(categorias)
    for k, valores in enumerate(series):
        cd.add_series(nomes[k] if k < len(nomes) else "", list(valores),
                      number_format=formato_serie(chart, k))
    chart.replace_data(cd)


def remover_slide(prs, slide) -> None:
    lista = prs.slides._sldIdLst
    for sld_id in list(lista):
        if prs.part.related_part(sld_id.rId) is slide.part:
            prs.part.drop_rel(sld_id.rId)
            lista.remove(sld_id)
            return


def clonar_caixa(shape, centro_x: int):
    """Duplica uma caixa de texto (mesmo estilo) centralizada em `centro_x`."""
    colecao = shape._parent                      # SlideShapes
    proximo_id = max(s.shape_id for s in colecao) + 1
    novo = copy.deepcopy(shape._element)
    shape._element.addnext(novo)
    clone = next(s for s in colecao if s._element is novo)
    clone._element.xpath("./p:nvSpPr/p:cNvPr")[0].set("id", str(proximo_id))
    clone.left = int(centro_x - clone.width / 2)
    return clone
