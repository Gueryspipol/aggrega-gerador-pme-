"""Camada de carga: preenche os slides do template (1, 13, 14-17 CC e 18-21 SC).

Os elementos são localizados pelo que são (texto, posição, tipo) e não por nomes ou ids,
assim o mesmo código serve aos slides CC e SC e sobrevive a pequenas mudanças no template.
Nada de layout é alterado: só texto, células de tabela e dados de gráficos nativos.
"""
from __future__ import annotations

import re
from typing import Optional

from pptx.enum.shapes import MSO_SHAPE_TYPE

from . import formatos as f
from .extract import FAIXA_INICIOS, FAIXA_ROTULOS
from .models import CenarioCalc, Demografia, Estudo, OperadoraCalc
from .pptx_writer import (clonar_caixa, definir_celula, definir_linhas, limpar,
                          substituir_dados_grafico, substituir_regex)

# número do slide (1-based) de cada parte do template
SLIDE_CAPA = 1
SLIDE_DEMOGRAFIA = 13
SLIDES_CENARIO = {
    "CC": {"resumo": 14, "financeira": 15, "faixas": [16, 17]},
    "SC": {"resumo": 18, "financeira": 19, "faixas": [20, 21]},
}
OPERADORAS_POR_SLIDE_FAIXAS = 2

_RE_VALOR = re.compile(r"^-?R\$\s?[\d.,]+$")
_RE_BADGE = re.compile(r"^[+-]?\d+%$")
_RE_MES = re.compile(r"^[A-ZÇ]{3}$")


def _caixas_texto(slide):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


# --------------------------------------------------------------------------- #
# Slide 1 – capa
# --------------------------------------------------------------------------- #
def preencher_capa(slide, cliente: str) -> None:
    alvo = max(_caixas_texto(slide), key=lambda s: s.top)   # nome do cliente: caixa mais baixa
    definir_linhas(alvo, [cliente])


# --------------------------------------------------------------------------- #
# Slide 13 – painel demográfico
# --------------------------------------------------------------------------- #
def _pct_inteiros(partes: list[int]) -> list[int]:
    total = sum(partes)
    if not total:
        return [0] * len(partes)
    pct = [int(round(p * 100 / total + 1e-9)) for p in partes]
    pct[-1] += 100 - sum(pct)          # garante soma 100
    return pct


def preencher_demografia(slide, demo: Demografia, cen: CenarioCalc) -> None:
    for sh in slide.shapes:
        if not getattr(sh, "has_chart", False) or not sh.has_chart:
            continue
        chart = sh.chart
        tipo = chart.chart_type
        nomes_series = [s.name.upper() for s in chart.plots[0].series]
        if "DOUGHNUT" in str(tipo):
            substituir_dados_grafico(chart, list(demo.titularidade),
                                     [list(demo.titularidade.values())])
        elif "PIE" in str(tipo):
            vidas = [a.vidas for a in cen.planos_atuais]
            total = sum(vidas) or 1
            substituir_dados_grafico(chart, [a.nome for a in cen.planos_atuais],
                                     [vidas, [v / total for v in vidas]])
        elif "FUNCIONÁRIOS" in nomes_series:
            t = [demo.faixa_etaria[i]["T"] or None for i in FAIXA_INICIOS]
            d = [demo.faixa_etaria[i]["D"] or None for i in FAIXA_INICIOS]
            substituir_dados_grafico(chart, FAIXA_ROTULOS, [t, d])
        else:   # distribuição do plano: valor atual por plano
            cats = [f"{a.nome}\n{a.vidas} VIDAS" for a in cen.planos_atuais]
            substituir_dados_grafico(chart, cats, [[a.total for a in cen.planos_atuais],
                                                   [None] * len(cats)])

    caixas = _caixas_texto(slide)
    masc = next((s for s in caixas if s.text_frame.text.strip().startswith("Masculino")), None)
    fem = next((s for s in caixas if s.text_frame.text.strip().startswith("Feminino")), None)
    m, fe = demo.genero["M"], demo.genero["F"]
    if masc is not None and fem is not None and (m + fe):
        definir_linhas(masc, ["Masculino", f"({m})"])
        definir_linhas(fem, [f"Feminino ({fe})"])
        pm, pf = _pct_inteiros([m, fe])
        for sh in caixas:
            if re.fullmatch(r"\d+%", sh.text_frame.text.strip()):
                perto = masc if abs(sh.left - masc.left) < abs(sh.left - fem.left) else fem
                definir_linhas(sh, [f"{pm if perto is masc else pf}%"])


# --------------------------------------------------------------------------- #
# Slides 14 / 18 – resumo executivo
# --------------------------------------------------------------------------- #
def preencher_resumo(slide, cen: CenarioCalc, estudo: Estudo, altura_slide: int,
                     avisos: list[str]) -> None:
    cartoes = sorted((s for s in slide.shapes
                      if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                      and s.height >= 0.3 * altura_slide and not s.text_frame.text.strip()),
                     key=lambda s: s.left)
    if len(cartoes) < 3:
        raise ValueError("Slide de resumo: cartões de coluna não encontrados no template")

    por_coluna: dict[int, list] = {i: [] for i in range(len(cartoes))}
    for sh in _caixas_texto(slide):
        if not (0.15 * altura_slide < sh.top < 0.92 * altura_slide):
            continue                                        # título e rodapés ficam intactos
        cx = sh.left + sh.width / 2
        for i, c in enumerate(cartoes):
            if c.left <= cx <= c.left + c.width:
                por_coluna[i].append(sh)
                break

    def achar(col, teste):
        return [s for s in por_coluna[col] if teste(s.text_frame.text.strip())]

    # colunas 0 (atual) e 1 (reajuste)
    for col, valor, mes in ((0, cen.fatura_atual, estudo.mes_atual),
                            (1, cen.fatura_reajustada, estudo.mes_reajuste)):
        for sh in achar(col, lambda t: bool(_RE_VALOR.match(t))):
            definir_linhas(sh, [f.brl_inteiro(valor)])
        for sh in achar(col, lambda t: t.startswith("Total de Vidas")):
            definir_linhas(sh, ["Total de Vidas", str(cen.total_vidas)])
        for sh in achar(col, lambda t: bool(_RE_MES.match(t.split("\n")[0]))):
            if mes:
                definir_linhas(sh, [mes[0], str(mes[1])])
            else:
                avisos.append(f"Slide de resumo {cen.tipo}: mês/ano da coluna {col + 1} não informado; mantido o do template")
    for sh in achar(1, lambda t: "Reajuste" in t):
        substituir_regex(sh, r"\d+(?:,\d+)?%", f.pct_2casas(cen.reajuste_pct))

    # colunas de operadoras
    n_cols = len(cartoes) - 2
    if len(cen.operadoras) > n_cols:
        avisos.append(f"Slide de resumo {cen.tipo}: template tem {n_cols} colunas de operadora; "
                      f"{len(cen.operadoras) - n_cols} operadora(s) não couberam")
    for j in range(n_cols):
        col = j + 2
        op: Optional[OperadoraCalc] = cen.operadoras[j] if j < len(cen.operadoras) else None
        if op is None:
            for sh in por_coluna[col]:
                limpar(sh)
            continue
        for sh in achar(col, lambda t: bool(_RE_VALOR.match(t))):
            definir_linhas(sh, [f.brl_inteiro(op.total_mensal)])
        for sh in achar(col, lambda t: bool(_RE_BADGE.match(t))):
            definir_linhas(sh, [f.pct_inteiro(op.reducao_pct)])
        for sh in achar(col, lambda t: t.startswith("Redução")):
            definir_linhas(sh, ["Redução Mensal", f.brl_delta(op.reducao_mensal)])


# --------------------------------------------------------------------------- #
# Slides 15 / 19 – análise financeira
# --------------------------------------------------------------------------- #
_ROTULOS_LINHA = ["Produto", "Número de Vidas", "Custo Médio", "Total"]


def _modelo_celula(tabela, k: int, c: int):
    for cc in range(c, -1, -1):
        cel = tabela.cell(k, cc)
        if cel.text_frame.paragraphs[0].runs:
            return cel
    return None


def preencher_financeira(slide, cen: CenarioCalc, avisos: list[str]) -> None:
    tabela = next(s for s in slide.shapes if s.has_table).table
    n_cols = len(tabela.columns)
    n_ops_cols = n_cols - 3
    if len(cen.operadoras) > n_ops_cols:
        avisos.append(f"Slide de análise financeira {cen.tipo}: tabela tem {n_ops_cols} colunas de operadora; "
                      f"{len(cen.operadoras) - n_ops_cols} não couberam")
    n_grupos = (len(tabela.rows) + 1) // 5
    if len(cen.planos_atuais) > n_grupos:
        avisos.append(f"Slide de análise financeira {cen.tipo}: tabela comporta {n_grupos} planos atuais; "
                      f"{len(cen.planos_atuais) - n_grupos} não couberam")

    for g in range(n_grupos):
        base = 5 * g
        usado = g < len(cen.planos_atuais)
        for k in range(4):
            if base + k >= len(tabela.rows):
                break
            valores = [""] * n_cols
            if usado:
                a = cen.planos_atuais[g]
                linhas = [
                    [a.nome, a.nome],
                    [str(a.vidas), str(a.vidas)],
                    [f.brl_centavos(a.custo_medio), f.brl_centavos(a.custo_medio_reaj)],
                    [f.brl_centavos(a.total), f.brl_centavos(a.total_reaj)],
                ]
                valores[0] = _ROTULOS_LINHA[k]
                valores[1], valores[2] = linhas[k]
                for j in range(n_ops_cols):
                    p = cen.operadoras[j].planos[g] if j < len(cen.operadoras) else None
                    if p is not None:
                        valores[3 + j] = [p.nome, str(p.vidas), f.brl_centavos(p.custo_medio),
                                          f.brl_centavos(p.total)][k]
            for c in range(n_cols):
                definir_celula(tabela.cell(base + k, c), valores[c],
                               _modelo_celula(tabela, k, c))


# --------------------------------------------------------------------------- #
# Slides 16-17 / 20-21 – custo por faixa etária
# --------------------------------------------------------------------------- #
def preencher_faixas(slide, cen: CenarioCalc, grupo: list[OperadoraCalc],
                     avisos: list[str]) -> None:
    quadro = next(s for s in slide.shapes if s.has_table)
    tabela = quadro.table
    n = len(cen.planos_atuais)
    n_valores = len(tabela.columns) - 1
    if len(grupo) * n > n_valores:
        avisos.append(f"Slide de faixas {cen.tipo}: {len(grupo) * n} colunas de plano para {n_valores} disponíveis")

    # coluna da tabela -> plano (ou None)
    alvo: dict[int, Optional[object]] = {c: None for c in range(1, n_valores + 1)}
    for k, op in enumerate(grupo):
        for j in range(n):
            c = 1 + k * n + j
            if c in alvo:
                alvo[c] = op.planos[j]

    n_faixas = min(len(cen.faixas), len(tabela.rows))
    for c in range(1, n_valores + 1):
        plano = alvo[c]
        for r in range(len(tabela.rows)):
            valor = ""
            if plano is not None and r < n_faixas and r < len(plano.por_faixa):
                valor = f.brl_centavos_espaco(plano.por_faixa[r])
            definir_celula(tabela.cell(r, c), valor, _modelo_celula(tabela, r, c))

    # rótulos (nome do plano) acima de cada coluna
    x = quadro.left
    centros = {0: x + tabela.columns[0].width / 2}
    for c in range(1, n_valores + 1):
        x_ini = quadro.left + sum(tabela.columns[i].width for i in range(c))
        centros[c] = x_ini + tabela.columns[c].width / 2
    rotulos = [s for s in _caixas_texto(slide)
               if quadro.top - 700000 <= s.top < quadro.top and not s.has_table]
    por_col: dict[int, object] = {}
    for sh in rotulos:
        cx = sh.left + sh.width / 2
        c = min(range(1, n_valores + 1), key=lambda i: abs(centros[i] - cx))
        por_col[c] = sh
    modelo = rotulos[0] if rotulos else None
    for c in range(1, n_valores + 1):
        plano = alvo[c]
        sh = por_col.get(c)
        if sh is None and plano is not None and modelo is not None:
            sh = clonar_caixa(modelo, centros[c])
        if sh is not None:
            definir_linhas(sh, [plano.nome if plano is not None else ""])
            if plano is not None and len(plano.nome) > 11:
                avisos.append(f"Slide de faixas {cen.tipo}: nome do plano '{plano.nome}' é longo e "
                              f"pode quebrar linha no rótulo; confira o slide")


# --------------------------------------------------------------------------- #
# Validação do template
# --------------------------------------------------------------------------- #
def _tem(slide, tipo: str) -> int:
    if tipo == "tabela":
        return sum(1 for s in slide.shapes if getattr(s, "has_table", False) and s.has_table)
    if tipo == "grafico":
        return sum(1 for s in slide.shapes if getattr(s, "has_chart", False) and s.has_chart)
    return sum(1 for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip())


def resumo_template(prs, ate: int = 24) -> str:
    """Uma linha por slide: o que o robô enxerga nele (para diagnosticar templates diferentes)."""
    linhas = [f"O template tem {len(prs.slides)} slides."]
    for i, sl in enumerate(prs.slides, 1):
        if i > ate:
            break
        textos = [s.text_frame.text.strip().split("\n")[0] for s in sl.shapes
                  if s.has_text_frame and s.text_frame.text.strip()]
        linhas.append(f"  slide {i}: {_tem(sl, 'tabela')} tabela(s), {_tem(sl, 'grafico')} gráfico(s), "
                      f"{len(textos)} texto(s) | {textos[0][:40] if textos else '(sem texto)'}")
    return "\n".join(linhas)


def validar_template(prs, tipos: list[str]) -> list[str]:
    """Problemas que impedem o preenchimento (lista vazia = template compatível)."""
    slides = list(prs.slides)
    precisa = [SLIDE_CAPA, SLIDE_DEMOGRAFIA] + [n for t in tipos for n in (
        SLIDES_CENARIO[t]["resumo"], SLIDES_CENARIO[t]["financeira"], *SLIDES_CENARIO[t]["faixas"])]
    if len(slides) < max(precisa):
        return [f"o template tem {len(slides)} slides, mas o robô espera ao menos {max(precisa)}"]
    problemas = []
    if _tem(slides[SLIDE_CAPA - 1], "texto") == 0:
        problemas.append(f"slide {SLIDE_CAPA} (capa): não tem caixa de texto para o nome do cliente")
    if _tem(slides[SLIDE_DEMOGRAFIA - 1], "grafico") == 0:
        problemas.append(f"slide {SLIDE_DEMOGRAFIA} (painel demográfico): não tem gráficos")
    for t in tipos:
        m = SLIDES_CENARIO[t]
        altura = prs.slide_height
        cartoes = [s for s in slides[m["resumo"] - 1].shapes
                   if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.height >= 0.3 * altura
                   and not s.text_frame.text.strip()]
        if len(cartoes) < 3:
            problemas.append(f"slide {m['resumo']} (resumo {t}): esperava as colunas (cartões) de valores")
        for n, nome in [(m["financeira"], "análise financeira"), *[(x, "custo por faixa etária") for x in m["faixas"]]]:
            if _tem(slides[n - 1], "tabela") == 0:
                problemas.append(f"slide {n} ({nome} {t}): esperava uma tabela e não achei")
    return problemas


# --------------------------------------------------------------------------- #
def preencher_estudo(prs, estudo: Estudo, avisos: list[str]) -> list:
    """Preenche todos os slides e devolve a lista de slides a remover."""
    slides = list(prs.slides)
    problemas = validar_template(prs, list(estudo.cenarios))
    if problemas:
        raise ValueError("O template não tem a estrutura esperada:\n- " + "\n- ".join(problemas)
                         + "\n\n" + resumo_template(prs))
    s = lambda n: slides[n - 1]   # noqa: E731
    remover = []

    preencher_capa(s(SLIDE_CAPA), estudo.cliente)
    primeiro = next(iter(estudo.cenarios.values()))
    preencher_demografia(s(SLIDE_DEMOGRAFIA), estudo.demografia, primeiro)

    for tipo, mapa in SLIDES_CENARIO.items():
        cen = estudo.cenarios.get(tipo)
        todos = [mapa["resumo"], mapa["financeira"], *mapa["faixas"]]
        if cen is None:
            remover += [s(n) for n in todos]
            continue
        preencher_resumo(s(mapa["resumo"]), cen, estudo, prs.slide_height, avisos)
        preencher_financeira(s(mapa["financeira"]), cen, avisos)
        for g, num in enumerate(mapa["faixas"]):
            ini = g * OPERADORAS_POR_SLIDE_FAIXAS
            grupo = cen.operadoras[ini:ini + OPERADORAS_POR_SLIDE_FAIXAS]
            if grupo:
                preencher_faixas(s(num), cen, grupo, avisos)
            else:
                remover.append(s(num))
    return remover
