"""Camada de transformação: replica em Python as regras do Excel de Gráficos.

Regras (conferidas contra as abas 'Análise Financeira' e 'Impacto Financeiro'):
- reajustado = custo médio atual * (1 + reajuste)
- total do plano = custo médio * nº de vidas do plano ATUAL
- impacto mensal = total da operadora - fatura reajustada
- impacto %     = total da operadora / fatura reajustada - 1
- ordem das operadoras: menor custo no cenário CC primeiro; o SC segue a MESMA
  ordem, mesmo que no SC a classificação mude.
"""
from __future__ import annotations

from typing import Optional, Sequence, Union

from .models import (CenarioCalc, CenarioMatriz, Estudo, Matriz, OperadoraCalc,
                     PlanoAtual, PlanoOperadora)

MAX_OPERADORAS = 4
Selecao = Sequence[Union[int, str]]


def _nome_operadora(matriz: Matriz, indice: int) -> str:
    for tipo in ("CC", "SC"):
        cen = matriz.cenarios.get(tipo)
        if cen and indice < len(cen.blocos) and cen.blocos[indice].nome:
            return cen.blocos[indice].nome
    return f"Operadora {indice}"


def listar_operadoras(matriz: Matriz) -> list[dict]:
    """Operadoras disponíveis (índice 1..n, sem a situação atual) com totais por cenário."""
    n = max(len(c.blocos) for c in matriz.cenarios.values())
    saida = []
    for i in range(1, n):
        item = {"indice": i, "nome": _nome_operadora(matriz, i)}
        for tipo, cen in matriz.cenarios.items():
            item[tipo] = cen.blocos[i].total_mensal if i < len(cen.blocos) else 0.0
        saida.append(item)
    return saida


def resolver_selecao(matriz: Matriz, selecao: Optional[Selecao]) -> list[int]:
    ops = listar_operadoras(matriz)
    por_indice = {o["indice"]: o for o in ops}
    if not selecao:
        escolhidas = [o["indice"] for o in ops
                      if any(o.get(t, 0) > 0 for t in matriz.cenarios)]
        if len(escolhidas) > MAX_OPERADORAS:
            raise ValueError(
                f"A Matriz tem {len(escolhidas)} operadoras com valores; selecione até "
                f"{MAX_OPERADORAS}. Disponíveis: "
                + ", ".join(f"{o['indice']}={o['nome']}" for o in ops if o['indice'] in escolhidas))
        return escolhidas

    resultado: list[int] = []
    for item in selecao:
        if isinstance(item, int) or str(item).strip().isdigit():
            idx = int(item)
            if idx not in por_indice:
                raise ValueError(f"Operadora de índice {idx} não existe (1..{max(por_indice)})")
        else:
            achadas = [o["indice"] for o in ops if str(item).strip().lower() in o["nome"].lower()]
            if len(achadas) != 1:
                raise ValueError(f"Nome '{item}' casa com {len(achadas)} operadoras; use o índice")
            idx = achadas[0]
        if idx not in resultado:
            resultado.append(idx)
    if not 1 <= len(resultado) <= MAX_OPERADORAS:
        raise ValueError(f"Selecione de 1 a {MAX_OPERADORAS} operadoras (recebi {len(resultado)})")
    return resultado


def _calc_cenario(cen: CenarioMatriz, indices: list[int], nome_op, avisos: list[str]) -> CenarioCalc:
    base = cen.blocos[0]
    slots = [i for i, p in enumerate(base.planos) if p.vidas > 0]
    if not slots:
        raise ValueError(f"Cenário {cen.tipo}: nenhum plano atual com vidas")

    fatura_atual = sum(base.planos[i].total for i in slots)
    reaj = cen.reajuste_pct
    if reaj is None:
        reaj = (cen.fatura_reajustada / cen.fatura_atual - 1) if cen.fatura_atual and cen.fatura_reajustada else 0.0
        avisos.append(f"{cen.tipo}: percentual de reajuste não encontrado no texto; calculado = {reaj:.4%}")
    if cen.fatura_atual and abs(cen.fatura_atual - fatura_atual) > 0.01:
        avisos.append(f"{cen.tipo}: 'Fatura Atual' da Matriz ({cen.fatura_atual:,.2f}) difere da soma dos planos ({fatura_atual:,.2f})")

    atuais = []
    for i in slots:
        p = base.planos[i]
        cm = p.custo_medio
        atuais.append(PlanoAtual(nome=p.nome or f"Plano {i + 1}", vidas=p.vidas, custo_medio=cm,
                                 total=p.total, custo_medio_reaj=cm * (1 + reaj),
                                 total_reaj=cm * (1 + reaj) * p.vidas))
    fatura_reaj = sum(a.total_reaj for a in atuais)

    ops = []
    for idx in indices:
        if idx >= len(cen.blocos):
            avisos.append(f"{cen.tipo}: operadora {idx} ausente na aba")
            continue
        bloco = cen.blocos[idx]
        planos: list[Optional[PlanoOperadora]] = []
        for k, slot in enumerate(slots):
            p = bloco.planos[slot] if slot < len(bloco.planos) else None
            if p is None or p.vidas <= 0:
                planos.append(None)
                avisos.append(f"{cen.tipo}: {nome_op(idx)} sem valores para o plano atual '{atuais[k].nome}'")
                continue
            if p.vidas != atuais[k].vidas:
                avisos.append(f"{cen.tipo}: {nome_op(idx)} / {p.nome}: {p.vidas} vidas na cotação x {atuais[k].vidas} no plano atual")
            planos.append(PlanoOperadora(nome=p.nome or f"Plano {slot + 1}", vidas=p.vidas,
                                         custo_medio=p.custo_medio, total=p.total, por_faixa=p.por_faixa))
        total = sum(p.total for p in planos if p)
        ops.append(OperadoraCalc(
            nome=nome_op(idx), bloco_indice=idx, total_mensal=total,
            reducao_mensal=total - fatura_reaj,
            reducao_pct=(total / fatura_reaj - 1) if fatura_reaj else 0.0,
            planos=planos))

    return CenarioCalc(tipo=cen.tipo, reajuste_pct=reaj, fatura_atual=fatura_atual,
                       fatura_reajustada=fatura_reaj, total_vidas=sum(a.vidas for a in atuais),
                       faixas=base.faixas, planos_atuais=atuais, operadoras=ops)


def calcular(matriz: Matriz, selecao: Optional[Selecao] = None, *,
             cliente: Optional[str] = None,
             mes_atual: Optional[tuple[str, int]] = None,
             mes_reajuste: Optional[tuple[str, int]] = None) -> Estudo:
    avisos: list[str] = []
    indices = resolver_selecao(matriz, selecao)

    # ordem definida pelo cenário de referência (CC; senão SC)
    ref = matriz.cenarios.get("CC") or matriz.cenarios["SC"]
    indices.sort(key=lambda i: (ref.blocos[i].total_mensal if i < len(ref.blocos) else float("inf"), i))

    nome_op = lambda i: _nome_operadora(matriz, i)  # noqa: E731
    cenarios = {}
    for tipo, cen in matriz.cenarios.items():
        if len(cen.blocos) != len(ref.blocos):
            avisos.append(f"{tipo}: número de blocos ({len(cen.blocos)}) difere do cenário de referência ({len(ref.blocos)})")
        cenarios[tipo] = _calc_cenario(cen, indices, nome_op, avisos)

    if "SC" not in cenarios:
        avisos.append("Matriz sem aba SC: apenas o cenário CC será preenchido")
    return Estudo(cliente=cliente or matriz.cliente or "", mes_atual=mes_atual,
                  mes_reajuste=mes_reajuste, demografia=matriz.demografia,
                  cenarios=cenarios, avisos=avisos)
