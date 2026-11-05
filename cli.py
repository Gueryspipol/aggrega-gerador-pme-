"""Linha de comando:

    python -m apresentador360 listar MATRIZ.xlsx
    python -m apresentador360 gerar MATRIZ.xlsx TEMPLATE.pptx -o saida.pptx --operadoras 1,2,3
"""
from __future__ import annotations

import argparse
import sys

from . import formatos as f
from .calc import listar_operadoras
from .extract import carregar_matriz
from .pipeline import gerar


def _listar(args) -> int:
    m = carregar_matriz(args.matriz)
    print(f"Cenários na Matriz: {', '.join(m.cenarios)}")
    print(f"{'#':>2}  {'Operadora':<28} " + " ".join(f"{t:>14}" for t in m.cenarios))
    for o in listar_operadoras(m):
        print(f"{o['indice']:>2}  {o['nome'] or '(sem nome)':<28} "
              + " ".join(f"{f.brl_centavos(o[t]):>14}" for t in m.cenarios))
    return 0


def _gerar(args) -> int:
    ops = [x.strip() for x in args.operadoras.split(",")] if args.operadoras else None
    r = gerar(args.matriz, args.template, args.saida, operadoras=ops, cliente=args.cliente,
              mes_atual=args.mes_atual, mes_reajuste=args.mes_reajuste)
    print(f"Gerado: {r.saida}")
    for cod, cen in r.estudo.cenarios.items():
        print(f"  {cod}: " + " | ".join(f"{o.nome} {f.brl_inteiro(o.total_mensal)} ({f.pct_inteiro(o.reducao_pct)})"
                                         for o in cen.operadoras))
    for a in r.avisos:
        print(f"  AVISO: {a}", file=sys.stderr)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="apresentador360")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("listar", help="lista as operadoras da Matriz")
    p1.add_argument("matriz")
    p1.set_defaults(fn=_listar)
    p2 = sub.add_parser("gerar", help="gera o PowerPoint")
    p2.add_argument("matriz")
    p2.add_argument("template")
    p2.add_argument("-o", "--saida", default="saida/apresentacao.pptx")
    p2.add_argument("--operadoras", help="índices ou nomes separados por vírgula (1 a 4)")
    p2.add_argument("--cliente")
    p2.add_argument("--mes-atual", help="ex.: JUN/2026 (padrão: lido do nome do arquivo)")
    p2.add_argument("--mes-reajuste", help="ex.: SET/2026")
    p2.set_defaults(fn=_gerar)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except (ValueError, FileNotFoundError) as e:
        print(f"ERRO: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
