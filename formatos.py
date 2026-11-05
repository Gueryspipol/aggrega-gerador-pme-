"""Formatação pt-BR dos valores que vão para os slides."""
from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from typing import Optional


def _arredonda(valor: float, casas: int) -> Decimal:
    q = Decimal(1).scaleb(-casas)
    return Decimal(str(valor)).quantize(q, rounding=ROUND_HALF_UP)


def _milhar(d: Decimal, casas: int) -> str:
    s = f"{abs(d):,.{casas}f}"                  # 1,234.56
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def brl_inteiro(valor: float) -> str:
    """R$18.220  (slide 14)"""
    d = _arredonda(valor, 0)
    return f"{'-' if d < 0 else ''}R${_milhar(d, 0)}"


def brl_delta(valor: float) -> str:
    """-R$14.239 / +R$2.155  (slide 14, 'Redução Mensal')"""
    d = _arredonda(valor, 0)
    sinal = "-" if d < 0 else ("+" if d > 0 else "")
    return f"{sinal}R${_milhar(d, 0)}"


def brl_centavos(valor: Optional[float]) -> str:
    """R$1.152,49  (slide 15)"""
    if valor is None:
        return ""
    d = _arredonda(valor, 2)
    return f"{'-' if d < 0 else ''}R${_milhar(d, 2)}"


def brl_centavos_espaco(valor: Optional[float]) -> str:
    """R$ 184,87  (slide 16)"""
    if valor is None:
        return ""
    return brl_centavos(valor).replace("R$", "R$ ", 1)


def pct_inteiro(frac: float) -> str:
    """-70%  (badge do slide 14)"""
    d = _arredonda(frac * 100, 0)
    sinal = "+" if d > 0 else ""
    return f"{sinal}{d}%"


def pct_2casas(frac: float) -> str:
    """11,83%"""
    return f"{_arredonda(frac * 100, 2)}".replace(".", ",") + "%"
