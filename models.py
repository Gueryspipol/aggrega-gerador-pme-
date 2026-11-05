"""Modelos de dados: o "contrato" entre as camadas de extração, cálculo e carga."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


# --------------------------------------------------------------------------- #
# Camada de extração (o que existe na Matriz)
# --------------------------------------------------------------------------- #
@dataclass
class Plano:
    nome: Optional[str]            # None quando o nome é só o placeholder "PLANO"
    vidas: int
    total: float                   # soma mensal do plano (linha "Total")
    por_faixa: list[Optional[float]] = field(default_factory=list)  # per capita, uma por faixa

    @property
    def custo_medio(self) -> float:
        return self.total / self.vidas if self.vidas else 0.0


@dataclass
class Bloco:
    """Um bloco 'OPERADORA' da aba da Matriz. O bloco 0 é sempre a situação atual."""
    indice: int
    nome: str
    linha: int
    faixas: list[str]
    planos: list[Plano]

    @property
    def total_mensal(self) -> float:
        return sum(p.total for p in self.planos if p.vidas > 0)

    @property
    def vazio(self) -> bool:
        return self.total_mensal <= 0


@dataclass
class CenarioMatriz:
    tipo: str                      # "CC" ou "SC"
    blocos: list[Bloco]
    reajuste_pct: Optional[float]  # 0.1183 = 11,83%
    fatura_atual: Optional[float]
    fatura_reajustada: Optional[float]


@dataclass
class Demografia:
    titularidade: dict[str, int]                 # {"Funcionários": 10, "Dependentes": 6}
    genero: dict[str, int]                       # {"M": 5, "F": 11}
    faixa_etaria: dict[int, dict[str, int]]      # idade inicial da faixa -> {"T": n, "D": n}


@dataclass
class Matriz:
    caminho: str
    cenarios: dict[str, CenarioMatriz]           # só as abas que existem
    demografia: Demografia
    cliente: Optional[str] = None
    mes_reajuste: Optional[str] = None


# --------------------------------------------------------------------------- #
# Camada de transformação (o que o PowerPoint precisa)
# --------------------------------------------------------------------------- #
@dataclass
class PlanoAtual:
    nome: str
    vidas: int
    custo_medio: float
    total: float
    custo_medio_reaj: float
    total_reaj: float


@dataclass
class PlanoOperadora:
    nome: str
    vidas: int
    custo_medio: float
    total: float
    por_faixa: list[Optional[float]]


@dataclass
class OperadoraCalc:
    nome: str
    bloco_indice: int
    total_mensal: float
    reducao_mensal: float          # negativo = economia
    reducao_pct: float
    planos: list[Optional[PlanoOperadora]]   # alinhados a CenarioCalc.planos_atuais


@dataclass
class CenarioCalc:
    tipo: str
    reajuste_pct: float
    fatura_atual: float
    fatura_reajustada: float
    total_vidas: int
    faixas: list[str]
    planos_atuais: list[PlanoAtual]
    operadoras: list[OperadoraCalc]


@dataclass
class Estudo:
    cliente: str
    mes_atual: Optional[tuple[str, int]]         # ("JUN", 2026)
    mes_reajuste: Optional[tuple[str, int]]
    demografia: Demografia
    cenarios: dict[str, CenarioCalc]
    avisos: list[str] = field(default_factory=list)
