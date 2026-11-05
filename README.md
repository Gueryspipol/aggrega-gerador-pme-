# Apresentador 360 — v2

Lê a **Matriz** (Excel de cotações) e preenche o **template PowerPoint** padrão, sem alterar layout,
fontes ou cores e sem tocar em logotipos/imagens.

```
Matriz.xlsx ──► extract.py ──► calc.py ──► slides.py / pptx_writer.py ──► apresentacao.pptx
 (CC, SC, Base)   leitura        regras      carga no template
                  dinâmica      do Excel
```

| Camada | Arquivo | Responsabilidade |
|---|---|---|
| Extração | `extract.py` | Lê as abas CC/SC/Base. Blocos de operadora e colunas de planos detectados por rótulo (a quantidade de planos varia por cliente). Aba SC ausente é tratada. |
| Transformação | `calc.py` | Reajuste, totais, redução mensal/%, ordenação por menor custo, seleção de 1 a 4 operadoras. |
| Carga | `slides.py`, `pptx_writer.py` | Troca só o texto (reaproveitando o run original), preenche tabelas e atualiza gráficos nativos. |
| Orquestração | `pipeline.py`, `cli.py` | `gerar(...)` e linha de comando. |

## Uso

```bash
pip install -r requirements.txt
export PYTHONPATH=src          # ou: pip install -e .

# 1) ver as operadoras da Matriz (índice, nome e total CC/SC)
python -m apresentador360 listar "Matriz.xlsx"

# 2) gerar (1 a 4 operadoras, pelo índice ou parte do nome)
python -m apresentador360 gerar "Matriz.xlsx" "templates/Template.pptx" \
    -o saida/apresentacao.pptx --operadoras 5,1,2 \
    --cliente "Shock Imports" --mes-atual JUN/2026 --mes-reajuste SET/2026
```

Interface web (Streamlit):

```bash
streamlit run app.py      # no Streamlit Cloud, Main file path = app.py
```

Coloque o template padrão em `templates/` (um `.pptx`) para ele aparecer na lista; senão o app pede o upload.

Em Python / Colab:

```python
from apresentador360.pipeline import gerar
r = gerar("Matriz.xlsx", "Template.pptx", "saida.pptx", operadoras=[5, 1, 2], mes_reajuste="SET/2026")
print(r.avisos)      # pendências para conferir
```

Cliente e mês atual são deduzidos do nome do arquivo quando não informados
(`Shock_Imports_Matriz_..._2026_06_v1.xlsx` → "Shock Imports", JUN/2026).
**O mês do reajuste não existe na Matriz** (o campo `MÊS DO REAJUSTE` vem como placeholder), por isso
deve ser informado em `--mes-reajuste`; se faltar, o robô mantém o texto do template e avisa.

## Mapeamento dos slides

| Slide | Conteúdo | Origem |
|---|---|---|
| 1 | Nome do cliente | `--cliente` / nome do arquivo |
| 13 | Painel demográfico (4 gráficos nativos + gênero) | aba `Base` |
| 14 / 18 | Resumo: valores, mês/ano, reajuste, % e R$ de redução, total de vidas — CC / SC | abas CC / SC |
| 15 / 19 | Análise financeira por plano atual × operadora — CC / SC | abas CC / SC |
| 16-17 / 20-21 | Custo por faixa etária (2 operadoras por slide) — CC / SC | abas CC / SC |
| 22+ | Rede credenciada | **manual** (ignorado) |

## Regras de negócio implementadas

* Reajustado = custo médio atual × (1 + reajuste); total do plano = custo médio × vidas do plano atual.
* Redução mensal = total da operadora − fatura reajustada; % = total ÷ fatura reajustada − 1.
* **Ordem**: menor custo no CC primeiro ("Menor Custo", depois "+ Opções"). O SC **segue a mesma
  ordem**, mesmo que no SC a classificação seja outra.
* Menos de 4 operadoras: colunas excedentes ficam em branco; slides inteiros sem nenhum dado
  (ex.: o 17 com 1-2 operadoras) são removidos. Matriz sem aba SC: slides 18-21 são removidos.
* Os cálculos são feitos em Python e foram validados contra as abas *Análise Financeira* e
  *Impacto Financeiro* do Excel de Gráficos (testes em `tests/`).

## Limites conhecidos (template atual)

* **Slide 14/18 tem 3 colunas de operadora**; a 4ª operadora entra nos slides 15/19 (tabela com 4 colunas)
  e 16-17/20-21, e o robô avisa que ela não coube no resumo.
* Logotipos das operadoras (e os cartões "c" atrás deles) são manuais, como combinado — ao trocar de
  operadoras, troque as imagens.
* Os gráficos do slide 13 estavam **vinculados** a um Excel em caminho de rede; o robô desfaz o vínculo e
  embute os dados (visual idêntico).
* Os nomes dos planos nos rótulos do slide 16 são centralizados sobre a coluna; nomes longos (> 11 caracteres)
  podem quebrar linha — o robô avisa.
* As linhas "Total c desconto / Custo Médio c desconto" da Matriz **não** são usadas (o Excel de
  Gráficos também usa o custo médio sem desconto).
* Tags de rodapé ("*Com coparticipação", "*Sem coparticipação | Amil com coparticipação para terapias")
  e textos fixos (telemedicina, etc.) permanecem os do template.

## Testes

```bash
pip install pytest
AP360_SAMPLES=/pasta/com/os/exemplos pytest      # sem a variável, os testes de integração são pulados
```
