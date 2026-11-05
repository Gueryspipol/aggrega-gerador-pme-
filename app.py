"""Interface Streamlit do Apresentador 360.

Executar:  streamlit run app.py
No Streamlit Community Cloud, o "Main file path" é:  app.py
"""
import sys
import tempfile
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).parent
sys.path.insert(0, str(RAIZ / "src"))
sys.path.insert(0, str(RAIZ))          # aceita também apresentador360/ direto na raiz

_na_pasta = ((RAIZ / "src" / "apresentador360" / "__init__.py").exists()
             or (RAIZ / "apresentador360" / "__init__.py").exists())
if not _na_pasta and (RAIZ / "__init__.py").exists() and (RAIZ / "pipeline.py").exists():
    # arquivos soltos na raiz do repositório (sem a pasta): a própria raiz vira o pacote
    import importlib.util
    _spec = importlib.util.spec_from_file_location(
        "apresentador360", RAIZ / "__init__.py", submodule_search_locations=[str(RAIZ)])
    _pacote = importlib.util.module_from_spec(_spec)
    sys.modules["apresentador360"] = _pacote
    _spec.loader.exec_module(_pacote)
    _na_pasta = True

if not _na_pasta:
    st.error("Pasta do robô não encontrada no repositório.")
    st.markdown(
        "O repositório precisa ter **`src/apresentador360/`** com estes 9 arquivos: "
        "`__init__.py`, `__main__.py`, `calc.py`, `cli.py`, `extract.py`, `formatos.py`, "
        "`models.py`, `pipeline.py`, `pptx_writer.py` e `slides.py`.")
    st.write("Arquivos encontrados na raiz:", sorted(p.name for p in RAIZ.iterdir()))
    st.stop()

from apresentador360 import formatos as f  # noqa: E402
from apresentador360.calc import MAX_OPERADORAS, listar_operadoras  # noqa: E402
from apresentador360.extract import carregar_matriz, diagnosticar  # noqa: E402
from apresentador360.pipeline import gerar, mes_do_arquivo  # noqa: E402

TEMPLATES = Path(__file__).parent / "templates"

st.set_page_config(page_title="Apresentador 360", layout="centered")
st.title("Apresentador 360")
st.caption("Matriz (Excel) → apresentação PowerPoint")

matriz_up = st.file_uploader("1) Matriz (.xlsx)", type=["xlsx"])

# O template é fixo: usa o .pptx do repositório (pasta templates/ ou raiz) sem perguntar.
template_padrao = sorted([*TEMPLATES.glob("*.pptx"), *RAIZ.glob("*.pptx")]) if TEMPLATES.exists() \
    else sorted(RAIZ.glob("*.pptx"))
template_up = None
if len(template_padrao) == 1:
    tpl_fixo = template_padrao[0]
    st.caption(f"Template: {tpl_fixo.name}")
elif len(template_padrao) > 1:
    tpl_fixo = next((p for p in template_padrao if "template" in p.name.lower()), None) \
        or st.selectbox("Template", template_padrao, format_func=lambda p: p.name)
else:
    tpl_fixo = None
    template_up = st.file_uploader("Template (.pptx) — não encontrei nenhum no repositório",
                                   type=["pptx"])

if matriz_up is not None:
    work = Path(tempfile.mkdtemp())
    matriz_path = work / matriz_up.name
    matriz_path.write_bytes(matriz_up.getvalue())

    try:
        matriz = carregar_matriz(matriz_path)
    except Exception as e:  # noqa: BLE001
        st.error("Não consegui ler a Matriz. Abaixo está o que o robô enxergou no arquivo "
                 "(me envie este texto para eu ajustar a leitura):")
        st.code(f"{e}\n\n{diagnosticar(matriz_path)}" if "Abas do arquivo" not in str(e) else str(e))
        st.stop()

    st.success(f"Cenários encontrados: {', '.join(matriz.cenarios)}")
    with st.expander("O que o robô reconheceu na Matriz"):
        st.code(diagnosticar(matriz_path))
    ops = [o for o in listar_operadoras(matriz) if any(o.get(t, 0) > 0 for t in matriz.cenarios)]
    rotulos = {o["indice"]: f"{o['indice']} - {o['nome'] or '(sem nome)'}  |  "
               + "  ".join(f"{t} {f.brl_centavos(o[t])}" for t in matriz.cenarios) for o in ops}
    escolha = st.multiselect(f"2) Operadoras (1 a {MAX_OPERADORAS})", list(rotulos),
                             format_func=rotulos.get, max_selections=MAX_OPERADORAS)

    col1, col2, col3 = st.columns(3)
    cliente = col1.text_input("Cliente", matriz.cliente or "")
    mes = mes_do_arquivo(matriz_path)
    mes_atual = col2.text_input("Mês atual", f"{mes[0]}/{mes[1]}" if mes else "", placeholder="JUN/2026")
    mes_reaj = col3.text_input("Mês do reajuste", "", placeholder="SET/2026")

    if st.button("Gerar apresentação", type="primary", disabled=not escolha):
        try:
            if template_up is not None:
                tpl = work / template_up.name
                tpl.write_bytes(template_up.getvalue())
            elif tpl_fixo is not None:
                tpl = tpl_fixo
            else:
                st.error("Envie o template .pptx.")
                st.stop()
            saida = work / f"{cliente or 'apresentacao'}.pptx"
            r = gerar(matriz_path, tpl, saida, operadoras=escolha, cliente=cliente or None,
                      mes_atual=mes_atual or None, mes_reajuste=mes_reaj or None)
        except Exception as e:  # noqa: BLE001
            st.error(f"Erro ao gerar ({type(e).__name__}): {e or 'sem mensagem'}")
            st.exception(e)
            try:
                from pptx import Presentation
                from apresentador360.slides import resumo_template
                st.code(resumo_template(Presentation(str(tpl))))
            except Exception:  # noqa: BLE001
                pass
        else:
            st.download_button("Baixar PowerPoint", saida.read_bytes(), file_name=saida.name,
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            for tipo, cen in r.estudo.cenarios.items():
                st.write(f"**{tipo}** — " + " | ".join(
                    f"{o.nome}: {f.brl_inteiro(o.total_mensal)} ({f.pct_inteiro(o.reducao_pct)})"
                    for o in cen.operadoras))
            if r.avisos:
                st.warning("Conferir antes de enviar:\n\n" + "\n".join(f"- {a}" for a in r.avisos))
