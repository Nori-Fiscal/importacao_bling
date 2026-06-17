"""
app.py — NF-e Editor: Importação DUIMP / XML F5 → Bling
Streamlit UI principal.

Execute:
    streamlit run app.py
"""

import io
import zipfile
import pandas as pd
import streamlit as st

from database import (
    init_db, get_all_eans, upsert_eans, salvar_ean_manual,
    get_db_stats, listar_todos, buscar_camex_por_ncm, get_camex_stats,
    buscar_ttd409_por_ncm, get_ttd409_stats,
)
from excel_loader import carregar_planilha_ean, carregar_planilha_xmlf5
from xml_service import (
    auditar_ttd409_xml,
    gerar_relatorio_faltantes,
    gerar_relatorio_ttd409,
    processar_xml,
)
from products_db import (
    init_db_produtos, get_produtos_stats, listar_produtos_base,
    listar_divergencias, marcar_divergencia_resolvida,
    conferir_ncm, registrar_divergencia_ncm,
    importar_produtos_2026, importar_relprodutos01,
)
from decimal import Decimal


# ---------------------------------------------------------------------------
# Configuração da página
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="NF-e Editor · Importação",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_db()
init_db_produtos()

# ---------------------------------------------------------------------------
# CSS customizado
# ---------------------------------------------------------------------------

st.markdown("""
<style>
/* Fonte e fundo geral */
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }

/* Remove padding topo */
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; }

/* Header principal */
.app-header {
    background: linear-gradient(135deg, #0f2244 0%, #1a3a6e 100%);
    border-radius: 12px;
    padding: 1.4rem 2rem;
    margin-bottom: 1.5rem;
    color: white;
}
.app-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; letter-spacing: -0.3px; }
.app-header p  { margin: 0.25rem 0 0; font-size: 0.85rem; opacity: 0.75; }

/* Cards de stat */
.stat-card {
    background: #f8f9fb;
    border: 1px solid #e4e8ef;
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    text-align: center;
}
.stat-card .num  { font-size: 2rem; font-weight: 700; color: #1a3a6e; line-height: 1; }
.stat-card .lbl  { font-size: 0.72rem; color: #6b7280; margin-top: 0.2rem; text-transform: uppercase; letter-spacing: 0.5px; }

/* Seção de etapa */
.step-badge {
    display: inline-block;
    background: #e8eef8;
    color: #1a3a6e;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    padding: 0.25rem 0.7rem;
    border-radius: 20px;
    margin-bottom: 0.5rem;
}

/* Aviso EAN */
.ean-warning {
    background: #fff8e1;
    border-left: 4px solid #f59e0b;
    border-radius: 6px;
    padding: 0.8rem 1.1rem;
    margin-bottom: 1rem;
}

.camex-warning {
    background: #eef6ff;
    border-left: 4px solid #2563eb;
    border-radius: 6px;
    padding: 0.8rem 1.1rem;
    margin-bottom: 1rem;
}

.ttd409-danger {
    background: #fff1f2;
    border-left: 5px solid #dc2626;
    border-radius: 6px;
    padding: 0.9rem 1.1rem;
    margin-bottom: 1rem;
    color: #7f1d1d;
}

/* Botão primário extra */
div.stButton > button[kind="primary"] {
    background: #1a3a6e !important;
    color: white !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}

/* Sidebar */
[data-testid="stSidebar"] { background: #f5f7fb; }
[data-testid="stSidebar"] .block-container { padding-top: 1rem !important; }

/* Expander header */
.streamlit-expanderHeader { font-weight: 600; }

/* Divider suave */
hr { border: none; border-top: 1px solid #e4e8ef; margin: 1.2rem 0; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

def _init_state():
    defaults = {
        "stage": "input",          # input | ttd409_audit | fill_ean | results
        "xmls_carregados": [],     # [(nome, bytes), ...]
        "faltantes": [],           # lista de dicts
        "resultados": [],          # lista processada
        "mapa_fiscal": {},
        "camex_alertas_pre": [],
        "ttd409_alertas_pre": [],
        "ttd409_auditorias": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ---------------------------------------------------------------------------
# Helpers UI
# ---------------------------------------------------------------------------

def _stat_card(num, label):
    return f"""<div class="stat-card"><div class="num">{num}</div><div class="lbl">{label}</div></div>"""


def _step_badge(text):
    return f'<span class="step-badge">{text}</span>'


def _limpar_digitos(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())


def _resumir_camex_matches(matches, limite=6):
    partes = []
    for m in matches[:limite]:
        ex = f" Ex {m.get('ex')}" if m.get("ex") else ""
        vig = ""
        if m.get("inicio_vigencia") or m.get("fim_vigencia"):
            vig = f" ({m.get('inicio_vigencia') or 'sem inicio'} a {m.get('fim_vigencia') or 'sem termino'})"
        ato = f" - {m.get('ato_legal')}" if m.get("ato_legal") else ""
        partes.append(f"{m.get('lista', '')}{ex}{vig}{ato}")
    if len(matches) > limite:
        partes.append(f"+{len(matches) - limite} lista(s)")
    return " | ".join(partes)


def _render_camex_alertas(alertas, titulo="Itens encontrados na base CAMEX/Gecex"):
    if not alertas:
        return
    st.markdown(f"""
    <div class="camex-warning">
        <strong>{len(alertas)} item(ns)</strong> do(s) XML(s) possuem NCM em lista CAMEX/Gecex vigente.
        Confira a lista, EX e vigencia antes de concluir a importacao.
    </div>
    """, unsafe_allow_html=True)
    with st.expander(titulo, expanded=True):
        st.dataframe(pd.DataFrame(alertas), use_container_width=True, hide_index=True)


def _render_ttd409_alertas(alertas, titulo="Itens com risco de bloqueio TTD409"):
    if not alertas:
        return
    st.markdown(f"""
    <div class="ttd409-danger">
        <strong>ALERTA GRAVE TTD409:</strong> {len(alertas)} item(ns) do(s) XML(s)
        batem com mercadorias do Anexo Unico do Decreto SC 2.128/2009.
        Esses itens nao devem entrar no TTD409 sem revisao fiscal.
    </div>
    """, unsafe_allow_html=True)
    with st.expander(titulo, expanded=True):
        st.dataframe(pd.DataFrame(alertas), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# SIDEBAR — Base de EANs
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## 🗄️ Base de EANs")

    db_stats = get_db_stats()
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.metric("SKUs cadastrados", db_stats["total"])
    with col_s2:
        ult = (db_stats["ultima_atualizacao"] or "—")[:10]
        st.metric("Última atualização", ult)

    st.divider()

    # Upload para atualizar base de EAN
    st.markdown("**Atualizar base por planilha**")
    st.caption("Colunas esperadas: SKU · EAN (GTIN)")
    plan_ean_file = st.file_uploader(
        "Planilha EAN (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="plan_ean_upload",
        label_visibility="collapsed",
    )

    if plan_ean_file:
        if st.button("⬆️ Importar para a base", use_container_width=True):
            try:
                registros = carregar_planilha_ean(plan_ean_file.getvalue(), plan_ean_file.name)
                if not registros:
                    st.warning("Nenhum SKU/EAN encontrado na planilha.")
                else:
                    ins, upd, err = upsert_eans(registros)
                    st.success(
                        f"✅ **{ins}** inseridos · **{upd}** atualizados"
                        + (f" · **{err}** erros" if err else "")
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler planilha: {e}")

    st.divider()

    # Visualizar base
    with st.expander("🔍 Ver base de EANs", expanded=False):
        registros_db = listar_todos()
        if registros_db:
            df_db = pd.DataFrame(registros_db)
            st.dataframe(df_db, use_container_width=True, height=250)
        else:
            st.info("Base vazia. Importe uma planilha de EANs.")

    st.divider()

    st.markdown("## \U0001f4e6 Base de Produtos")
    prod_stats = get_produtos_stats()
    c1, c2 = st.columns(2)
    with c1: st.metric("Produtos", prod_stats["total"])
    with c2: st.metric("C/ descricao", prod_stats["com_descricao_padrao"])
    c3, c4 = st.columns(2)
    with c3: st.metric("NCM Albema", prod_stats["com_ncm_al"])
    with c4: st.metric("NCM Tonina", prod_stats["com_ncm_tonina"])
    st.caption(f"Divergencias NCM: {prod_stats['divergencias_pendentes']}")

    plan_prod = st.file_uploader("produtos_2026", type=["xlsx","xls"], key="plan_prod", label_visibility="collapsed")
    if plan_prod:
        if st.button("Importar descricoes", use_container_width=True, key="btn_prod"):
            try:
                import xlrd
                wb = xlrd.open_workbook(file_contents=plan_prod.getvalue())
                ws = wb.sheet_by_index(0)
                rows = []
                for r in range(1, ws.nrows):
                    sku = str(ws.cell_value(r,1)).strip().replace("\t","")
                    desc = str(ws.cell_value(r,2)).strip()
                    und = str(ws.cell_value(r,3)).strip()
                    ncm = str(ws.cell_value(r,4)).strip()
                    if sku: rows.append({"sku":sku,"descricao":desc,"unidade":und,"ncm":ncm})
                ins,upd,err = importar_produtos_2026(rows)
                st.success(f"{ins} ins, {upd} upd" + (f", {err} err" if err else ""))
                st.rerun()
            except Exception as e: st.error(f"Erro: {e}")

    plan_ncm = st.file_uploader("NCM Tonina", type=["xlsx","xls"], key="plan_ncm", label_visibility="collapsed")
    if plan_ncm:
        if st.button("Importar NCMs", use_container_width=True, key="btn_ncm"):
            try:
                import xlrd
                wb = xlrd.open_workbook(file_contents=plan_ncm.getvalue())
                ws = wb.sheet_by_index(0)
                rows = []
                ncols = ws.ncols
                h = [str(ws.cell_value(0,c)).strip().lower() for c in range(min(ncols,5))]
                if any("codigo interno" in x for x in h) or ncols <= 5:
                    for r in range(1, ws.nrows):
                        cod = str(ws.cell_value(r,0)).strip()
                        ncm = str(ws.cell_value(r,1)).strip()
                        if cod: rows.append({"codigo":cod,"ncm":ncm})
                else:
                    for r in range(1, ws.nrows):
                        cod = str(ws.cell_value(r,2)).strip()
                        ncm = str(ws.cell_value(r,24)).strip()
                        if cod: rows.append({"codigo":cod,"ncm":ncm})
                atu,sem,err = importar_relprodutos01(rows)
                st.success(f"{atu} NCMs, {sem} sem match" + (f", {err} err" if err else ""))
                st.rerun()
            except Exception as e: st.error(f"Erro: {e}")

    with st.expander("Ver base", expanded=False):
        prods = listar_produtos_base()
        if prods: st.dataframe(pd.DataFrame(prods), use_container_width=True, height=200)
        else: st.info("Vazia.")
    with st.expander(f"Divergencias NCM ({prod_stats['divergencias_pendentes']})", expanded=False):
        divs = listar_divergencias(apenas_pendentes=True)
        if divs:
            st.dataframe(pd.DataFrame(divs), use_container_width=True, height=150)
            if st.button("Resolver todas", use_container_width=True):
                for d in divs: marcar_divergencia_resolvida(d["id"])
                st.rerun()
        else: st.info("Nenhuma.")

    st.divider()

    st.markdown("## Base CAMEX/Gecex")
    camex_stats = get_camex_stats()
    c_cam_1, c_cam_2 = st.columns(2)
    with c_cam_1:
        st.metric("NCMs vigentes", camex_stats["ncm_vigentes"])
    with c_cam_2:
        st.metric("Registros", camex_stats["registros_vigentes"])
    st.caption(
        f"{camex_stats['fontes']} fonte(s) oficiais - atualizado em "
        f"{camex_stats['ultima_atualizacao'] or 'sem data'}"
    )

    st.divider()

    st.markdown("## TTD409")
    ttd409_stats = get_ttd409_stats()
    c_ttd_1, c_ttd_2 = st.columns(2)
    with c_ttd_1:
        st.metric("Itens legais", ttd409_stats["itens_legais"])
    with c_ttd_2:
        st.metric("NCMs/prefixos", ttd409_stats["total_registros"])
    st.caption(f"Decreto SC 2.128/2009 - atualizado em {ttd409_stats['ultima_atualizacao'] or 'sem data'}")




# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

st.markdown("""
<div class="app-header">
    <h1>NF-e Editor · Importação DUIMP / XML F5</h1>
    <p>Ajustes automáticos para emissão no Bling · EAN · ICMS zerado · IPI · PIS/COFINS</p>
</div>
""", unsafe_allow_html=True)


# ===========================================================================
# ETAPA 1 — INPUT
# ===========================================================================

if st.session_state.stage == "input":

    st.markdown(_step_badge("Etapa 1 de 3 · Carregar arquivos"), unsafe_allow_html=True)
    st.markdown("### Carregue os XMLs e a Planilha XML F5")

    col1, col2 = st.columns([2, 1])

    with col1:
        uploaded_xmls = st.file_uploader(
            "XMLs NF-e (um ou mais)",
            type=["xml"],
            accept_multiple_files=True,
            help="Selecione todos os XMLs que deseja processar de uma vez.",
        )

    with col2:
        plan_fiscal = st.file_uploader(
            "Planilha XML F5",
            type=["xlsx", "xls", "csv"],
            help="Planilha F5 ou espelho HAGN008 com SKU/Part Number, Siscomex, AFRMM e, se houver, Base/Alq/Valor de IBS e CBS.",
        )

    st.divider()

    col_ttd409, col_ttd409_info = st.columns([1, 2])
    with col_ttd409:
        auditar_ttd409 = st.button(
            "Conferir TTD409",
            use_container_width=True,
            type="primary",
            disabled=not uploaded_xmls,
        )

    with col_ttd409_info:
        if uploaded_xmls:
            st.info("Conferencia TTD409 disponivel para os XMLs carregados, sem alterar os arquivos.")
        else:
            st.info("Carregue um ou mais XMLs para conferir TTD409.")

    col_btn, col_info = st.columns([1, 2])
    with col_btn:
        analisar = st.button(
            "🔍 Analisar XMLs",
            use_container_width=True,
            type="primary",
            disabled=(not uploaded_xmls or not plan_fiscal),
        )

    with col_info:
        if not uploaded_xmls:
            st.info("Aguardando XMLs...")
        elif not plan_fiscal:
            st.info("Aguardando Planilha XML F5...")
        else:
            st.success(f"✅ {len(uploaded_xmls)} XML(s) · Planilha XML F5 pronta")

    if auditar_ttd409:
        xmls_carregados = [(f.name, f.getvalue()) for f in uploaded_xmls]
        auditorias = []
        with st.spinner("Conferindo NCMs na base TTD409..."):
            for nome, xml_bytes in xmls_carregados:
                itens, stats = auditar_ttd409_xml(xml_bytes, nome, buscar_ttd409_por_ncm)
                auditorias.append({
                    "nome_original": nome,
                    "itens": itens,
                    "stats": stats,
                })

        st.session_state.xmls_carregados = xmls_carregados
        st.session_state.ttd409_auditorias = auditorias
        st.session_state.stage = "ttd409_audit"
        st.rerun()

    if analisar:
        # Carrega planilha fiscal
        try:
            mapa_fiscal = carregar_planilha_xmlf5(plan_fiscal.getvalue(), plan_fiscal.name)
        except Exception as e:
            st.error(f"Erro ao ler Planilha XML F5: {e}")
            st.stop()

        # Carrega EANs da base
        mapa_ean = get_all_eans()

        # Pré-analisa XMLs para encontrar itens sem EAN
        xmls_carregados = [(f.name, f.getvalue()) for f in uploaded_xmls]
        faltantes = []
        camex_alertas_pre = []
        ttd409_alertas_pre = []

        for nome, xml_bytes in xmls_carregados:
            try:
                from lxml import etree
                parser = etree.XMLParser(recover=True, huge_tree=True)
                tree = etree.parse(io.BytesIO(xml_bytes), parser)
                root = tree.getroot()
                for det in root.xpath(".//*[local-name()='det']"):
                    prod = next(iter(det.xpath("./*[local-name()='prod']")), None)
                    if prod is None:
                        continue
                    cprod = next(iter(prod.xpath("./*[local-name()='cProd']/text()")), "").strip()
                    xprod = next(iter(prod.xpath("./*[local-name()='xProd']/text()")), "").strip()
                    ncm = next(iter(prod.xpath("./*[local-name()='NCM']/text()")), "").strip()
                    extipi = next(iter(prod.xpath("./*[local-name()='EXTIPI']/text()")), "").strip()

                    if ncm:
                        matches_ttd409 = buscar_ttd409_por_ncm(ncm)
                        if matches_ttd409:
                            ttd409_alertas_pre.append({
                                "Gravidade": "GRAVE",
                                "Arquivo XML": nome,
                                "nItem": det.get("nItem", ""),
                                "SKU": cprod,
                                "NCM": ncm,
                                "Descricao": xprod,
                                "Regra TTD409": " | ".join(
                                    f"Item {m.get('item')} - {m.get('descricao_legal', '')}"
                                    for m in matches_ttd409
                                ),
                                "Acao recomendada": "Nao incluir no TTD409 sem revisar descricao legal, NCM e excecoes aplicaveis.",
                            })

                        matches_camex = buscar_camex_por_ncm(ncm, extipi)
                        if matches_camex:
                            camex_alertas_pre.append({
                                "Arquivo XML": nome,
                                "nItem": det.get("nItem", ""),
                                "SKU": cprod,
                                "NCM": ncm,
                                "EXTIPI": extipi,
                                "Descricao": xprod,
                                "Qtd. listas": len(matches_camex),
                                "Listas CAMEX": _resumir_camex_matches(matches_camex),
                            })

                    if not cprod:
                        continue
                    if not mapa_ean.get(cprod):
                        cean = next(iter(prod.xpath("./*[local-name()='cEAN']/text()")), "").strip()
                        faltantes.append({
                            "Arquivo XML": nome,
                            "nItem": det.get("nItem", ""),
                            "SKU": cprod,
                            "Descrição": xprod,
                            "cEAN no XML": cean,
                            "EAN (preencher)": "",
                        })
            except Exception as e:
                st.warning(f"Não foi possível pré-analisar {nome}: {e}")

        st.session_state.xmls_carregados = xmls_carregados
        st.session_state.mapa_fiscal = mapa_fiscal
        st.session_state.faltantes = faltantes
        st.session_state.camex_alertas_pre = camex_alertas_pre
        st.session_state.ttd409_alertas_pre = ttd409_alertas_pre

        if faltantes:
            st.session_state.stage = "fill_ean"
        else:
            # Sem faltantes: processa direto
            resultados = []
            with st.spinner("Processando XMLs..."):
                mapa_ean = get_all_eans()
                for nome, xml_bytes in xmls_carregados:
                    xml_out, stats = processar_xml(
                        xml_bytes, mapa_ean, nome, st.session_state.mapa_fiscal,
                        buscar_camex_por_ncm, buscar_ttd409_por_ncm,
                        conferir_ncm, registrar_divergencia_ncm,
                    )
                    resultados.append({"nome_original": nome, "xml_processado": xml_out, "stats": stats})
            st.session_state.resultados
            st.session_state.stage = "results"

        st.rerun()


# ===========================================================================
# ETAPA 2 — PREENCHIMENTO MANUAL DE EAN
# ===========================================================================

elif st.session_state.stage == "ttd409_audit":

    auditorias = st.session_state.get("ttd409_auditorias", [])
    todos_itens = []
    erros = []
    for auditoria in auditorias:
        todos_itens.extend(auditoria.get("itens", []))
        stats = auditoria.get("stats", {})
        for erro in stats.get("erros", []):
            erros.append({
                "Arquivo XML": stats.get("arquivo", auditoria.get("nome_original", "")),
                "Erro": erro,
            })

    bloqueios = [r for r in todos_itens if r.get("Gravidade") == "GRAVE"]
    sem_ncm = [r for r in todos_itens if r.get("Status TTD409") == "SEM NCM NO XML"]
    ncm_distintos = len({str(r.get("NCM", "")).strip() for r in todos_itens if str(r.get("NCM", "")).strip()})

    st.markdown(_step_badge("Auditoria TTD409"), unsafe_allow_html=True)
    st.markdown("### Conferencia de NCMs contra TTD409")

    c1, c2, c3, c4, c5 = st.columns(5)
    for col, num, lbl in [
        (c1, len(auditorias), "XMLs analisados"),
        (c2, len(todos_itens), "Itens analisados"),
        (c3, len(bloqueios), "Bloqueios graves"),
        (c4, len(sem_ncm), "Itens sem NCM"),
        (c5, ncm_distintos, "NCMs distintos"),
    ]:
        col.markdown(_stat_card(num, lbl), unsafe_allow_html=True)

    st.divider()

    _render_ttd409_alertas(bloqueios, "Bloqueios graves TTD409")

    if sem_ncm:
        st.markdown(f"""
        <div class="ean-warning">
            <strong>{len(sem_ncm)} item(ns)</strong> estao sem NCM no XML.
            Esses itens precisam de revisao manual antes de qualquer enquadramento no TTD409.
        </div>
        """, unsafe_allow_html=True)
        with st.expander("Itens sem NCM", expanded=True):
            st.dataframe(pd.DataFrame(sem_ncm), use_container_width=True, hide_index=True)

    if erros:
        with st.expander(f"{len(erros)} XML(s) com erro de leitura", expanded=True):
            st.dataframe(pd.DataFrame(erros), use_container_width=True, hide_index=True)

    if todos_itens:
        ordenados = sorted(
            todos_itens,
            key=lambda r: (
                0 if r.get("Gravidade") == "GRAVE" else 1,
                r.get("Arquivo XML", ""),
                str(r.get("nItem", "")),
            ),
        )
        with st.expander("Todos os itens analisados", expanded=False):
            st.dataframe(pd.DataFrame(ordenados), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum item encontrado nos XMLs analisados.")

    st.markdown("#### Resumo por arquivo")
    for auditoria in auditorias:
        stats = auditoria.get("stats", {})
        nome = auditoria.get("nome_original", stats.get("arquivo", ""))
        itens = auditoria.get("itens", [])
        bloqueios_arquivo = [r for r in itens if r.get("Gravidade") == "GRAVE"]

        with st.container(border=True):
            status_icon = "🔴" if bloqueios_arquivo or stats.get("erros") else "🟢"
            st.markdown(f"**{status_icon} {nome}**")
            c_arq_1, c_arq_2, c_arq_3 = st.columns(3)
            c_arq_1.write(f"Itens: {stats.get('itens_analisados', 0)}")
            c_arq_2.write(f"Bloqueios TTD409: {stats.get('bloqueios_ttd409', 0)}")
            c_arq_3.write(f"Sem NCM: {stats.get('itens_sem_ncm', 0)}")

            if bloqueios_arquivo:
                with st.expander("Bloqueios deste arquivo", expanded=True):
                    st.dataframe(pd.DataFrame(bloqueios_arquivo), use_container_width=True, hide_index=True)

    st.divider()

    col_dl, col_voltar, col_novo = st.columns(3)
    with col_dl:
        rel_ttd409 = gerar_relatorio_ttd409(auditorias)
        st.download_button(
            label="Baixar relatorio TTD409 (.xlsx)",
            data=rel_ttd409,
            file_name="relatorio_ttd409_xmls_importacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_voltar:
        if st.button("Voltar", use_container_width=True):
            st.session_state.stage = "input"
            st.rerun()
    with col_novo:
        if st.button("Processar novos XMLs", use_container_width=True):
            st.session_state.stage = "input"
            st.session_state.xmls_carregados = []
            st.session_state.faltantes = []
            st.session_state.resultados = []
            st.session_state.camex_alertas_pre = []
            st.session_state.ttd409_alertas_pre = []
            st.session_state.ttd409_auditorias = []
            st.rerun()


elif st.session_state.stage == "fill_ean":

    faltantes = st.session_state.faltantes

    st.markdown(_step_badge("Etapa 2 de 3 · Preencher EANs ausentes"), unsafe_allow_html=True)
    st.markdown("### Itens sem EAN na base")
    _render_ttd409_alertas(st.session_state.get("ttd409_alertas_pre", []))
    _render_camex_alertas(st.session_state.get("camex_alertas_pre", []))

    st.markdown(f"""
    <div class="ean-warning">
        ⚠️ <strong>{len(faltantes)} item(ns)</strong> não encontrados na base de EANs.
        Preencha os EANs abaixo e eles serão salvos automaticamente na base.
        Campos deixados em branco serão preservados como estão no XML.
    </div>
    """, unsafe_allow_html=True)

    df_faltantes = pd.DataFrame(faltantes)

    edited = st.data_editor(
        df_faltantes,
        column_config={
            "EAN (preencher)": st.column_config.TextColumn(
                "EAN (preencher)",
                help="Digite o EAN/GTIN de 8, 12, 13 ou 14 dígitos.",
                width="medium",
            ),
            "Arquivo XML": st.column_config.TextColumn(width="medium"),
            "nItem": st.column_config.TextColumn(width="small"),
            "SKU": st.column_config.TextColumn(width="medium"),
            "Descrição": st.column_config.TextColumn(width="large"),
            "cEAN no XML": st.column_config.TextColumn(width="medium"),
        },
        disabled=["Arquivo XML", "nItem", "SKU", "Descrição", "cEAN no XML"],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
    )

    st.divider()

    col_b1, col_b2, _ = st.columns([1, 1, 2])

    with col_b1:
        processar_btn = st.button("✅ Salvar EANs e Processar", type="primary", use_container_width=True)
    with col_b2:
        if st.button("⬅️ Voltar", use_container_width=True):
            st.session_state.stage = "input"
            st.rerun()

    if processar_btn:
        # Salva EANs preenchidos manualmente
        salvos = 0
        for _, row in edited.iterrows():
            ean_raw = str(row.get("EAN (preencher)", "")).strip()
            if ean_raw:
                ean_digits = _limpar_digitos(ean_raw)
                if ean_digits:
                    salvar_ean_manual(row["SKU"], ean_digits, row.get("Descrição", ""))
                    salvos += 1

        if salvos:
            st.toast(f"{salvos} EAN(s) salvos na base!", icon="✅")

        # Processa XMLs
        mapa_ean = get_all_eans()
        resultados = []
        with st.spinner("Processando XMLs..."):
            for nome, xml_bytes in st.session_state.xmls_carregados:
                xml_out, stats = processar_xml(
                    xml_bytes, mapa_ean, nome, st.session_state.mapa_fiscal,
                    buscar_camex_por_ncm, buscar_ttd409_por_ncm,
                )
                resultados.append({"nome_original": nome, "xml_processado": xml_out, "stats": stats})

        st.session_state.resultados = resultados
        st.session_state.stage = "results"
        st.rerun()


# ===========================================================================
# ETAPA 3 — RESULTADOS
# ===========================================================================

elif st.session_state.stage == "results":

    resultados = st.session_state.resultados

    st.markdown(_step_badge("Etapa 3 de 3 · Resultados"), unsafe_allow_html=True)
    st.markdown("### Processamento concluído")

    # Totalizadores rápidos
    total_xml = len(resultados)
    total_ok  = sum(1 for r in resultados if not r["stats"]["erros"] and r["xml_processado"])
    total_ean_criados = sum(r["stats"].get("ean_criados", 0) for r in resultados)
    total_ean_ausentes = sum(r["stats"].get("ean_ausentes", 0) for r in resultados)
    total_icms = sum(r["stats"].get("icms_zerados", 0) for r in resultados)
    total_ibscbs = sum(r["stats"].get("ibscbs_itens_gerados", 0) for r in resultados)
    total_camex = sum(r["stats"].get("camex_alertas", 0) for r in resultados)
    total_ttd409 = sum(r["stats"].get("ttd409_bloqueios", 0) for r in resultados)
    total_desc = sum(r["stats"].get("descricoes_padronizadas", 0) for r in resultados)
    total_ncm_div = sum(r["stats"].get("ncm_divergencias", 0) for r in resultados)
    total_ncm_sem = sum(len(r["stats"].get("ncm_sem_base", [])) for r in resultados)

    c1, c2, c3, c4, c5, c6, c7, c10 = st.columns(10)
    for col, num, lbl in [
        (c1, total_xml, "XMLs processados"),
        (c2, total_ok, "Com sucesso"),
        (c3, total_ean_criados, "EANs inseridos"),
        (c4, total_ean_ausentes, "Itens sem EAN"),
        (c5, total_desc, "Descricoes padronizadas"),
        (c6, total_ncm_div, "Divergencias NCM"),
        (c7, total_ncm_sem, "SKUs sem base"),
        (c8, total_ttd409, "Bloqueios TTD409"),
        (c9, total_camex, "Alertas CAMEX"),
        (c10, total_icms, "ICMS zerados"),
    ]:
        col.markdown(_stat_card(num, lbl), unsafe_allow_html=True)

    st.divider()

    # Alerta se ainda há itens sem EAN
    ttd409_alertas_resultados = []
    for r in resultados:
        ttd409_alertas_resultados.extend(r["stats"].get("ttd409_itens_detalhado", []))
    _render_ttd409_alertas(ttd409_alertas_resultados, "Detalhe TTD409 por item")

    camex_alertas_resultados = []
    for r in resultados:
        camex_alertas_resultados.extend(r["stats"].get("camex_itens_detalhado", []))
    _render_camex_alertas(camex_alertas_resultados, "Detalhe CAMEX/Gecex por item")

    if total_ean_ausentes > 0:
        st.markdown(f"""
        <div class="ean-warning">
            ⚠️ <strong>{total_ean_ausentes} item(ns)</strong> ainda sem EAN na base.
            O XML foi preservado (cEAN/cEANTrib originais mantidos).
            Baixe o relatório Excel abaixo para ver o detalhe.
        </div>
        """, unsafe_allow_html=True)

    # Por arquivo
    st.markdown("#### Detalhes por arquivo")

    for r in resultados:
        nome  = r["nome_original"]
        stats = r["stats"]
        xml_out = r["xml_processado"]

        with st.container(border=True):
            col_nome, col_dl = st.columns([4, 1])
            with col_nome:
                status_icon = "🔴" if stats["erros"] or stats.get("ttd409_bloqueios", 0) else ("🟡" if stats["ean_ausentes"] or stats.get("camex_alertas", 0) else "🟢")
                st.markdown(f"**{status_icon} {nome}**")

            with col_dl:
                if xml_out:
                    base = nome.rsplit(".", 1)[0]
                    st.download_button(
                        label="⬇️ Download XML",
                        data=xml_out,
                        file_name=f"{base}_processado.xml",
                        mime="application/xml",
                        use_container_width=True,
                        key=f"dl_{nome}",
                    )

            if stats["erros"]:
                for e in stats["erros"]:
                    st.error(e)
                continue

            # Stats em colunas compactas
            s = stats
            col_a, col_b, col_c, col_d = st.columns(4)
            with col_a:
                st.markdown("**EAN**")
                st.write(f"Criados: {s['ean_criados']} · Atualizados: {s['ean_atualizados']}")
                st.write(f"Preservados (sem base): {s['ean_preservados']} · Ausentes: {s['ean_ausentes']}")
            with col_b:
                st.markdown("**ICMS / Fiscal**")
                st.write(f"CFOP ajustados: {s.get('cfop_ajustados', 0)} · ICMS zerados: {s['icms_zerados']}")
                st.write(f"infCpl limpos: {s['infcpl_limpa']} · cFabricante: {s['cfabricante_inseridos']} ins / {s['cfabricante_atualizados']} upd")
            with col_c:
                st.markdown("**IPI / PIS / COFINS / IBS-CBS**")
                st.write(f"IPI CST→03: {s['ipi_cst_alterados']+s['ipi_cst_criados']}")
                st.write(f"PIS CST→50: {s['pis_cst_alterados']} · COFINS: {s['cofins_cst_alterados']}")
                st.write(f"IBS/CBS: {s.get('ibscbs_itens_gerados', 0)} item(ns) · total: {s.get('ibscbs_totais_gerados', 0)}")

            with col_d:
                st.markdown("**TTD409 / CAMEX**")
                st.write(f"Bloqueios TTD409: {s.get('ttd409_bloqueios', 0)}")
                st.write(f"Alertas CAMEX: {s.get('camex_alertas', 0)}")

            # Itens sem EAN deste arquivo
            falt = s.get("faltantes_detalhado", [])
            if falt:
                with st.expander(f"⚠️ {len(falt)} item(ns) sem EAN neste arquivo"):
                    st.dataframe(pd.DataFrame(falt), use_container_width=True, hide_index=True)

            ttd409_itens = s.get("ttd409_itens_detalhado", [])
            if ttd409_itens:
                with st.expander(f"ALERTA GRAVE: {len(ttd409_itens)} item(ns) com risco TTD409 neste arquivo", expanded=True):
                    st.dataframe(pd.DataFrame(ttd409_itens), use_container_width=True, hide_index=True)

            camex_itens = s.get("camex_itens_detalhado", [])
            if camex_itens:
                with st.expander(f"{len(camex_itens)} item(ns) em lista CAMEX/Gecex neste arquivo"):
                    st.dataframe(pd.DataFrame(camex_itens), use_container_width=True, hide_index=True)

            desc_nao = s.get("descricoes_nao_encontradas", [])
            if desc_nao:
                with st.expander(f"\u26a0\ufe0f {len(desc_nao)} descricao(oes) nao encontrada(s)"):
                    st.dataframe(pd.DataFrame(desc_nao), use_container_width=True, hide_index=True)

            ncm_div = s.get("ncm_divergencias_detalhado", [])
            if ncm_div:
                with st.expander(f"\u26a0\ufe0f {len(ncm_div)} divergencia(s) de NCM", expanded=True):
                    st.dataframe(pd.DataFrame(ncm_div), use_container_width=True, hide_index=True)

            # Avisos
            if s["avisos"]:
                with st.expander(f"💬 {len(s['avisos'])} aviso(s)"):
                    for a in s["avisos"][:50]:
                        st.write(f"– {a}")
                    if len(s["avisos"]) > 50:
                        st.write(f"... +{len(s['avisos'])-50} avisos omitidos")

    st.divider()

    # Downloads globais
    col_z1, col_z2, col_z3 = st.columns(3)

    with col_z1:
        if len(resultados) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for r in resultados:
                    if r["xml_processado"]:
                        base = r["nome_original"].rsplit(".", 1)[0]
                        zf.writestr(f"{base}_processado.xml", r["xml_processado"])
            zip_buf.seek(0)
            st.download_button(
                label="⬇️ Baixar todos os XMLs (.zip)",
                data=zip_buf,
                file_name="xmls_processados.zip",
                mime="application/zip",
                use_container_width=True,
            )

    with col_z2:
        rel_bytes = gerar_relatorio_faltantes(resultados)
        st.download_button(
            label="📊 Relatório Excel (SKUs sem EAN)",
            data=rel_bytes,
            file_name="relatorio_skus_sem_ean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_z3:
        if st.button("🔄 Processar novos XMLs", use_container_width=True):
            st.session_state.stage = "input"
            st.session_state.xmls_carregados = []
            st.session_state.faltantes = []
            st.session_state.resultados = []
            st.session_state.camex_alertas_pre = []
            st.session_state.ttd409_alertas_pre = []
            st.session_state.ttd409_auditorias = []
            st.rerun()
