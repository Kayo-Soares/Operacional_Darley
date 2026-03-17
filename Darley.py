import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(
    page_title="Painel de Expedição",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# ESTILO
# =========================
st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(180deg, #0b1220 0%, #111827 100%);
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    .hero {
        padding: 1.2rem 1.4rem;
        border-radius: 18px;
        background: linear-gradient(135deg, rgba(249,115,22,0.18), rgba(14,165,233,0.12));
        border: 1px solid rgba(255,255,255,0.08);
        margin-bottom: 1rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
    }
    .hero h1 {
        font-size: 2rem;
        margin: 0;
        color: #f8fafc;
    }
    .hero p {
        font-size: 0.95rem;
        color: #cbd5e1;
        margin-top: 0.35rem;
        margin-bottom: 0;
    }
    .kpi-card {
        background: rgba(17,24,39,0.78);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 18px;
        padding: 1rem 1rem 0.8rem 1rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.15);
    }
    .kpi-title {
        color: #94a3b8;
        font-size: 0.88rem;
        margin-bottom: 0.2rem;
    }
    .kpi-value {
        color: #f8fafc;
        font-size: 1.7rem;
        font-weight: 700;
        line-height: 1.1;
    }
    .kpi-sub {
        color: #cbd5e1;
        font-size: 0.82rem;
        margin-top: 0.35rem;
    }
    div[data-testid="stMetric"] {
        background: rgba(17,24,39,0.78);
        border: 1px solid rgba(255,255,255,0.08);
        padding: 14px;
        border-radius: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# FUNÇÕES
# =========================

def padronizar_texto(s):
    if pd.isna(s):
        return np.nan
    s = str(s).strip()
    return s if s else np.nan

@st.cache_data(show_spinner=False)
def carregar_dados_multiplos(arquivos):
    bases_car, bases_prob, bases_bi, bases_dina = [], [], [], []
    abas_identificadas = set()
    fontes = []

    for idx, arquivo in enumerate(arquivos, start=1):
        if arquivo is None:
            continue

        nome_fonte = getattr(arquivo, "name", f"Base {idx}")
        xls = pd.ExcelFile(arquivo)
        abas = xls.sheet_names
        abas_identificadas.update(abas)

        df_car = pd.read_excel(arquivo, sheet_name="CARREGAMENTO") if "CARREGAMENTO" in abas else pd.DataFrame()
        df_prob = pd.read_excel(arquivo, sheet_name="PROBLEMATICOS") if "PROBLEMATICOS" in abas else pd.DataFrame()
        df_bi = pd.read_excel(arquivo, sheet_name="BI") if "BI" in abas else pd.DataFrame()
        df_dina = pd.read_excel(arquivo, sheet_name="DINA") if "DINA" in abas else pd.DataFrame()

        for nome_df, df in {
            "CARREGAMENTO": df_car,
            "PROBLEMATICOS": df_prob,
            "BI": df_bi,
            "DINA": df_dina,
        }.items():
            if not df.empty:
                df["Arquivo origem"] = nome_fonte
                df["Base arquivo"] = f"Base {idx}"
                df["Aba origem"] = nome_df

        for df in [df_car, df_prob]:
            if not df.empty:
                for col in [
                    "Tempo de digitalização",
                    "Tempo de upload",
                    "Saída do dia",
                ]:
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col], errors="coerce")

                for col in [
                    "Digitalizador",
                    "Base de escaneamento",
                    "Base Destino",
                    "Município de Destino",
                    "Estado da cidade de destino",
                    "Nome da linha",
                    "Tipo problemático",
                    "Descrição de Pacote Problemático",
                    "Número do lote",
                    "Número de pedido JMS",
                ]:
                    if col in df.columns:
                        df[col] = df[col].apply(padronizar_texto)

        bases_car.append(df_car)
        bases_prob.append(df_prob)
        bases_bi.append(df_bi)
        bases_dina.append(df_dina)
        fontes.append(nome_fonte)

    df_car = pd.concat([df for df in bases_car if not df.empty], ignore_index=True) if any(not df.empty for df in bases_car) else pd.DataFrame()
    df_prob = pd.concat([df for df in bases_prob if not df.empty], ignore_index=True) if any(not df.empty for df in bases_prob) else pd.DataFrame()
    df_bi = pd.concat([df for df in bases_bi if not df.empty], ignore_index=True) if any(not df.empty for df in bases_bi) else pd.DataFrame()
    df_dina = pd.concat([df for df in bases_dina if not df.empty], ignore_index=True) if any(not df.empty for df in bases_dina) else pd.DataFrame()

    if not df_car.empty and "Tempo de digitalização" in df_car.columns:
        df_car["Data"] = df_car["Tempo de digitalização"].dt.date
        df_car["Hora"] = df_car["Tempo de digitalização"].dt.hour
        df_car["DiaHora"] = df_car["Tempo de digitalização"].dt.floor("H")
    else:
        df_car["Data"] = pd.NaT
        df_car["Hora"] = np.nan
        df_car["DiaHora"] = pd.NaT

    if not df_prob.empty and "Tempo de digitalização" in df_prob.columns:
        df_prob["Data"] = df_prob["Tempo de digitalização"].dt.date
        df_prob["Hora"] = df_prob["Tempo de digitalização"].dt.hour
    else:
        df_prob["Data"] = pd.NaT
        df_prob["Hora"] = np.nan

    return df_car, df_prob, df_bi, df_dina, sorted(abas_identificadas), fontes


def aplicar_filtros(df_car, df_prob):
    st.sidebar.markdown("## Filtros")

    datas = sorted([d for d in df_car["Data"].dropna().unique().tolist()]) if "Data" in df_car.columns else []
    base_opts = sorted(df_car.get("Base Destino", pd.Series(dtype=str)).dropna().unique().tolist())
    mun_opts = sorted(df_car.get("Município de Destino", pd.Series(dtype=str)).dropna().unique().tolist())
    dig_opts = sorted(df_car.get("Digitalizador", pd.Series(dtype=str)).dropna().unique().tolist())
    linha_opts = sorted(df_car.get("Nome da linha", pd.Series(dtype=str)).dropna().unique().tolist())
    arquivo_opts = sorted(df_car.get("Base arquivo", pd.Series(dtype=str)).dropna().unique().tolist())

    data_sel = st.sidebar.multiselect("Data", datas, default=datas)
    arquivo_sel = st.sidebar.multiselect("Arquivo / base carregada", arquivo_opts, default=arquivo_opts)
    base_sel = st.sidebar.multiselect("Base destino", base_opts)
    mun_sel = st.sidebar.multiselect("Município destino", mun_opts)
    dig_sel = st.sidebar.multiselect("Digitalizador", dig_opts)
    linha_sel = st.sidebar.multiselect("Linha", linha_opts)

    def filtrar(df):
        if df.empty:
            return df
        out = df.copy()
        if data_sel:
            out = out[out["Data"].isin(data_sel)]
        if arquivo_sel and "Base arquivo" in out.columns:
            out = out[out["Base arquivo"].isin(arquivo_sel)]
        if base_sel and "Base Destino" in out.columns:
            out = out[out["Base Destino"].isin(base_sel)]
        if mun_sel and "Município de Destino" in out.columns:
            out = out[out["Município de Destino"].isin(mun_sel)]
        if dig_sel and "Digitalizador" in out.columns:
            out = out[out["Digitalizador"].isin(dig_sel)]
        if linha_sel and "Nome da linha" in out.columns:
            out = out[out["Nome da linha"].isin(linha_sel)]
        return out

    return filtrar(df_car), filtrar(df_prob)


def card(titulo, valor, subtitulo=""):
    st.markdown(
        f"""
        <div class='kpi-card'>
            <div class='kpi-title'>{titulo}</div>
            <div class='kpi-value'>{valor}</div>
            <div class='kpi-sub'>{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def grafico_vazio(texto="Sem dados para os filtros selecionados"):
    fig = go.Figure()
    fig.add_annotation(text=texto, x=0.5, y=0.5, showarrow=False, font=dict(size=18))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    fig.update_layout(height=350, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    return fig


# =========================
# CABEÇALHO
# =========================
st.markdown(
    """
    <div class='hero'>
        <h1>📦 Painel Gerencial de Expedição</h1>
        <p>Acompanhamento de volume expedido, produtividade por digitalizador, destinos, lotes e pacotes problemáticos em um layout executivo.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

arquivo_default = Path("/mnt/data/EXPEDIÇÃO.xlsx")

st.sidebar.markdown("### Bases de entrada")
arquivo_1 = st.sidebar.file_uploader("Base 1 - planilha de inserido em lote", type=["xlsx"], key="base_1")
arquivo_2 = st.sidebar.file_uploader("Base 2 - planilha de bipe de pacote problematico", type=["xlsx"], key="base_2")

arquivos_para_carga = [arq for arq in [arquivo_1, arquivo_2] if arq is not None]
if not arquivos_para_carga and arquivo_default.exists():
    arquivos_para_carga = [arquivo_default]

if not arquivos_para_carga:
    st.warning("Envie pelo menos uma planilha .xlsx para iniciar a análise.")
    st.stop()

# =========================
# CARGA
# =========================
df_car, df_prob, df_bi, df_dina, abas, fontes = carregar_dados_multiplos(arquivos_para_carga)

if df_car.empty:
    st.error("A aba CARREGAMENTO não foi encontrada ou está vazia.")
    st.stop()

car_f, prob_f = aplicar_filtros(df_car, df_prob)

# =========================
# KPIs
# =========================
vol_total = len(car_f)
pedidos_unicos = car_f["Número de pedido JMS"].nunique() if "Número de pedido JMS" in car_f.columns else 0
lotes = car_f["Número do lote"].nunique() if "Número do lote" in car_f.columns else 0
digitalizadores = car_f["Digitalizador"].nunique() if "Digitalizador" in car_f.columns else 0
bases_dest = car_f["Base Destino"].nunique() if "Base Destino" in car_f.columns else 0
prob_total = len(prob_f)
prob_rate = (prob_total / vol_total * 100) if vol_total else 0
media_dig = (vol_total / digitalizadores) if digitalizadores else 0

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    card("Volumes expedidos", f"{vol_total:,.0f}".replace(",", "."), "Registros somados das bases carregadas")
with c2:
    card("Pedidos únicos", f"{pedidos_unicos:,.0f}".replace(",", "."), "Número de pedido JMS")
with c3:
    card("Lotes", f"{lotes:,.0f}".replace(",", "."), "Lotes movimentados")
with c4:
    card("Digitalizadores", f"{digitalizadores:,.0f}".replace(",", "."), f"Média de {media_dig:,.0f}".replace(",", ".") + " por pessoa")
with c5:
    card("Bases destino", f"{bases_dest:,.0f}".replace(",", "."), "Cobertura operacional")
with c6:
    card("% problemáticos", f"{prob_rate:.2f}%", f"{prob_total:,.0f}".replace(",", ".") + " ocorrências")

# =========================
# RESUMO DE CARGA
# =========================
if fontes:
    resumo_fontes = " | ".join([f"{i+1}: {nome}" for i, nome in enumerate(fontes)])
    st.caption(f"Bases carregadas: {resumo_fontes}")

# =========================
# INSIGHTS RÁPIDOS
# =========================
ins1, ins2, ins3 = st.columns(3)
with ins1:
    if not car_f.empty and "Digitalizador" in car_f.columns:
        top_dig = car_f["Digitalizador"].value_counts().head(1)
        nome = top_dig.index[0] if not top_dig.empty else "-"
        qtd = int(top_dig.iloc[0]) if not top_dig.empty else 0
        st.info(f"**Maior produtividade:** {nome} com **{qtd:,}** bipagens.".replace(",", "."))
with ins2:
    if not car_f.empty and "Base Destino" in car_f.columns:
        top_base = car_f["Base Destino"].value_counts().head(1)
        nome = top_base.index[0] if not top_base.empty else "-"
        qtd = int(top_base.iloc[0]) if not top_base.empty else 0
        st.info(f"**Base com maior volume:** {nome} com **{qtd:,}** pacotes.".replace(",", "."))
with ins3:
    if not prob_f.empty and "Tipo problemático" in prob_f.columns:
        top_prob = prob_f["Tipo problemático"].value_counts().head(1)
        nome = top_prob.index[0] if not top_prob.empty else "-"
        qtd = int(top_prob.iloc[0]) if not top_prob.empty else 0
        st.info(f"**Principal causa crítica:** {nome} com **{qtd:,}** registros.".replace(",", "."))

# =========================
# ABAS
# =========================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Visão Executiva",
    "Produtividade",
    "Destinos",
    "Problemáticos",
    "Detalhamento",
])

with tab1:
    g1, g2 = st.columns([1.4, 1])

    with g1:
        if not car_f.empty and "DiaHora" in car_f.columns:
            vol_hora = (
                car_f.dropna(subset=["DiaHora"])
                .groupby("DiaHora")
                .size()
                .reset_index(name="Volumes")
                .sort_values("DiaHora")
            )
            fig = px.area(
                vol_hora,
                x="DiaHora",
                y="Volumes",
                title="Evolução do volume expedido ao longo do tempo",
            )
            fig.update_layout(height=380)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with g2:
        fig = go.Figure(
            go.Indicator(
                mode="gauge+number",
                value=prob_rate,
                number={"suffix": "%"},
                title={"text": "Taxa de problemáticos"},
                gauge={
                    "axis": {"range": [None, max(10, round(prob_rate * 1.6, 2) + 1)]},
                    "bar": {"thickness": 0.35},
                    "steps": [
                        {"range": [0, 2], "color": "rgba(34,197,94,0.35)"},
                        {"range": [2, 5], "color": "rgba(245,158,11,0.30)"},
                        {"range": [5, 100], "color": "rgba(239,68,68,0.25)"},
                    ],
                },
            )
        )
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)

    l1, l2 = st.columns(2)
    with l1:
        if not car_f.empty and "Base Destino" in car_f.columns:
            top_bases = car_f["Base Destino"].value_counts().head(10).reset_index()
            top_bases.columns = ["Base Destino", "Volumes"]
            fig = px.bar(top_bases, x="Volumes", y="Base Destino", orientation="h", title="Top 10 bases destino")
            fig.update_layout(height=420, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)
    with l2:
        if not car_f.empty and "Digitalizador" in car_f.columns:
            top_digs = car_f["Digitalizador"].value_counts().head(10).reset_index()
            top_digs.columns = ["Digitalizador", "Volumes"]
            fig = px.bar(top_digs, x="Volumes", y="Digitalizador", orientation="h", title="Top 10 digitalizadores")
            fig.update_layout(height=420, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

with tab2:
    p1, p2 = st.columns([1.2, 1])
    with p1:
        if not car_f.empty and "Hora" in car_f.columns:
            prod_hora = car_f.dropna(subset=["Hora"]).groupby("Hora").size().reset_index(name="Volumes")
            fig = px.bar(prod_hora, x="Hora", y="Volumes", title="Volume por hora da digitalização")
            fig.update_layout(height=380)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with p2:
        if not car_f.empty and "Digitalizador" in car_f.columns:
            por_dig = car_f["Digitalizador"].value_counts().reset_index()
            por_dig.columns = ["Digitalizador", "Volumes"]
            por_dig["% Part."] = por_dig["Volumes"] / por_dig["Volumes"].sum() * 100
            st.dataframe(por_dig.head(15), use_container_width=True, hide_index=True)
        else:
            st.dataframe(pd.DataFrame())

    if not car_f.empty and {"Digitalizador", "Hora"}.issubset(car_f.columns):
        heat = car_f.dropna(subset=["Digitalizador", "Hora"]).groupby(["Digitalizador", "Hora"]).size().reset_index(name="Volumes")
        heat_piv = heat.pivot(index="Digitalizador", columns="Hora", values="Volumes").fillna(0)
        fig = px.imshow(
            heat_piv,
            aspect="auto",
            title="Heatmap de produtividade por digitalizador e hora",
            text_auto=True,
        )
        fig.update_layout(height=520)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.plotly_chart(grafico_vazio(), use_container_width=True)

with tab3:
    d1, d2 = st.columns(2)
    with d1:
        if not car_f.empty and "Município de Destino" in car_f.columns:
            top_mun = car_f["Município de Destino"].value_counts().head(15).reset_index()
            top_mun.columns = ["Município", "Volumes"]
            fig = px.bar(top_mun, x="Volumes", y="Município", orientation="h", title="Top municípios destino")
            fig.update_layout(height=450, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with d2:
        if not car_f.empty and "Estado da cidade de destino" in car_f.columns:
            uf = car_f["Estado da cidade de destino"].value_counts().head(10).reset_index()
            uf.columns = ["UF", "Volumes"]
            fig = px.pie(uf, values="Volumes", names="UF", hole=0.55, title="Participação por UF destino")
            fig.update_layout(height=450)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    if not car_f.empty and {"Base Destino", "Município de Destino"}.issubset(car_f.columns):
        cruz = (
            car_f.groupby(["Base Destino", "Município de Destino"]).size().reset_index(name="Volumes")
            .sort_values("Volumes", ascending=False)
            .head(30)
        )
        st.dataframe(cruz, use_container_width=True, hide_index=True)

with tab4:
    pr1, pr2 = st.columns([1.1, 1])
    with pr1:
        if not prob_f.empty and "Tipo problemático" in prob_f.columns:
            tipo = prob_f["Tipo problemático"].value_counts().reset_index()
            tipo.columns = ["Tipo problemático", "Ocorrências"]
            fig = px.bar(tipo, x="Ocorrências", y="Tipo problemático", orientation="h", title="Distribuição dos problemáticos")
            fig.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem registros de pacotes problemáticos"), use_container_width=True)

    with pr2:
        if not prob_f.empty and "Base Destino" in prob_f.columns:
            base_prob = prob_f["Base Destino"].value_counts().head(12).reset_index()
            base_prob.columns = ["Base Destino", "Ocorrências"]
            fig = px.bar(base_prob, x="Base Destino", y="Ocorrências", title="Bases com mais ocorrências")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    if not prob_f.empty:
        cols_exibir = [
            c for c in [
                "Número de pedido JMS",
                "Tipo problemático",
                "Descrição de Pacote Problemático",
                "Base Destino",
                "Município de Destino",
                "Tempo de digitalização",
            ] if c in prob_f.columns
        ]
        st.dataframe(prob_f[cols_exibir].head(300), use_container_width=True, hide_index=True)

with tab5:
    st.markdown("### Base analítica")
    modo = st.radio("Tabela", ["CARREGAMENTO", "PROBLEMATICOS"], horizontal=True)
    if modo == "CARREGAMENTO":
        st.dataframe(car_f, use_container_width=True, hide_index=True, height=520)
        csv = car_f.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Baixar CARREGAMENTO filtrado", csv, "carregamento_filtrado.csv", "text/csv")
    else:
        st.dataframe(prob_f, use_container_width=True, hide_index=True, height=520)
        csv = prob_f.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Baixar PROBLEMATICOS filtrado", csv, "problematicos_filtrado.csv", "text/csv")

st.caption(f"Abas identificadas nas bases carregadas: {', '.join(abas)}")
