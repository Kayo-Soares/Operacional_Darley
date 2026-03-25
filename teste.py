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

st.markdown(
    """
    <style>
    .main { background: linear-gradient(180deg, #0b1220 0%, #111827 100%); }
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    .hero {
        padding: 1.2rem 1.4rem; border-radius: 18px;
        background: linear-gradient(135deg, rgba(249,115,22,0.18), rgba(14,165,233,0.12));
        border: 1px solid rgba(255,255,255,0.08); margin-bottom: 1rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
    }
    .hero h1 { font-size: 2rem; margin: 0; color: #f8fafc; }
    .hero p { font-size: 0.95rem; color: #cbd5e1; margin-top: 0.35rem; margin-bottom: 0; }
    .kpi-card {
        background: rgba(17,24,39,0.78); border: 1px solid rgba(255,255,255,0.08);
        border-radius: 18px; padding: 1rem 1rem 0.8rem 1rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.15);
    }
    .kpi-title { color: #94a3b8; font-size: 0.88rem; margin-bottom: 0.2rem; }
    .kpi-value { color: #f8fafc; font-size: 1.7rem; font-weight: 700; line-height: 1.1; }
    .kpi-sub { color: #cbd5e1; font-size: 0.82rem; margin-top: 0.35rem; }
    div[data-testid="stMetric"] {
        background: rgba(17,24,39,0.78); border: 1px solid rgba(255,255,255,0.08);
        padding: 14px; border-radius: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

COLS_TEXTO = [
    "Digitalizador", "Base de escaneamento", "Base Destino", "Município de Destino",
    "Estado da cidade de destino", "Nome da linha", "Tipo problemático",
    "Descrição de Pacote Problemático", "Número do lote", "Número de pedido JMS",
    "Tipo de bipagem",
]
COLS_DATA = ["Tempo de digitalização", "Tempo de upload", "Saída do dia"]


def padronizar_texto(s):
    if pd.isna(s):
        return np.nan
    s = str(s).strip()
    return s if s else np.nan


def padronizar_dataframe(df, nome_fonte, idx, aba):
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df["Arquivo origem"] = nome_fonte
    df["Base arquivo"] = f"Base {idx}"
    df["Aba origem"] = aba

    for col in COLS_DATA:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in COLS_TEXTO:
        if col in df.columns:
            df[col] = df[col].apply(padronizar_texto)

    if "Tempo de digitalização" in df.columns:
        df["Data"] = df["Tempo de digitalização"].dt.date
        df["Hora"] = df["Tempo de digitalização"].dt.hour
        df["DiaHora"] = df["Tempo de digitalização"].dt.floor("h")
    else:
        df["Data"] = pd.NaT
        df["Hora"] = np.nan
        df["DiaHora"] = pd.NaT

    return df


def detectar_problematicos(df):
    if df.empty:
        return pd.DataFrame()

    mask = pd.Series(False, index=df.index)

    if "Tipo problemático" in df.columns:
        mask = mask | df["Tipo problemático"].notna()

    if "Descrição de Pacote Problemático" in df.columns:
        mask = mask | df["Descrição de Pacote Problemático"].notna()

    if "Tipo de bipagem" in df.columns:
        bip = df["Tipo de bipagem"].astype(str).str.lower()
        mask = mask | bip.str.contains("problem", na=False) | bip.str.contains("problemático", na=False)

    return df.loc[mask].copy()


@st.cache_data(show_spinner=False)
def carregar_dados_multiplos(arquivos):
    bases_car, bases_prob = [], []
    abas_identificadas = set()
    fontes = []
    modo_detectado = []

    for idx, arquivo in enumerate(arquivos, start=1):
        if arquivo is None:
            continue

        nome_fonte = getattr(arquivo, "name", f"Base {idx}")
        xls = pd.ExcelFile(arquivo)
        abas = xls.sheet_names
        abas_identificadas.update(abas)
        fontes.append(nome_fonte)

        abas_upper = {a.upper(): a for a in abas}
        tem_modelo_completo = any(a in abas_upper for a in ["CARREGAMENTO", "PROBLEMATICOS"])

        if tem_modelo_completo:
            modo_detectado.append(f"Base {idx}: modo multiaba")
            if "CARREGAMENTO" in abas_upper:
                df_car = pd.read_excel(xls, sheet_name=abas_upper["CARREGAMENTO"])
                df_car = padronizar_dataframe(df_car, nome_fonte, idx, abas_upper["CARREGAMENTO"])
                bases_car.append(df_car)
            if "PROBLEMATICOS" in abas_upper:
                df_prob = pd.read_excel(xls, sheet_name=abas_upper["PROBLEMATICOS"])
                df_prob = padronizar_dataframe(df_prob, nome_fonte, idx, abas_upper["PROBLEMATICOS"])
                bases_prob.append(df_prob)
        else:
            primeira_aba = abas[0]
            df_base = pd.read_excel(xls, sheet_name=primeira_aba)
            df_base = padronizar_dataframe(df_base, nome_fonte, idx, primeira_aba)
            bases_car.append(df_base)
            bases_prob.append(detectar_problematicos(df_base))
            modo_detectado.append(f"Base {idx}: modo aba única ({primeira_aba})")

    df_car = pd.concat([df for df in bases_car if not df.empty], ignore_index=True) if any(not df.empty for df in bases_car) else pd.DataFrame()
    df_prob = pd.concat([df for df in bases_prob if not df.empty], ignore_index=True) if any(not df.empty for df in bases_prob) else pd.DataFrame()

    return df_car, df_prob, sorted(abas_identificadas), fontes, modo_detectado


def aplicar_filtros(df_car, df_prob):
    st.sidebar.markdown("## Filtros")

    datas = sorted([d for d in df_car.get("Data", pd.Series(dtype=object)).dropna().unique().tolist()])
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
        if data_sel and "Data" in out.columns:
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


st.markdown(
    """
    <div class='hero'>
        <h1>📦 Painel Gerencial de Expedição</h1>
        <p>Acompanhamento de volume, produtividade, destinos e exceções em um layout executivo e pronto para uso diário.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

arquivo_default = Path("/mnt/data/EXPEDIÇÃO.xlsx")

st.sidebar.markdown("### Bases de entrada")
arquivo_1 = st.sidebar.file_uploader("Base 1 - planilha de bipe", type=["xlsx"], key="base_1")
arquivo_2 = st.sidebar.file_uploader("Base 2 - planilha de problematico", type=["xlsx"], key="base_2")

arquivos_para_carga = [arq for arq in [arquivo_1, arquivo_2] if arq is not None]
if not arquivos_para_carga and arquivo_default.exists():
    arquivos_para_carga = [arquivo_default]

if not arquivos_para_carga:
    st.info("Envie pelo menos uma planilha para iniciar a análise.")
    st.stop()

car_df, prob_df, abas, fontes, modos = carregar_dados_multiplos(arquivos_para_carga)
car_f, prob_f = aplicar_filtros(car_df, prob_df)

vol_total = len(car_f)
prob_total = len(prob_f)
prob_rate = round((prob_total / vol_total) * 100, 2) if vol_total else 0
bases_destino = car_f["Base Destino"].nunique() if "Base Destino" in car_f.columns else 0
digitalizadores = car_f["Digitalizador"].nunique() if "Digitalizador" in car_f.columns else 0
lotes = car_f["Número do lote"].nunique() if "Número do lote" in car_f.columns else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    card("Volumes", f"{vol_total:,}".replace(",", "."), "Registros analisados")
with c2:
    card("Problemáticos", f"{prob_total:,}".replace(",", "."), "Registros críticos detectados")
with c3:
    card("Taxa crítica", f"{prob_rate}%", "Problemáticos sobre o total")
with c4:
    card("Bases destino", f"{bases_destino:,}".replace(",", "."), "Cobertura operacional")
with c5:
    card("Digitalizadores", f"{digitalizadores:,}".replace(",", "."), f"Lotes únicos: {lotes:,}".replace(",", "."))

ins1, ins2, ins3 = st.columns(3)
with ins1:
    if not car_f.empty and "Digitalizador" in car_f.columns:
        top_dig = car_f["Digitalizador"].value_counts().head(1)
        if not top_dig.empty:
            st.info(f"**Maior produtividade:** {top_dig.index[0]} com **{int(top_dig.iloc[0]):,}** bipagens.".replace(",", "."))
with ins2:
    if not car_f.empty and "Base Destino" in car_f.columns:
        top_base = car_f["Base Destino"].value_counts().head(1)
        if not top_base.empty:
            st.info(f"**Base com maior volume:** {top_base.index[0]} com **{int(top_base.iloc[0]):,}** registros.".replace(",", "."))
with ins3:
    if not prob_f.empty and "Tipo problemático" in prob_f.columns:
        top_prob = prob_f["Tipo problemático"].value_counts().head(1)
        if not top_prob.empty:
            st.info(f"**Principal causa crítica:** {top_prob.index[0]} com **{int(top_prob.iloc[0]):,}** registros.".replace(",", "."))
        else:
            st.info("**Registros críticos detectados**, mas sem classificação de tipo problemático.")
    elif not prob_f.empty:
        st.info(f"**Registros críticos detectados:** {len(prob_f):,}.".replace(",", "."))


tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Visão Executiva",
    "Produtividade",
    "Destinos",
    "Exceções",
    "Pódio",
    "Detalhamento",
])

with tab1:
    g1, g2 = st.columns([1.4, 1])

    with g1:
        if not car_f.empty and "DiaHora" in car_f.columns:
            vol_hora = (
                car_f.dropna(subset=["DiaHora"]).groupby("DiaHora").size().reset_index(name="Volumes").sort_values("DiaHora")
            )
            if not vol_hora.empty:
                fig = px.area(vol_hora, x="DiaHora", y="Volumes", title="Evolução do volume ao longo do tempo")
                fig.update_layout(height=380)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.plotly_chart(grafico_vazio(), use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with g2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=prob_rate,
            number={"suffix": "%"},
            title={"text": "Taxa de exceções"},
            gauge={
                "axis": {"range": [None, max(10, round(prob_rate * 1.6, 2) + 1)]},
                "bar": {"thickness": 0.35},
                "steps": [
                    {"range": [0, 2], "color": "rgba(34,197,94,0.35)"},
                    {"range": [2, 5], "color": "rgba(245,158,11,0.30)"},
                    {"range": [5, 100], "color": "rgba(239,68,68,0.25)"},
                ],
            },
        ))
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)

    l1, l2 = st.columns(2)
    with l1:
        if not car_f.empty and "Base Destino" in car_f.columns:
            ranking_bases = car_f["Base Destino"].value_counts().reset_index()
            ranking_bases.columns = ["Base Destino", "Bipagens"]
            ranking_bases = ranking_bases.sort_values("Bipagens", ascending=True)
            altura_bases = max(420, len(ranking_bases) * 28)
            fig = px.bar(
                ranking_bases,
                x="Bipagens",
                y="Base Destino",
                orientation="h",
                title="Classificação geral de bases destino",
                text="Bipagens",
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(
                height=altura_bases,
                yaxis={"categoryorder": "total ascending"},
                xaxis_title="Número de bipagens",
                margin=dict(r=80),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with l2:
        if not car_f.empty and "Digitalizador" in car_f.columns:
            ranking_digs = car_f["Digitalizador"].value_counts().reset_index()
            ranking_digs.columns = ["Digitalizador", "Bipagens"]
            ranking_digs = ranking_digs.sort_values("Bipagens", ascending=True)
            altura_digs = max(420, len(ranking_digs) * 28)
            fig = px.bar(
                ranking_digs,
                x="Bipagens",
                y="Digitalizador",
                orientation="h",
                title="Classificação geral de digitalizadores",
                text="Bipagens",
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(
                height=altura_digs,
                yaxis={"categoryorder": "total ascending"},
                xaxis_title="Número de bipagens",
                margin=dict(r=80),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

with tab2:
    p1, p2 = st.columns([1.2, 1])
    with p1:
        if not car_f.empty and "Hora" in car_f.columns:
            prod_hora = car_f.dropna(subset=["Hora"]).groupby("Hora").size().reset_index(name="Volumes")
            if not prod_hora.empty:
                fig = px.bar(prod_hora, x="Hora", y="Volumes", title="Volume por hora da digitalização")
                fig.update_layout(height=380)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.plotly_chart(grafico_vazio(), use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    with p2:
        if not car_f.empty and "Digitalizador" in car_f.columns:
            por_dig = car_f["Digitalizador"].value_counts().reset_index()
            por_dig.columns = ["Digitalizador", "Volumes"]
            por_dig["% Part."] = (por_dig["Volumes"] / por_dig["Volumes"].sum() * 100).round(2)
            st.dataframe(por_dig.head(15), use_container_width=True, hide_index=True)
        else:
            st.dataframe(pd.DataFrame())

    if not car_f.empty and {"Digitalizador", "Hora"}.issubset(car_f.columns):
        heat = car_f.dropna(subset=["Digitalizador", "Hora"]).groupby(["Digitalizador", "Hora"]).size().reset_index(name="Volumes")
        if not heat.empty:
            heat_piv = heat.pivot(index="Digitalizador", columns="Hora", values="Volumes").fillna(0)
            fig = px.imshow(heat_piv, aspect="auto", title="Heatmap de produtividade por digitalizador e hora", text_auto=True)
            fig.update_layout(height=520)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)
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
            uf = car_f["Estado da cidade de destino"].value_counts().reset_index()
            uf.columns = ["UF", "Bipagens"]
            uf = uf.sort_values("Bipagens", ascending=True)
            altura_uf = max(450, len(uf) * 30)
            fig = px.bar(
                uf,
                x="Bipagens",
                y="UF",
                orientation="h",
                title="Classificação geral por UF destino",
                text="Bipagens",
            )
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_layout(
                height=altura_uf,
                yaxis={"categoryorder": "total ascending"},
                xaxis_title="Número de bipagens",
                margin=dict(r=80),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    if not car_f.empty and {"Base Destino", "Município de Destino"}.issubset(car_f.columns):
        cruz = car_f.groupby(["Base Destino", "Município de Destino"]).size().reset_index(name="Volumes").sort_values("Volumes", ascending=False).head(30)
        st.dataframe(cruz, use_container_width=True, hide_index=True)

with tab4:
    pr1, pr2 = st.columns([1.1, 1])
    with pr1:
        if not prob_f.empty and "Tipo problemático" in prob_f.columns and prob_f["Tipo problemático"].notna().any():
            tipo = prob_f["Tipo problemático"].value_counts().reset_index()
            tipo.columns = ["Tipo problemático", "Ocorrências"]
            fig = px.bar(tipo, x="Ocorrências", y="Tipo problemático", orientation="h", title="Distribuição das exceções")
            fig.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        elif not prob_f.empty and "Base Destino" in prob_f.columns:
            resumo = prob_f["Base Destino"].value_counts().head(15).reset_index()
            resumo.columns = ["Base Destino", "Ocorrências"]
            fig = px.bar(resumo, x="Ocorrências", y="Base Destino", orientation="h", title="Exceções por base destino")
            fig.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio("Sem registros críticos identificados"), use_container_width=True)

    with pr2:
        if not prob_f.empty and "Base Destino" in prob_f.columns:
            base_prob = prob_f["Base Destino"].value_counts().head(12).reset_index()
            base_prob.columns = ["Base Destino", "Ocorrências"]
            fig = px.bar(base_prob, x="Base Destino", y="Ocorrências", title="Bases com mais exceções")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.plotly_chart(grafico_vazio(), use_container_width=True)

    if not prob_f.empty:
        cols_exibir = [c for c in [
            "Número de pedido JMS", "Tipo de bipagem", "Tipo problemático", "Descrição de Pacote Problemático",
            "Base Destino", "Município de Destino", "Tempo de digitalização", "Base arquivo"
        ] if c in prob_f.columns]
        st.dataframe(prob_f[cols_exibir].head(300), use_container_width=True, hide_index=True)

with tab5:
    st.markdown("### Pódio de produtividade")

    if not car_f.empty and "Digitalizador" in car_f.columns:
        ranking = car_f["Digitalizador"].value_counts().reset_index()
        ranking.columns = ["Digitalizador", "Bipagens"]
        ranking["Participacao"] = (ranking["Bipagens"] / ranking["Bipagens"].sum() * 100).round(2)
        top3 = ranking.head(3).copy()

        st.markdown(
            """
            <style>
            .podio-card {
                border-radius: 20px;
                padding: 1.2rem 1rem;
                text-align: center;
                border: 1px solid rgba(255,255,255,0.08);
                background: linear-gradient(180deg, rgba(15,23,42,0.95), rgba(10,15,30,0.98));
                box-shadow: 0 10px 28px rgba(0,0,0,0.28);
                height: 100%;
            }
            .podio-medalha { font-size: 2rem; margin-bottom: 0.2rem; }
            .podio-posicao { color: #f8fafc; font-weight: 800; font-size: 2.1rem; line-height: 1; }
            .podio-nome { color: #f8fafc; font-weight: 700; font-size: 1rem; margin-top: 0.6rem; }
            .podio-valor { color: #7dd3fc; font-weight: 800; font-size: 1.8rem; margin-top: 0.55rem; }
            .podio-part { color: #cbd5e1; font-size: 0.9rem; margin-top: 0.35rem; }
            .podio-prata { margin-top: 90px; min-height: 290px; }
            .podio-ouro { min-height: 380px; border: 1px solid rgba(250,204,21,0.28); }
            .podio-bronze { margin-top: 130px; min-height: 250px; }
            </style>
            """,
            unsafe_allow_html=True,
        )

        def podio_item(coluna, item, posicao, medalha, classe_extra=""):
            with coluna:
                nome = item["Digitalizador"]
                bipagens = f"{int(item['Bipagens']):,}".replace(",", ".")
                particip = str(item["Participacao"]).replace(".", ",")
                st.markdown(
                    f"""
                    <div class="podio-card {classe_extra}">
                        <div class="podio-medalha">{medalha}</div>
                        <div class="podio-posicao">{posicao}º</div>
                        <div class="podio-nome">{nome}</div>
                        <div class="podio-valor">{bipagens} bipagens</div>
                        <div class="podio-part">{particip}% do total geral</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        c_esq, c_ctr, c_dir = st.columns([1, 1.15, 1])
        if len(top3) >= 2:
            podio_item(c_esq, top3.iloc[1], 2, "🥈", "podio-prata")
        else:
            with c_esq:
                st.plotly_chart(grafico_vazio("Sem 2º lugar disponível"), use_container_width=True)

        if len(top3) >= 1:
            podio_item(c_ctr, top3.iloc[0], 1, "🥇", "podio-ouro")
        else:
            with c_ctr:
                st.plotly_chart(grafico_vazio("Sem dados de produtividade"), use_container_width=True)

        if len(top3) >= 3:
            podio_item(c_dir, top3.iloc[2], 3, "🥉", "podio-bronze")
        else:
            with c_dir:
                st.plotly_chart(grafico_vazio("Sem 3º lugar disponível"), use_container_width=True)

        st.markdown("### Classificação geral de digitalizadores")
        ranking_plot = ranking.sort_values("Bipagens", ascending=True)
        altura_ranking = max(420, len(ranking_plot) * 28)
        fig = px.bar(
            ranking_plot,
            x="Bipagens",
            y="Digitalizador",
            orientation="h",
            text="Bipagens",
            title="Ranking geral por número de bipagens",
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(
            height=altura_ranking,
            yaxis={"categoryorder": "total ascending"},
            xaxis_title="Número de bipagens",
            margin=dict(r=80),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.plotly_chart(grafico_vazio("Sem dados de digitalizador para montar o pódio"), use_container_width=True)

with tab6:
    st.markdown("### Base analítica")
    modo = st.radio("Tabela", ["GERAL", "EXCEÇÕES"], horizontal=True)
    if modo == "GERAL":
        st.dataframe(car_f, use_container_width=True, hide_index=True, height=520)
        csv = car_f.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Baixar base geral filtrada", csv, "base_geral_filtrada.csv", "text/csv")
    else:
        st.dataframe(prob_f, use_container_width=True, hide_index=True, height=520)
        csv = prob_f.to_csv(index=False).encode("utf-8-sig")
        st.download_button("Baixar base de exceções filtrada", csv, "base_excecoes_filtrada.csv", "text/csv")

st.caption(f"Abas identificadas: {', '.join(abas)}")
st.caption(f"Modo de leitura: {' | '.join(modos)}")
