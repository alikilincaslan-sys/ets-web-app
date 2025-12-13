# ============================================================
# ETS GELİŞTİRME MODÜLÜ V001 – IEA STYLE INFOGRAPHIC DASHBOARD
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="ETS Development Module | Policy Dashboard",
    layout="wide"
)

st.title("ETS Development Module – Policy Dashboard")


# ============================================================
# STYLE (IEA-like clean UI) Görsel Kutucuk yazu stilleri
# ============================================================
st.markdown("""
<style>
.kpi {
    background: #ffffff;
    border-radius: 14px;
    padding: 14px;
    border: 1px solid rgba(0,0,0,0.10);
    box-shadow: 0 8px 22px rgba(0,0,0,0.10);
}
.kpi-label {
    font-size: 0.8rem;
    color: rgba(0,0,0,0.65) !important;
}
.kpi-value {
    font-size: 1.6rem;
    font-weight: 700;
    color: rgba(0,0,0,0.88) !important;
}
</style>
""", unsafe_allow_html=True)


def kpi(label, value):
    st.markdown(f"""
    <div class="kpi">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("Scenario Parameters")

price_min, price_max = st.sidebar.slider(
    "Carbon price range (€/tCO₂)",
    0, 200, (5, 20), step=1
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    0.0, 1.0, 1.0, step=0.05
)

benchmark_top_pct = st.sidebar.select_slider(
    "Benchmark (best plants %)",
    options=[10,20,30,40,50,60,70,80,90,100],
    value=100
)

price_method = st.sidebar.selectbox(
    "Carbon price method",
    ["Market Clearing", "Average Compliance Cost"]
)

fx_rate = st.sidebar.number_input(
    "FX rate (TL/€)",
    value=50.0,
    step=1.0
)


# ============================================================
# DATA UPLOAD
# ============================================================
uploaded = st.file_uploader(
    "Upload ETS input Excel file",
    type=["xlsx"]
)

if uploaded is None:
    st.info("Please upload an Excel file to start.")
    st.stop()


def read_all_sheets(file):
    xls = pd.ExcelFile(file)
    frames = []
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sh)
        df["FuelType"] = sh
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


df_raw = read_all_sheets(uploaded)
df = clean_ets_input(df_raw)


# ============================================================
# RUN MODEL
# ============================================================
if st.button("Run ETS Simulation"):

    sonuc_df, benchmark_map, clearing_price = ets_hesapla(
        df,
        price_min,
        price_max,
        agk,
        benchmark_top_pct=int(benchmark_top_pct),
        price_method=price_method
    )

    # ========================================================
    # KPI ROW (POLICY SNAPSHOT)
    # ========================================================
    total_gen = df["Generation_MWh"].sum()
    total_emis = df["Emissions_tCO2"].sum() / 1e6

    avg_cost_tl_mwh = (
        (sonuc_df["ets_net_cashflow_€/MWh"] * fx_rate * sonuc_df["Generation_MWh"]).sum()
        / sonuc_df["Generation_MWh"].sum()
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("Total generation (2024)", f"{total_gen:,.0f} MWh")
    with c2: kpi("Total emissions (2024)", f"{total_emis:,.2f} MtCO₂")
    with c3: kpi("Carbon price (2026–27)", f"{clearing_price:.2f} €/tCO₂")
    with c4: kpi("Avg ETS impact", f"{avg_cost_tl_mwh:,.1f} TL/MWh")


    st.markdown("---")

    # ========================================================
    # IEA STYLE INFOGRAPHIC BAR
    # ========================================================
    st.subheader("Net ETS impact on electricity generation costs")

    df_plot = sonuc_df.copy()
    df_plot["TL_per_MWh"] = df_plot["ets_net_cashflow_€/MWh"] * fx_rate

    # IEA-style focus: Top 30 most affected plants
    df_plot = df_plot.sort_values("TL_per_MWh").head(30)

    df_plot["Impact"] = np.where(
        df_plot["TL_per_MWh"] >= 0,
        "Cost increase",
        "Cost reduction"
    )

    fig = px.bar(
        df_plot,
        x="TL_per_MWh",
        y="Plant",
        orientation="h",
        color="Impact",
        color_discrete_map={
            "Cost increase": "#c0392b",
            "Cost reduction": "#2980b9"
        },
        labels={
            "TL_per_MWh": "Net ETS impact (TL/MWh)",
            "Plant": ""
        },
        title="Net ETS impact on electricity generation costs (TL/MWh)<br><sup>Top 30 plants, sorted</sup>"
    )

    fig.update_layout(
        template="simple_white",
        height=750,
        bargap=0.18,
        title_x=0.01,
        legend_orientation="h",
        legend_y=1.08,
        legend_x=0.01,
        xaxis=dict(
            zeroline=True,
            zerolinecolor="black",
            gridcolor="rgba(0,0,0,0.05)"
        ),
        yaxis=dict(tickfont=dict(size=11))
    )

    fig.update_traces(
        hovertemplate=
        "<b>%{y}</b><br>" +
        "Net ETS impact: %{x:.1f} TL/MWh<br>" +
        "<extra></extra>"
    )

    st.plotly_chart(fig, use_container_width=True)

    # ========================================================
    # FULL TABLE (OPTIONAL – COLLAPSIBLE)
    # ========================================================
    with st.expander("Show full plant-level results table"):
        st.dataframe(sonuc_df, use_container_width=True)
