# ============================================================
# ETS GELİŞTİRME MODÜLÜ V001 – FINAL (Infographic + Word Note)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px

from datetime import datetime
from io import BytesIO

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel


# ============================================================
# DEFAULTS
# ============================================================
DEFAULTS = {
    "price_range": (5, 20),
    "agk": 1.00,
    "benchmark_top_pct": 100,
    "benchmark_method": "En iyi tesis dilimi (üretim payı)",
    "price_method": "Market Clearing",
    "slope_bid": 150,
    "slope_ask": 150,
    "spread": 1.0,
    "do_clean": False,
    "lower_pct": 1.0,
    "upper_pct": 2.0,
    "fx_rate": 50.0,  # <<< EURO KURU SABİT 50 TL
    "trf": 0.0,
}

st.set_page_config(
    page_title="ETS Geliştirme Modülü V001",
    layout="wide"
)

st.title("ETS Geliştirme Modülü V001")


# ============================================================
# INFOGRAPHIC CSS
# ============================================================

st.markdown("""
<style>
  .kpi {
    background: rgba(255,255,255,0.95);
    border: 1px solid rgba(0,0,0,0.10);
    border-radius: 18px;
    padding: 14px 16px;
    box-shadow: 0 10px 28px rgba(0,0,0,0.10);
  }
  .kpi .label { font-size: 0.85rem; color: rgba(0,0,0,0.65); }
  .kpi .value { font-size: 1.55rem; font-weight: 750; color: rgba(0,0,0,0.90); line-height: 1.1; }
  .kpi .sub { font-size: 0.8rem; color: rgba(0,0,0,0.60); }
</style>
""", unsafe_allow_html=True)


def kpi_card(label, value, sub=""):
    st.markdown(f"""
    <div class="kpi">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# SIDEBAR – PARAMETERS
# ============================================================
st.sidebar.header("Model Parametreleri")

price_min, price_max = st.sidebar.slider(
    "Karbon Fiyat Aralığı (€/tCO₂)",
    0, 200,
    st.session_state.get("price_range", DEFAULTS["price_range"]),
    step=1,
    key="price_range"
)

agk = st.sidebar.slider(
    "Adil Geçiş Katsayısı (AGK)",
    0.0, 1.0,
    float(st.session_state.get("agk", DEFAULTS["agk"])),
    step=0.05,
    key="agk"
)

# -------------------------
# Benchmark belirleme yöntemi
# -------------------------
benchmark_method_tr = st.sidebar.selectbox(
    "Benchmark belirleme yöntemi",
    options=[
        "Üretim ağırlıklı benchmark",
        "Kurulu güç ağırlıklı benchmark",
        "En iyi tesis dilimi (üretim payı)",
    ],
    index=[
        "Üretim ağırlıklı benchmark",
        "Kurulu güç ağırlıklı benchmark",
        "En iyi tesis dilimi (üretim payı)",
    ].index(st.session_state.get("benchmark_method", DEFAULTS.get("benchmark_method", "En iyi tesis dilimi (üretim payı)"))),
    key="benchmark_method",
    help="Benchmark, yakıt türü bazında hesaplanır. Seçilen yöntem sadece benchmark değerini (B_fuel) değiştirir; AGK formülü aynı kalır."
)

# Yöntem kodu (model tarafına giden)
BENCH_METHOD_MAP = {
    "Üretim ağırlıklı benchmark": "generation_weighted",
    "Kurulu güç ağırlıklı benchmark": "capacity_weighted",
    "En iyi tesis dilimi (üretim payı)": "best_plants",
}
benchmark_method = BENCH_METHOD_MAP.get(benchmark_method_tr, "best_plants")

# En iyi tesis dilimi seçildiyse yüzde parametresi aktif
if benchmark_method == "best_plants":
    benchmark_top_pct = st.sidebar.select_slider(
        "En iyi üretim dilimi (%)",
        options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
        key="benchmark_top_pct",
        help="Yakıt bazında intensity düşük olan santrallerden başlayarak, toplam üretimin seçilen yüzdesi dolana kadar seçim yapılır. Benchmark, seçilen dilimin üretim-ağırlıklı ortalamasıdır."
    )
else:
    # Üretim ağırlıklı / kurulu güç ağırlıklı seçeneklerinde yüzde parametresi kullanılmaz
    benchmark_top_pct = 100
    st.sidebar.caption("Not: Bu yöntemde 'en iyi tesis yüzdesi' kullanılmaz (otomatik 100%).")

if benchmark_method == "capacity_weighted":
    st.sidebar.info("Kurulu güç ağırlıklı benchmark için Excel'de 'InstalledCapacity_MW' kolonu bulunmalıdır (her santral satırında).")

price_method = st.sidebar.selectbox(
    "Fiyat Hesaplama Yöntemi",
    ["Market Clearing", "Average Compliance Cost"],
    index=0
)

slope_bid = st.sidebar.slider("Talep Eğimi (β_bid)", 10, 500, DEFAULTS["slope_bid"], step=10)
slope_ask = st.sidebar.slider("Arz Eğimi (β_ask)", 10, 500, DEFAULTS["slope_ask"], step=10)
spread = st.sidebar.slider("Bid/Ask Spread", 0.0, 10.0, DEFAULTS["spread"], step=0.5)

fx_rate = st.sidebar.number_input(
    "Euro Kuru (TL/€)",
    min_value=0.0,
    value=float(DEFAULTS["fx_rate"]),
    step=1.0
)

trf = st.sidebar.slider(
    "Geçiş Dönemi Telafi Katsayısı (TRF)",
    min_value=0.0,
    max_value=1.0,
    value=float(DEFAULTS.get("trf", 0.0)),
    step=0.05,
    help="Pilot dönemde, benchmark nedeniyle oluşan ilave yükün ne kadarının ücretsiz telafi edileceğini gösterir. "
         "TRF=0 → telafi yok; TRF=1 → (I−B) farkının tamamı telafi edilir (sadece I>B olan tesisler için)."
)


# ============================================================
# FILE UPLOAD
# ============================================================
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("Lütfen Excel dosyası yükleyin.")
    st.stop()


def read_all_sheets(file):
    xls = pd.ExcelFile(file)
    frames = []
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sh)
        df["FuelType"] = sh
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


df_all_raw = read_all_sheets(uploaded)
df_all = clean_ets_input(df_all_raw)


# ============================================================
# RUN MODEL
# ============================================================
if st.button("Run ETS Model"):

    sonuc_df, benchmark_map, clearing_price = ets_hesapla(
        df_all,
        price_min,
        price_max,
        agk,
        slope_bid=slope_bid,
        slope_ask=slope_ask,
        spread=spread,
        benchmark_top_pct=int(benchmark_top_pct),
        benchmark_method=benchmark_method,
        price_method=price_method,
        trf=float(trf)
    )

    st.success(f"Carbon Price: {clearing_price:.2f} €/tCO₂")

    # ========================================================
    # INFOGRAPHIC KPI ROW
    # ========================================================
    total_gen = df_all["Generation_MWh"].sum()
    total_emis = df_all["Emissions_tCO2"].sum() / 1e6

    if "ets_net_cashflow_€/MWh" in sonuc_df.columns:
        avg_tl_mwh = (
            (sonuc_df["ets_net_cashflow_€/MWh"] * fx_rate * sonuc_df["Generation_MWh"]).sum()
            / sonuc_df["Generation_MWh"].sum()
        )
    else:
        avg_tl_mwh = np.nan

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi_card("Total Generation", f"{total_gen:,.0f} MWh", "2024")
    with c2: kpi_card("Total Emissions", f"{total_emis:,.2f} MtCO₂", "2024")
    with c3: kpi_card("Carbon Price", f"{clearing_price:.2f} €/tCO₂", price_method)
    with c4: kpi_card("FX Rate", f"{fx_rate:.0f} TL/€", "Scenario")
    with c5: kpi_card("Avg ETS Impact", f"{avg_tl_mwh:,.2f} TL/MWh", "Gen.-weighted")

    # ========================================================
    # INFOGRAPHIC – SINGLE CLEAN CHART
    # ========================================================
    st.subheader("Santral Bazlı Net ETS Etkisi (TL/MWh)")

    df_plot = sonuc_df.copy()
    df_plot["TL_per_MWh"] = df_plot["ets_net_cashflow_€/MWh"] * fx_rate
    df_plot = df_plot.sort_values("TL_per_MWh")

    
    # IEA-style interactive infographic (Plotly)
    df_plot["Impact_Type"] = np.where(
        df_plot["TL_per_MWh"] >= 0,
        "Cost increase",
        "Cost reduction",
    )

    # Focus view: show top N plants by absolute impact for readability
    top_n = 30
    df_view = df_plot.copy()
    if len(df_view) > top_n:
        df_view = df_view.reindex(
            df_view["TL_per_MWh"].abs().sort_values(ascending=False).head(top_n).index
        ).sort_values("TL_per_MWh")

    fig = px.bar(
        df_view,
        x="TL_per_MWh",
        y="Plant",
        orientation="h",
        color="Impact_Type",
        color_discrete_map={
            "Cost increase": "#c0392b",
            "Cost reduction": "#2980b9",
        },
        labels={
            "TL_per_MWh": "Net ETS impact (TL/MWh)",
            "Plant": "",
            "Impact_Type": "",
        },
        title="Net ETS impact on electricity generation costs (TL/MWh)<br><sup>Most affected plants, sorted</sup>",
    )

    fig.update_layout(
        template="simple_white",
        height=750,
        bargap=0.18,
        title_x=0.01,
        font=dict(family="Arial, sans-serif", size=13),
        legend_orientation="h",
        legend_y=1.08,
        legend_x=0.01,
        xaxis=dict(
            zeroline=True,
            zerolinecolor="black",
            gridcolor="rgba(0,0,0,0.06)",
            title="Net ETS impact (TL/MWh)",
        ),
        yaxis=dict(tickfont=dict(size=11), title=""),
    )

    fig.update_traces(
        hovertemplate="<b>%{y}</b><br>Net ETS impact: %{x:.1f} TL/MWh<extra></extra>"
    )

    st.plotly_chart(fig, use_container_width=True)
    # ========================================================
    # RAW TABLE
    # ========================================================
    st.subheader("Tüm Sonuçlar (Ham Tablo)")
    st.dataframe(sonuc_df, use_container_width=True)

    # ========================================================
    # CSV DOWNLOAD
    # ========================================================
    st.download_button(
        "Download results as CSV",
        data=sonuc_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="ets_results.csv",
        mime="text/csv"
    )
