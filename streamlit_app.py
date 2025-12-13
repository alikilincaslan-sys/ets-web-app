# ============================================================
# ETS GELİŞTİRME MODÜLÜ V001 – FINAL (Infographic + Word Note)
# (KPI grey + KG + Scope sliders + Auction Clearing)
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
    "price_method": "Market Clearing",
    "slope_bid": 150,
    "slope_ask": 150,
    "spread": 1.0,
    "do_clean": False,
    "lower_pct": 1.0,
    "upper_pct": 2.0,
    "fx_rate": 50.0,  # EURO KURU SABİT 50 TL
    "trf": 0.0,
}

st.set_page_config(
    page_title="ETS Geliştirme Modülü V001",
    layout="wide"
)

st.title("ETS Geliştirme Modülü v002")


# ============================================================
# INFOGRAPHIC CSS (Grey KPI + equal size)
# ============================================================
st.markdown("""
<style>
  .kpi {
    background: #f1f3f5;
    border: 1px solid rgba(0,0,0,0.12);
    border-radius: 16px;
    padding: 14px 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.08);
    min-height: 112px;
    display: flex;
    flex-direction: column;
    justify-content: center;
  }
  .kpi .label { font-size: 0.80rem; color: rgba(0,0,0,0.65); margin-bottom: 4px; }
  .kpi .value { font-size: 1.30rem; font-weight: 750; color: rgba(0,0,0,0.90); line-height: 1.15; word-break: break-word; }
  .kpi .sub { font-size: 0.75rem; color: rgba(0,0,0,0.60); margin-top: 4px; }
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
benchmark_method = st.sidebar.selectbox(
    "Benchmark belirleme yöntemi",
    [
        "Üretim ağırlıklı benchmark",
        "Kurulu güç ağırlıklı benchmark",
        "En iyi tesis dilimi (üretim payı)",
    ],
    index=0,
    key="benchmark_method",
    help="Benchmark (B_fuel) yakıt bazında hesaplanır. Seçilen yöntem, B_fuel'in nasıl belirleneceğini tanımlar.",
)
st.sidebar.caption("Not: Kurulu güç ağırlıklı yöntemde Excel'de InstalledCapacity_MW kolonu gerekir.")

# Yönteme bağlı parametre: En iyi tesis dilimi (%)
benchmark_top_pct = int(st.session_state.get("benchmark_top_pct", DEFAULTS.get("benchmark_top_pct", 100)))
if benchmark_method == "En iyi tesis dilimi (üretim payı)":
    benchmark_top_pct = st.sidebar.select_slider(
        "En iyi tesis dilimi (%)",
        options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        value=int(st.session_state.get("benchmark_top_pct", DEFAULTS.get("benchmark_top_pct", 100))),
        key="benchmark_top_pct",
        help="Yakıt grubu içinde intensity düşük olanlardan başlayarak toplam üretimin belirtilen yüzdesine kadar olan dilim seçilir; benchmark bu dilimin üretim-ağırlıklı ortalamasıdır.",
    )
else:
    st.session_state["benchmark_top_pct"] = 100
    benchmark_top_pct = 100


# ============================================================
# BENCHMARK SCOPE (BY FUEL) – OPTIONAL DROPPING
# ============================================================
st.sidebar.markdown("### Benchmark scope (by fuel)")

SCOPE_OPTIONS = [
    "Include all plants",
    "Exclude 5 plants with LOWEST EI",
    "Exclude 5 plants with HIGHEST EI",
]

scope_dg = st.sidebar.selectbox("DG Plants", SCOPE_OPTIONS, index=0, key="scope_dg")
scope_import = st.sidebar.selectbox("Imported Coal Plants", SCOPE_OPTIONS, index=0, key="scope_import")
scope_lignite = st.sidebar.selectbox("Lignite Plants", SCOPE_OPTIONS, index=0, key="scope_lignite")

def _fuel_mask(series: pd.Series, patterns):
    s = series.fillna("").astype(str).str.lower()
    m = False
    for p in patterns:
        m = m | s.str.contains(p, regex=False)
    return m

FUEL_GROUPS = {
    "DG": {"patterns": ["dg", "doğalgaz", "dogalgaz", "natural gas", "gas"], "scope": scope_dg},
    "IMPORT": {"patterns": ["ithal", "import", "imported"], "scope": scope_import},
    "LIGNITE": {"patterns": ["linyit", "lignite"], "scope": scope_lignite},
}

# -------------------------
# Price method
# -------------------------
price_method = st.sidebar.selectbox(
    "Fiyat Hesaplama Yöntemi",
    ["Market Clearing", "Average Compliance Cost", "Auction Clearing"],
    index=0
)

# Auction slider (only when Auction Clearing selected)
auction_supply_share = 1.0
if price_method == "Auction Clearing":
    auction_supply_share = st.sidebar.slider(
        "Auction supply (% of total demand)",
        min_value=10,
        max_value=200,
        value=100,
        step=10,
        help="Compliance demand sabit kabul edilir. Auction supply, toplam talebin yüzdesi olarak allowance arzını belirler. "
             "100% = demand kadar arz. 50% = kıtlık (fiyat yükselir). 150% = bolluk (fiyat tabana yaklaşır)."
    ) / 100.0

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

# UI'daki Türkçe seçimi, ets_hesapla'nın beklediği koda çevir
BENCHMARK_METHOD_MAP = {
    "Üretim ağırlıklı benchmark": "generation_weighted",
    "Kurulu güç ağırlıklı benchmark": "capacity_weighted",
    "En iyi tesis dilimi (üretim payı)": "best_plants",
}
benchmark_method_code = BENCHMARK_METHOD_MAP.get(benchmark_method, "best_plants")


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
# APPLY SCOPE DROPPING BEFORE RUN (plants may be removed entirely)
# ============================================================
df_run = df_all.copy()
dropped_plants = []

PLANT_COL = "Plant" if "Plant" in df_run.columns else ("PlantName" if "PlantName" in df_run.columns else None)

def _apply_scope(group_key: str, cfg: dict):
    global df_run, dropped_plants
    scope = cfg["scope"]
    pats = cfg["patterns"]

    if scope == "Include all plants":
        return

    mask = _fuel_mask(df_run["FuelType"], pats)
    sub = df_run.loc[mask].copy()
    if sub.empty:
        return

    sub["__EI__"] = pd.to_numeric(sub["Emissions_tCO2"], errors="coerce") / pd.to_numeric(sub["Generation_MWh"], errors="coerce")
    sub = sub.replace([np.inf, -np.inf], np.nan).dropna(subset=["__EI__"])
    if sub.empty:
        return

    n = 5
    if scope == "Exclude 5 plants with LOWEST EI":
        pick = sub.nsmallest(n, "__EI__")
    else:
        pick = sub.nlargest(n, "__EI__")

    if PLANT_COL:
        ids = pick[PLANT_COL].astype(str).tolist()
        dropped_plants.extend([(group_key, x) for x in ids])
        df_run.drop(index=pick.index, inplace=True)
    else:
        ids = pick.index.astype(int).tolist()
        dropped_plants.extend([(group_key, f"row#{i}") for i in ids])
        df_run.drop(index=pick.index, inplace=True)

for gk, cfg in FUEL_GROUPS.items():
    _apply_scope(gk, cfg)

if dropped_plants:
    with st.sidebar.expander("Dropped plants (scope)"):
        for gk, pid in dropped_plants:
            st.write(f"- {gk}: {pid}")


# ============================================================
# RUN MODEL
# ============================================================
if st.button("Run ETS Model"):

    sonuc_df, benchmark_map, clearing_price = ets_hesapla(
        df_run,
        price_min,
        price_max,
        agk,
        slope_bid=slope_bid,
        slope_ask=slope_ask,
        spread=spread,
        benchmark_method=benchmark_method_code,
        benchmark_top_pct=int(benchmark_top_pct),
        cap_col="InstalledCapacity_MW",
        price_method=price_method,
        trf=float(trf),
        auction_supply_share=float(auction_supply_share),
    )

    st.success(f"Carbon Price: {clearing_price:.2f} €/tCO₂")

    # ========================================================
    # INFOGRAPHIC KPI ROW
    # ========================================================
    total_gen = df_all["Generation_MWh"].sum()
    total_capacity = df_all["InstalledCapacity_MW"].sum() if "InstalledCapacity_MW" in df_all.columns else np.nan
    total_emis = df_all["Emissions_tCO2"].sum() / 1e6

    if "ets_net_cashflow_€/MWh" in sonuc_df.columns and "Generation_MWh" in sonuc_df.columns:
        avg_tl_mwh = (
            (sonuc_df["ets_net_cashflow_€/MWh"] * fx_rate * sonuc_df["Generation_MWh"]).sum()
            / sonuc_df["Generation_MWh"].sum()
        )
    else:
        avg_tl_mwh = np.nan

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: kpi_card("Electricity Generation", f"{total_gen:,.0f} MWh", "2024")
    with c2: kpi_card("Installed Capacity (KG)", f"{total_capacity:,.0f} MW", "Total")
    with c3: kpi_card("Total Emissions", f"{total_emis:,.2f} MtCO₂", "2024")
    with c4: kpi_card("Carbon Price", f"{clearing_price:.2f} €/tCO₂", price_method)
    with c5: kpi_card("FX Rate", f"{fx_rate:.0f} TL/€", "Scenario")
    with c6: kpi_card("Avg ETS Impact", f"{avg_tl_mwh:,.2f} TL/MWh", "Gen.-weighted")

    # ========================================================
    # CLEAN CHART
    # ========================================================
    st.subheader("Santral Bazlı Net ETS Etkisi (TL/MWh)")

    df_plot = sonuc_df.copy()
    df_plot["TL_per_MWh"] = df_plot["ets_net_cashflow_€/MWh"] * fx_rate
    df_plot = df_plot.sort_values("TL_per_MWh")

    df_plot["Impact_Type"] = np.where(
        df_plot["TL_per_MWh"] >= 0,
        "Cost increase",
        "Cost reduction",
    )

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
