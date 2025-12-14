# ============================================================
# ETS GELİŞTİRME MODÜLÜ V001 – FINAL (Infographic + Word Note)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go

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
    "price_range": (1, 15),
    "agk": 0.50,
    "benchmark_top_pct": 100,
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
    page_title="ETS Geliştirme Modülü V002",
    layout="wide"
)

st.title("ETS Geliştirme Modülü v002")


# ============================================================
# INFOGRAPHIC CSS
# ============================================================
st.markdown("""
<style>
  .kpi {
    background: #f2f2f2;
    border: 1px solid rgba(0,0,0,0.12);
    border-radius: 16px;
    padding: 14px 16px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.08);
    min-height: 115px;
    display: flex;
    flex-direction: column;
    justify-content: center;
  }
  .kpi .label {
    font-size: 0.80rem;
    color: rgba(0,0,0,0.65);
    margin-bottom: 4px;
  }
  .kpi .value {
    font-size: 1.30rem;
    font-weight: 750;
    color: rgba(0,0,0,0.90);
    line-height: 1.15;
    word-break: break-word;
  }
  .kpi .sub {
    font-size: 0.75rem;
    color: rgba(0,0,0,0.60);
    margin-top: 4px;
  }
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

# -------------------------
# Price method (ADD: Auction Clearing)
# -------------------------
price_method = st.sidebar.selectbox(
    "Fiyat Hesaplama Yöntemi",
    ["Market Clearing", "Average Compliance Cost", "Auction Clearing"],
    index=0
)

# ADD: Auction slider only when Auction Clearing selected
auction_supply_share = 1.0
if price_method == "Auction Clearing":
    auction_supply_share = st.sidebar.slider(
        "Auction supply (% of total demand)",
        min_value=10,
        max_value=200,
        value=100,
        step=10,
        help="Compliance demand (net ETS > 0) sabit kabul edilir. Auction supply toplam talebin yüzdesi kadar arzı ifade eder. "
             "100% = demand kadar arz; <100% kıtlık; >100% bolluk."
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

# FIX: UI'daki Türkçe seçimi, ets_model.ets_hesapla'nın beklediği koda çevir
BENCHMARK_METHOD_MAP = {
    "Üretim ağırlıklı benchmark": "generation_weighted",
    "Kurulu güç ağırlıklı benchmark": "capacity_weighted",
    "En iyi tesis dilimi (üretim payı)": "best_plants",
}
benchmark_method_code = BENCHMARK_METHOD_MAP.get(benchmark_method, "best_plants")

# ============================================================
# BENCHMARK SCOPE (OPTIONAL FILTERING BY FUEL GROUP)
# ============================================================
st.sidebar.subheader("Benchmark scope (by fuel)")

SCOPE_OPTIONS = [
    "Include all plants",
    "Exclude 5 plants with LOWEST EI",
    "Exclude 5 plants with HIGHEST EI",
]

scope_dg = st.sidebar.selectbox("DG Plants", SCOPE_OPTIONS, index=0, key="scope_dg")
scope_import = st.sidebar.selectbox("Imported Coal Plants", SCOPE_OPTIONS, index=0, key="scope_import")
scope_lignite = st.sidebar.selectbox("Lignite Plants", SCOPE_OPTIONS, index=0, key="scope_lignite")


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
# APPLY SCOPE FILTER (DROPS PLANTS FROM CALCULATION IF CHOSEN)
# ============================================================
def _fuel_group_of(ft: str) -> str:
    s = str(ft).strip().lower()
    if any(k in s for k in ["dg", "doğalgaz", "dogalgaz", "natural gas", "gas", "ng"]):
        return "DG"
    if any(k in s for k in ["ithal", "import", "imported"]):
        return "IMPORT_COAL"
    if any(k in s for k in ["linyit", "lignite"]):
        return "LIGNITE"
    return "OTHER"

def _apply_scope(df: pd.DataFrame, group_code: str, option: str, n: int = 5):
    if option == "Include all plants":
        return df, []
    dfg = df[df["FuelType"].apply(_fuel_group_of) == group_code].copy()
    if dfg.empty:
        return df, []

    agg = (
        dfg.groupby("Plant", as_index=False)[["Emissions_tCO2", "Generation_MWh"]]
        .sum()
    )
    agg["EI"] = agg["Emissions_tCO2"] / agg["Generation_MWh"].replace(0, np.nan)
    agg = agg.dropna(subset=["EI"])

    if agg.empty:
        return df, []

    asc = option == "Exclude 5 plants with LOWEST EI"
    picks = agg.sort_values("EI", ascending=asc).head(n)["Plant"].tolist()

    mask_group = df["FuelType"].apply(_fuel_group_of) == group_code
    df2 = df[~(mask_group & df["Plant"].isin(picks))].copy()
    return df2, picks

df_scoped = df_all.copy()
dropped = {"DG": [], "IMPORT_COAL": [], "LIGNITE": []}

df_scoped, dropped["DG"] = _apply_scope(df_scoped, "DG", st.session_state.get("scope_dg", "Include all plants"))
df_scoped, dropped["IMPORT_COAL"] = _apply_scope(df_scoped, "IMPORT_COAL", st.session_state.get("scope_import", "Include all plants"))
df_scoped, dropped["LIGNITE"] = _apply_scope(df_scoped, "LIGNITE", st.session_state.get("scope_lignite", "Include all plants"))

df_all = df_scoped

if any(len(v) > 0 for v in dropped.values()):
    st.sidebar.caption("Dropped plants (by scope):")
    if dropped["DG"]:
        st.sidebar.write("DG:", ", ".join(dropped["DG"]))
    if dropped["IMPORT_COAL"]:
        st.sidebar.write("Imported coal:", ", ".join(dropped["IMPORT_COAL"]))
    if dropped["LIGNITE"]:
        st.sidebar.write("Lignite:", ", ".join(dropped["LIGNITE"]))


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
    total_capacity = df_all["InstalledCapacity_MW"].sum()
    total_emis = df_all["Emissions_tCO2"].sum() / 1e6

    if "ets_net_cashflow_€/MWh" in sonuc_df.columns:
        avg_tl_mwh = (
            (sonuc_df["ets_net_cashflow_€/MWh"] * fx_rate * sonuc_df["Generation_MWh"]).sum()
            / sonuc_df["Generation_MWh"].sum()
        )
    else:
        avg_tl_mwh = np.nan

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:
        kpi_card("Electricity Generation", f"{total_gen:,.0f} MWh", "2024")
    with c2:
        kpi_card("Installed Capacity (KG)", f"{total_capacity:,.0f} MW", "Toplam")
    with c3:
        kpi_card("Total Emissions", f"{total_emis:,.2f} MtCO₂", "2024")
    with c4:
        kpi_card("Carbon Price", f"{clearing_price:.2f} €/tCO₂", price_method)
    with c5:
        kpi_card("FX Rate", f"{fx_rate:.0f} TL/€", "Scenario")
    with c6:
        kpi_card("Avg ETS Impact", f"{avg_tl_mwh:,.2f} TL/MWh", "Gen.-weighted")

    # ========================================================
    # IEA VISUALS – Market Summary / Price Formation / Bid-Ask / Benchmark Distribution
    # ========================================================
    st.subheader("IEA-style market visuals")

    # ---------- 1) Market summary cards ----------
    demand_tco2 = float(sonuc_df.loc[sonuc_df["net_ets"] > 0, "net_ets"].sum())
    if price_method == "Auction Clearing":
        supply_tco2 = demand_tco2 * float(auction_supply_share)
        supply_label = "Auction supply"
    else:
        supply_tco2 = float((-sonuc_df.loc[sonuc_df["net_ets"] < 0, "net_ets"]).sum())
        supply_label = "Available supply (surplus)"

    traded_tco2 = float(min(demand_tco2, supply_tco2)) if (demand_tco2 > 0 and supply_tco2 > 0) else 0.0
    shortage_tco2 = float(max(demand_tco2 - supply_tco2, 0.0))
    shortage_pct = (shortage_tco2 / demand_tco2 * 100.0) if demand_tco2 > 0 else 0.0

    s1, s2, s3, s4 = st.columns(4)
    with s1:
        kpi_card("Compliance demand", f"{demand_tco2:,.0f} tCO₂", "Σ(net_ets > 0)")
    with s2:
        kpi_card(supply_label, f"{supply_tco2:,.0f} tCO₂", "Policy / surplus based")
    with s3:
        kpi_card("Traded volume", f"{traded_tco2:,.0f} tCO₂", "min(demand, supply)")
    with s4:
        kpi_card("Shortage", f"{shortage_tco2:,.0f} tCO₂", f"{shortage_pct:.1f}% of demand")

    st.caption("Note: In Auction Clearing, demand is assumed inelastic and supply is set as a share of total compliance demand.")


    # ---------- NEW: Price formation chart (method-specific, X = allowances) ----------
    def _demand_at_price(buyers_df: pd.DataFrame, p: float, pmin: float) -> float:
        if buyers_df.empty:
            return 0.0
        q0 = buyers_df["net_ets"].to_numpy(dtype=float)
        p_bid = buyers_df["p_bid"].to_numpy(dtype=float)
        denom = np.maximum(p_bid - pmin, 1e-9)
        frac = 1.0 - (p - pmin) / denom
        return float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))

    def _supply_at_price(sellers_df: pd.DataFrame, p: float, pmax: float) -> float:
        if sellers_df.empty:
            return 0.0
        q0 = (-sellers_df["net_ets"]).to_numpy(dtype=float)  # supply quantity
        p_ask = sellers_df["p_ask"].to_numpy(dtype=float)
        denom = np.maximum(pmax - p_ask, 1e-9)
        frac = (p - p_ask) / denom
        return float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))

    with st.expander("Price formation (how the carbon price is formed)", expanded=True):
        buyers = sonuc_df.loc[sonuc_df["net_ets"] > 0, ["net_ets", "p_bid"]].copy()
        sellers = sonuc_df.loc[sonuc_df["net_ets"] < 0, ["net_ets", "p_ask"]].copy()

        fig_pf = go.Figure()

        if price_method == "Market Clearing":
            # Step-like inverse curves (context)
            if not buyers.empty:
                b = buyers.sort_values("p_bid", ascending=False).copy()
                b["cum_q"] = b["net_ets"].cumsum()
                fig_pf.add_trace(go.Scatter(
                    x=b["cum_q"], y=b["p_bid"],
                    mode="lines", name="Demand curve (bids)"
                ))

            if not sellers.empty:
                s = sellers.sort_values("p_ask", ascending=True).copy()
                s["supply_q"] = (-s["net_ets"]).astype(float)
                s["cum_q"] = s["supply_q"].cumsum()
                fig_pf.add_trace(go.Scatter(
                    x=s["cum_q"], y=s["p_ask"],
                    mode="lines", name="Supply curve (asks)"
                ))

            # Intersection markers using continuous activation logic
            qd = _demand_at_price(buyers, float(clearing_price), float(price_min))
            qs = _supply_at_price(sellers, float(clearing_price), float(price_max))
            q_star = float(min(qd, qs))

            fig_pf.add_hline(
                y=float(clearing_price),
                line_dash="dash",
                annotation_text=f"Clearing price: {clearing_price:.2f} €/tCO₂",
                annotation_position="top left",
            )
            fig_pf.add_vline(
                x=q_star,
                line_dash="dash",
                annotation_text=f"Traded volume: {q_star:,.0f} tCO₂",
                annotation_position="top right",
            )

            subtitle = "Market Clearing: price is where available supply meets compliance demand (intersection)."

        elif price_method == "Auction Clearing":
            if buyers.empty:
                st.info("No buyers (net_ets > 0) in the current scope — cannot draw auction price formation.")
                subtitle = "Auction Clearing: insufficient buyers in current scope."
            else:
                demand = float(buyers["net_ets"].sum())
                supply = float(demand * float(auction_supply_share))

                b = buyers.sort_values("p_bid", ascending=False).copy()
                b["cum_q"] = b["net_ets"].cumsum()

                fig_pf.add_trace(go.Scatter(
                    x=b["cum_q"], y=b["p_bid"],
                    mode="lines", name="Bid stack (descending bids)"
                ))

                fig_pf.add_vline(
                    x=supply,
                    line_dash="dash",
                    annotation_text=f"Auction supply: {supply:,.0f} tCO₂",
                    annotation_position="top right",
                )
                fig_pf.add_hline(
                    y=float(clearing_price),
                    line_dash="dash",
                    annotation_text=f"Clearing price: {clearing_price:.2f} €/tCO₂",
                    annotation_position="top left",
                )

                subtitle = "Auction Clearing: fixed supply is sold to highest bids; marginal winning bid sets the price."

        else:  # Average Compliance Cost
            if buyers.empty:
                st.info("No buyers (net_ets > 0) in the current scope — cannot draw ACC price formation.")
                subtitle = "Average Compliance Cost: insufficient buyers in current scope."
            else:
                b = buyers.sort_values("p_bid", ascending=False).copy()
                b["cum_q"] = b["net_ets"].cumsum()

                fig_pf.add_trace(go.Scatter(
                    x=b["cum_q"], y=b["p_bid"],
                    mode="lines", name="Bid stack (buyers)"
                ))
                fig_pf.add_hline(
                    y=float(clearing_price),
                    line_dash="dash",
                    annotation_text=f"ACC price: {clearing_price:.2f} €/tCO₂",
                    annotation_position="top left",
                )

                subtitle = "Average Compliance Cost: price is the demand-weighted average of buyers’ bids (no strict intersection)."

        fig_pf.update_layout(
            template="simple_white",
            height=460,
            margin=dict(l=10, r=10, t=40, b=10),
            legend_orientation="h",
            legend_y=1.08,
            legend_x=0.01,
            xaxis_title="Allowances (tCO₂)",
            yaxis_title="Price (€/tCO₂)",
        )
        fig_pf.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")
        fig_pf.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")

        st.plotly_chart(fig_pf, use_container_width=True)
        st.caption(subtitle)


    # ---------- 2) Bid–Ask curves + clearing price ----------
    with st.expander("Bid–Ask curves and clearing price", expanded=True):
        buyers = sonuc_df.loc[sonuc_df["net_ets"] > 0, ["net_ets", "p_bid"]].copy()
        sellers = sonuc_df.loc[sonuc_df["net_ets"] < 0, ["net_ets", "p_ask"]].copy()

        if buyers.empty:
            st.info("No buyers (net_ets > 0) in the current scope — bid curve not available.")
        else:
            buyers = buyers.sort_values("p_bid", ascending=True)
            buyers["cum_q"] = buyers["net_ets"].cumsum()

            fig_curves = px.line(
                buyers,
                x="cum_q",
                y="p_bid",
                template="simple_white",
                labels={"cum_q": "Allowances (tCO₂)", "p_bid": "Price (€/tCO₂)"},
            )
            fig_curves.update_traces(name="Demand (bids)", line=dict(color="#1f77b4"))

            if not sellers.empty:
                sellers = sellers.sort_values("p_ask", ascending=True)
                sellers["supply_q"] = (-sellers["net_ets"]).astype(float)
                sellers["cum_q"] = sellers["supply_q"].cumsum()

                fig_s = px.line(
                    sellers,
                    x="cum_q",
                    y="p_ask",
                    template="simple_white",
                    labels={"cum_q": "Allowances (tCO₂)", "p_ask": "Price (€/tCO₂)"},
                )
                fig_s.update_traces(name="Supply (asks)", line=dict(color="#d62728"))
                for tr in fig_s.data:
                    fig_curves.add_trace(tr)

            fig_curves.add_hline(
                y=float(clearing_price),
                line_dash="dash",
                line_color="black",
                annotation_text=f"Clearing: {clearing_price:.2f} €/tCO₂",
                annotation_position="top left",
            )

            fig_curves.update_layout(
                height=420,
                legend_orientation="h",
                legend_y=1.08,
                legend_x=0.01,
                margin=dict(l=10, r=10, t=40, b=10),
            )
            fig_curves.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")
            fig_curves.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")

            st.plotly_chart(fig_curves, use_container_width=True)

    # ---------- 3) Fuel-wise intensity distribution + benchmark ----------
    with st.expander("Emission intensity distribution vs benchmark (by fuel)", expanded=True):
        if "intensity" not in sonuc_df.columns:
            st.warning("Model output does not contain 'intensity' column — cannot draw distribution.")
        else:
            fig_box = px.box(
                sonuc_df,
                x="FuelType",
                y="intensity",
                points="all",
                template="simple_white",
                labels={"FuelType": "", "intensity": "tCO₂ / MWh"},
            )
            fig_box.update_layout(
                height=520,
                margin=dict(l=10, r=10, t=40, b=10),
                showlegend=False,
            )
            fig_box.update_xaxes(showgrid=False)
            fig_box.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")

            for fuel, b in (benchmark_map or {}).items():
                try:
                    b_val = float(b)
                except Exception:
                    continue
                if np.isfinite(b_val):
                    fig_box.add_hline(
                        y=b_val,
                        line_dash="dash",
                        line_color="black",
                        opacity=0.55,
                        annotation_text=f"{fuel} benchmark",
                        annotation_position="top left",
                    )

            st.plotly_chart(fig_box, use_container_width=True)

            st.caption(
                "How to read: Each box shows the distribution of plant-level emission intensities (tCO₂/MWh) within a fuel group. "
                "The line inside the box is the median; the box spans the 25th–75th percentiles; dots are individual plants "
                "(outliers may appear as isolated points). "
                "The dashed horizontal line is the fuel-specific benchmark (B_fuel). "
                "Plants above the benchmark are net buyers (higher ETS cost), while plants below the benchmark are net sellers "
                "(potential surplus)."
            )

    st.divider()

    # ========================================================
    # INFOGRAPHIC – SINGLE CLEAN CHART
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
