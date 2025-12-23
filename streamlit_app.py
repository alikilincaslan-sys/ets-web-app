# ============================================================
# ETS GELİŞTİRME MODÜLÜ V002 – Scenario Runner + IEA Visuals
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from io import BytesIO

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input


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
    "fx_rate": 50.0,  # EURO KURU
    "trf": 0.0,
}

st.set_page_config(page_title="ETS Geliştirme Modülü V002", layout="wide")
st.title("ETS Geliştirme Modülü v002")


# ============================================================
# INFOGRAPHIC CSS
# ============================================================
st.markdown(
    """
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
  .kpi .label { font-size: 0.80rem; color: rgba(0,0,0,0.65); margin-bottom: 4px; }
  .kpi .value { font-size: 1.30rem; font-weight: 750; color: rgba(0,0,0,0.90); line-height: 1.15; word-break: break-word; }
  .kpi .sub { font-size: 0.75rem; color: rgba(0,0,0,0.60); margin-top: 4px; }
</style>
""",
    unsafe_allow_html=True,
)


def kpi_card(label, value, sub=""):
    st.markdown(
        f"""
    <div class="kpi">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>
    """,
        unsafe_allow_html=True,
    )


# ============================================================
# HELPERS
# ============================================================
def read_all_sheets(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sh)
        df["FuelType"] = sh
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def _fuel_group_of(ft: str) -> str:
    s = str(ft).strip().lower()
    if any(k in s for k in ["dg", "doğalgaz", "dogalgaz", "natural gas", "gas", "ng"]):
        return "DG"
    if any(k in s for k in ["ithal", "import", "imported"]):
        return "IMPORT_COAL"
    if any(k in s for k in ["linyit", "lignite"]):
        return "LIGNITE"
    return "OTHER"


def _fuel_label(ft: str, default=None) -> str:
    g = _fuel_group_of(ft)
    if g == "DG":
        return "Natural Gas"
    if g == "IMPORT_COAL":
        return "Import Coal"
    if g == "LIGNITE":
        return "Lignite"
    return default if default is not None else str(ft)


def _apply_scope(df: pd.DataFrame, group_code: str, option: str, n: int = 5):
    if option == "Include all plants":
        return df, []
    dfg = df[df["FuelType"].apply(_fuel_group_of) == group_code].copy()
    if dfg.empty:
        return df, []

    agg = dfg.groupby("Plant", as_index=False)[["Emissions_tCO2", "Generation_MWh"]].sum()
    agg["EI"] = agg["Emissions_tCO2"] / agg["Generation_MWh"].replace(0, np.nan)
    agg = agg.dropna(subset=["EI"])
    if agg.empty:
        return df, []

    asc = option == "Exclude 5 plants with LOWEST EI"
    picks = agg.sort_values("EI", ascending=asc).head(n)["Plant"].tolist()

    mask_group = df["FuelType"].apply(_fuel_group_of) == group_code
    df2 = df[~(mask_group & df["Plant"].isin(picks))].copy()
    return df2, picks


def _bench_group_map(benchmark_map: dict) -> dict:
    """Return fuel-group -> list of (label, benchmark). Supports two-tier."""
    out = {"Natural Gas": [], "Import Coal": [], "Lignite": []}
    if not benchmark_map:
        return out

    for k, v in benchmark_map.items():
        try:
            b = float(v)
        except Exception:
            continue
        if not np.isfinite(b):
            continue

        key = str(k)
        grp = _fuel_label(key, None)
        if grp not in out:
            continue

        lbl = "benchmark"
        if "Best tier" in key:
            lbl = "Best-tier benchmark"
        elif "Worst tier" in key:
            lbl = "Worst-tier benchmark"

        out[grp].append((lbl, b))

    def _order(t):
        lbl, _ = t
        if "Best-tier" in lbl:
            return 0
        if "Worst-tier" in lbl:
            return 1
        return 2

    for g in out:
        out[g] = sorted(out[g], key=_order)
    return out


# ============================================================
# SIDEBAR – COMMON PARAMETERS
# ============================================================
st.sidebar.header("Model Parametreleri (Common)")

price_min, price_max = st.sidebar.slider(
    "Karbon Fiyat Aralığı (€/tCO₂)",
    0,
    200,
    st.session_state.get("price_range", DEFAULTS["price_range"]),
    step=1,
    key="price_range",
)

agk_common = st.sidebar.slider(
    "Adil Geçiş Katsayısı (AGK) (common)",
    0.0,
    1.0,
    float(st.session_state.get("agk_common", DEFAULTS["agk"])),
    step=0.05,
    key="agk_common",
)

benchmark_method_ui = st.sidebar.selectbox(
    "Benchmark belirleme yöntemi (common)",
    [
        "Üretim ağırlıklı benchmark",
        "Kurulu güç ağırlıklı benchmark",
        "En iyi tesis dilimi (üretim payı)",
        "Two-tier benchmark (Best vs Worst, by plant count)",
    ],
    index=0,
    key="benchmark_method_common",
)
st.sidebar.caption("Not: Kurulu güç ağırlıklı yöntemde Excel'de InstalledCapacity_MW kolonu gerekir.")

benchmark_top_pct_common = int(st.session_state.get("benchmark_top_pct_common", DEFAULTS.get("benchmark_top_pct", 100)))
if benchmark_method_ui == "En iyi tesis dilimi (üretim payı)":
    benchmark_top_pct_common = st.sidebar.select_slider(
        "En iyi tesis dilimi (%) (common)",
        options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        value=int(st.session_state.get("benchmark_top_pct_common", DEFAULTS.get("benchmark_top_pct", 100))),
        key="benchmark_top_pct_common",
    )
else:
    st.session_state["benchmark_top_pct_common"] = 100
    benchmark_top_pct_common = 100

# Two-tier benchmark split (common)
tier_best_pct_common = 50
if benchmark_method_ui == "Two-tier benchmark (Best vs Worst, by plant count)":
    tier_best_pct_common = st.sidebar.select_slider(
        "Best tier share (% of plants) (common)",
        options=[10, 20, 30, 40, 50, 60, 70, 80, 90],
        value=int(st.session_state.get("tier_best_pct_common", 50)),
        key="tier_best_pct_common",
        help="Within each fuel group, plants are ranked by EI and split into Best/Worst tiers by plant count. "
             "Each tier gets its own benchmark (generation-weighted EI).",
    )
else:
    st.session_state["tier_best_pct_common"] = 50
    tier_best_pct_common = 50


# ============================================================
# SIDEBAR – SCENARIO PARAMETERS (TOP)
# ============================================================
st.sidebar.divider()
st.sidebar.subheader("Scenario 1 (Reference)")
agk_ref = st.sidebar.slider("AGK (Ref)", 0.0, 1.0, float(agk_common), 0.05, key="agk_ref")
price_method_ref = st.sidebar.selectbox(
    "Fiyat Hesaplama (Ref)",
    ["Market Clearing", "Average Compliance Cost", "Auction Clearing"],
    index=0,
    key="price_method_ref",
)
auction_supply_share_ref = 1.0
if price_method_ref == "Auction Clearing":
    auction_supply_share_ref = st.sidebar.slider(
        "Auction supply (% of demand) (Ref)",
        min_value=10,
        max_value=200,
        value=100,
        step=10,
        key="auction_supply_ref",
    ) / 100.0

st.sidebar.subheader("Scenario 2")
agk_sc2 = st.sidebar.slider("AGK (Sc2)", 0.0, 1.0, float(agk_common), 0.05, key="agk_sc2")
price_method_sc2 = st.sidebar.selectbox(
    "Fiyat Hesaplama (Sc2)",
    ["Market Clearing", "Average Compliance Cost", "Auction Clearing"],
    index=0,
    key="price_method_sc2",
)
auction_supply_share_sc2 = 1.0
if price_method_sc2 == "Auction Clearing":
    auction_supply_share_sc2 = st.sidebar.slider(
        "Auction supply (% of demand) (Sc2)",
        min_value=10,
        max_value=200,
        value=100,
        step=10,
        key="auction_supply_sc2",
    ) / 100.0

st.sidebar.divider()
slope_bid = st.sidebar.slider("Talep Eğimi (β_bid)", 10, 500, DEFAULTS["slope_bid"], step=10, key="slope_bid")
slope_ask = st.sidebar.slider("Arz Eğimi (β_ask)", 10, 500, DEFAULTS["slope_ask"], step=10, key="slope_ask")
spread = st.sidebar.slider("Bid/Ask Spread", 0.0, 10.0, DEFAULTS["spread"], step=0.5, key="spread")

fx_rate = st.sidebar.number_input(
    "Euro Kuru (TL/€)",
    min_value=0.0,
    value=float(DEFAULTS["fx_rate"]),
    step=1.0,
    key="fx_rate",
)

trf = st.sidebar.slider(
    "Geçiş Dönemi Telafi Katsayısı (TRF)",
    min_value=0.0,
    max_value=1.0,
    value=float(DEFAULTS.get("trf", 0.0)),
    step=0.05,
    key="trf",
)

BENCHMARK_METHOD_MAP = {
    "Üretim ağırlıklı benchmark": "generation_weighted",
    "Kurulu güç ağırlıklı benchmark": "capacity_weighted",
    "En iyi tesis dilimi (üretim payı)": "best_plants",
    "Two-tier benchmark (Best vs Worst, by plant count)": "two_tier",
}
benchmark_method_code = BENCHMARK_METHOD_MAP.get(benchmark_method_ui, "best_plants")


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

df_all_raw = read_all_sheets(uploaded)
df_all = clean_ets_input(df_all_raw)

# Apply scope filters (drops plants from calculation if chosen)
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
# RUN BOTH SCENARIOS
# ============================================================
st.divider()
run = st.button("Run BOTH Scenarios (Reference + Scenario 2)")

if run:
    # Scenario 1
    sonuc_ref, bench_ref, price_ref = ets_hesapla(
        df_all,
        price_min,
        price_max,
        agk_ref,
        slope_bid=slope_bid,
        slope_ask=slope_ask,
        spread=spread,
        benchmark_method=benchmark_method_code,
        benchmark_top_pct=int(benchmark_top_pct_common),
        tier_best_pct=int(tier_best_pct_common),
        cap_col="InstalledCapacity_MW",
        price_method=price_method_ref,
        trf=float(trf),
        auction_supply_share=float(auction_supply_share_ref),
    )

    # Scenario 2
    sonuc_sc2, bench_sc2, price_sc2 = ets_hesapla(
        df_all,
        price_min,
        price_max,
        agk_sc2,
        slope_bid=slope_bid,
        slope_ask=slope_ask,
        spread=spread,
        benchmark_method=benchmark_method_code,
        benchmark_top_pct=int(benchmark_top_pct_common),
        tier_best_pct=int(tier_best_pct_common),
        cap_col="InstalledCapacity_MW",
        price_method=price_method_sc2,
        trf=float(trf),
        auction_supply_share=float(auction_supply_share_sc2),
    )

    st.subheader("Scenario headline results")
    c1, c2, c3 = st.columns(3)
    with c1:
        kpi_card("Reference carbon price", f"{price_ref:,.2f} €/tCO₂", price_method_ref)
    with c2:
        kpi_card("Scenario 2 carbon price", f"{price_sc2:,.2f} €/tCO₂", price_method_sc2)
    with c3:
        kpi_card("Δ Price (Sc2 - Ref)", f"{(price_sc2-price_ref):,.2f} €/tCO₂", "difference")

    # ========================================================
    # COMPARISON CHART (TL/MWh) with fuel filter
    # ========================================================
    st.subheader("Plant-level comparison (TL/MWh)")

    fuel_choice = st.selectbox(
        "Fuel filter",
        options=["All", "Natural Gas", "Import Coal", "Lignite"],
        index=0,
        key="fuel_filter_comp",
    )

    def _apply_fuel_filter(df, fuel_label):
        if fuel_label == "All":
            return df
        return df[df["FuelType"].apply(lambda z: _fuel_label(z, "Other")) == fuel_label].copy()

    ref = _apply_fuel_filter(sonuc_ref.copy(), fuel_choice)
    sc2 = _apply_fuel_filter(sonuc_sc2.copy(), fuel_choice)

    ref["TL_per_MWh_Ref"] = ref["ets_net_cashflow_€/MWh"] * float(fx_rate)
    sc2["TL_per_MWh_Sc2"] = sc2["ets_net_cashflow_€/MWh"] * float(fx_rate)

    comp = ref.merge(sc2[["Plant", "TL_per_MWh_Sc2"]], on="Plant", how="inner")
    comp["Δ_TL_per_MWh"] = comp["TL_per_MWh_Sc2"] - comp["TL_per_MWh_Ref"]

    comp = comp.sort_values("Δ_TL_per_MWh")

    top_n = 30
    if len(comp) > top_n:
        comp = comp.reindex(comp["Δ_TL_per_MWh"].abs().sort_values(ascending=False).head(top_n).index)

    # side-by-side bars
    comp_melt = comp.melt(
        id_vars=["Plant"],
        value_vars=["TL_per_MWh_Ref", "TL_per_MWh_Sc2"],
        var_name="Scenario",
        value_name="TL_per_MWh",
    )
    comp_melt["Scenario"] = comp_melt["Scenario"].replace(
        {"TL_per_MWh_Ref": "Reference", "TL_per_MWh_Sc2": "Scenario 2"}
    )

    fig_comp = px.bar(
        comp_melt.sort_values(["Plant", "Scenario"]),
        x="TL_per_MWh",
        y="Plant",
        orientation="h",
        color="Scenario",
        barmode="group",
        labels={"TL_per_MWh": "Net ETS impact (TL/MWh)", "Plant": "", "Scenario": ""},
        template="simple_white",
    )
    fig_comp.update_layout(height=750, margin=dict(l=10, r=10, t=40, b=10))
    fig_comp.update_xaxes(gridcolor="rgba(0,0,0,0.06)")
    fig_comp.update_yaxes(showgrid=False)
    st.plotly_chart(fig_comp, use_container_width=True)

    # ========================================================
    # Effective EI ranked chart (AGK included) + scenario toggles
    # ========================================================
    st.subheader("Effective emission intensity (EI_eff) ranked (AGK included)")

    show_ref = st.checkbox("Show Reference", value=True)
    show_sc2 = st.checkbox("Show Scenario 2", value=True)

    # build ranked series by plant, using EI_eff mean (already includes AGK + benchmark)
    def _rank_df(df_out: pd.DataFrame, label: str):
        d = df_out.copy()
        d["FuelGroup"] = d["FuelType"].apply(lambda z: _fuel_label(z, "Other"))
        if fuel_choice != "All":
            d = d[d["FuelGroup"] == fuel_choice].copy()

        # sort by EI_eff
        d = d.sort_values("EI_eff", ascending=True).reset_index(drop=True)
        d["Rank"] = np.arange(1, len(d) + 1)
        d["Scenario"] = label
        return d

    lines = []
    if show_ref:
        lines.append(_rank_df(sonuc_ref, "Reference"))
    if show_sc2:
        lines.append(_rank_df(sonuc_sc2, "Scenario 2"))

    if lines:
        plot_df = pd.concat(lines, ignore_index=True)

        fig_ei = px.line(
            plot_df,
            x="Rank",
            y="EI_eff",
            color="Scenario",
            markers=True,
            template="simple_white",
            labels={"Rank": "Plants ranked by effective intensity (low → high)", "EI_eff": "Effective emission intensity (tCO₂/MWh)"},
        )
        fig_ei.update_layout(height=520, margin=dict(l=10, r=10, t=40, b=10))
        fig_ei.update_xaxes(showgrid=False)
        fig_ei.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)")

        # benchmark lines (fuel-specific). Use reference benchmark map for guidance.
        bm_draw = _bench_group_map(bench_ref)

        BM_COLOR = {
            "Natural Gas": "blue",
            "Import Coal": "gold",
            "Imported Coal": "gold",
            "Lignite": "green",
        }

        groups_to_draw = ["Natural Gas", "Import Coal", "Lignite"] if fuel_choice == "All" else [fuel_choice]
        for grp in groups_to_draw:
            items = bm_draw.get(grp, [])
            if not items:
                continue
            for idx, (lbl, bval) in enumerate(items):
                if not np.isfinite(bval):
                    continue
                dash = "dash" if "Best-tier" in lbl else ("dot" if "Worst-tier" in lbl else "dash")
                fig_ei.add_hline(
                    y=float(bval),
                    line_dash=dash,
                    line_width=2,
                    line_color=BM_COLOR.get(grp, "gray"),
                    opacity=0.55,
                    annotation_text=f"{grp} {lbl}",
                    annotation_position="top left",
                )

        st.plotly_chart(fig_ei, use_container_width=True)

        st.caption(
            "EI_eff = B_fuel + (EI − B_fuel) × (1 − AGK). "
            "AGK increases reduce dispersion across plants around the benchmark, reflecting the fairness/transition mechanism."
        )
    else:
        st.info("Select at least one scenario to display (Reference and/or Scenario 2).")

    st.divider()

    # ========================================================
    # RAW TABLES + EXCEL DOWNLOAD (same columns as shown)
    # ========================================================
    st.subheader("Tüm Sonuçlar (Ham Tablo)")

    # add TL columns + keep capacity next to plant if exists
    def _enrich(df_out: pd.DataFrame, scenario_name: str) -> pd.DataFrame:
        d = df_out.copy()
        d["Scenario"] = scenario_name
        d["ETS_TL_total"] = d["ets_net_cashflow_€"] * float(fx_rate)
        d["ETS_TL_per_MWh"] = d["ets_net_cashflow_€/MWh"] * float(fx_rate)
        return d

    ref_x = _enrich(sonuc_ref, "Reference")
    sc2_x = _enrich(sonuc_sc2, "Scenario 2")

    out_all = pd.concat([ref_x, sc2_x], ignore_index=True)

    # reorder columns: Plant, InstalledCapacity_MW then rest
    front = ["Plant"]
    if "InstalledCapacity_MW" in out_all.columns:
        front += ["InstalledCapacity_MW"]
    front += ["Scenario"]
    rest = [c for c in out_all.columns if c not in front]
    out_all = out_all[front + rest]

    st.dataframe(out_all, use_container_width=True)

    # Excel download (same columns as above)
    def _to_excel_bytes(df_res: pd.DataFrame, bm_ref: dict, bm_sc2: dict):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_res.to_excel(writer, index=False, sheet_name="Results_All")
            pd.DataFrame({"BenchmarkKey": list(bm_ref.keys()), "Value": list(bm_ref.values())}).to_excel(
                writer, index=False, sheet_name="Benchmark_Ref"
            )
            pd.DataFrame({"BenchmarkKey": list(bm_sc2.keys()), "Value": list(bm_sc2.values())}).to_excel(
                writer, index=False, sheet_name="Benchmark_Sc2"
            )
        buf.seek(0)
        return buf.getvalue()

    excel_bytes = _to_excel_bytes(out_all, bench_ref, bench_sc2)

    st.download_button(
        "Download results as Excel (.xlsx)",
        data=excel_bytes,
        file_name="ets_results_scenarios.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
