# ============================================================
# ETS GELİŞTİRME MODÜLÜ – SCENARIO COMPARISON (Reference vs Scenario 2)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

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
    "fx_rate": 50.0,
    "trf": 0.0,
}

st.set_page_config(page_title="ETS Geliştirme Modülü – Scenario Compare", layout="wide")
st.title("ETS Geliştirme Modülü – Scenario Comparison (Reference vs Scenario 2)")


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


def add_cost_columns(df_in: pd.DataFrame, fx_rate: float, clearing_price: float) -> pd.DataFrame:
    df = df_in.copy()

    # capacity merge safety (if missing)
    if "InstalledCapacity_MW" not in df.columns:
        df["InstalledCapacity_MW"] = np.nan

    # TL/MWh
    if "ets_net_cashflow_€/MWh" in df.columns:
        df["ets_net_cashflow_TL/MWh"] = pd.to_numeric(df["ets_net_cashflow_€/MWh"], errors="coerce") * float(fx_rate)

    # Total €
    if "net_ets" in df.columns:
        df["ets_net_cashflow_€"] = pd.to_numeric(df["net_ets"], errors="coerce") * float(clearing_price)
        df["ets_net_cashflow_TL"] = df["ets_net_cashflow_€"] * float(fx_rate)

    # fallback TL/MWh from totals
    if "ets_net_cashflow_TL/MWh" not in df.columns and "ets_net_cashflow_TL" in df.columns and "Generation_MWh" in df.columns:
        gen = pd.to_numeric(df["Generation_MWh"], errors="coerce").replace(0, np.nan)
        df["ets_net_cashflow_TL/MWh"] = df["ets_net_cashflow_TL"] / gen

    # reorder: Plant then InstalledCapacity_MW
    cols = list(df.columns)
    if "Plant" in cols and "InstalledCapacity_MW" in cols:
        cols.remove("InstalledCapacity_MW")
        plant_idx = cols.index("Plant")
        cols.insert(plant_idx + 1, "InstalledCapacity_MW")
        df = df[cols]

    return df


def to_excel_bytes(ref_df: pd.DataFrame,
                   sc2_df: pd.DataFrame,
                   comp_df: pd.DataFrame,
                   bm_map_ref: dict,
                   bm_map_sc2: dict,
                   params_ref: dict,
                   params_sc2: dict) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        ref_df.to_excel(writer, index=False, sheet_name="Results_Reference")
        sc2_df.to_excel(writer, index=False, sheet_name="Results_Scenario2")
        comp_df.to_excel(writer, index=False, sheet_name="Comparison")

        # Benchmarks
        if isinstance(bm_map_ref, dict) and len(bm_map_ref) > 0:
            pd.DataFrame([{"FuelType": k, "Benchmark_Reference": v} for k, v in bm_map_ref.items()]) \
              .to_excel(writer, index=False, sheet_name="Benchmarks_Ref")
        if isinstance(bm_map_sc2, dict) and len(bm_map_sc2) > 0:
            pd.DataFrame([{"FuelType": k, "Benchmark_Scenario2": v} for k, v in bm_map_sc2.items()]) \
              .to_excel(writer, index=False, sheet_name="Benchmarks_Sc2")

        # Params
        pd.DataFrame([params_ref]).to_excel(writer, index=False, sheet_name="Params_Ref")
        pd.DataFrame([params_sc2]).to_excel(writer, index=False, sheet_name="Params_Sc2")

    return out.getvalue()


# ============================================================
# SIDEBAR – GLOBAL SETTINGS (common for both scenarios)
# ============================================================
st.sidebar.header("Global Settings (apply to both scenarios)")

# Benchmark method (common)
benchmark_method_ui = st.sidebar.selectbox(
    "Benchmark belirleme yöntemi (common)",
    [
        "Üretim ağırlıklı benchmark",
        "Kurulu güç ağırlıklı benchmark",
        "En iyi tesis dilimi (üretim payı)",
    ],
    index=0,
    key="benchmark_method_common",
    help="Benchmark (B_fuel) yakıt bazında hesaplanır. Seçilen yöntem, B_fuel'in nasıl belirleneceğini tanımlar.",
)
st.sidebar.caption("Not: Kurulu güç ağırlıklı yöntemde Excel'de InstalledCapacity_MW kolonu gerekir.")

BENCHMARK_METHOD_MAP = {
    "Üretim ağırlıklı benchmark": "generation_weighted",
    "Kurulu güç ağırlıklı benchmark": "capacity_weighted",
    "En iyi tesis dilimi (üretim payı)": "best_plants",
}
benchmark_method_code = BENCHMARK_METHOD_MAP.get(benchmark_method_ui, "best_plants")

benchmark_top_pct = 100
if benchmark_method_ui == "En iyi tesis dilimi (üretim payı)":
    benchmark_top_pct = st.sidebar.select_slider(
        "En iyi tesis dilimi (%) (common)",
        options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
        value=int(st.session_state.get("benchmark_top_pct_common", DEFAULTS["benchmark_top_pct"])),
        key="benchmark_top_pct_common",
        help="Yakıt grubu içinde intensity düşük olanlardan başlayarak toplam üretimin belirtilen yüzdesine kadar olan dilim seçilir; benchmark bu dilimin üretim-ağırlıklı ortalamasıdır.",
    )
else:
    st.session_state["benchmark_top_pct_common"] = 100
    benchmark_top_pct = 100


# Scope controls (common)
st.sidebar.subheader("Benchmark scope (common, by fuel)")
SCOPE_OPTIONS = [
    "Include all plants",
    "Exclude 5 plants with LOWEST EI",
    "Exclude 5 plants with HIGHEST EI",
]
scope_dg = st.sidebar.selectbox("DG Plants", SCOPE_OPTIONS, index=0, key="scope_dg")
scope_import = st.sidebar.selectbox("Imported Coal Plants", SCOPE_OPTIONS, index=0, key="scope_import")
scope_lignite = st.sidebar.selectbox("Lignite Plants", SCOPE_OPTIONS, index=0, key="scope_lignite")


# ============================================================
# SIDEBAR – SCENARIO PARAMETERS (Reference vs Scenario 2)
# ============================================================
st.sidebar.header("Scenario Parameters")

tabs = st.sidebar.tabs(["Reference", "Scenario 2"])

def scenario_controls(prefix: str, default_overrides: dict = None):
    d = DEFAULTS.copy()
    if default_overrides:
        d.update(default_overrides)

    with st.sidebar:
        pass

    # controls inside the active tab (caller context)
    price_min, price_max = st.slider(
        f"Karbon Fiyat Aralığı (€/tCO₂) [{prefix}]",
        0, 200,
        st.session_state.get(f"{prefix}_price_range", d["price_range"]),
        step=1,
        key=f"{prefix}_price_range"
    )

    agk = st.slider(
        f"Adil Geçiş Katsayısı (AGK) [{prefix}]",
        0.0, 1.0,
        float(st.session_state.get(f"{prefix}_agk", d["agk"])),
        step=0.05,
        key=f"{prefix}_agk"
    )

    price_method = st.selectbox(
        f"Fiyat Hesaplama Yöntemi [{prefix}]",
        ["Market Clearing", "Average Compliance Cost", "Auction Clearing"],
        index=0 if d["price_method"] == "Market Clearing" else (1 if d["price_method"] == "Average Compliance Cost" else 2),
        key=f"{prefix}_price_method"
    )

    auction_supply_share = 1.0
    if price_method == "Auction Clearing":
        auction_supply_share = st.slider(
            f"Auction supply (% of total demand) [{prefix}]",
            min_value=10,
            max_value=200,
            value=int(st.session_state.get(f"{prefix}_auction_supply_pct", 100)),
            step=10,
            key=f"{prefix}_auction_supply_pct",
            help="Demand inelastic varsayımıyla, arzı toplam talebin yüzdesi olarak ayarlar."
        ) / 100.0

    slope_bid = st.slider(f"Talep Eğimi (β_bid) [{prefix}]", 10, 500, int(d["slope_bid"]), step=10, key=f"{prefix}_slope_bid")
    slope_ask = st.slider(f"Arz Eğimi (β_ask) [{prefix}]", 10, 500, int(d["slope_ask"]), step=10, key=f"{prefix}_slope_ask")
    spread = st.slider(f"Bid/Ask Spread [{prefix}]", 0.0, 10.0, float(d["spread"]), step=0.5, key=f"{prefix}_spread")

    fx_rate = st.number_input(
        f"Euro Kuru (TL/€) [{prefix}]",
        min_value=0.0,
        value=float(st.session_state.get(f"{prefix}_fx_rate", d["fx_rate"])),
        step=1.0,
        key=f"{prefix}_fx_rate"
    )

    trf = st.slider(
        f"Geçiş Dönemi Telafi Katsayısı (TRF) [{prefix}]",
        min_value=0.0,
        max_value=1.0,
        value=float(st.session_state.get(f"{prefix}_trf", d["trf"])),
        step=0.05,
        key=f"{prefix}_trf"
    )

    return {
        "price_min": price_min,
        "price_max": price_max,
        "agk": agk,
        "price_method": price_method,
        "auction_supply_share": float(auction_supply_share),
        "slope_bid": slope_bid,
        "slope_ask": slope_ask,
        "spread": spread,
        "fx_rate": float(fx_rate),
        "trf": float(trf),
    }

with tabs[0]:
    ref_params = scenario_controls("REF")

with tabs[1]:
    sc2_params = scenario_controls("SC2")


# ============================================================
# FILE UPLOAD
# ============================================================
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("Lütfen Excel dosyası yükleyin.")
    st.stop()

df_all_raw = read_all_sheets(uploaded)
df_all = clean_ets_input(df_all_raw)


# ============================================================
# APPLY COMMON SCOPE FILTER
# ============================================================
df_scoped = df_all.copy()
dropped = {"DG": [], "IMPORT_COAL": [], "LIGNITE": []}
df_scoped, dropped["DG"] = _apply_scope(df_scoped, "DG", scope_dg)
df_scoped, dropped["IMPORT_COAL"] = _apply_scope(df_scoped, "IMPORT_COAL", scope_import)
df_scoped, dropped["LIGNITE"] = _apply_scope(df_scoped, "LIGNITE", scope_lignite)
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
run = st.button("Run BOTH Scenarios (Reference + Scenario 2)")

if run:
    # --- Reference ---
    ref_out, ref_bm_map, ref_price = ets_hesapla(
        df_all,
        ref_params["price_min"],
        ref_params["price_max"],
        ref_params["agk"],
        slope_bid=ref_params["slope_bid"],
        slope_ask=ref_params["slope_ask"],
        spread=ref_params["spread"],
        benchmark_method=benchmark_method_code,
        benchmark_top_pct=int(benchmark_top_pct),
        cap_col="InstalledCapacity_MW",
        price_method=ref_params["price_method"],
        trf=float(ref_params["trf"]),
        auction_supply_share=float(ref_params["auction_supply_share"]),
    )

    # --- Scenario 2 ---
    sc2_out, sc2_bm_map, sc2_price = ets_hesapla(
        df_all,
        sc2_params["price_min"],
        sc2_params["price_max"],
        sc2_params["agk"],
        slope_bid=sc2_params["slope_bid"],
        slope_ask=sc2_params["slope_ask"],
        spread=sc2_params["spread"],
        benchmark_method=benchmark_method_code,
        benchmark_top_pct=int(benchmark_top_pct),
        cap_col="InstalledCapacity_MW",
        price_method=sc2_params["price_method"],
        trf=float(sc2_params["trf"]),
        auction_supply_share=float(sc2_params["auction_supply_share"]),
    )

    # Add TL and ordering
    ref_df = add_cost_columns(ref_out, fx_rate=ref_params["fx_rate"], clearing_price=float(ref_price))
    sc2_df = add_cost_columns(sc2_out, fx_rate=sc2_params["fx_rate"], clearing_price=float(sc2_price))

    # Basic headline
    st.subheader("Scenario headline results")
    h1, h2, h3 = st.columns(3)
    with h1:
        kpi_card("Reference carbon price", f"{ref_price:.2f} €/tCO₂", ref_params["price_method"])
    with h2:
        kpi_card("Scenario 2 carbon price", f"{sc2_price:.2f} €/tCO₂", sc2_params["price_method"])
    with h3:
        kpi_card("Δ Price (Sc2 - Ref)", f"{(sc2_price - ref_price):+.2f} €/tCO₂", "difference")

    # Compare plant-level TL/MWh
    st.subheader("Plant-level comparison (TL/MWh)")
    # choose best available columns
    ref_col = "ets_net_cashflow_TL/MWh" if "ets_net_cashflow_TL/MWh" in ref_df.columns else None
    sc2_col = "ets_net_cashflow_TL/MWh" if "ets_net_cashflow_TL/MWh" in sc2_df.columns else None

    comp = ref_df[["Plant", "InstalledCapacity_MW"]].copy()
    comp = comp.merge(ref_df[["Plant"] + ([ref_col] if ref_col else [])], on="Plant", how="left")
    comp = comp.merge(sc2_df[["Plant"] + ([sc2_col] if sc2_col else [])], on="Plant", how="left", suffixes=("_Ref", "_Sc2"))

    # rename safely
    if ref_col:
        comp = comp.rename(columns={ref_col: "TL_per_MWh_Ref"})
    else:
        comp["TL_per_MWh_Ref"] = np.nan

    if sc2_col:
        comp = comp.rename(columns={sc2_col: "TL_per_MWh_Sc2"})
    else:
        comp["TL_per_MWh_Sc2"] = np.nan

    comp["Δ_TL_per_MWh"] = comp["TL_per_MWh_Sc2"] - comp["TL_per_MWh_Ref"]

    # sort by absolute delta
    view = comp.copy()
    view["absΔ"] = view["Δ_TL_per_MWh"].abs()
    view = view.sort_values("absΔ", ascending=False).head(30).sort_values("Δ_TL_per_MWh")

    fig_delta = px.bar(
        view,
        x="Δ_TL_per_MWh",
        y="Plant",
        orientation="h",
        template="simple_white",
        labels={"Δ_TL_per_MWh": "Δ Net ETS impact (TL/MWh) [Scenario2 - Reference]", "Plant": ""},
        title="Top 30 plants by |Δ TL/MWh|"
    )
    fig_delta.update_layout(height=750, margin=dict(l=10, r=10, t=60, b=10))
    fig_delta.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.06)", zeroline=True, zerolinecolor="black")
    fig_delta.update_yaxes(showgrid=False)
    st.plotly_chart(fig_delta, use_container_width=True)

    # Show full tables side by side
    st.subheader("Raw results (same columns) – Reference vs Scenario 2")
    cL, cR = st.columns(2)
    with cL:
        st.markdown("### Reference – Results")
        st.dataframe(ref_df, use_container_width=True, height=520)
    with cR:
        st.markdown("### Scenario 2 – Results")
        st.dataframe(sc2_df, use_container_width=True, height=520)

    st.subheader("Comparison table (Δ)")
    st.dataframe(comp.drop(columns=["absΔ"], errors="ignore"), use_container_width=True)

    # Excel download
    params_ref = {
        **ref_params,
        "benchmark_method": benchmark_method_code,
        "benchmark_top_pct": int(benchmark_top_pct),
        "scope_dg": scope_dg,
        "scope_import": scope_import,
        "scope_lignite": scope_lignite,
    }
    params_sc2 = {
        **sc2_params,
        "benchmark_method": benchmark_method_code,
        "benchmark_top_pct": int(benchmark_top_pct),
        "scope_dg": scope_dg,
        "scope_import": scope_import,
        "scope_lignite": scope_lignite,
    }

    excel_bytes = to_excel_bytes(
        ref_df=ref_df,
        sc2_df=sc2_df,
        comp_df=comp.drop(columns=["absΔ"], errors="ignore"),
        bm_map_ref=ref_bm_map,
        bm_map_sc2=sc2_bm_map,
        params_ref=params_ref,
        params_sc2=params_sc2,
    )

    st.download_button(
        "Download Scenario Results as Excel (.xlsx)",
        data=excel_bytes,
        file_name="ets_scenario_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
