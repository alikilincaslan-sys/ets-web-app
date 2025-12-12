import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

import altair as alt
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

from ets_model import ets_hesapla

# Temizleme modülü opsiyonel (repo'da data_cleaning.py varsa kullanır)
try:
    from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel
    HAS_CLEANING = True
except Exception:
    HAS_CLEANING = False


# -------------------------
# Default values (V001 Stable)
# -------------------------
DEFAULTS = {
    "price_range": (5, 20),
    "agk": 1.00,
    "benchmark_top_pct": 100,
    "slope_bid": 150,
    "slope_ask": 150,
    "spread": 1.0,
    "do_clean": False,
    "lower_pct": 1.0,
    "upper_pct": 2.0,
    "price_method": "Market Clearing",
}


def reset_to_defaults() -> None:
    st.session_state["price_range"] = DEFAULTS["price_range"]
    st.session_state["agk"] = DEFAULTS["agk"]
    st.session_state["benchmark_top_pct"] = DEFAULTS["benchmark_top_pct"]
    st.session_state["slope_bid"] = DEFAULTS["slope_bid"]
    st.session_state["slope_ask"] = DEFAULTS["slope_ask"]
    st.session_state["spread"] = DEFAULTS["spread"]
    st.session_state["do_clean"] = DEFAULTS["do_clean"]
    st.session_state["lower_pct"] = DEFAULTS["lower_pct"]
    st.session_state["upper_pct"] = DEFAULTS["upper_pct"]
    st.session_state["price_method"] = DEFAULTS["price_method"]
    st.rerun()


# -------------------------
# App
# -------------------------
st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")
st.title("ETS Geliştirme Modülü V001")

st.write(
    """
### ETS Geliştirme Modülü – Model Açıklaması

Bu arayüz, elektrik üretim sektörüne yönelik **tesis bazlı ve piyasa tutarlı** bir **Emisyon Ticaret Sistemi (ETS)**
simülasyonu oluşturmak için tasarlanmıştır.

**Veri girişi**
- Excel’deki **tüm sekmeleri** okur ve tek bir veri setinde birleştirir.
- Her sekmenin adı otomatik olarak **FuelType** olarak eklenir.
- Beklenen kolonlar: **Plant**, **Generation_MWh**, **Emissions_tCO2**.

**Benchmark & Tahsis**
- Her yakıt türü için üretim ağırlıklı **benchmark (B_fuel)** hesaplanır.
- **Benchmark Top %**: Benchmark’ı belirlerken “en temiz” dilimi seçer.
  - %100 → tüm tesisler
  - %10 → en iyi %10’luk (en düşük yoğunluk) tesisler
- **AGK (Just Transition Coefficient)**, tahsis yoğunluğunu tesis yoğunluğu ile benchmark arasında harmanlar:

\[
T_i = I_i + AGK \cdot (B_{fuel}-I_i)
\]

- AGK = 0 → **Tesis yoğunluğu (I_i)** (tam yumuşak)
- AGK = 1 → **Benchmark (B_fuel)** (tam benchmark)

**Piyasa / Fiyat**
- Tüm tesisler **tek ETS piyasasında** clearing price üretir (yakıt bazında ayrı fiyat yok).
- **Price range (min–max)**: Piyasa clearing price bu aralık içinde aranır.
- **β_bid / β_ask / Spread**: BID–ASK eğrilerinin şekli ve ayrışmasını kalibre eder.

**Opsiyonel veri temizleme**
- (Varsa) veri temizleme modülü ile temel temizlik + outlier filtresi uygulanabilir.
- Outlier: tesis yoğunluğu, yakıt bazlı benchmark bandının dışında ise çıkarılır.

**AGK etkisi karşılaştırması (Yeni)**
- Model, mevcut AGK ile **AGK=1 (tam benchmark)** senaryosunu karşılaştırır:
  - Tesis bazında **ΔT (tahsis yoğunluğu farkı)**,
  - **Δ Net ETS**, **Δ Cost** ve **Δ Net Cashflow** (aynı clearing price varsayımıyla)
  - Ekranda ve Excel raporunda gösterir.

**Çıktılar**
- Uygulama içinde sonuç tabloları + piyasa grafikleri
- Tek tıkla **Excel rapor (grafikli)** + CSV
"""
)

# -------------------------
# Sidebar
# -------------------------
st.sidebar.header("Model Parameters")

if st.sidebar.button("Reset to defaults"):
    reset_to_defaults()

price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=st.session_state.get("price_range", DEFAULTS["price_range"]),
    step=1,
    key="price_range",
    help=f"Default: {DEFAULTS['price_range']}. Clearing price bu aralık içinde bulunur.",
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("agk", DEFAULTS["agk"])),
    step=0.05,
    key="agk",
    help=f"Default: {DEFAULTS['agk']}. AGK=1→Benchmark, AGK=0→Tesis yoğunluğu.",
)

benchmark_top_pct = st.sidebar.slider(
    "Benchmark Top % (best performers)",
    min_value=10,
    max_value=100,
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    step=10,
    key="benchmark_top_pct",
    help=f"Default: {DEFAULTS['benchmark_top_pct']}. Benchmark hesaplarında en düşük yoğunluk dilimi.",
)

st.sidebar.subheader("Price Method")

price_method = st.sidebar.selectbox(
    "Carbon Price Method",
    options=["Market Clearing", "Average Compliance Cost (ACC)"],
    index=["Market Clearing", "Average Compliance Cost (ACC)"].index(
        st.session_state.get("price_method", DEFAULTS["price_method"])
    ),
    key="price_method",
    help=(
        "Market Clearing: arz artan, talep azalan; kesişim fiyatı.\n"
        "ACC: pozitif net ETS yükümlülüklerinin ortalama uyum maliyeti (model yaklaşımı)."
    ),
)

st.sidebar.subheader("Market Calibration")

slope_bid = st.sidebar.slider(
    "Bid Slope (β_bid)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_bid", DEFAULTS["slope_bid"])),
    step=10,
    key="slope_bid",
    help=f"Default: {DEFAULTS['slope_bid']}. Alıcıların ödeme isteği hassasiyeti.",
)

slope_ask = st.sidebar.slider(
    "Ask Slope (β_ask)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_ask", DEFAULTS["slope_ask"])),
    step=10,
    key="slope_ask",
    help=f"Default: {DEFAULTS['slope_ask']}. Satıcıların satış isteği hassasiyeti.",
)

spread = st.sidebar.slider(
    "Bid/Ask Spread (€/tCO₂)",
    min_value=0.0,
    max_value=10.0,
    value=float(st.session_state.get("spread", DEFAULTS["spread"])),
    step=0.5,
    key="spread",
    help=f"Default: {DEFAULTS['spread']}. Spread bid/ask ayrışmasını artırır.",
)

st.sidebar.divider()
st.sidebar.caption("Excel'de beklenen kolonlar: Plant, Generation_MWh, Emissions_tCO2")
st.sidebar.caption("Sekme adı FuelType olarak alınır.")

# -------------------------
# Cleaning controls (optional)
# -------------------------
st.sidebar.subheader("Data Cleaning")

if HAS_CLEANING:
    do_clean = st.sidebar.toggle(
        "Apply cleaning rules?",
        value=bool(st.session_state.get("do_clean", DEFAULTS["do_clean"])),
        key="do_clean",
        help=f"Default: {'ON' if DEFAULTS['do_clean'] else 'OFF'}. Kapalıysa ham veriyle devam edilir.",
    )

    lower_pct = st.sidebar.slider(
        "Lower bound vs Benchmark (L)",
        min_value=0.0,
        max_value=1.0,
        value=float(st.session_state.get("lower_pct", DEFAULTS["lower_pct"])),
        step=0.05,
        key="lower_pct",
        help=f"Default: {DEFAULTS['lower_pct']}. Alt sınır = (1-L)*B.",
    )

    upper_pct = st.sidebar.slider(
        "Upper bound vs Benchmark (U)",
        min_value=0.0,
        max_value=2.0,
        value=float(st.session_state.get("upper_pct", DEFAULTS["upper_pct"])),
        step=0.05,
        key="upper_pct",
        help=f"Default: {DEFAULTS['upper_pct']}. Üst sınır = (1+U)*B.",
    )
else:
    do_clean = False
    lower_pct = DEFAULTS["lower_pct"]
    upper_pct = DEFAULTS["upper_pct"]
    st.sidebar.info("data_cleaning.py bulunamadı → temizleme devre dışı.")


# -------------------------
# Helpers
# -------------------------
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])


def read_all_sheets(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df["FuelType"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def build_market_curve(sonuc_df: pd.DataFrame, pmin: float, pmax: float, step: int = 1) -> pd.DataFrame:
    """Smooth (lineer) toplam arz-talep eğrileri."""
    prices = np.arange(pmin, pmax + step, step)

    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    rows = []
    for p in prices:
        # Demand
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()
            p_bid = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid - pmin, 1e-6)
            frac = 1.0 - (p - pmin) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            demand = 0.0

        # Supply
        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask = sellers["p_ask"].to_numpy()
            denom = np.maximum(pmax - p_ask, 1e-6)
            frac = (p - p_ask) / denom
            supply = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            supply = 0.0

        rows.append({"Price": float(p), "Demand": demand, "Supply": supply})

    return pd.DataFrame(rows)


def build_step_curves(sonuc_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """AB ETS mantığında step (kümülatif) eğriler: demand azalan, supply artan."""
    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["p_bid", "net_ets"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["p_ask", "net_ets"]].copy()

    demand_step = pd.DataFrame(columns=["Price", "CumQty"])
    supply_step = pd.DataFrame(columns=["Price", "CumQty"])

    if not buyers.empty:
        buyers = buyers.sort_values("p_bid", ascending=False)
        buyers["qty"] = buyers["net_ets"].clip(lower=0)
        buyers["CumQty"] = buyers["qty"].cumsum()
        demand_step = buyers.rename(columns={"p_bid": "Price"})[["Price", "CumQty"]]

    if not sellers.empty:
        sellers = sellers.sort_values("p_ask", ascending=True)
        sellers["qty"] = (-sellers["net_ets"]).clip(lower=0)
        sellers["CumQty"] = sellers["qty"].cumsum()
        supply_step = sellers.rename(columns={"p_ask": "Price"})[["Price", "CumQty"]]

    return demand_step, supply_step


def make_altair_market_charts(curve_df: pd.DataFrame, demand_step: pd.DataFrame, supply_step: pd.DataFrame, clearing_price: float):
    # Smooth chart
    long_df = curve_df.melt(id_vars=["Price"], value_vars=["Demand", "Supply"], var_name="Curve", value_name="value")
    smooth_lines = alt.Chart(long_df).mark_line().encode(
        x=alt.X("Price:Q", title="Price (€/tCO₂)"),
        y=alt.Y("value:Q", title="Volume (tCO₂)"),
        color=alt.Color("Curve:N", legend=alt.Legend(title="Curve")),
        tooltip=["Price", "Curve", "value"],
    )

    vline = alt.Chart(pd.DataFrame({"x": [clearing_price]})).mark_rule().encode(x="x:Q")

    smooth = (smooth_lines + vline).properties(
        height=320,
        title="Supply–Demand (Smooth) + Clearing Price",
    )

    # Step chart
    parts = []
    if not demand_step.empty:
        parts.append(
            alt.Chart(demand_step).mark_line(interpolate="step-after").encode(
                x=alt.X("Price:Q", title="Price (€/tCO₂)"),
                y=alt.Y("CumQty:Q", title="Cumulative volume (tCO₂)"),
                color=alt.value("#7fdbff"),
                tooltip=["Price", "CumQty"],
            )
        )
    if not supply_step.empty:
        parts.append(
            alt.Chart(supply_step).mark_line(interpolate="step-after").encode(
                x=alt.X("Price:Q", title="Price (€/tCO₂)"),
                y=alt.Y("CumQty:Q", title="Cumulative volume (tCO₂)"),
                color=alt.value("#0074d9"),
                tooltip=["Price", "CumQty"],
            )
        )

    step = alt.layer(*parts, vline).properties(
        height=320,
        title="AB ETS Step Curves (Bids/Asks) + Clearing Price",
    )

    return smooth, step


def agk_compare(sonuc_df: pd.DataFrame, clearing_price: float) -> pd.DataFrame:
    """
    Mevcut AGK ile AGK=1 (tam benchmark) karşılaştırması.
    Fiyat aynı clearing price kabul edilir (counterfactual).
    """
    df = sonuc_df.copy()

    if "intensity" not in df.columns:
        df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # Current already in tahsis_intensity
    df["tahsis_intensity_current"] = df.get("tahsis_intensity", np.nan)

    # AGK=1 => T = B_fuel
    df["tahsis_intensity_agk1"] = df["B_fuel"]

    # Free allocation comparison
    df["free_alloc_current"] = df.get("free_alloc", df["Generation_MWh"] * df["tahsis_intensity_current"])
    df["free_alloc_agk1"] = df["Generation_MWh"] * df["tahsis_intensity_agk1"]

    # Net ETS
    df["net_ets_current"] = df.get("net_ets", df["Emissions_tCO2"] - df["free_alloc_current"])
    df["net_ets_agk1"] = df["Emissions_tCO2"] - df["free_alloc_agk1"]

    # Cost/Revenue with SAME price (counterfactual)
    df["ets_cost_agk1_€"] = df["net_ets_agk1"].clip(lower=0) * clearing_price
    df["ets_rev_agk1_€"] = (-df["net_ets_agk1"]).clip(lower=0) * clearing_price
    df["ets_net_cashflow_agk1_€"] = df["ets_rev_agk1_€"] - df["ets_cost_agk1_€"]

    df["ets_cost_current_€"] = df.get("ets_cost_total_€", df["net_ets_current"].clip(lower=0) * clearing_price)
    df["ets_rev_current_€"] = df.get("ets_revenue_total_€", (-df["net_ets_current"]).clip(lower=0) * clearing_price)
    df["ets_net_cashflow_current_€"] = df.get("ets_net_cashflow_€", df["ets_rev_current_€"] - df["ets_cost_current_€"])

    # Deltas (Current - AGK=1)
    df["delta_tahsis_intensity"] = df["tahsis_intensity_current"] - df["tahsis_intensity_agk1"]
    df["delta_free_alloc"] = df["free_alloc_current"] - df["free_alloc_agk1"]
    df["delta_net_ets"] = df["net_ets_current"] - df["net_ets_agk1"]
    df["delta_cost_€"] = df["ets_cost_current_€"] - df["ets_cost_agk1_€"]
    df["delta_net_cashflow_€"] = df["ets_net_cashflow_current_€"] - df["ets_net_cashflow_agk1_€"]

    keep = [
        "Plant", "FuelType",
        "intensity", "B_fuel",
        "tahsis_intensity_current", "tahsis_intensity_agk1", "delta_tahsis_intensity",
        "free_alloc_current", "free_alloc_agk1", "delta_free_alloc",
        "net_ets_current", "net_ets_agk1", "delta_net_ets",
        "ets_cost_current_€", "ets_cost_agk1_€", "delta_cost_€",
        "ets_net_cashflow_current_€", "ets_net_cashflow_agk1_€", "delta_net_cashflow_€",
    ]
    keep = [c for c in keep if c in df.columns]
    return df[keep]


# -------------------------
# Load data
# -------------------------
if uploaded is None:
    st.info("Lütfen bir Excel yükleyin.")
    st.stop()

try:
    df_all_raw = read_all_sheets(uploaded)
except Exception as e:
    st.error(f"Excel okunurken hata oluştu: {e}")
    st.stop()

st.subheader("Yüklenen veri (ham / birleştirilmiş)")
st.dataframe(df_all_raw.head(50), use_container_width=True)

# -------------------------
# Cleaning (optional)
# -------------------------
df_all = df_all_raw.copy()
if do_clean and HAS_CLEANING:
    st.subheader("Veri Temizleme (aktif)")
    cleaned_frames = []
    reports_basic = []

    for ft in df_all["FuelType"].dropna().unique():
        part = df_all[df_all["FuelType"] == ft].copy()
        cleaned, rep = clean_ets_input(part, fueltype=ft)
        rep["FuelType"] = ft
        reports_basic.append(rep)
        cleaned_frames.append(cleaned)

    df_clean = pd.concat(cleaned_frames, ignore_index=True)
    rep_basic_df = pd.DataFrame(reports_basic)
    st.write("Temel temizlik özeti (sekme bazında):")
    st.dataframe(rep_basic_df, use_container_width=True)

    before = len(df_clean)
    df_clean2, rep_out = filter_intensity_outliers_by_fuel(
        df_clean, lower_pct=lower_pct, upper_pct=upper_pct
    )
    after = len(df_clean2)

    st.info(
        f"Outlier filtresi: {rep_out.get('outliers_dropped', 0)} satır çıkarıldı "
        f"({before:,} → {after:,}). Band: [{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]"
    )

    df_all = df_clean2

    clean_out = BytesIO()
    with pd.ExcelWriter(clean_out, engine="openpyxl") as w:
        df_all.to_excel(w, index=False, sheet_name="Cleaned_Data")
    clean_out.seek(0)

    st.download_button(
        "Download Cleaned Data (Excel)",
        data=clean_out,
        file_name="ETS_Cleaned_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
elif do_clean and not HAS_CLEANING:
    st.warning("Temizleme açık seçildi ama data_cleaning.py bulunamadı → ham veriyle devam.")
else:
    st.caption("Temizleme kapalı: ham veriyle devam ediliyor.")

st.subheader("Modelde kullanılacak veri (ilk 50 satır)")
st.dataframe(df_all.head(50), use_container_width=True)

# -------------------------
# Run model
# -------------------------
if st.button("Run ETS Model"):
    try:
        sonuc_df, benchmark_map, clearing_price = ets_hesapla(
            df_all,
            float(price_min),
            float(price_max),
            float(agk),
            slope_bid=float(slope_bid),
            slope_ask=float(slope_ask),
            spread=float(spread),
            benchmark_top_pct=int(benchmark_top_pct),
            price_method=str(price_method),
        )

        st.success(f"Clearing Price: {clearing_price:.2f} €/tCO₂  |  Method: {price_method}")

        # Benchmark table
        st.subheader("Benchmark (yakıt bazında)")
        bench_df = (
            pd.DataFrame(
                [{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()]
            )
            .sort_values("FuelType")
            .reset_index(drop=True)
        )
        st.dataframe(bench_df, use_container_width=True)

        # KPI summary
        total_cost = float(sonuc_df.get("ets_cost_total_€", pd.Series(dtype=float)).sum())
        total_revenue = float(sonuc_df.get("ets_revenue_total_€", pd.Series(dtype=float)).sum())
        net_cashflow = float(sonuc_df.get("ets_net_cashflow_€", pd.Series(dtype=float)).sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam ETS Maliyeti (€)", f"{total_cost:,.0f}")
        c2.metric("Toplam ETS Geliri (€)", f"{total_revenue:,.0f}")
        c3.metric("Net Nakit Akışı (€)", f"{net_cashflow:,.0f}")

        # -------------------------
        # AGK impact vs AGK=1 (Benchmark baseline)
        # -------------------------
        st.subheader("AGK Etkisi (Mevcut AGK vs AGK=1)")
        comp_df = agk_compare(sonuc_df, float(clearing_price))

        # Fuel-level summary (weighted by generation if available)
        fuel_summary = None
        if "Generation_MWh" in sonuc_df.columns and "FuelType" in sonuc_df.columns:
            tmp = sonuc_df.copy()
            if "tahsis_intensity" in tmp.columns and "B_fuel" in tmp.columns:
                tmp["tahsis_intensity_agk1"] = tmp["B_fuel"]
                fuel_summary = (
                    tmp.groupby("FuelType", dropna=False)
                    .apply(lambda g: pd.Series({
                        "Gen_MWh": float(g["Generation_MWh"].sum()),
                        "Avg_I": float((g["Emissions_tCO2"].sum() / g["Generation_MWh"].sum()) if g["Generation_MWh"].sum() > 0 else np.nan),
                        "Avg_T_current": float(np.average(g["tahsis_intensity"], weights=g["Generation_MWh"])) if g["Generation_MWh"].sum() > 0 else np.nan,
                        "Avg_T_AGK1": float(np.average(g["tahsis_intensity_agk1"], weights=g["Generation_MWh"])) if g["Generation_MWh"].sum() > 0 else np.nan,
                    }))
                    .reset_index()
                )
                fuel_summary["Delta_T"] = fuel_summary["Avg_T_current"] - fuel_summary["Avg_T_AGK1"]

        left, right = st.columns([1, 2])

        with left:
            plant_list = comp_df["Plant"].dropna().unique().tolist()
            selected = st.selectbox("Santral seç (örnek gösterim)", options=plant_list[:500] if plant_list else ["—"])
            if plant_list:
                row = comp_df[comp_df["Plant"] == selected].iloc[0].to_dict()
                st.metric("I (tesis yoğunluğu)", f"{row.get('intensity', np.nan):.4f}")
                st.metric("B_fuel (benchmark)", f"{row.get('B_fuel', np.nan):.4f}")
                st.metric("T_current (AGK)", f"{row.get('tahsis_intensity_current', np.nan):.4f}")
                st.metric("T_AGK=1", f"{row.get('tahsis_intensity_agk1', np.nan):.4f}")
                st.metric("ΔT (current - AGK1)", f"{row.get('delta_tahsis_intensity', np.nan):.4f}")

        with right:
            if fuel_summary is not None and not fuel_summary.empty:
                st.write("Yakıt bazında (üretim ağırlıklı) ortalama tahsis yoğunluğu karşılaştırması:")
                fs_long = fuel_summary.melt(
                    id_vars=["FuelType"],
                    value_vars=["Avg_I", "Avg_T_current", "Avg_T_AGK1"],
                    var_name="Metric",
                    value_name="Value",
                )
                chart = (
                    alt.Chart(fs_long)
                    .mark_bar()
                    .encode(
                        x=alt.X("FuelType:N", title="FuelType"),
                        y=alt.Y("Value:Q", title="tCO₂/MWh"),
                        color=alt.Color("Metric:N", legend=alt.Legend(title="")),
                        tooltip=["FuelType", "Metric", "Value"],
                    )
                    .properties(height=260)
                )
                st.altair_chart(chart, use_container_width=True)
            else:
                st.info("Fuel summary üretilemedi (kolonlar eksik olabilir).")

        st.dataframe(comp_df.head(50), use_container_width=True)

        # -------------------------
        # Market charts (app)
        # -------------------------
        st.subheader("📈 Piyasa Grafikleri (Uygulama içi)")

        curve_df = build_market_curve(sonuc_df, float(price_min), float(price_max), step=1)
        demand_step, supply_step = build_step_curves(sonuc_df)
        smooth_chart, step_chart = make_altair_market_charts(curve_df, demand_step, supply_step, float(clearing_price))

        st.altair_chart(smooth_chart, use_container_width=True)
        st.altair_chart(step_chart, use_container_width=True)

        # -------------------------
        # Tables
        # -------------------------
        st.subheader("ETS Sonuçları – Alıcılar (Net ETS > 0)")
        buyers_df = sonuc_df[sonuc_df["net_ets"] > 0].copy()
        st.dataframe(buyers_df, use_container_width=True)

        st.subheader("ETS Sonuçları – Satıcılar (Net ETS < 0)")
        sellers_df = sonuc_df[sonuc_df["net_ets"] < 0].copy()
        st.dataframe(sellers_df, use_container_width=True)

        st.subheader("Tüm Sonuçlar (ham tablo)")
        st.dataframe(sonuc_df, use_container_width=True)

        # -------------------------
        # Excel report + charts
        # -------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df = pd.DataFrame(
                {
                    "Metric": [
                        "Clearing Price (€/tCO₂)",
                        "Price Method",
                        "Total ETS Cost (€)",
                        "Total ETS Revenue (€)",
                        "Net Cashflow (€)",
                        "Price Min",
                        "Price Max",
                        "AGK",
                        "Benchmark Top %",
                        "Bid Slope",
                        "Ask Slope",
                        "Spread",
                        "Cleaning Applied",
                        "Outlier Band (lower, upper)",
                    ],
                    "Value": [
                        clearing_price,
                        price_method,
                        total_cost,
                        total_revenue,
                        net_cashflow,
                        price_min,
                        price_max,
                        agk,
                        benchmark_top_pct,
                        slope_bid,
                        slope_ask,
                        spread,
                        str(do_clean),
                        f"[{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]" if do_clean else "N/A",
                    ],
                }
            )
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            bench_df.to_excel(writer, sheet_name="Benchmarks", index=False)
            sonuc_df.to_excel(writer, sheet_name="All_Plants", index=False)

            # Market curve (smooth)
            curve_df.to_excel(writer, sheet_name="Market_Curve", index=False)

            # AGK comparison
            comp_df.to_excel(writer, sheet_name="AGK_Compare", index=False)
            if fuel_summary is not None:
                fuel_summary.to_excel(writer, sheet_name="Fuel_Summary", index=False)

            # ----- OpenPyXL charts -----
            wb = writer.book

            # 1) Supply–Demand line chart (Market_Curve)
            ws_curve = wb["Market_Curve"]
            ws_curve["D1"] = "Clearing_Price"
            for r in range(2, ws_curve.max_row + 1):
                ws_curve[f"D{r}"] = float(clearing_price)

            line = LineChart()
            line.title = "Market Supply–Demand Curve"
            line.y_axis.title = "Volume (tCO₂)"
            line.x_axis.title = "Price (€/tCO₂)"
            data = Reference(ws_curve, min_col=2, min_row=1, max_col=3, max_row=ws_curve.max_row)
            cats = Reference(ws_curve, min_col=1, min_row=2, max_row=ws_curve.max_row)
            line.add_data(data, titles_from_data=True)
            line.set_categories(cats)
            line.add_data(Reference(ws_curve, min_col=4, min_row=1, max_row=ws_curve.max_row), titles_from_data=True)
            line.height = 12
            line.width = 24
            ws_curve.add_chart(line, "F2")

            # 2) AGK delta cost bar chart (top 20)
            ws_cmp = wb["AGK_Compare"]
            ws_cmp["Z1"] = "Plant"
            ws_cmp["AA1"] = "Delta_Cost_€"

            headers = [ws_cmp.cell(row=1, column=c).value for c in range(1, ws_cmp.max_column + 1)]
            try:
                _ = headers.index("Plant")
                _ = headers.index("delta_cost_€")
                comp_top = comp_df.sort_values("delta_cost_€", ascending=False).head(20)
                for i, (_, rr) in enumerate(comp_top.iterrows(), start=2):
                    ws_cmp[f"Z{i}"] = rr["Plant"]
                    ws_cmp[f"AA{i}"] = float(rr["delta_cost_€"])

                bar = BarChart()
                bar.type = "col"
                bar.title = "AGK Impact vs AGK=1 (Top 20 Δ Cost €)"
                bar.y_axis.title = "Δ Cost (€)"
                bar.x_axis.title = "Plant"
                data_b = Reference(ws_cmp, min_col=27, min_row=1, max_col=27, max_row=21)  # AA
                cats_b = Reference(ws_cmp, min_col=26, min_row=2, max_row=21)  # Z
                bar.add_data(data_b, titles_from_data=True)
                bar.set_categories(cats_b)
                bar.height = 12
                bar.width = 28
                bar.dataLabels = DataLabelList()
                bar.dataLabels.showVal = False
                ws_cmp.add_chart(bar, "AD2")
            except Exception:
                pass

        output.seek(0)

        st.download_button(
            label="Download ETS Report (Excel + Charts)",
            data=output,
            file_name="ETS_Report_Stable_WithCharts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        csv_bytes = sonuc_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download results as CSV",
            data=csv_bytes,
            file_name="ets_results.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Model çalışırken hata oluştu: {e}")
