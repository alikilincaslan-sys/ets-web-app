import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel


# -------------------------
# Default values (V001 Stable)
# -------------------------
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
    # Excel-only AGK sensitivity (no UI chart)
    "agk_list_excel": None,
}


def build_agk_excel_table(sonuc_df: pd.DataFrame, agk_list: list[float]) -> pd.DataFrame:
    """AGK duyarlılık tablosu (sadece Excel çıktısı için, arayüzde grafik yok).

    Beklenen kolonlar: Plant, FuelType, intensity, B_fuel

    Tahsis yoğunluğu formülü:
      T_i(alpha) = I_i + alpha * (B_fuel - I_i)

    Çıktı: Plant bazında, seçilen her alpha için ayrı kolon.
    """
    if sonuc_df.empty:
        return pd.DataFrame()

    need = {"Plant", "FuelType", "intensity", "B_fuel"}
    missing = [c for c in need if c not in sonuc_df.columns]
    if missing:
        raise ValueError(f"AGK Excel tablosu için eksik kolon(lar): {missing}")

    base = sonuc_df[["Plant", "FuelType", "intensity", "B_fuel"]].copy()

    # güvenli + tekilleştir
    base = base.drop_duplicates(subset=["Plant", "FuelType"])

    for a in agk_list:
        col = f"AGK_{a:.2f}".replace(".", "p")
        base[col] = base["intensity"] + float(a) * (base["B_fuel"] - base["intensity"])

    # Sıralama: ilk alpha kolonu baz alınır
    sort_col = f"AGK_{agk_list[0]:.2f}".replace(".", "p")
    base = base.sort_values(sort_col, ascending=True).reset_index(drop=True)
    return base

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

# -------------------------
# Model açıklaması (tek blok - düzeltilmiş)
# -------------------------
with st.expander("📌 Model Açıklaması / Sliderlar neyi değiştiriyor?", expanded=True):
    st.markdown(
        """
### ETS Geliştirme Modülü – Model Açıklaması

Bu arayüz, elektrik üretim sektörüne yönelik **tesis bazlı** ve **piyasa tutarlı** bir **ETS (Emisyon Ticaret Sistemi)** simülasyonu oluşturur.

**Veri girişi**
- Excel’deki **tüm sekmeleri** okur ve birleştirir (**FuelType = sekme adı**).
- Beklenen kolonlar: `Plant`, `Generation_MWh`, `Emissions_tCO2`

**Benchmark (yakıt bazlı)**
- Yakıt türü içinde üretim ağırlıklı benchmark hesaplanır.
- **Benchmark Top %**: Yakıt içindeki “en düşük intensity” dilimini seçer:
  - 100% = tüm tesisler (varsayılan)
  - 10% / 20% = en temiz dilim (daha sıkı benchmark)

**AGK (Adil Geçiş Katsayısı)**
- Tahsis yoğunluğu formülü:
  - **Tᵢ = Iᵢ + AGK × (B_fuel − Iᵢ)**
- AGK=1 → Benchmark’a tam yaklaşır (varsayılan)
- AGK=0 → Tesis kendi yoğunluğunda kalır

**Karbon fiyatı (tek piyasa)**
- Tüm tesisler tek piyasada birleşir ve **tek karbon fiyatı** oluşur.
- **Price Method**
  - Market Clearing: arz-talep kesişimi
  - ACC: alıcıların p_bid değerlerinin (net yükümlülükle ağırlıklı) ortalaması
- **Carbon Price Range (min–max)**: fiyat bu aralıkta kalır.

**Market Calibration**
- β_bid: alıcıların fiyat hassasiyeti
- β_ask: satıcıların fiyat hassasiyeti
- Spread: BID/ASK ayrışması için ek fark

**Veri Temizleme (opsiyonel)**
- Cleaning OFF ise sadece temel temizlik yapılır.
- Cleaning ON ise intensity outlier’lar benchmark bandına göre filtrelenir:
  - lo = B × (1 − L)
  - hi = B × (1 + U)

**Çıktılar**
- Sonuç tabloları + Excel rapor (çok sayfalı) + grafikler (Supply–Demand ve Top-20 cashflow)
"""
    )

# -------------------------
# Sidebar: Reset
# -------------------------
st.sidebar.header("Model Parameters")

if st.sidebar.button("🔄 Reset to Default"):
    st.session_state["price_range"] = DEFAULTS["price_range"]
    st.session_state["agk"] = DEFAULTS["agk"]
    st.session_state["benchmark_top_pct"] = DEFAULTS["benchmark_top_pct"]
    st.session_state["price_method"] = DEFAULTS["price_method"]
    st.session_state["slope_bid"] = DEFAULTS["slope_bid"]
    st.session_state["slope_ask"] = DEFAULTS["slope_ask"]
    st.session_state["spread"] = DEFAULTS["spread"]
    st.session_state["do_clean"] = DEFAULTS["do_clean"]
    st.session_state["lower_pct"] = DEFAULTS["lower_pct"]
    st.session_state["upper_pct"] = DEFAULTS["upper_pct"]
    st.rerun()

# -------------------------
# Sidebar: sliders (session_state bağlı)
# -------------------------
price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=st.session_state.get("price_range", DEFAULTS["price_range"]),
    step=1,
    key="price_range",
    help="Clearing price bu aralık içinde bulunur.",
)
st.sidebar.caption("Default: (5, 20)")

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("agk", DEFAULTS["agk"])),
    step=0.05,
    key="agk",
    help="AGK=1→Benchmark, AGK=0→Tesis yoğunluğu. Tᵢ = Iᵢ + AGK×(B − Iᵢ)",
)
st.sidebar.caption("Default: AGK = 1.00")

# Excel-only: AGK senaryo listesi (arayüzde grafik yok, sadece Excel 'AGK_SONUC' sayfasına yazılır)
agk_excel_options = [round(x, 2) for x in np.arange(0.0, 1.0001, 0.05)]
default_excel_list = st.session_state.get("agk_list_excel")
if not default_excel_list:
    default_excel_list = [float(agk)]

agk_list_excel = st.sidebar.multiselect(
    "AGK senaryoları (Excel çıktısı)",
    options=agk_excel_options,
    default=default_excel_list,
    key="agk_list_excel",
    help="Grafik yok. Seçilen AGK değerleri için tahsis yoğunluğu (T_i) Excel'de AGK_SONUC sayfasına yazılır.",
)

if not agk_list_excel:
    agk_list_excel = [float(agk)]

st.sidebar.subheader("Benchmark Settings")
benchmark_top_pct = st.sidebar.select_slider(
    "Benchmark = Best plants (by intensity) %",
    options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    key="benchmark_top_pct",
    help="Yakıt bazında benchmark, intensity düşük olan en iyi dilimden hesaplanır. 100=tüm tesisler.",
)
st.sidebar.caption("Default: 100")

st.sidebar.subheader("Carbon Price Method")
_methods = ["Market Clearing", "Average Compliance Cost"]
_default_method = st.session_state.get("price_method", DEFAULTS["price_method"])
if _default_method not in _methods:
    _default_method = "Market Clearing"

price_method = st.sidebar.selectbox(
    "Price calculation method",
    options=_methods,
    index=_methods.index(_default_method),
    key="price_method",
    help="Market Clearing: arz-talep kesişimi. ACC: alıcıların p_bid (net_ets ile ağırlıklı) ortalaması.",
)
st.sidebar.caption("Default: Market Clearing")

st.sidebar.subheader("Market Calibration")

slope_bid = st.sidebar.slider(
    "Bid Slope (β_bid)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_bid", DEFAULTS["slope_bid"])),
    step=10,
    key="slope_bid",
    help="Alıcıların (kirli tesis) ödeme isteği hassasiyeti.",
)
st.sidebar.caption("Default: 150")

slope_ask = st.sidebar.slider(
    "Ask Slope (β_ask)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_ask", DEFAULTS["slope_ask"])),
    step=10,
    key="slope_ask",
    help="Satıcıların (temiz tesis) satış isteği hassasiyeti.",
)
st.sidebar.caption("Default: 150")

spread = st.sidebar.slider(
    "Bid/Ask Spread (€/tCO₂)",
    min_value=0.0,
    max_value=10.0,
    value=float(st.session_state.get("spread", DEFAULTS["spread"])),
    step=0.5,
    key="spread",
    help="Spread eklemek bid/ask aynı görünmesini azaltır.",
)
st.sidebar.caption("Default: 1.0")

st.sidebar.divider()
st.sidebar.caption("Excel'de beklenen kolonlar: Plant, Generation_MWh, Emissions_tCO2")
st.sidebar.caption("Sekme adı FuelType olarak alınır.")

# -------------------------
# Data Cleaning Controls
# -------------------------
st.sidebar.subheader("Data Cleaning")

do_clean = st.sidebar.toggle(
    "Apply cleaning rules?",
    value=bool(st.session_state.get("do_clean", DEFAULTS["do_clean"])),
    key="do_clean",
    help="Kapalıysa (Hayır), outlier filtresi uygulanmaz.",
)
st.sidebar.caption("Default: OFF")

lower_pct = st.sidebar.slider(
    "Lower bound vs Benchmark (L)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("lower_pct", DEFAULTS["lower_pct"])),
    step=0.05,
    key="lower_pct",
    help="lo = B*(1-L). L=1.0 => lo=0. L=0.5 => lo=0.5B.",
)
st.sidebar.caption("Default: 1.0")

upper_pct = st.sidebar.slider(
    "Upper bound vs Benchmark (U)",
    min_value=0.0,
    max_value=2.0,
    value=float(st.session_state.get("upper_pct", DEFAULTS["upper_pct"])),
    step=0.05,
    key="upper_pct",
    help="hi = B*(1+U). U=1.0 => hi=2B. U=2.0 => hi=3B.",
)
st.sidebar.caption("Default: 2.0")

# -------------------------
# Excel upload
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


def build_market_curve(sonuc_df: pd.DataFrame, price_min: int, price_max: int, step: int = 1) -> pd.DataFrame:
    prices = np.arange(price_min, price_max + step, step)

    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    rows = []
    for p in prices:
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()
            p_bid_arr = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid_arr - price_min, 1e-6)
            frac = 1.0 - (p - price_min) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            demand = 0.0

        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask_arr = sellers["p_ask"].to_numpy()
            denom = np.maximum(price_max - p_ask_arr, 1e-6)
            frac = (p - p_ask_arr) / denom
            supply = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            supply = 0.0

        rows.append({"Price": float(p), "Total_Demand": demand, "Total_Supply": supply})

    return pd.DataFrame(rows)


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
# Cleaning
# -------------------------
st.subheader("Veri Temizleme (opsiyonel)")

df_all = df_all_raw.copy()

try:
    df_all = clean_ets_input(df_all)
except Exception as e:
    st.error(f"Temel temizlikte hata: {e}")
    st.stop()

removed_df = pd.DataFrame()

if do_clean:
    before = len(df_all)
    try:
        df_all, removed_df = filter_intensity_outliers_by_fuel(
            df_all, lower_pct=lower_pct, upper_pct=upper_pct
        )
    except Exception as e:
        st.error(f"Outlier filtresinde hata: {e}")
        st.stop()

    after = len(df_all)
    st.info(
        f"Outlier filtresi: {before - after} satır çıkarıldı "
        f"({before:,} → {after:,}). Band: [{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]"
    )
    if not removed_df.empty:
        with st.expander("Çıkarılan outlier satırlar (önizleme)"):
            st.dataframe(removed_df.head(200), use_container_width=True)
else:
    st.warning("Temizleme kapalı: (sadece temel temizlik yapıldı)")

st.subheader("Modelde kullanılacak veri (ilk 50 satır)")
st.dataframe(df_all.head(50), use_container_width=True)

# -------------------------
# Run model
# -------------------------
if st.button("Run ETS Model"):
    try:
        sonuc_df, benchmark_map, clearing_price = ets_hesapla(
            df_all,
            price_min,
            price_max,
            agk,
            slope_bid=slope_bid,
            slope_ask=slope_ask,
            spread=spread,
            benchmark_top_pct=int(benchmark_top_pct),
            price_method=price_method,
        )

        st.success(f"Carbon Price ({price_method}): {clearing_price:.2f} €/tCO₂")
        st.caption(f"Benchmark method: Best {benchmark_top_pct}% (production-share, by lowest intensity)")

        st.subheader("Benchmark (yakıt bazında)")
        bench_df = (
            pd.DataFrame([{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()])
            .sort_values("FuelType")
            .reset_index(drop=True)
        )
        st.dataframe(bench_df, use_container_width=True)

        total_cost = float(sonuc_df["ets_cost_total_€"].sum())
        total_revenue = float(sonuc_df["ets_revenue_total_€"].sum())
        net_cashflow = float(sonuc_df["ets_net_cashflow_€"].sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam ETS Maliyeti (€)", f"{total_cost:,.0f}")
        c2.metric("Toplam ETS Geliri (€)", f"{total_revenue:,.0f}")
        c3.metric("Net Nakit Akışı (€)", f"{net_cashflow:,.0f}")

        st.subheader("ETS Sonuçları – Alıcılar (Net ETS > 0)")
        buyers_df = sonuc_df[sonuc_df["net_ets"] > 0].copy()
        st.dataframe(
            buyers_df[
                [
                    "Plant",
                    "FuelType",
                    "net_ets",
                    "carbon_price",
                    "ets_cost_total_€",
                    "ets_cost_€/MWh",
                    "ets_net_cashflow_€",
                    "ets_net_cashflow_€/MWh",
                ]
            ],
            use_container_width=True,
        )

        st.subheader("ETS Sonuçları – Satıcılar (Net ETS < 0)")
        sellers_df = sonuc_df[sonuc_df["net_ets"] < 0].copy()
        st.dataframe(
            sellers_df[
                [
                    "Plant",
                    "FuelType",
                    "net_ets",
                    "carbon_price",
                    "ets_revenue_total_€",
                    "ets_revenue_€/MWh",
                    "ets_net_cashflow_€",
                    "ets_net_cashflow_€/MWh",
                ]
            ],
            use_container_width=True,
        )

        st.subheader("Tüm Sonuçlar (ham tablo)")
        st.dataframe(sonuc_df, use_container_width=True)

        curve_df = build_market_curve(sonuc_df, price_min, price_max, step=1)

        cashflow_top20 = (
            sonuc_df[["Plant", "FuelType", "ets_net_cashflow_€"]]
            .copy()
            .sort_values("ets_net_cashflow_€", ascending=False)
            .head(20)
        )

        # Excel-only AGK sensitivity table (no UI chart)
        agk_excel_df = build_agk_excel_table(sonuc_df, agk_list_excel)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df = pd.DataFrame(
                {
                    "Metric": [
                        "Carbon Price (€/tCO₂)",
                        "Price Method",
                        "Total ETS Cost (€)",
                        "Total ETS Revenue (€)",
                        "Net Cashflow (€)",
                        "Price Min",
                        "Price Max",
                        "AGK",
                        "AGK Excel Scenarios",
                        "Benchmark Top %",
                        "Bid Slope",
                        "Ask Slope",
                        "Spread",
                        "Cleaning Applied",
                        "Outlier Band",
                        "Rows (raw)",
                        "Rows (used)",
                        "Rows removed (outlier)",
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
                        ", ".join([str(round(x, 2)) for x in agk_list_excel]),
                        int(benchmark_top_pct),
                        slope_bid,
                        slope_ask,
                        spread,
                        str(do_clean),
                        f"[{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]" if do_clean else "N/A",
                        len(df_all_raw),
                        len(df_all),
                        0 if removed_df.empty else len(removed_df),
                    ],
                }
            )
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            bench_df.to_excel(writer, sheet_name="Benchmarks", index=False)
            sonuc_df.to_excel(writer, sheet_name="All_Plants", index=False)
            buyers_df.to_excel(writer, sheet_name="Buyers", index=False)
            sellers_df.to_excel(writer, sheet_name="Sellers", index=False)
            curve_df.to_excel(writer, sheet_name="Market_Curve", index=False)
            cashflow_top20.to_excel(writer, sheet_name="Cashflow_Top20", index=False)
            if not agk_excel_df.empty:
                agk_excel_df.to_excel(writer, sheet_name="AGK_SONUC", index=False)
            if not removed_df.empty:
                removed_df.to_excel(writer, sheet_name="Removed_Outliers", index=False)

            wb = writer.book

            ws_curve = wb["Market_Curve"]
            line = LineChart()
            line.title = "Market Supply–Demand Curve"
            line.y_axis.title = "Volume (tCO₂)"
            line.x_axis.title = "Price (€/tCO₂)"

            data = Reference(ws_curve, min_col=2, min_row=1, max_col=3, max_row=ws_curve.max_row)
            cats = Reference(ws_curve, min_col=1, min_row=2, max_row=ws_curve.max_row)
            line.add_data(data, titles_from_data=True)
            line.set_categories(cats)
            line.height = 12
            line.width = 24

            ws_curve["D1"] = "Carbon_Price"
            for r in range(2, ws_curve.max_row + 1):
                ws_curve[f"D{r}"] = float(clearing_price)

            line.add_data(
                Reference(ws_curve, min_col=4, min_row=1, max_row=ws_curve.max_row),
                titles_from_data=True,
            )
            ws_curve.add_chart(line, "E2")

            ws_cf = wb["Cashflow_Top20"]
            bar = BarChart()
            bar.type = "col"
            bar.title = "Top 20 Plants – ETS Net Cashflow (€)"
            bar.y_axis.title = "€"
            bar.x_axis.title = "Plant"

            data_cf = Reference(ws_cf, min_col=3, min_row=1, max_row=ws_cf.max_row)
            cats_cf = Reference(ws_cf, min_col=1, min_row=2, max_row=ws_cf.max_row)
            bar.add_data(data_cf, titles_from_data=True)
            bar.set_categories(cats_cf)
            bar.height = 12
            bar.width = 28

            bar.dataLabels = DataLabelList()
            bar.dataLabels.showVal = False

            ws_cf.add_chart(bar, "E2")

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
