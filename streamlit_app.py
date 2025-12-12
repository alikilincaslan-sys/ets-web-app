import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

import matplotlib.pyplot as plt

from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

from ets_model import ets_hesapla

# -------------------------
# data_cleaning import (opsiyonel, yoksa app düşmesin)
# -------------------------
try:
    from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel
    HAS_CLEANING = True
except Exception:
    HAS_CLEANING = False

    def clean_ets_input(df: pd.DataFrame):
        """Fallback: minimum temizlik + kolon standardizasyonu."""
        df = df.copy()

        # Bazı olası kolon isimlerini normalize et
        rename_map = {}
        if "Emissions" in df.columns and "Emissions_tCO2" not in df.columns:
            rename_map["Emissions"] = "Emissions_tCO2"
        if "Generation" in df.columns and "Generation_MWh" not in df.columns:
            rename_map["Generation"] = "Generation_MWh"
        if rename_map:
            df = df.rename(columns=rename_map)

        # Zorunlu kolonlar
        required = ["Plant", "Generation_MWh", "Emissions_tCO2", "FuelType"]
        for c in required:
            if c not in df.columns:
                raise ValueError(f"Excel kolon eksik: {c}")

        # Tip dönüşümü
        df["Generation_MWh"] = pd.to_numeric(df["Generation_MWh"], errors="coerce")
        df["Emissions_tCO2"] = pd.to_numeric(df["Emissions_tCO2"], errors="coerce")

        # Null/0 temizliği
        df = df.dropna(subset=["Plant", "FuelType", "Generation_MWh", "Emissions_tCO2"])
        df = df[df["Generation_MWh"] > 0]
        df = df[df["Emissions_tCO2"] >= 0]

        return df

    def filter_intensity_outliers_by_fuel(df: pd.DataFrame, lower_pct: float, upper_pct: float):
        """
        Fallback: Yakıt bazında benchmark'a göre intensity bandı dışını at.
        Band:
          lo = (1 - L) * B
          hi = (1 + U) * B
        """
        df = df.copy()
        df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

        removed = []
        kept = []

        for ft, part in df.groupby("FuelType", dropna=False):
            B = part["Emissions_tCO2"].sum() / part["Generation_MWh"].sum()
            lo = (1.0 - lower_pct) * B
            hi = (1.0 + upper_pct) * B

            mask = (part["intensity"] >= lo) & (part["intensity"] <= hi)
            kept.append(part[mask].copy())
            removed.append(part[~mask].copy())

        kept_df = pd.concat(kept, ignore_index=True) if kept else df.iloc[0:0].copy()
        removed_df = pd.concat(removed, ignore_index=True) if removed else df.iloc[0:0].copy()

        return kept_df, removed_df


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
}

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")
st.title("ETS Geliştirme Modülü V001")

# -------------------------
# Açıklama (tek blok, düzgün)
# -------------------------
st.write(
    """
### ETS Geliştirme Modülü – Model Açıklaması

Bu arayüz, elektrik üretim sektörüne yönelik **tesis bazlı** bir ETS (Emisyon Ticaret Sistemi) simülasyonu çalıştırır ve sonuçları **Excel raporu + grafikler** olarak indirmenizi sağlar.

**Model akışı**
1) Excel’deki tüm sekmeler okunur ve birleştirilir (**FuelType = sekme adı**)  
2) Yakıt bazında benchmark hesaplanır (**Benchmark Top %** ile “en iyi dilim” seçilebilir)  
3) **AGK** ile tahsis yoğunluğu hesaplanır: **Tᵢ = Iᵢ + AGK×(B − Iᵢ)**  
4) Ücretsiz tahsis → net ETS pozisyonu (alıcı/satıcı) bulunur  
5) Tüm tesisler tek piyasada toplanır ve **tek bir karbon fiyatı** üretilir  
6) Maliyet/gelir/net nakit akışı hesaplanır  
7) Çıktılar Excel/CSV olarak indirilir

**Fiyat yöntemi (tek seçim)**
- **Market Clearing:** Toplam arz-talep eğrilerinin kesiştiği fiyat (AB ETS mantığına en yakın)  
- **Average Compliance Cost (ACC):** Sadece alıcıların (NetETS>0) ödeme isteğine dayalı ağırlıklı ortalama fiyat (band içinde kırpılır)

**Varsayılan (Default)**
- Price range: **(5, 20)**, AGK: **1.00**, Benchmark Top %: **100**
- β_bid: **150**, β_ask: **150**, Spread: **1**
- Cleaning: **OFF**, L: **1.0**, U: **2.0**
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
    help="AGK yönü: AGK=1→Benchmark, AGK=0→Tesis yoğunluğu (Tᵢ = Iᵢ + AGK×(B − Iᵢ))",
)
st.sidebar.caption("Default: 1.00")

st.sidebar.subheader("Benchmark Settings")
benchmark_top_pct = st.sidebar.select_slider(
    "Benchmark = Best plants (by intensity) %",
    options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    key="benchmark_top_pct",
    help="Yakıt bazında benchmark, intensity düşük olan en iyi dilimden hesaplanır. 100 = tüm tesisler.",
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
    help="Market Clearing: arz-talep kesişimi. ACC: alıcıların p_bid değerlerinin (NetETS ile ağırlıklı) ortalaması.",
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
    "Apply outlier filter (optional)?",
    value=bool(st.session_state.get("do_clean", DEFAULTS["do_clean"])),
    key="do_clean",
    help="OFF ise outlier filtresi uygulanmaz. (Temel temizlik yine de uygulanır.)",
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

if not HAS_CLEANING:
    st.sidebar.warning("data_cleaning.py bulunamadı: fallback temizlik kullanılıyor.")

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

def build_market_curve(sonuc_df: pd.DataFrame, pmin: int, pmax: int, step: int = 1) -> pd.DataFrame:
    """
    Aynı lineer BID/ASK mantığıyla her fiyat seviyesinde toplam arz ve talebi üretir.
    Streamlit + Excel grafiği için kullanılır.
    """
    prices = np.arange(pmin, pmax + step, step)

    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    rows = []
    for p in prices:
        # Demand
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()
            p_bid_arr = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid_arr - pmin, 1e-6)
            frac = 1.0 - (p - pmin) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            demand = 0.0

        # Supply
        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask_arr = sellers["p_ask"].to_numpy()
            denom = np.maximum(pmax - p_ask_arr, 1e-6)
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
    st.warning("Temizleme kapalı: ham veri (temel temizlik sonrası) ile devam ediliyor.")

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
        st.caption(f"Benchmark: Best {benchmark_top_pct}% (by lowest intensity, fuel-based)")

        # Benchmark table
        st.subheader("Benchmark (yakıt bazında)")
        bench_df = (
            pd.DataFrame([{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()])
            .sort_values("FuelType")
            .reset_index(drop=True)
        )
        st.dataframe(bench_df, use_container_width=True)

        # KPIs
        total_cost = float(sonuc_df["ets_cost_total_€"].sum())
        total_revenue = float(sonuc_df["ets_revenue_total_€"].sum())
        net_cashflow = float(sonuc_df["ets_net_cashflow_€"].sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam ETS Maliyeti (€)", f"{total_cost:,.0f}")
        c2.metric("Toplam ETS Geliri (€)", f"{total_revenue:,.0f}")
        c3.metric("Net Nakit Akışı (€)", f"{net_cashflow:,.0f}")

        # Buyers / Sellers
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

        # Curve + Top20
        curve_df = build_market_curve(sonuc_df, price_min, price_max, step=1)

        cashflow_top20 = (
            sonuc_df[["Plant", "FuelType", "ets_net_cashflow_€"]]
            .copy()
            .sort_values("ets_net_cashflow_€", ascending=False)
            .head(20)
        )

        # -------------------------
        # ✅ Streamlit grafikler (daha düzgün)
        # -------------------------
        st.subheader("Grafikler")

        # Supply–Demand (daha “ETS-like” görünüm: step)
        fig1 = plt.figure()
        plt.step(curve_df["Price"], curve_df["Total_Demand"], where="post", label="Total Demand (BID)")
        plt.step(curve_df["Price"], curve_df["Total_Supply"], where="post", label="Total Supply (ASK)")
        plt.axvline(float(clearing_price), linestyle="--", label=f"Carbon Price = {clearing_price:.2f}")
        plt.xlabel("Price (€/tCO₂)")
        plt.ylabel("Volume (tCO₂)")
        plt.title("Market Supply–Demand (Step Curves)")
        plt.legend()
        st.pyplot(fig1, clear_figure=True)

        # Top20 cashflow bar
        fig2 = plt.figure()
        plt.bar(cashflow_top20["Plant"].astype(str), cashflow_top20["ets_net_cashflow_€"])
        plt.xticks(rotation=75, ha="right")
        plt.ylabel("€")
        plt.title("Top 20 Plants – ETS Net Cashflow (€)")
        st.pyplot(fig2, clear_figure=True)

        # -------------------------
        # Excel report + charts
        # -------------------------
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

            if not removed_df.empty:
                removed_df.to_excel(writer, sheet_name="Removed_Outliers", index=False)

            wb = writer.book

            # Excel Supply–Demand (LineChart)
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

            # Excel Cashflow chart
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
