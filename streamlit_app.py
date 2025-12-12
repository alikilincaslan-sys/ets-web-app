import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

# Plotly (grafikler için) – yoksa uygulama çökmesin
try:
    import plotly.graph_objects as go
    PLOTLY_OK = True
except Exception:
    PLOTLY_OK = False

from ets_model import ets_hesapla

# ✅ Temizleme modülü opsiyonel (repo'da data_cleaning.py varsa kullan)
CLEANING_OK = True
try:
    from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel
except Exception:
    CLEANING_OK = False

# -------------------------
# Helpers
# -------------------------
def read_all_sheets(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df["FuelType"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def build_market_curve(sonuc_df: pd.DataFrame, price_min: int, price_max: int, step: int = 1) -> pd.DataFrame:
    """
    Aynı lineer BID/ASK mantığıyla her fiyat seviyesinde toplam arz ve talebi üretir.
    (Uygulama içi & Excel Supply–Demand grafiği için)
    """
    prices = np.arange(price_min, price_max + step, step)

    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    rows = []
    for p in prices:
        # Demand
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()
            p_bid = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid - price_min, 1e-6)
            frac = 1.0 - (p - price_min) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            demand = 0.0

        # Supply
        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask = sellers["p_ask"].to_numpy()
            denom = np.maximum(price_max - p_ask, 1e-6)
            frac = (p - p_ask) / denom
            supply = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            supply = 0.0

        rows.append({"Price": float(p), "Total_Demand": demand, "Total_Supply": supply})

    return pd.DataFrame(rows)


def build_step_curves(sonuc_df: pd.DataFrame):
    """
    AB ETS step curve benzeri gösterim için:
    - Demand: p_bid azalan sırada birikimli miktar
    - Supply: p_ask artan sırada birikimli miktar
    """
    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    if not buyers.empty:
        buyers = buyers.sort_values("p_bid", ascending=False)
        buyers["q"] = buyers["net_ets"]
        buyers["cum_q"] = buyers["q"].cumsum()
    else:
        buyers = pd.DataFrame(columns=["p_bid", "cum_q"])

    if not sellers.empty:
        sellers = sellers.sort_values("p_ask", ascending=True)
        sellers["q"] = (-sellers["net_ets"])
        sellers["cum_q"] = sellers["q"].cumsum()
    else:
        sellers = pd.DataFrame(columns=["p_ask", "cum_q"])

    return buyers, sellers


def compute_allocation_intensity(intensity: pd.Series, b_fuel: pd.Series, agk: float) -> pd.Series:
    """
    AGK yönü:
    AGK=1 => Benchmark (B)
    AGK=0 => Tesis yoğunluğu (I)
    T = I + AGK*(B - I)
    """
    return intensity + agk * (b_fuel - intensity)


# -------------------------
# Page
# -------------------------
st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

st.write(
    """
Bu arayüz:

- Excel dosyasındaki **tüm sekmeleri** okur ve birleştirir (**FuelType = sekme adı**).
- Seçilen **Benchmark Top %** kuralına göre (örn. %10, %20, …, %100) yakıt bazında **üretim-ağırlıklı benchmark** hesaplar.
- **AGK (Adil Geçiş Katsayısı)** ile tesisin ücretsiz tahsis için kullanılacak **tahsis yoğunluğunu** hesaplar  
  (**AGK=1 → Benchmark**, **AGK=0 → tesisin kendi yoğunluğu**).
- Tüm tesisleri **tek ETS piyasasında** birleştirir ve seçilen fiyat yöntemiyle **clearing price** hesaplar:
  - **Market Clearing**: AB ETS mantığıyla arz (ASK) ve talebin (BID) kesişimi.
  - **Average Compliance Cost (ACC)**: uyum maliyeti temelli ortalama fiyat yaklaşımı.
- (Opsiyonel) **Veri temizleme** uygular (temizleme modülü varsa).
- Sonuçları **Excel rapor + grafik** (ve CSV) olarak indirir.
"""
)

# -------------------------
# Defaults + Reset
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

def reset_defaults():
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

st.sidebar.header("Model Parameters")
if st.sidebar.button("Reset to Defaults"):
    reset_defaults()

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
    help=f"Default: {DEFAULTS['agk']}. AGK=1→Benchmark, AGK=0→Tesis yoğunluğu. T = I + AGK*(B - I)",
)

benchmark_top_pct = st.sidebar.slider(
    "Benchmark Top % (best performers)",
    min_value=10,
    max_value=100,
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    step=10,
    key="benchmark_top_pct",
    help=f"Default: {DEFAULTS['benchmark_top_pct']}. %10 = en iyi %10 (yakıt içinde), %100 = tüm tesisler.",
)

price_method = st.sidebar.selectbox(
    "Price Method",
    options=["Market Clearing", "Average Compliance Cost (ACC)"],
    index=["Market Clearing", "Average Compliance Cost (ACC)"].index(
        st.session_state.get("price_method", DEFAULTS["price_method"])
    ),
    key="price_method",
    help="Default: Market Clearing. Tek seçim: karışıklık olmasın diye.",
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
    help=f"Default: {DEFAULTS['spread']}. Spread eklemek bid/ask aynı görünmesini azaltır.",
)

# -------------------------
# Data Cleaning Controls
# -------------------------
st.sidebar.subheader("Data Cleaning")

if not CLEANING_OK:
    st.sidebar.info("data_cleaning.py bulunamadı / import edilemedi. Temizleme otomatik kapalı çalışır.")
    st.session_state["do_clean"] = False

do_clean = st.sidebar.toggle(
    "Apply cleaning rules?",
    value=bool(st.session_state.get("do_clean", DEFAULTS["do_clean"])) if CLEANING_OK else False,
    key="do_clean",
    help=f"Default: {'ON' if DEFAULTS['do_clean'] else 'OFF'}. Kapalıysa ham veriyle devam eder.",
    disabled=not CLEANING_OK,
)

lower_pct = st.sidebar.slider(
    "Lower bound vs Benchmark (L)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("lower_pct", DEFAULTS["lower_pct"])),
    step=0.05,
    key="lower_pct",
    help=f"Default: {DEFAULTS['lower_pct']}. 1.0 => alt sınır 0 (B*(1-1)=0). 0.5 => alt sınır 0.5B.",
    disabled=not CLEANING_OK,
)

upper_pct = st.sidebar.slider(
    "Upper bound vs Benchmark (U)",
    min_value=0.0,
    max_value=2.0,
    value=float(st.session_state.get("upper_pct", DEFAULTS["upper_pct"])),
    step=0.05,
    key="upper_pct",
    help=f"Default: {DEFAULTS['upper_pct']}. 1.0 => üst sınır 2B. 2.0 => üst sınır 3B.",
    disabled=not CLEANING_OK,
)

st.sidebar.divider()
st.sidebar.caption("Excel kolonları: Plant, Generation_MWh, Emissions_tCO2 (FuelType sekme adından gelir)")

# -------------------------
# Excel upload
# -------------------------
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

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
st.subheader("Veri Temizleme (opsiyonel)")
df_all = df_all_raw.copy()

clean_report_df = None
outlier_report = None

if do_clean and CLEANING_OK:
    cleaned_frames = []
    reports_basic = []

    for ft in df_all["FuelType"].unique():
        part = df_all[df_all["FuelType"] == ft].copy()
        cleaned, rep = clean_ets_input(part, fueltype=ft)
        rep["FuelType"] = ft
        reports_basic.append(rep)
        cleaned_frames.append(cleaned)

    df_clean = pd.concat(cleaned_frames, ignore_index=True)
    clean_report_df = pd.DataFrame(reports_basic)

    st.write("Temel temizlik özeti (sekme bazında):")
    st.dataframe(clean_report_df, use_container_width=True)

    before = len(df_clean)
    df_clean2, rep_out = filter_intensity_outliers_by_fuel(df_clean, lower_pct=lower_pct, upper_pct=upper_pct)
    after = len(df_clean2)
    outlier_report = rep_out

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
else:
    st.warning("Temizleme kapalı: ham veriyle devam ediliyor.")

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
            price_method="Average Compliance Cost" if "ACC" in price_method else "Market Clearing",
        )

        st.success(f"Clearing Price: {clearing_price:.2f} €/tCO₂  |  Method: {price_method}")

        # Benchmarks
        st.subheader("Benchmark (yakıt bazında)")
        bench_df = (
            pd.DataFrame([{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()])
            .sort_values("FuelType")
            .reset_index(drop=True)
        )
        st.dataframe(bench_df, use_container_width=True)

        # KPI
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
                ["Plant","FuelType","net_ets","carbon_price","ets_cost_total_€","ets_cost_€/MWh","ets_net_cashflow_€","ets_net_cashflow_€/MWh"]
            ],
            use_container_width=True,
        )

        st.subheader("ETS Sonuçları – Satıcılar (Net ETS < 0)")
        sellers_df = sonuc_df[sonuc_df["net_ets"] < 0].copy()
        st.dataframe(
            sellers_df[
                ["Plant","FuelType","net_ets","carbon_price","ets_revenue_total_€","ets_revenue_€/MWh","ets_net_cashflow_€","ets_net_cashflow_€/MWh"]
            ],
            use_container_width=True,
        )

        # Raw table
        st.subheader("Tüm Sonuçlar (ham tablo)")
        st.dataframe(sonuc_df, use_container_width=True)

        # -------------------------
        # Market Charts (app)
        # -------------------------
        st.subheader("📈 Piyasa Grafikleri (Uygulama içi)")

        curve_df = build_market_curve(sonuc_df, price_min, price_max, step=1)
        step_buy, step_sell = build_step_curves(sonuc_df)

        if PLOTLY_OK:
            # Smooth
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=curve_df["Price"], y=curve_df["Total_Demand"], mode="lines", name="Demand"))
            fig1.add_trace(go.Scatter(x=curve_df["Price"], y=curve_df["Total_Supply"], mode="lines", name="Supply"))
            fig1.add_vline(x=float(clearing_price), line_width=2)
            fig1.update_layout(
                title="Supply–Demand (Smooth) + Clearing Price",
                xaxis_title="Price (€/tCO₂)",
                yaxis_title="Volume (tCO₂)",
                height=380,
            )
            st.plotly_chart(fig1, use_container_width=True)

            # Step
            fig2 = go.Figure()
            if not step_buy.empty:
                fig2.add_trace(go.Scatter(x=step_buy["p_bid"], y=step_buy["cum_q"], mode="lines", line_shape="hv", name="Demand (step)"))
            if not step_sell.empty:
                fig2.add_trace(go.Scatter(x=step_sell["p_ask"], y=step_sell["cum_q"], mode="lines", line_shape="hv", name="Supply (step)"))
            fig2.add_vline(x=float(clearing_price), line_width=2)
            fig2.update_layout(
                title="AB ETS Step Curves (Bids/Asks) + Clearing Price",
                xaxis_title="Price (€/tCO₂)",
                yaxis_title="Cumulative Volume (tCO₂)",
                height=380,
            )
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Plotly yok. Basit çizim için st.line_chart kullanılıyor.")
            st.line_chart(curve_df.set_index("Price")[["Total_Demand","Total_Supply"]])

        # -------------------------
        # AGK – Benchmark Etkisi (Sade Grafik)
        # -------------------------
        st.subheader("📉 AGK – Benchmark Etkisi (Sade Grafik)")

        # Zorunlu kolonlar (ets_model çıktısında zaten var)
        if not all(c in sonuc_df.columns for c in ["Plant", "FuelType", "intensity", "B_fuel", "tahsis_intensity"]):
            st.warning("Bu grafik için gerekli kolonlar yok: Plant, FuelType, intensity, B_fuel, tahsis_intensity")
        else:
            # Santral seçimi (Tüm Santraller dahil)
            plants = sorted(sonuc_df["Plant"].astype(str).unique().tolist())
            plant_choice = st.selectbox(
                "Santral seç",
                ["Tüm Santraller"] + plants,
                index=0,
                key="agk_benchmark_plant_choice",
            )

            # Tesis bazında değerler
            df_plot = sonuc_df[["Plant","FuelType","intensity","B_fuel","tahsis_intensity"]].copy()
            df_plot["tahsis_intensity_agk"] = df_plot["tahsis_intensity"]
            df_plot["tahsis_intensity_agk1"] = compute_allocation_intensity(df_plot["intensity"], df_plot["B_fuel"], 1.0)  # = B_fuel

            if plant_choice == "Tüm Santraller":
                # daha dengeli görünüm için sıralama (yoğunluğa göre)
                df_plot = df_plot.sort_values("intensity", ascending=True).reset_index(drop=True)

                if PLOTLY_OK:
                    figA = go.Figure()
                    figA.add_trace(go.Scatter(x=df_plot["Plant"], y=df_plot["intensity"], mode="lines", name="Yoğunluk I (tCO₂/MWh)"))
                    figA.add_trace(go.Scatter(x=df_plot["Plant"], y=df_plot["B_fuel"], mode="lines", name="Benchmark B (tCO₂/MWh)"))
                    figA.add_trace(go.Scatter(x=df_plot["Plant"], y=df_plot["tahsis_intensity_agk"], mode="lines", name=f"Tahsis Yoğunluğu T(AGK={agk:.2f})"))
                    figA.update_layout(
                        title="Tüm Santraller – I vs B vs T(AGK)",
                        xaxis_title="Santral (Yoğunluğa göre sıralı)",
                        yaxis_title="tCO₂/MWh",
                        height=520,
                    )
                    st.plotly_chart(figA, use_container_width=True)
                else:
                    st.line_chart(df_plot.set_index("Plant")[["intensity","B_fuel","tahsis_intensity_agk"]])

                st.caption("Not: AGK=1 çizgisi (T(AGK=1)) = Benchmark B ile aynı olduğu için ayrıca çizdirilmedi (aynı seri).")

            else:
                row = df_plot[df_plot["Plant"].astype(str) == str(plant_choice)].iloc[0]
                one = pd.DataFrame(
                    {
                        "Series": ["Yoğunluk I", "Benchmark B (=AGK1)", f"Tahsis T(AGK={agk:.2f})"],
                        "Value": [float(row["intensity"]), float(row["B_fuel"]), float(row["tahsis_intensity_agk"])],
                    }
                )

                cL, cR = st.columns([1, 2])
                with cL:
                    st.write("Seçili santral – sayısal değerler")
                    st.dataframe(one, use_container_width=True)

                with cR:
                    if PLOTLY_OK:
                        figB = go.Figure()
                        figB.add_trace(go.Scatter(x=one["Series"], y=one["Value"], mode="lines+markers"))
                        figB.update_layout(
                            title=f"{plant_choice} – I vs B vs T(AGK)",
                            xaxis_title="",
                            yaxis_title="tCO₂/MWh",
                            height=420,
                        )
                        st.plotly_chart(figB, use_container_width=True)
                    else:
                        st.line_chart(one.set_index("Series")["Value"])

        # -------------------------
        # Excel report (+charts)
        # -------------------------
        cashflow_top20 = (
            sonuc_df[["Plant", "FuelType", "ets_net_cashflow_€"]]
            .copy()
            .sort_values("ets_net_cashflow_€", ascending=False)
            .head(20)
        )

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
            buyers_df.to_excel(writer, sheet_name="Buyers", index=False)
            sellers_df.to_excel(writer, sheet_name="Sellers", index=False)
            curve_df.to_excel(writer, sheet_name="Market_Curve", index=False)
            cashflow_top20.to_excel(writer, sheet_name="Cashflow_Top20", index=False)

            # ✅ AGK Impact sheet (grafik verisi)
            if all(c in sonuc_df.columns for c in ["Plant","FuelType","intensity","B_fuel","tahsis_intensity"]):
                agk_impact = sonuc_df[["Plant","FuelType","intensity","B_fuel","tahsis_intensity"]].copy()
                agk_impact = agk_impact.sort_values("intensity", ascending=True).reset_index(drop=True)
                agk_impact.rename(
                    columns={
                        "intensity": "I_intensity",
                        "B_fuel": "B_benchmark",
                        "tahsis_intensity": f"T_alloc_AGK_{agk:.2f}",
                    },
                    inplace=True,
                )
                agk_impact.to_excel(writer, sheet_name="AGK_Impact", index=False)

            wb = writer.book

            # 1) Supply–Demand Line Chart
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
            ws_curve["D1"] = "Clearing_Price"
            for r in range(2, ws_curve.max_row + 1):
                ws_curve[f"D{r}"] = float(clearing_price)
            line.add_data(Reference(ws_curve, min_col=4, min_row=1, max_row=ws_curve.max_row), titles_from_data=True)
            ws_curve.add_chart(line, "E2")

            # 2) Cashflow bar
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

            # 3) AGK Impact chart (Excel) – basit line (3 seri)
            if "AGK_Impact" in wb.sheetnames:
                ws_ai = wb["AGK_Impact"]
                # Kolonlar: Plant, FuelType, I_intensity, B_benchmark, T_alloc...
                ai_line = LineChart()
                ai_line.title = "AGK Impact – I vs B vs T(AGK)"
                ai_line.y_axis.title = "tCO₂/MWh"
                ai_line.x_axis.title = "Plant (sorted by I)"
                ai_data = Reference(ws_ai, min_col=3, min_row=1, max_col=5, max_row=ws_ai.max_row)
                ai_cats = Reference(ws_ai, min_col=1, min_row=2, max_row=ws_ai.max_row)
                ai_line.add_data(ai_data, titles_from_data=True)
                ai_line.set_categories(ai_cats)
                ai_line.height = 14
                ai_line.width = 30
                ws_ai.add_chart(ai_line, "G2")

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
