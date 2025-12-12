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
    "slope_bid": 150,
    "slope_ask": 150,
    "spread": 1.0,
    "do_clean": False,
    "lower_pct": 1.0,
    "upper_pct": 2.0,
}


st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.write(
    """
### ETS Geliştirme Modülü V001 — Ne yapar?

Bu arayüz, tek bir Excel dosyasıyla **çok yakıtlı termik santraller için ETS** (Emisyon Ticaret Sistemi) simülasyonu yapar ve sonuçları **Excel rapor + grafik** olarak indirmenizi sağlar.

#### 1) Veri okuma ve birleştirme
- Excel dosyasındaki **tüm sekmeleri** okur.
- Her sekme adı otomatik olarak **FuelType** kabul edilir.
- Tüm sekmeler **tek bir veri setinde (DataFrame)** birleştirilir.
- Beklenen kolonlar: **Plant, Generation_MWh, Emissions_tCO2**.

#### 2) Benchmark hesaplama (yakıt bazında)
- Her **FuelType** için benchmark (B_fuel) hesaplanır.
- **Benchmark Settings (Best plants %)** ile benchmark’ın hangi “en iyi” dilimden hesaplanacağı seçilir:
  - **100%** → o yakıttaki **tüm tesisler** benchmark’a girer (varsayılan/nötr yaklaşım).
  - **10–90%** → intensity’si (Emissions/Generation) en düşük olan “en iyi” tesislerden başlayarak,
    toplam üretimin seçilen yüzdesi dolana kadar alınır ve **üretim ağırlıklı** benchmark hesaplanır.

#### 3) AGK ile tahsis yoğunluğu (Just Transition)
- Her tesisin gerçek emisyon yoğunluğu: **Iᵢ = Emissions_tCO2 / Generation_MWh**
- Yakıt benchmark’ı: **B_fuel**
- Tahsis yoğunluğu (Allocation Intensity) şu şekilde hesaplanır:
  - **Tᵢ = Iᵢ + AGK × (B_fuel − Iᵢ)**
- Yorum:
  - **AGK = 1.0** → Tᵢ tamamen **benchmark’a** eşitlenir (daha sıkı/benchmark bazlı yaklaşım).
  - **AGK = 0.0** → Tᵢ tamamen **tesis yoğunluğuna** yaklaşır (daha yumuşak/tesise yakın yaklaşım).

#### 4) Ücretsiz tahsis ve net ETS pozisyonu
- Ücretsiz tahsis: **FreeAllocᵢ = Generation_MWh × Tᵢ**
- Net ETS pozisyonu:
  - **NetETSᵢ = Emissions_tCO2 − FreeAllocᵢ**
  - **NetETS > 0** → tesis **alıcıdır** (yükümlülük)
  - **NetETS < 0** → tesis **satıcıdır** (fazla tahsis)

#### 5) Tek piyasa clearing price (tüm tesisler birlikte)
- Tüm tesisler **tek bir piyasada** toplanır.
- Her tesis için **BID/ASK fiyatları** üretilir ve toplam arz-talep eğrileri ile **clearing price** bulunur.
- Clearing price, **Carbon Price Range** içinde (min–max) bulunur ve tüm tesisler için **aynı fiyat** kullanılır.

#### 6) Maliyet ve gelir hesapları
- Alıcılar (NetETS>0): **Cost = NetETS × Price**
- Satıcılar (NetETS<0): **Revenue = |NetETS| × Price**
- Ayrıca €/MWh bazında maliyet/gelir ve net nakit akışı raporlanır.

#### 7) Veri temizleme (opsiyonel)
- **Apply cleaning rules?** kapalıysa temizleme yapılmaz.
- Açık ise:
  - Temel temizlik uygulanır (sayı olmayan değerler, eksikler vb. düzeltilir/elenir).
  - “Intensity outlier” filtresi ile, yakıt bazlı benchmark’a göre band dışındaki tesisler veriden çıkarılabilir.
  - Band parametreleri:
    - Alt sınır: **lo = B × (1 − L)**
    - Üst sınır: **hi = B × (1 + U)**

#### 8) Raporlama / çıktı
- Sonuçlar ekranda tablo olarak gösterilir:
  - Benchmark tablosu
  - Alıcılar / Satıcılar
  - Tüm tesis sonuçları
- Ayrıca tek tuşla **Excel raporu** indirilir:
  - Summary, Benchmarks, All_Plants, Buyers, Sellers
  - Market_Curve (arz-talep)
  - Cashflow_Top20
  - Grafikler: Supply–Demand eğrisi ve Top 20 cashflow bar grafiği
---

### Slider’lar neyi değiştirir?

**Carbon Price Range (€/tCO₂)**
- Clearing price aramasının yapılacağı minimum ve maksimum fiyat aralığını belirler.
- Piyasa fiyatı bu bandın dışına çıkamaz.

**AGK (Just Transition Coefficient)**
- Tahsis yoğunluğunu benchmark’a yaklaştırma derecesini belirler.
- AGK artarsa tahsis benchmark’a yaklaşır; azalırsa tesisin kendi yoğunluğuna yaklaşır.

**Benchmark = Best plants %**
- Yakıt bazlı benchmark’ın hangi “en iyi” dilimden hesaplanacağını belirler.
- Daha düşük yüzde → daha sıkı benchmark (genellikle daha yüksek yükümlülük).

**Bid Slope (β_bid)**
- “Kirli” tesislerin ödeme isteğinin (bid) intensity farkına duyarlılığını belirler.
- Artarsa bid’ler daha ayrışır (talep davranışı daha keskinleşir).

**Ask Slope (β_ask)**
- “Temiz” tesislerin satış isteğinin (ask) intensity farkına duyarlılığını belirler.
- Artarsa ask’ler daha ayrışır (arz davranışı daha keskinleşir).

**Bid/Ask Spread**
- Bid ve ask fiyatları arasına sabit bir ayrım ekler.
- Piyasa mikro yapısını daha gerçekçi yapar (bid/ask üst üste binmesini azaltır).

**Apply cleaning rules?**
- Açık ise outlier filtresi uygulanır; kapalı ise ham veriyle devam edilir.

**L ve U (Outlier band)**
- Temizleme açıksa benchmark etrafında izin verilen intensity bandını belirler.
- U büyürse üst band genişler (daha az tesis elenir), L büyürse alt band gevşer (0’a yaklaşır).
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
    help="AGK yönü: AGK=1→Benchmark, AGK=0→Tesis yoğunluğu (T_i = I + AGK*(B - I))",
)
st.sidebar.caption("Default: AGK = 1.00")

st.sidebar.subheader("Benchmark Settings")
benchmark_top_pct = st.sidebar.select_slider(
    "Benchmark = Best plants (by intensity) %",
    options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    key="benchmark_top_pct",
    help="Yakıt bazında benchmark, intensity düşük olan en iyi dilimden (production-share) hesaplanır. 100 = tüm tesisler.",
)
st.sidebar.caption("Default: 100")

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
    help="0 bırakabilirsin. Spread eklemek bid/ask aynı görünmesini azaltır.",
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
    help="Kapalıysa (Hayır), veri temizleme/outlier filtresi uygulanmaz.",
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
# Cleaning (basic always + optional outlier)
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
    if len(removed_df) > 0:
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
        )

        st.success(f"Clearing Price: {clearing_price:.2f} €/tCO₂")
        st.caption(f"Benchmark method: Best {benchmark_top_pct}% (production-share, by lowest intensity)")

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

        curve_df = build_market_curve(sonuc_df, price_min, price_max, step=1)

        cashflow_top20 = (
            sonuc_df[["Plant", "FuelType", "ets_net_cashflow_€"]]
            .copy()
            .sort_values("ets_net_cashflow_€", ascending=False)
            .head(20)
        )

        # -------------------------
        # Excel report + charts
        # -------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df = pd.DataFrame(
                {
                    "Metric": [
                        "Clearing Price (€/tCO₂)",
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

            line.add_data(
                Reference(ws_curve, min_col=4, min_row=1, max_row=ws_curve.max_row),
                titles_from_data=True
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
