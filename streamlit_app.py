import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

# Sayfa ayarları
st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

# Başlık
st.title("ETS Geliştirme Modülü V001")

st.write(
    "Bu arayüz: Excel'deki tüm sekmeleri okuyarak yakıt türüne göre benchmark hesaplar, "
    "Adil Geçiş Katsayısı (AGK) ile tahsis yoğunluğunu belirler ve tüm tesisler için "
    "birleşik ETS piyasasında tek bir karbon fiyatı üretir."
)

# -----------------------------
# PARAMETRELER (Sol panel)
# -----------------------------
st.sidebar.header("Model Parameters")

# Karbon fiyat aralığı (min–max)
price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=(0, 100),
    step=1,
    help="ETS clearing price bu aralıkta aranacak. Örnek: 0–100 €/tCO₂"
)

# Adil Geçiş Katsayısı
agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=0.50,
    step=0.05,
    help="Tahsis Yoğunluğuᵢ = B_yakıt + AGK × (Iᵢ − B_yakıt). 1 → saf benchmark, 0 → santral yoğunluğu."
)
st.sidebar.subheader("Market Calibration")

slope_bid = st.sidebar.slider(
    "Bid Slope (β_bid)",
    min_value=10,
    max_value=500,
    value=150,
    step=10,
    help="Alıcıların (kirli tesis) ödeme isteği hassasiyeti. Yükseldikçe p_bid artar."
)

slope_ask = st.sidebar.slider(
    "Ask Slope (β_ask)",
    min_value=10,
    max_value=500,
    value=150,
    step=10,
    help="Satıcıların (temiz tesis) satış isteği hassasiyeti. Yükseldikçe p_ask artar."
)

spread = st.sidebar.slider(
    "Bid/Ask Spread (€/tCO₂)",
    min_value=0.0,
    max_value=10.0,
    value=0.0,
    step=0.5,
    help="İstersen piyasa spread'i ekler. 0 bırakabilirsin."
)

# -----------------------------
# EXCEL YÜKLEME
# -----------------------------
uploaded_file = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Lütfen bir Excel dosyası yükleyin.")
else:
    # Tüm sekmeleri oku ve birleştir
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = []

        for sheet in xls.sheet_names:
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)
            df_sheet["FuelType"] = sheet  # Sheet adı yakıt türü olarak eklendi
            all_sheets.append(df_sheet)

        df_all = pd.concat(all_sheets, ignore_index=True)

    except Exception as e:
        st.error(f"Excel okunurken hata oluştu: {e}")
        st.stop()

    st.subheader("Tüm Tesisler (Birleştirilmiş Veri)")
    st.dataframe(df_all, use_container_width=True)

    # -----------------------------
    # MODEL ÇALIŞTIRMA
    # -----------------------------
    if st.button("Run ETS Model"):
        try:
           sonuc_df, benchmark_map, clearing_price = ets_hesapla(
    df_all,
    price_min,
    price_max,
    agk,
    slope_bid,
    slope_ask,
    spread,
)

        except Exception as e:
            st.error(f"Model çalışırken hata oluştu: {e}")
            st.stop()

        # Clearing Price
        st.success(f"Market Clearing Price: {clearing_price:.2f} €/tCO₂")

        # Benchmark tablosu
        st.subheader("Fuel-Type Benchmark Intensities (B_yakıt)")
        bench_rows = [
            {"FuelType": ft, "B_yakıt (tCO₂/MWh)": val}
            for ft, val in benchmark_map.items()
        ]
        bench_df = pd.DataFrame(bench_rows)
        st.table(bench_df)

        # Sonuçlar
        st.subheader("Plant-Level ETS Results")
        st.dataframe(sonuc_df, use_container_width=True)

        # CSV indirme
        csv = sonuc_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download results as CSV",
            data=csv,
            file_name="ets_results.csv",
            mime="text/csv",
        )
