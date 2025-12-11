import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

st.write(
    "Bu arayüz: Excel'deki tüm sekmeleri okuyarak yakıt türüne göre benchmark hesaplar, "
    "Adil Geçiş Katsayısı (AGK) ile tahsis yoğunluğunu belirler ve tüm tesisler için birleşik ETS piyasasında tek bir karbon fiyatı üretir."
)

# --- PARAMETRELER ---
st.sidebar.header("Model Parameters")

# 1) Karbon fiyat aralığı (minimum–maksimum)
price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=(0, 100),
    step=1,
    help="ETS clearing price bu aralıkta aranacak. Örnek: 0–100 €/tCO₂"
)

# 2) AGK katsayısı
agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=0.50,
    step=0.05,
    help="Tahsis Yoğunluğuᵢ = B_yakıt + AGK × (Iᵢ − B_yakıt). 0 → saf benchmark, 1 → santral yoğunluğu."
)

# 3) Excel dosyası yükleme
uploaded_file = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Lütfen bir Excel dosyası yükleyin.")
else:
    # --- TÜM SEKME VERİLERİNİ OKU VE BİRLEŞTİR ---
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = []

        for sheet in xls.sheet_names:
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)
            df_sheet["FuelType"] = sheet  # Sheet adı yakıt türü
            all_sheets.append(df_sheet)

        df_all = pd.concat(all_sheets, ignore_index=True)

    except Exception as e:
       st.error(f"Excel okunurken hata oluştu: {e}")
