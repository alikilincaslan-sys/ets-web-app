import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001 — Fuel-Type Benchmark + Unified Market Clearing")

st.markdown("""
Bu model:

- Excel'deki **tüm sekmeleri** otomatik okur,
- Her satıra, geldiği sayfa adını `FuelType` olarak ekler,
- Aynı yakıt türündeki tesisler için üretim ağırlıklı **benchmark yoğunluğu (B_yakıt)** hesaplar,
- **Adil Geçiş Katsayısı (AGK)** ile formüldeki gibi tahsis yoğunluğunu yumuşatır:

> Tahsis Yoğunluğuᵢ = B_yakıt + AGK × (Iᵢ − B_yakıt)

- `Free Allocation Ratio` ile ücretsiz tahsis oranını belirler,
- Tüm tesisleri tek bir ETS piyasasında clearing’e sokar ve **tek bir karbon fiyatı** bulur,
- Her santral için ETS maliyetlerini hesaplar.
""")

# --- Yan panel: Parametreler ---
st.sidebar.header("Model Parameters")

price_cap = st.sidebar.number_input(
    "Carbon Price Cap (€/tCO₂)",
    min_value=0.0,
    max_value=200.0,
    value=50.0,
    step=1.0,
)

free_alloc_ratio = st.sidebar.number_input(
    "Free Allocation Ratio",
    min_value=0.0,
    max_value=1.0,
    value=0.15,
    step=0.01,
    help="0 = full free allocation, 0.15 = 15% reduction on the computed allocation."
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=0.50,
    step=0.05,
    help="Used in: Tahsis Yoğunluğuᵢ = B_yakıt + AGK × (Iᵢ − B_yakıt). 0 → pure fuel benchmark, 1 → plant intensity."
)

uploaded_file = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = []

        for sheet in xls.sheet_names:
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)
            df_sheet["FuelType"] = sheet  # sheet adı yakıt türü olarak
            all_sheets.append(df_sheet)

        df_all = pd.concat(all_sheets, ignore_index=True)
    except Exception as e:
        st.error(f"Excel okunurken hata oluştu: {e}")
        st.stop()

    st.subheader("Tüm Tesisler (Tüm Sekmeler Birleştirildi)")
    st.dataframe(df_all, use_container_width=True)

    if st.button("Run ETS Model"):
        try:
            sonuc_df, benchmark_map, clearing_price = ets_hesapla(
                df_all, price_cap, free_alloc_ratio, a
