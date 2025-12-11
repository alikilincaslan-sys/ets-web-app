import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001 — Fuel-Type Benchmark + Unified Market Clearing")

st.markdown("""
Bu model:
- Excel'deki **tüm sekmeleri** otomatik okur,
- Her santral için `FuelType` bilgisi ekler (sheet adı olarak),
- **Yakıt türüne göre ayrı benchmark** hesaplar,
- Tüm santraller birlikte **tek bir ETS piyasasında** clearing price oluşturur,
- Her santrale **aynı piyasa fiyatını** uygular.
""")

# --- Parametreler ---
st.sidebar.header("Model Parametreleri")

tavan_fiyat = st.sidebar.number_input(
    "Tavan Fiyat (€/tCO₂)", min_value=0.0, max_value=200.0, value=50.0, step=1.0
)

agk_orani = st.sidebar.number_input(
    "AGK Oranı (ör: 0.15)", min_value=0.0, max_value=1.0, value=0.15, step=0.01
)

alpha = st.sidebar.slider(
    "Benchmark Smoothing (α)", min_value=0.0, max_value=1.0, value=0.75, step=0.05
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

    st.subheader("Tüm Tesisler (Bütün Sekmeler Birleştirildi)")
    st.dataframe(df_all, use_container_width=True)

    if st.button("Modeli Çalıştır"):
        try:
            sonuc_df, benchmark_map, clearing_price = ets_hesapla(
                df_all, tavan_fiyat, agk_orani, alpha
            )
        except Exception as e:
            st.error(f"Model çalışırken hata oluştu: {e}")
            st.stop()

        # --- Clearing Price ---
        st.success(f"**Piyasa Clearing Price: {clearing_price:.2f} €/tCO₂**")

        # --- Fuel-Type Benchmark Tablosu ---
        st.subheader("Yakıt Türü Bazında Benchmark Değerleri")
        bench_df = pd.DataFrame([
            {"FuelType": ft, "Benchmark (tCO₂/MWh)": val}
            for ft, val in benchmark_map.items()
        ])
        st.table(bench_df)

        # --- Sonuç Tablosu ---
        st.subheader("Santral Bazında ETS Sonuçları")
        st.dataframe(sonuc_df, use_container_width=True)

        # CSV indir
        csv = sonuc_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Sonuçları CSV olarak indir",
            data=csv,
            file_name="ets_sonuc.csv",
            mime="text/csv",
        )

else:
    st.info("Lütfen bir Excel yükleyin.")
