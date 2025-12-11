import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

st.write(
    "Bu arayüz: Excel'deki tüm sekmeleri okuyarak yakıt türüne göre benchmark hesaplar, "
    "Adil Geçiş Katsayısı (AGK) ile tahsis yoğunluğunu belirler ve tüm tesisler için birleşik ETS piyasasında tek bir karbon fiyatı üretir."
)

# -----------------------------
# PARAMETRELER
# -----------------------------
st.sidebar.header("Model Parameters")

price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=(0, 100),
    step=1,
    help="ETS clearing price bu aralıkta aranacak. Örnek: 0–100 €/tCO₂"
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=0.50,
    step=0.05,
    help="Tahsis Yoğunluğuᵢ = B_yakıt + AGK × (Iᵢ − B_yakıt)."
)

# -----------------------------
# EXCEL YÜKLEME
# -----------------------------
uploaded_file = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Lütfen bir Excel dosyası yükleyin.")
else:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = []

        for sheet in xls.sheet_names:
            df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)
            df_sheet["FuelType"] = sheet
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
            )
        except Exception as e:
            st.error(f"Model çalışırken hata oluştu: {e}")
            st.stop()

        st.success(f"Market Clearing Price: {clearing_price:.2f} €/tCO₂")

        st.subheader("Fuel-Type Benchmark Intensities (B_yakıt)")
        bench_rows = [{"FuelType": ft, "B_yakıt (tCO₂/MWh)": val} for ft, val in benchmark_map.items()]
        bench_df = pd.DataFrame(bench_rows)
        st.table(bench_df)

        st.subheader("Plant-Level ETS Results")
        st.dataframe(sonuc_df, use_container_width=True)

        csv = sonuc_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download results as CSV",
            data=csv,
            file_name="ets_results.csv",
            mime="text/csv",
        )
