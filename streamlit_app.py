# streamlit_app.py
import streamlit as st
import pandas as pd
from ets_model import ets_hesapla

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")
st.markdown(
    """
Bu arayüz ile:

- Excel dosyasından santral verilerini yükleyebilir,
- Tavan fiyat, AGK oranı ve α (benchmark smoothing) parametrelerini ayarlayabilir,
- Santral bazında ETS maliyetlerini hesaplayabilirsiniz.
"""
)

# --- Yan panel: Parametreler ---
st.sidebar.header("Model Parametreleri")

tavan_fiyat = st.sidebar.number_input(
    "Tavan Fiyat (€/tCO₂)", min_value=0.0, max_value=1000.0, value=15.0, step=1.0
)

agk_orani = st.sidebar.number_input(
    "AGK Oranı (ör. 0.15)", min_value=0.0, max_value=1.0, value=0.15, step=0.01
)

alpha = st.sidebar.slider(
    "Benchmark Smoothing Katsayısı (α)", min_value=0.0, max_value=1.0, value=0.75, step=0.05
)

st.sidebar.markdown("---")
st.sidebar.write("Excel'deki beklenen minimum kolonlar:")
st.sidebar.code("Plant, Emissions_tCO2, Generation_MWh", language="text")

# --- Dosya yükleme ---
uploaded_file = st.file_uploader(
    "Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Excel dosyası okunurken hata oluştu: {e}")
        st.stop()

    st.subheader("Yüklenen Ham Veri")
    st.dataframe(df, use_container_width=True)

    if st.button("Modeli Çalıştır"):
        try:
            sonuc_df = ets_hesapla(df, tavan_fiyat, agk_orani, alpha)
        except Exception as e:
            st.error(f"Model çalışırken hata oluştu: {e}")
        else:
            st.success("Hesaplama tamamlandı.")
            st.subheader("Sonuç Tablosu")
            st.dataframe(sonuc_df, use_container_width=True)

            # CSV indirme
            csv = sonuc_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Sonuçları CSV olarak indir",
                data=csv,
                file_name="ets_sonuc.csv",
                mime="text/csv",
            )
else:
    st.info("Lütfen önce bir Excel dosyası yükleyin.")
