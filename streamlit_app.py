import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="ETS AGK Excel Çıktısı", layout="wide")

st.title("AGK (α) Sonuçları — Sadece Excel Çıktısı")

st.markdown(
    """
Bu sayfada **AGK grafiği yoktur**.  
Seçtiğin **AGK (α)** değerleri için **tüm santrallerin** emisyon yoğunluğu/benchmark değerleri hazırlanır,
santraller **düşükten yükseğe** sıralanır ve sonuçlar **Excel** olarak indirilir.
"""
)

# ---------------------------------------------------------------------
# 0) SENİN VERİN: df isimli dataframe, en az şu kolonları içermeli:
#    - plant_name : santral adı
#    - emission_intensity : (örnek) hesaplanan emisyon yoğunluğu (tCO2/MWh vb.)
#
# Eğer senin uygulamada df başka isimdeyse, aşağıdaki satırı kendi df'inle değiştir.
# ---------------------------------------------------------------------
try:
    df  # noqa: F821
except NameError:
    st.warning(
        "Bu dosyada örnek amaçlı bir iskelet var. Uygulamanın ana kısmında oluşturduğun "
        "`df` dataframe'i bu sayfaya/alıma aktarılmalı (plant_name ve emission_intensity içermeli)."
    )
    st.stop()

# ---------------------------------------------------------------------
# 1) AGK (alpha) seçimi
# ---------------------------------------------------------------------
alpha_list = st.multiselect(
    "Excel'de gösterilecek AGK (α) değerlerini seçin",
    options=[0.25, 0.5, 0.75, 0.9, 1.25, 1.5, 2.0],
    default=[0.5, 0.75],
)

if len(alpha_list) == 0:
    st.info("En az bir AGK (α) değeri seçin.")
    st.stop()


# ---------------------------------------------------------------------
# 2) SENİN MODEL HESABIN: compute_intensity_by_alpha
#    Bu fonksiyon mutlaka şu formatta dönmeli:
#      plant_name | intensity
#    intensity: AGK=alpha için santral bazında değer (emisyon yoğunluğu / benchmark)
# ---------------------------------------------------------------------
def compute_intensity_by_alpha(alpha: float) -> pd.DataFrame:
    """
    ÇIKTI:
      plant_name: santral adı
      intensity : AGK (alpha) senaryosuna göre değer
    NOT:
      Aşağıdaki hesap, senin gerçek model fonksiyonunla değiştirilmeli.
    """
    # ---- PLACEHOLDER (örnek) ----
    # Burayı, ETS modülündeki gerçek hesap fonksiyonunla değiştir.
    out = df.groupby("plant_name")["emission_intensity"].mean().reset_index()
    out = out.rename(columns={"emission_intensity": "intensity"})
    # -----------------------------
    return out[["plant_name", "intensity"]]


def build_agk_table(alpha_list_) -> pd.DataFrame:
    """AGK senaryolarını yan yana sütunlayıp (wide format) sıralı tablo üretir."""
    frames = []
    for a in alpha_list_:
        tmp = compute_intensity_by_alpha(a).copy()
        tmp = tmp.rename(columns={"intensity": f"AGK_{a}"})
        frames.append(tmp.set_index("plant_name"))

    out = pd.concat(frames, axis=1).reset_index()

    # Santralleri ilk seçilen AGK sütununa göre düşükten yükseğe sırala
    base_col = f"AGK_{alpha_list_[0]}"
    if base_col in out.columns:
        out = out.sort_values(base_col, ascending=True)

    return out


df_agk_excel = build_agk_table(alpha_list)

st.subheader("Önizleme (Sadece Tablo)")
st.dataframe(df_agk_excel, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------
# 3) Excel üret + indirme butonu
# ---------------------------------------------------------------------
def to_excel_bytes(df_out: pd.DataFrame) -> bytes:
    bio = BytesIO()
    # openpyxl yoksa xlsxwriter'a düş
    engine = "openpyxl"
    try:
        import openpyxl  # noqa: F401
    except Exception:
        engine = "xlsxwriter"

    with pd.ExcelWriter(bio, engine=engine) as writer:
        df_out.to_excel(writer, sheet_name="AGK_SONUC", index=False)
    return bio.getvalue()


excel_bytes = to_excel_bytes(df_agk_excel)

st.download_button(
    label="📥 AGK_SONUC Excel'i indir",
    data=excel_bytes,
    file_name="AGK_SONUC.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
