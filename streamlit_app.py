import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

from ets_model import ets_hesapla

# -------------------------------------------------
# Sayfa ayarı
# -------------------------------------------------
st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

st.write(
    """
### ETS Geliştirme Modülü – Model Açıklaması

Bu arayüz:

- Excel dosyasındaki **tüm sekmeleri** okur ve birleştirir (FuelType = sekme adı),
- Yakıt türüne göre **üretim-ağırlıklı benchmark** hesaplar,
- **AGK (Adil Geçiş Katsayısı)** ile tahsis yoğunluğunu belirler,
- Tüm tesisleri **tek ETS piyasasında** birleştirerek **BID / ASK** eğrileri üzerinden **clearing price** üretir,
- **AGK değiştikçe benchmark etkisini** santral bazında veya tüm sistem için görselleştirir.
"""
)

# -------------------------------------------------
# Sidebar – Parametreler
# -------------------------------------------------
st.sidebar.header("Model Parameters")

price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    0, 200, (5, 20), step=1
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    0.0, 1.0, 1.0, step=0.05,
    help="AGK=1 → benchmark, AGK=0 → tesis yoğunluğu"
)

benchmark_top_pct = st.sidebar.slider(
    "Benchmark – Best X% Plants",
    10, 100, 100, step=10,
    help="Benchmark hesabında en iyi X% tesis kullanılır"
)

slope_bid = st.sidebar.slider("β_bid", 50, 300, 150, step=10)
slope_ask = st.sidebar.slider("β_ask", 50, 300, 150, step=10)
spread = st.sidebar.slider("Bid/Ask Spread", 0.0, 5.0, 1.0, step=0.5)

# -------------------------------------------------
# Excel yükleme
# -------------------------------------------------
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])

def read_all_sheets(file):
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df["FuelType"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True)

if uploaded is None:
    st.info("Lütfen bir Excel dosyası yükleyin.")
    st.stop()

df_all = read_all_sheets(uploaded)

st.subheader("Yüklenen veri (ilk 50 satır)")
st.dataframe(df_all.head(50), use_container_width=True)

# -------------------------------------------------
# Model çalıştır
# -------------------------------------------------
if st.button("Run ETS Model"):

    sonuc_df, benchmark_map, clearing_price = ets_hesapla(
        df_all,
        price_min,
        price_max,
        agk,
        slope_bid=slope_bid,
        slope_ask=slope_ask,
        spread=spread,
        benchmark_top_pct=benchmark_top_pct
    )

    st.success(f"Clearing Price = {clearing_price:.2f} €/tCO₂")

    # -------------------------------------------------
    # AGK – Benchmark Karşılaştırma Grafiği
    # -------------------------------------------------
    st.subheader("AGK – Benchmark Etkisi (Sade Grafik)")

    # Sistem üretim-ağırlıklı benchmark
    tmp = sonuc_df.copy()
    tmp["intensity"] = tmp["Emissions_tCO2"] / tmp["Generation_MWh"]

    system_benchmark = (
        tmp.groupby("FuelType")
        .apply(lambda g: g["Emissions_tCO2"].sum() / g["Generation_MWh"].sum())
        .to_dict()
    )

    plant_options = ["Tüm Santraller"] + sorted(sonuc_df["Plant"].unique())
    selected_plant = st.selectbox("Santral seç", plant_options)

    agk_grid = np.round(np.arange(0.0, 1.01, 0.05), 2)

    if selected_plant == "Tüm Santraller":
        I = tmp["Emissions_tCO2"].sum() / tmp["Generation_MWh"].sum()
        B_current = np.average(tmp["B_fuel"], weights=tmp["Generation_MWh"])
        B_prod = np.average(
            tmp["FuelType"].map(system_benchmark),
            weights=tmp["Generation_MWh"]
        )
    else:
        row = sonuc_df[sonuc_df["Plant"] == selected_plant].iloc[0]
        I = row["Emissions_tCO2"] / row["Generation_MWh"]
        B_current = row["B_fuel"]
        B_prod = system_benchmark[row["FuelType"]]

    curve = pd.DataFrame({"AGK": agk_grid})
    curve["Tahsis_Yogunlugu"] = I + curve["AGK"] * (B_current - I)

    base = alt.Chart(curve).mark_line(color="#4cc9f0").encode(
        x="AGK",
        y="Tahsis_Yogunlugu",
        tooltip=["AGK", alt.Tooltip("Tahsis_Yogunlugu", format=".4f")]
    )

    point_selected = alt.Chart(
        pd.DataFrame({"AGK": [agk], "T": [I + agk * (B_current - I)]})
    ).mark_point(size=80, color="orange").encode(
        x="AGK", y="T"
    )

    point_agk1 = alt.Chart(
        pd.DataFrame({"AGK": [1.0], "T": [B_current]})
    ).mark_point(size=80, color="red").encode(
        x="AGK", y="T"
    )

    hline = alt.Chart(
        pd.DataFrame({"y": [B_prod]})
    ).mark_rule(strokeDash=[6, 4], color="white").encode(
        y="y"
    )

    st.altair_chart(base + point_selected + point_agk1 + hline, use_container_width=True)

    st.caption(
        "Mavi çizgi: AGK’ye bağlı tahsis yoğunluğu | "
        "Turuncu nokta: seçili AGK | "
        "Kırmızı nokta: AGK=1 | "
        "Beyaz kesikli çizgi: üretim-ağırlıklı benchmark"
    )

    # -------------------------------------------------
    # Sonuç tablosu
    # -------------------------------------------------
    st.subheader("Tüm Sonuçlar (ham tablo)")
    st.dataframe(sonuc_df, use_container_width=True)
