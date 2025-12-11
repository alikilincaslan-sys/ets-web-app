import streamlit as st
import pandas as pd

from ets_model import ets_hesapla


st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")
st.write(
    """
Bu arayüz:
- Excel dosyasındaki **tüm sekmeleri** okur ve birleştirir (FuelType = sekme adı),
- Yakıt türüne göre benchmark hesaplar,
- AGK ile tahsis yoğunluğunu belirler,
- Tüm tesisleri tek piyasada birleştirip **BID/ASK** eğrileriyle **clearing price** üretir.
"""
)

# -------------------------
# Sidebar: Parametreler
# -------------------------
st.sidebar.header("Model Parameters")

price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=(0, 20),
    step=1,
    help="Clearing price bu aralık içinde bulunur."
)

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=0.50,
    step=0.05,
    help="T_i = B + AGK*(I - B)"
)

st.sidebar.subheader("Market Calibration")

slope_bid = st.sidebar.slider(
    "Bid Slope (β_bid)",
    min_value=10,
    max_value=500,
    value=150,
    step=10,
    help="Alıcıların (kirli tesis) ödeme isteği hassasiyeti."
)

slope_ask = st.sidebar.slider(
    "Ask Slope (β_ask)",
    min_value=10,
    max_value=500,
    value=150,
    step=10,
    help="Satıcıların (temiz tesis) satış isteği hassasiyeti."
)

spread = st.sidebar.slider(
    "Bid/Ask Spread (€/tCO₂)",
    min_value=0.0,
    max_value=10.0,
    value=0.0,
    step=0.5,
    help="0 bırakabilirsin. Spread eklemek bid/ask aynı görünmesini azaltır."
)

st.sidebar.divider()
st.sidebar.caption("Excel'de beklenen kolonlar: Plant, Generation_MWh, Emissions_tCO2")
st.sidebar.caption("Sekme adı FuelType olarak alınır.")

# -------------------------
# Excel yükleme
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

if uploaded is None:
    st.info("Lütfen bir Excel yükleyin.")
    st.stop()

try:
    df_all = read_all_sheets(uploaded)
except Exception as e:
    st.error(f"Excel okunurken hata oluştu: {e}")
    st.stop()

st.subheader("Yüklenen veri (birleştirilmiş)")
st.dataframe(df_all.head(50), use_container_width=True)

# -------------------------
# Model çalıştır
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
        )

        st.success(f"Clearing Price: {clearing_price:.2f} €/tCO₂")

        st.subheader("Benchmark (yakıt bazında)")
        bench_df = pd.DataFrame(
            [{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()]
        ).sort_values("FuelType")
        st.dataframe(bench_df, use_container_width=True)

        st.subheader("Sonuçlar")
        st.dataframe(sonuc_df, use_container_width=True)

        csv_bytes = sonuc_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download results as CSV",
            data=csv_bytes,
            file_name="ets_results.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Model çalışırken hata oluştu: {e}")
