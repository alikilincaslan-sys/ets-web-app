# ============================================================
# ETS GELİŞTİRME MODÜLÜ V001 – FINAL (Infographic + Word Note)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.express as px

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input

# ============================================================
# PAGE
# ============================================================
st.set_page_config(page_title="ETS Geliştirme Modülü V002", layout="wide")
st.title("ETS Geliştirme Modülü v002")

# ============================================================
# CSS (IEA style)
# ============================================================
st.markdown("""
<style>
.kpi {
    background:#f2f2f2; border-radius:14px; padding:14px;
    border:1px solid rgba(0,0,0,.12);
}
.kpi .label {font-size:.8rem;color:#555}
.kpi .value {font-size:1.35rem;font-weight:700}
</style>
""", unsafe_allow_html=True)

def kpi_card(label, value, sub=""):
    st.markdown(
        f"""<div class='kpi'>
        <div class='label'>{label}</div>
        <div class='value'>{value}</div>
        <div style='font-size:.7rem;color:#777'>{sub}</div>
        </div>""",
        unsafe_allow_html=True
    )

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("Model Parametreleri")

price_min, price_max = st.sidebar.slider("Carbon price range (€/tCO₂)",0,200,(1,15))
agk = st.sidebar.slider("AGK",0.0,1.0,0.5,0.05)

price_method = st.sidebar.selectbox(
    "Price method",
    ["Market Clearing","Average Compliance Cost","Auction Clearing"]
)

auction_supply_share = 1.0
if price_method=="Auction Clearing":
    auction_supply_share = st.sidebar.slider(
        "Auction supply (% of demand)",10,200,100,10
    )/100

# ============================================================
# FILE UPLOAD
# ============================================================
uploaded = st.file_uploader("Upload Excel (.xlsx)",type="xlsx")
if not uploaded:
    st.stop()

xls = pd.ExcelFile(uploaded)
df_all = []
for sh in xls.sheet_names:
    d = pd.read_excel(xls,sh)
    d["FuelType"]=sh
    df_all.append(d)

df_all = clean_ets_input(pd.concat(df_all,ignore_index=True))

# ============================================================
# RUN MODEL
# ============================================================
if st.button("Run ETS Model"):

    sonuc_df, benchmark_map, clearing_price = ets_hesapla(
        df_all,
        price_min,price_max,agk,
        price_method=price_method,
        auction_supply_share=auction_supply_share
    )

    # ========================================================
    # KPI ROW
    # ========================================================
    total_gen = df_all["Generation_MWh"].sum()
    total_em = df_all["Emissions_tCO2"].sum()/1e6
    total_cap = df_all.get("InstalledCapacity_MW",pd.Series()).sum()

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi_card("Generation",f"{total_gen:,.0f} MWh")
    with c2: kpi_card("Installed capacity",f"{total_cap:,.0f} MW")
    with c3: kpi_card("Emissions",f"{total_em:.2f} MtCO₂")
    with c4: kpi_card("Carbon price",f"{clearing_price:.2f} €/t")

    # ========================================================
    # === IEA VISUAL 1: AUCTION / MARKET SUMMARY
    # ========================================================
    st.subheader("Market summary")

    demand = sonuc_df.loc[sonuc_df.net_ets>0,"net_ets"].sum()
    supply = demand*auction_supply_share if price_method=="Auction Clearing" else np.nan
    traded = min(demand,supply) if price_method=="Auction Clearing" else demand

    a1,a2,a3 = st.columns(3)
    with a1: kpi_card("Total demand",f"{demand:,.0f} tCO₂")
    with a2: kpi_card("Auction supply",f"{supply:,.0f} tCO₂","Auction only")
    with a3: kpi_card("Traded volume",f"{traded:,.0f} tCO₂")

    # ========================================================
    # === IEA VISUAL 2: BID–ASK CURVES
    # ========================================================
    st.subheader("Bid–Ask curves and clearing price")

    buyers = sonuc_df[sonuc_df.net_ets>0].sort_values("p_bid")
    sellers = sonuc_df[sonuc_df.net_ets<0].sort_values("p_ask")

    buyers["cum_q"] = buyers["net_ets"].cumsum()
    sellers["cum_q"] = (-sellers["net_ets"]).cumsum()

    fig,ax = plt.subplots(figsize=(7,4))
    ax.plot(buyers["cum_q"],buyers["p_bid"],label="Demand (bids)",color="#1f77b4")
    ax.plot(sellers["cum_q"],sellers["p_ask"],label="Supply (asks)",color="#d62728")
    ax.axhline(clearing_price,color="black",ls="--",label="Clearing price")

    ax.set_xlabel("Allowances (tCO₂)")
    ax.set_ylabel("Price (€/tCO₂)")
    ax.legend()
    ax.grid(alpha=.15)

    st.pyplot(fig)

    # ========================================================
    # === IEA VISUAL 3: FUEL BENCHMARK DISTRIBUTION
    # ========================================================
    st.subheader("Emission intensity vs benchmark (by fuel)")

    df_plot = sonuc_df.copy()
    fig2 = px.box(
        df_plot,
        x="FuelType",
        y="intensity",
        points="all",
        color="FuelType",
        template="simple_white",
        labels={"intensity":"tCO₂ / MWh","FuelType":""}
    )

    for fuel,b in benchmark_map.items():
        fig2.add_hline(
            y=b,
            line_dash="dash",
            annotation_text=f"{fuel} benchmark",
            annotation_position="top left"
        )

    fig2.update_layout(showlegend=False)
    st.plotly_chart(fig2,use_container_width=True)

    # ========================================================
    # RESULTS TABLE
    # ========================================================
    st.subheader("Results table")
    st.dataframe(sonuc_df,use_container_width=True)
