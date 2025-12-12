import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.subheader("AGK (α) Değişimine Göre Emisyon Yoğunlukları — Tüm Santraller (Sıralı)")

# 1) Kullanıcıdan birden fazla AGK (alpha) seçtirelim
#    (İstersen bunu multiselect yerine slider + liste de yapabiliriz)
alpha_list = st.multiselect(
    "Grafikte gösterilecek AGK (α) değerlerini seçin",
    options=[0.25, 0.5, 0.75, 0.9, 1.25, 1.5, 2.0],
    default=[0.5, 0.75]
)

if len(alpha_list) == 0:
    st.info("En az bir AGK (α) değeri seçin.")
    st.stop()

# 2) SENİN MODELDEN GELEN VERİ: santral listesi ve ID/name
#    df_plants: en azından 'plant_name' sütunu olsun.
#    Aşağıdaki satırı kendi dataframe'inle değiştir:
#    df_plants = results_df[['plant_name']].drop_duplicates().copy()
df_plants = df[['plant_name']].drop_duplicates().copy()  # <-- kendi df ismine göre düzelt

# 3) Burada her alpha için santral bazında emisyon yoğunluğu/benchmark değerini üretiyoruz.
#    Bu fonksiyonu senin mevcut hesap fonksiyonunla değiştir.
def compute_intensity_by_alpha(alpha: float) -> pd.DataFrame:
    """
    ÇIKTI: plant_name, intensity_alpha
    intensity_alpha: AGK=alpha için santral bazında emisyon yoğunluğu / benchmark değer(ler)i
    """
    # ---- BURASI SENİN HESABIN ----
    # Örnek: model fonksiyonun şöyle bir seri döndürüyor olsun:
    # series = model_calc_intensity(df, alpha=alpha)  # index=plant_name
    # out = series.reset_index().rename(columns={0:"intensity"})
    #
    # Şimdilik örnek bir placeholder:
    out = df.groupby("plant_name")["emission_intensity"].mean().reset_index()
    out = out.rename(columns={"emission_intensity": "intensity"})
    # ---- BURAYA KADAR ----

    out["alpha"] = alpha
    return out[["plant_name", "intensity", "alpha"]]

frames = [compute_intensity_by_alpha(a) for a in alpha_list]
df_plot = pd.concat(frames, ignore_index=True)

# 4) Sıralama: “seçilen” AGK’ya göre (ben ilk seçileni baz aldım)
alpha_sort = alpha_list[0]
order = (
    df_plot[df_plot["alpha"] == alpha_sort]
    .sort_values("intensity", ascending=True)["plant_name"]
    .tolist()
)

# 5) Grafik: X=Santral sırası (rank), Y=Yoğunluk, her alpha ayrı çizgi
#    İstersen X eksenini plant_name yazdırabiliriz ama çok kalabalık olursa rank daha okunur.
plant_to_rank = {p: i+1 for i, p in enumerate(order)}
df_plot["rank"] = df_plot["plant_name"].map(plant_to_rank)
df_plot = df_plot.dropna(subset=["rank"]).sort_values(["alpha", "rank"])

fig = go.Figure()

for a in alpha_list:
    dfa = df_plot[df_plot["alpha"] == a].sort_values("rank")
    fig.add_trace(
        go.Scatter(
            x=dfa["rank"],
            y=dfa["intensity"],
            mode="lines",
            name=f"AGK α={a}",
            hovertemplate="Sıra: %{x}<br>Yoğunluk: %{y:.4f}<extra></extra>"
        )
    )

fig.update_layout(
    height=600,
    xaxis_title=f"Santraller (düşük→yüksek sıralı) — sıralama α={alpha_sort}",
    yaxis_title="Emisyon Yoğunluğu / Benchmark Değeri",
    legend_title="AGK Senaryosu",
    hovermode="x unified"
)

st.plotly_chart(fig, use_container_width=True)

# 6) İsteğe bağlı: hangi sıraya göre dizildiğini kullanıcı görsün
with st.expander("Sıralanan santral listesi"):
    st.dataframe(pd.DataFrame({"rank": range(1, len(order)+1), "plant_name": order}))
