# ets_model.py
import pandas as pd
import numpy as np

def ets_hesapla(df, tavan_fiyat, agk_orani, alpha):
    """
    ETS Geliştirme Modülü V001 ana hesaplama fonksiyonu.

    Parametreler:
        df           : Excel'den okunmuş ham veri (her satır bir santral)
        tavan_fiyat  : €/tCO2 cinsinden fiyat tavanı
        agk_orani    : Örneğin 0.15 gibi, ücretsiz tahsis oranı
        alpha        : Benchmark smoothing katsayısı (0–1)

    Dönüş:
        sonuc_df     : Hesaplanmış kolonları içeren DataFrame
    """

    # Gerekli kolon isimleri:
    required_cols = ["Plant", "Emissions_tCO2", "Generation_MWh"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(
                f"Excel dosyasında '{col}' isimli kolon bulunamadı. "
                f"Lütfen kolon adlarını kontrol edin."
            )

    df = df.copy()

    # 1) Emisyon yoğunluğu (tCO2/MWh)
    df["intensity_tCO2_per_MWh"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # 2) Basit benchmark (ortalama yoğunluk → alpha ile smoothing)
    ort_intensity = df["Emissions_tCO2"].sum() / df["Generation_MWh"].sum()
    benchmark_intensity = alpha * ort_intensity
    df["benchmark_intensity"] = benchmark_intensity

    # 3) Ücretsiz tahsis ve net ETS yükümlülüğü
    df["free_allocation_tCO2"] = df["benchmark_intensity"] * df["Generation_MWh"] * (1 - agk_orani)
    df["net_ets_obligation_tCO2"] = df["Emissions_tCO2"] - df["free_allocation_tCO2"]
    df["net_ets_obligation_tCO2"] = df["net_ets_obligation_tCO2"].clip(lower=0)

    # 4) Fiyat ve maliyet
    carbon_price = tavan_fiyat
    df["carbon_price_€/tCO2"] = carbon_price
    df["ets_cost_€"] = df["net_ets_obligation_tCO2"] * df["carbon_price_€/tCO2"]
    df["ets_cost_€/MWh"] = df["ets_cost_€"] / df["Generation_MWh"]

    # TL cinsinden (örnek kur: 35 TL/€)
    kur = 35
    df["ets_cost_TL/MWh"] = df["ets_cost_€/MWh"] * kur

    show_cols = [
        "Plant",
        "Emissions_tCO2",
        "Generation_MWh",
        "intensity_tCO2_per_MWh",
        "benchmark_intensity",
        "free_allocation_tCO2",
        "net_ets_obligation_tCO2",
        "carbon_price_€/tCO2",
        "ets_cost_€/MWh",
        "ets_cost_TL/MWh",
    ]

    sonuc_df = df[show_cols]

    return sonuc_df
