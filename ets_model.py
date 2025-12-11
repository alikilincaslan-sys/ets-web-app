import pandas as pd
import numpy as np


def market_clearing_price(net_positions, price_min=0, price_max=100, step=1.0):
    prices = np.arange(price_min, price_max + step, step)

    """
    Arz-talep dengesine göre clearing price hesaplar.
    net_positions > 0 → talep
    net_positions < 0 → arz
    """
    prices = np.arange(0, price_cap + step, step)

    for price in prices:
        demand = net_positions[net_positions > 0].sum()
        supply = -net_positions[net_positions < 0].sum()

        if supply >= demand:
            return price

    return price_cap


def ets_hesapla(df, price_cap, agk):
    """
    Tahsis Yoğunluğuᵢ = B_yakıt + AGK * (Iᵢ - B_yakıt)

    - B_yakıt: yakıt bazlı üretim ağırlıklı benchmark yoğunluğu
    - AGK    : Adil Geçiş Katsayısı (0–1)
    - Ücretsiz tahsis: Generation_MWh × Tahsis Yoğunluğu
    """

    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    df = df.copy()

    # 1) Gerçek yoğunluk (Iᵢ)
    df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # 2) Yakıt bazlı benchmark (B_yakıt): üretim ağırlıklı yoğunluk
    benchmark_map = {}
    for ft in df["FuelType"].unique():
        subset = df[df["FuelType"] == ft]
        B_fuel = subset["Emissions_tCO2"].sum() / subset["Generation_MWh"].sum()
        benchmark_map[ft] = B_fuel

    df["B_fuel"] = df["FuelType"].map(benchmark_map)

    # 3) Tahsis yoğunluğu (senin formülün)
    df["tahsis_intensity"] = df["B_fuel"] + agk * (df["intensity"] - df["B_fuel"])

    # 4) Ücretsiz tahsis
    df["free_alloc"] = df["Generation_MWh"] * df["tahsis_intensity"]

    # 5) Net ETS pozisyonu
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # 6) Clearing price – tüm tesisler birlikte
    clearing_price = market_clearing_price(df["net_ets"], price_cap)

    # 7) ETS maliyetleri (sadece pozitif net_ets için)
    df["carbon_price"] = clearing_price
    df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price
    df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]

    return df, benchmark_map, clearing_price
