import pandas as pd
import numpy as np


def market_clearing_price(net_positions, price_cap=100, step=0.1):
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


def ets_hesapla(df, price_cap, free_alloc_ratio, agk):
    """
    Tahsis Yoğunluğu_i = B_yakıt + AGK * (I_i - B_yakıt)

    - B_yakıt: yakıt bazlı üretim ağırlıklı benchmark
    - AGK: Adil geçiş katsayısı (0–1)
    - free_alloc_ratio: Free Allocation Ratio
    """

    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    df = df.copy()

    # 1) Gerçek yoğunluk (I_i)
    df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # 2) YAKIT TÜREVİNE GÖRE BENCHMARK (B_yakıt)
    benchmark_map = {}
    for ft in df["FuelType"].unique():
        subset = df[df["FuelType"] == ft]
        B_fuel = subset["Emissions_tCO2"].sum() / subset["Generation_MWh"].sum()
        benchmark_map[ft] = B_fuel

    df["B_fuel"] = df["FuelType"].map(benchmark_map)

    # 3) TAHSiS YOĞUNLUĞU FORMÜLÜ
    df["tahsis_intensity"] = df["B_fuel"] + agk * (df["intensity"] - df["B_fuel"])

    # 4) ÜCRETSİZ TAHSiS (Free Allocation Ratio uygulanır)
    df["free_alloc"] = df["Generation_MWh"] * df["tahsis_intensity"] * (1 - free_alloc_ratio)

    # 5) NET ETS POZISYONU
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # 6) CLEARING PRICE – tüm tesisler bir arada
    clearing_price = market_clearing_price(df["net_ets"], price_cap)

    # 7) MALIYETLER
    df["carbon_price"] = clearing_price
    df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price
    df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]

    return df, benchmark_map, clearing_price
