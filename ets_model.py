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


def ets_hesapla(df, price_cap, agk_orani, alpha):
    """
    - Benchmark: fuel-type based
    - Clearing price: all plants combined
    - Fiyat: unified clearing price
    """
    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    df = df.copy()

    # Emisyon yoğunluğu
    df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # FUEL-TYPE SPECIFIC BENCHMARK
    benchmark_map = {}
    for ft in df["FuelType"].unique():
        subset = df[df["FuelType"] == ft]
        avg_intensity = subset["Emissions_tCO2"].sum() / subset["Generation_MWh"].sum()
        benchmark_map[ft] = alpha * avg_intensity

    # Santral bazında benchmark ata
    df["benchmark"] = df["FuelType"].map(benchmark_map)

    # Ücretsiz tahsis
    df["free_alloc"] = df["Generation_MWh"] * df["benchmark"] * (1 - agk_orani)

    # Net ETS pozisyonu
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # CLEARING PRICE → ALL PLANTS TOGETHER
    clearing_price = market_clearing_price(df["net_ets"].values, price_cap)

    # Maliyet
    df["carbon_price"] = clearing_price
    df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price
    df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]

    return df, benchmark_map, clearing_price
