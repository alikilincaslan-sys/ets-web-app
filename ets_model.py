import numpy as np
import pandas as pd


def market_clearing_price_linear(df, price_min, price_max, step=1):
    if price_max <= price_min:
        raise ValueError("price_max must be greater than price_min")

    prices = np.arange(price_min, price_max + step, step)

    buyers = df[df["net_ets"] > 0]
    sellers = df[df["net_ets"] < 0]

    for p in prices:
        total_demand = 0.0
        total_supply = 0.0

        # Demand (buyers)
        if not buyers.empty:
            q0 = buyers["net_ets"].values
            p_bid = buyers["p_bid"].values
            denom = np.maximum(p_bid - price_min, 1e-6)
            frac = 1 - (p - price_min) / denom
            total_demand = np.sum(q0 * np.clip(frac, 0, 1))

        # Supply (sellers)
        if not sellers.empty:
            q0 = (-sellers["net_ets"]).values
            p_ask = sellers["p_ask"].values
            denom = np.maximum(price_max - p_ask, 1e-6)
            frac = (p - p_ask) / denom
            total_supply = np.sum(q0 * np.clip(frac, 0, 1))

        if total_supply >= total_demand:
            return float(p)

    return float(price_max)


def ets_hesapla(
    df,
    price_min,
    price_max,
    agk,
    slope_bid=150.0,
    slope_ask=150.0,
    spread=0.0,
):
    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"Missing column: {col}")

    if not (0 <= agk <= 1):
        raise ValueError("AGK must be between 0 and 1")

    df = df.copy()

    # 1) Gerçek yoğunluk (I)
    df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # 2) Yakıt bazlı benchmark (B)
    benchmark_map = {}
    for ft in df["FuelType"].unique():
        sub = df[df["FuelType"] == ft]
        benchmark_map[ft] = sub["Emissions_tCO2"].sum() / sub["Generation_MWh"].sum()

    df["B_fuel"] = df["FuelType"].map(benchmark_map)

    # 3) ✅ DOĞRU AGK FORMÜLÜ (TERSİ DÜZELTİLDİ)
    # T_i = I_i + AGK * (B - I)
    df["tahsis_intensity"] = (
        df["intensity"] + agk * (df["B_fuel"] - df["intensity"])
    )

    # 4) Ücretsiz tahsis
    df["free_alloc"] = df["Generation_MWh"] * df["tahsis_intensity"]

    # 5) Net ETS pozisyonu
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # 6) BID / ASK
    delta = df["intensity"] - df["B_fuel"]

    p_bid = price_min + slope_bid * np.maximum(delta, 0)
    p_ask = price_min + slope_ask * np.maximum(-delta, 0)

    p_bid = p_bid + spread / 2
    p_ask = np.maximum(p_ask - spread / 2, price_min)

    df["p_bid"] = p_bid.clip(price_min, price_max)
    df["p_ask"] = p_ask.clip(price_min, price_max)

    # 7) Clearing price
    clearing_price = market_clearing_price_linear(
        df[["net_ets", "p_bid", "p_ask"]],
        price_min,
        price_max,
    )

    # 8) ETS nakit akışları
    df["carbon_price"] = clearing_price
    df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price
    df["ets_revenue_total_€"] = (-df["net_ets"]).clip(lower=0) * clearing_price
    df["ets_net_cashflow_€"] = df["ets_revenue_total_€"] - df["ets_cost_total_€"]

    df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]
    df["ets_revenue_€/MWh"] = df["ets_revenue_total_€"] / df["Generation_MWh"]
    df["ets_net_cashflow_€/MWh"] = df["ets_net_cashflow_€"] / df["Generation_MWh"]

    return df, benchmark_map, clearing_price
