import numpy as np
import pandas as pd


def market_clearing_price_linear(df_positions: pd.DataFrame, price_min: int, price_max: int, step: int = 1) -> float:
    """
    Lineer BID/ASK yaklaşımı ile piyasa clearing fiyatı.

    Buyers (net_ets > 0): fiyat arttıkça talep lineer azalır ve p_bid'de sıfırlanır.
    Sellers (net_ets < 0): fiyat arttıkça arz lineer artar ve p_ask'ta 0'dan başlar.

    Clearing: toplam_arz >= toplam_talep olduğu ilk fiyat.
    """
    if price_max <= price_min:
        raise ValueError("price_max, price_min'den büyük olmalı.")

    prices = np.arange(price_min, price_max + step, step)

    buyers = df_positions[df_positions["net_ets"] > 0].copy()
    sellers = df_positions[df_positions["net_ets"] < 0].copy()

    for p in prices:
        total_demand = 0.0
        total_supply = 0.0

        # DEMAND (BUYERS)
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()  # pozitif
            p_bid = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid - price_min, 1e-9)
            frac = 1.0 - (p - price_min) / denom
            q = q0 * np.clip(frac, 0.0, 1.0)
            total_demand = float(np.sum(q))

        # SUPPLY (SELLERS)
        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()  # pozitif kapasite
            p_ask = sellers["p_ask"].to_numpy()
            denom = np.maximum(price_max - p_ask, 1e-9)
            frac = (p - p_ask) / denom
            q = q0 * np.clip(frac, 0.0, 1.0)
            total_supply = float(np.sum(q))

        if total_supply >= total_demand:
            return float(p)

    return float(price_max)


def ets_hesapla(
    df: pd.DataFrame,
    price_min: int,
    price_max: int,
    agk: float,
    slope_bid: float = 150.0,
    slope_ask: float = 150.0,
    spread: float = 0.0,
):
    """
    1) Yakıt bazlı benchmark (B_yakıt)
    2) Tahsis yoğunluğu: T_i = B_yakıt + AGK*(I_i - B_yakıt)
    3) Net ETS: net_ets = Em - Gen*T_i
    4) BID/ASK fonksiyonları ile birleşik piyasada clearing price (lineer)

    Not:
    - slope_bid, slope_ask ve spread opsiyoneldir (default değerleri var).
    """

    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    if price_max <= price_min:
        raise ValueError("price_max, price_min'den büyük olmalı.")
    if not (0.0 <= agk <= 1.0):
        raise ValueError("AGK 0 ile 1 arasında olmalı.")
    if slope_bid <= 0 or slope_ask <= 0:
        raise ValueError("slope_bid ve slope_ask pozitif olmalı.")
    if spread < 0:
        raise ValueError("spread negatif olamaz.")

    df = df.copy()

    # 1) Gerçek yoğunluk (Iᵢ)
    df["intensity"] = df["Emissions_tCO2"] / df["Generation_MWh"]

    # 2) Yakıt bazlı benchmark (üretim ağırlıklı)
    benchmark_map = {}
    for ft in df["FuelType"].unique():
        sub = df[df["FuelType"] == ft]
        B_fuel = sub["Emissions_tCO2"].sum() / sub["Generation_MWh"].sum()
        benchmark_map[ft] = float(B_fuel)

    df["B_fuel"] = df["FuelType"].map(benchmark_map)

    # 3) Tahsis yoğunluğu (AGK senin tanımınla)
    df["tahsis_intensity"] = df["B_fuel"] + agk * (df["intensity"] - df["B_fuel"])

    # 4) Ücretsiz tahsis
    df["free_alloc"] = df["Generation_MWh"] * df["tahsis_intensity"]

    # 5) Net ETS pozisyonu
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # 6) BID/ASK (ayrı)
    delta = df["intensity"] - df["B_fuel"]

    p_bid = price_min + slope_bid * np.maximum(delta, 0.0)       # kirli → daha yüksek bid
    p_ask = price_min + slope_ask * np.maximum(-delta, 0.0)      # temiz → daha yüksek ask

    # Spread uygula (istersen 0 bırak)
    p_bid = p_bid + spread / 2.0
    p_ask = np.maximum(p_ask - spread / 2.0, price_min)

    # Aralığa kırp
    df["p_bid"] = p_bid.clip(lower=price_min, upper=price_max)
    df["p_ask"] = p_ask.clip(lower=price_min, upper=price_max)

    # 7) Clearing price
    clearing_price = market_clearing_price_linear(
        df[["net_ets", "p_bid", "p_ask"]],
        price_min,
        price_max,
        step=1,
    )

    # 8) ETS maliyetleri
    df["carbon_price"] = clearing_price

# Alıcı maliyeti (net_ets > 0)
df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price

# Satıcı geliri (net_ets < 0)
df["ets_revenue_total_€"] = (-df["net_ets"]).clip(lower=0) * clearing_price

# Net nakit akışı: + gelir, - maliyet
df["ets_net_cashflow_€"] = df["ets_revenue_total_€"] - df["ets_cost_total_€"]

# MWh bazına indirgeme
df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]
df["ets_revenue_€/MWh"] = df["ets_revenue_total_€"] / df["Generation_MWh"]
df["ets_net_cashflow_€/MWh"] = df["ets_net_cashflow_€"] / df["Generation_MWh"]


    return df, benchmark_map, clearing_price
