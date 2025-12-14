import numpy as np
import pandas as pd


def _compute_benchmark_top_percent(subset: pd.DataFrame, top_pct: int) -> float:
    """
    Benchmark = seçilen grubun üretim ağırlıklı yoğunluğu = sum(E)/sum(G)

    Seçim mantığı:
    - intensity düşük olanlardan başlayarak sıralar
    - toplam üretimin top_pct kadarı dolana kadar seçer (production-share based)
    """
    if subset.empty:
        return np.nan

    sub = subset.copy()
    sub["Generation_MWh"] = pd.to_numeric(sub["Generation_MWh"], errors="coerce")
    sub["Emissions_tCO2"] = pd.to_numeric(sub["Emissions_tCO2"], errors="coerce")
    sub = sub.dropna(subset=["Generation_MWh", "Emissions_tCO2"])
    sub = sub[(sub["Generation_MWh"] > 0) & (sub["Emissions_tCO2"] >= 0)]

    if sub.empty:
        return np.nan

    sub["intensity"] = sub["Emissions_tCO2"] / sub["Generation_MWh"]
    sub = sub.sort_values("intensity", ascending=True)

    # 100% => tüm tesisler
    if top_pct >= 100:
        return float(sub["Emissions_tCO2"].sum() / sub["Generation_MWh"].sum())

    total_gen = float(sub["Generation_MWh"].sum())
    if total_gen <= 0:
        return np.nan

    target_gen = total_gen * (top_pct / 100.0)

    sub["cum_gen"] = sub["Generation_MWh"].cumsum()
    chosen = sub[sub["cum_gen"] <= target_gen].copy()

    # Güvenlik: en az 1 satır seç
    if chosen.empty:
        chosen = sub.head(1)

    return float(chosen["Emissions_tCO2"].sum() / chosen["Generation_MWh"].sum())


def _compute_benchmark_generation_weighted(subset: pd.DataFrame) -> float:
    """Benchmark = yakıt grubunun üretim-ağırlıklı yoğunluğu = sum(E)/sum(G)."""
    if subset.empty:
        return np.nan
    sub = subset.copy()
    sub["Generation_MWh"] = pd.to_numeric(sub["Generation_MWh"], errors="coerce")
    sub["Emissions_tCO2"] = pd.to_numeric(sub["Emissions_tCO2"], errors="coerce")
    sub = sub.dropna(subset=["Generation_MWh", "Emissions_tCO2"])
    sub = sub[(sub["Generation_MWh"] > 0) & (sub["Emissions_tCO2"] >= 0)]
    if sub.empty:
        return np.nan
    return float(sub["Emissions_tCO2"].sum() / sub["Generation_MWh"].sum())


def _compute_benchmark_capacity_weighted(subset: pd.DataFrame, cap_col: str = "InstalledCapacity_MW") -> float:
    """Benchmark = yakıt grubunun kurulu güç-ağırlıklı yoğunluğu = sum(intensity*Cap)/sum(Cap)."""
    if subset.empty:
        return np.nan
    sub = subset.copy()
    sub["Generation_MWh"] = pd.to_numeric(sub["Generation_MWh"], errors="coerce")
    sub["Emissions_tCO2"] = pd.to_numeric(sub["Emissions_tCO2"], errors="coerce")
    if cap_col not in sub.columns:
        raise ValueError(f"Kurulu güç ağırlıklı benchmark için Excel kolon eksik: {cap_col}")
    sub[cap_col] = pd.to_numeric(sub[cap_col], errors="coerce")
    sub = sub.dropna(subset=["Generation_MWh", "Emissions_tCO2", cap_col])
    sub = sub[(sub["Generation_MWh"] > 0) & (sub["Emissions_tCO2"] >= 0) & (sub[cap_col] > 0)]
    if sub.empty:
        return np.nan
    sub["intensity"] = sub["Emissions_tCO2"] / sub["Generation_MWh"]
    w = sub[cap_col].to_numpy()
    denom = float(np.sum(w))
    if denom <= 0:
        return np.nan
    return float(np.sum(sub["intensity"].to_numpy() * w) / denom)


def _compute_benchmarks(
    df: pd.DataFrame,
    benchmark_method: str = "best_plants",
    benchmark_top_pct: int = 100,
    cap_col: str = "InstalledCapacity_MW",
) -> dict:
    """Yakıt bazında benchmark (B_fuel). Yönteme göre hesaplanır."""
    bench = {}
    for ft, g in df.groupby("FuelType"):
        if benchmark_method == "generation_weighted":
            bench[ft] = _compute_benchmark_generation_weighted(g)
        elif benchmark_method == "capacity_weighted":
            bench[ft] = _compute_benchmark_capacity_weighted(g, cap_col=cap_col)
        else:  # "best_plants"
            bench[ft] = _compute_benchmark_top_percent(g, int(benchmark_top_pct))
    return bench


def _build_bid_ask(
    df: pd.DataFrame,
    price_min: float,
    price_max: float,
    slope_bid: float,
    slope_ask: float,
    spread: float,
) -> pd.DataFrame:
    """
    ✅ SOFT MAPPING (Option A):
    Hard clip yerine yumuşak doyum:
      f(z) = 1 - exp(-k*z)   (z>=0)

    - delta = intensity - B_fuel
    - "kirli" (delta>0) => p_bid artar (price_max'a yaklaşır, yapışmaz)
    - "temiz" (delta<0) => p_ask düşer (price_min'e yaklaşır, yapışmaz)

    slope_bid / slope_ask artık "doyum hızı" gibi davranır.
    Eski lineer yaklaşımın ölçeğini korumak için:
      k ≈ slope / (price_max - price_min)
    """
    out = df.copy()
    delta = (out["intensity"] - out["B_fuel"]).fillna(0.0)

    rng = float(price_max - price_min)
    rng = max(rng, 1e-9)

    # "doyum hızı" (k): slope büyüdükçe daha hızlı tavana/zemine yaklaşır ama "yapışmaz"
    k_bid = float(max(slope_bid, 0.0)) / rng
    k_ask = float(max(slope_ask, 0.0)) / rng

    # Pozitif/negatif delta ayrıştır
    dpos = np.maximum(delta.to_numpy(dtype=float), 0.0)
    dneg = np.maximum((-delta).to_numpy(dtype=float), 0.0)

    # 0..1 arası yumuşak sıkıştırma
    bid_frac = 1.0 - np.exp(-k_bid * dpos)
    ask_frac = 1.0 - np.exp(-k_ask * dneg)

    # Fiyatlar (bounded ama clip yok: doğal olarak [min,max] içinde)
    out["p_bid"] = price_min + rng * bid_frac + spread / 2.0
    out["p_ask"] = price_max - rng * ask_frac - spread / 2.0

    # Çok nadir numeric taşmaları güvenliğe al
    out["p_bid"] = np.minimum(np.maximum(out["p_bid"], price_min), price_max)
    out["p_ask"] = np.minimum(np.maximum(out["p_ask"], price_min), price_max)

    return out


def _market_clearing_price(df: pd.DataFrame, price_min: float, price_max: float, step: float = 1.0) -> float:
    """
    Tek piyasa clearing:
    - Demand: fiyat arttıkça azalır (p_bid ile ölçek)
    - Supply: fiyat arttıkça artar (p_ask ile ölçek)
    İlk Supply >= Demand olduğu fiyat clearing.
    """
    prices = np.arange(price_min, price_max + step, step)

    buyers = df[df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = df[df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    for p in prices:
        # Demand
        if buyers.empty:
            demand = 0.0
        else:
            q0 = buyers["net_ets"].to_numpy()
            p_bid = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid - price_min, 1e-9)
            frac = 1.0 - (p - price_min) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))

        # Supply
        if sellers.empty:
            supply = 0.0
        else:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask = sellers["p_ask"].to_numpy()
            denom = np.maximum(price_max - p_ask, 1e-9)
            frac = (p - p_ask) / denom
            supply = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))

        if supply >= demand:
            return float(p)

    return float(price_max)


def _average_compliance_cost_price(df: pd.DataFrame, price_min: float, price_max: float) -> float:
    """
    Average Compliance Cost (ACC):
      P = sum(NetETS_i * p_bid_i) / sum(NetETS_i)   (NetETS>0)
    Sonra [price_min, price_max] bandına kırpılır.
    """
    buyers = df[df["net_ets"] > 0].copy()
    if buyers.empty:
        return float(price_min)

    w = buyers["net_ets"].to_numpy()
    p = buyers["p_bid"].to_numpy()
    denom = float(np.sum(w))
    if denom <= 0:
        return float(price_min)

    acc = float(np.sum(w * p) / denom)
    return float(np.clip(acc, price_min, price_max))


def _auction_clearing_price(
    df: pd.DataFrame,
    price_min: float,
    price_max: float,
    auction_supply_share: float = 1.0,
) -> float:
    """
    Auction Clearing (talep sabit, yıl sonu compliance):
      Demand = sum(net_ets) for buyers (net_ets > 0)
      Supply = Demand * auction_supply_share
      En yüksek p_bid verenlerden başlayarak Supply kadar tahsis edilir.
      Clearing price = marjinal kazanan teklifin p_bid'i.
    """
    buyers = df[df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    if buyers.empty:
        return float(price_min)

    demand = float(buyers["net_ets"].sum())
    if not np.isfinite(demand) or demand <= 0:
        return float(price_min)

    supply = demand * float(auction_supply_share)
    if not np.isfinite(supply) or supply <= 0:
        return float(price_max)

    if supply >= demand:
        return float(price_min)

    buyers = buyers.sort_values("p_bid", ascending=False)
    cum = buyers["net_ets"].cumsum()

    idx = int(min(max(cum.searchsorted(supply, side="left"), 0), len(buyers) - 1))
    clearing = float(buyers.iloc[idx]["p_bid"])
    return float(np.clip(clearing, price_min, price_max))


def ets_hesapla(
    df: pd.DataFrame,
    price_min: float,
    price_max: float,
    agk: float,
    slope_bid: float = 150,
    slope_ask: float = 150,
    spread: float = 0.0,
    benchmark_method: str = "best_plants",
    cap_col: str = "InstalledCapacity_MW",
    benchmark_top_pct: int = 100,
    free_alloc_share: float = 100.0,
    trf: float = 0.0,
    price_method: str = "Market Clearing",
    auction_supply_share: float = 1.0,
):
    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    x = df.copy()

    x["Emissions_tCO2"] = pd.to_numeric(x["Emissions_tCO2"], errors="coerce")
    x["Generation_MWh"] = pd.to_numeric(x["Generation_MWh"], errors="coerce")
    x = x.dropna(subset=["Emissions_tCO2", "Generation_MWh", "Plant", "FuelType"])
    x = x[(x["Generation_MWh"] > 0) & (x["Emissions_tCO2"] >= 0)]

    # 1) Gerçek yoğunluk
    x["intensity"] = x["Emissions_tCO2"] / x["Generation_MWh"]

    # 2) Benchmark (yakıt bazında)
    benchmark_map = _compute_benchmarks(
        x,
        benchmark_method=benchmark_method,
        benchmark_top_pct=int(benchmark_top_pct),
        cap_col=cap_col
    )
    x["B_fuel"] = x["FuelType"].map(benchmark_map)

    # 3) Tahsis yoğunluğu (AGK)
    x["tahsis_intensity"] = x["intensity"] + float(agk) * (x["B_fuel"] - x["intensity"])

    # 4) Ücretsiz tahsis
    x["free_alloc"] = x["Generation_MWh"] * x["tahsis_intensity"]
    x["free_alloc"] = x["free_alloc"] * (float(free_alloc_share) / 100.0)

    # TRF telafisi
    trf_val = float(trf) if trf is not None else 0.0
    x["ilave_tahsis_trf"] = np.maximum(x["intensity"] - x["B_fuel"], 0.0) * x["Generation_MWh"] * trf_val
    x["free_alloc_total"] = x["free_alloc"] + x["ilave_tahsis_trf"]

    # 5) Net ETS pozisyonu
    x["net_ets"] = x["Emissions_tCO2"] - x["free_alloc_total"]

    # 6) BID/ASK (✅ soft mapping)
    x = _build_bid_ask(x, price_min, price_max, slope_bid, slope_ask, spread)

    # 7) Price
    if price_method == "Average Compliance Cost":
        clearing_price = _average_compliance_cost_price(x, price_min, price_max)
    elif price_method == "Auction Clearing":
        clearing_price = _auction_clearing_price(
            x, price_min, price_max, auction_supply_share=float(auction_supply_share)
        )
    else:
        clearing_price = _market_clearing_price(x, price_min, price_max, step=1.0)

    # 8) Maliyet / gelir
    x["carbon_price"] = clearing_price

    x["ets_cost_total_€"] = x["net_ets"].clip(lower=0) * clearing_price
    x["ets_revenue_total_€"] = (-x["net_ets"].clip(upper=0)) * clearing_price

    x["ets_cost_€/MWh"] = x["ets_cost_total_€"] / x["Generation_MWh"]
    x["ets_revenue_€/MWh"] = x["ets_revenue_total_€"] / x["Generation_MWh"]

    x["ets_net_cashflow_€"] = x["ets_revenue_total_€"] - x["ets_cost_total_€"]
    x["ets_net_cashflow_€/MWh"] = x["ets_net_cashflow_€"] / x["Generation_MWh"]

    return x, benchmark_map, clearing_price
