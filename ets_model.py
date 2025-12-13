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


def _compute_benchmarks(df: pd.DataFrame, benchmark_top_pct: int = 100) -> dict:
    """Yakıt bazında benchmark (B_fuel). Seçim: en iyi top% (intensity düşük)."""
    bench = {}
    for ft, g in df.groupby("FuelType"):
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
    Basit BID/ASK üretimi:
    - intensity > B => daha "kirli" => p_bid yukarı
    - intensity < B => daha "temiz" => p_ask aşağı
    """
    out = df.copy()
    delta = (out["intensity"] - out["B_fuel"]).fillna(0.0)

    out["p_bid"] = np.clip(
        price_min + slope_bid * np.maximum(delta, 0.0) + spread / 2.0,
        price_min,
        price_max,
    )

    out["p_ask"] = np.clip(
        price_max - slope_ask * np.maximum(-delta, 0.0) - spread / 2.0,
        price_min,
        price_max,
    )

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
    Sadece alıcıların (NetETS>0) p_bid değerlerini, net yükümlülükle ağırlıklandırarak ortalama fiyat üretir:

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


def ets_hesapla(
    df: pd.DataFrame,
    price_min: float,
    price_max: float,
    agk: float,
    slope_bid: float = 150,
    slope_ask: float = 150,
    spread: float = 0.0,
    benchmark_top_pct: int = 100,
    free_alloc_share: float = 100.0,
    trf: float = 0.0,
    price_method: str = "Market Clearing",  # ✅ yeni
):
    """
    AGK yönü:
      AGK=1  -> tahsis yoğunluğu benchmark'a yaklaşır
      AGK=0  -> tahsis yoğunluğu tesis yoğunluğuna yaklaşır

    Formül:
      T_i = I_i + AGK * (B_fuel - I_i)

    Benchmark:
      benchmark_top_pct=10 => en iyi %10 üretim dilimi
      benchmark_top_pct=100 => tüm tesisler

    TRF (Geçiş Dönemi Telafi Katsayısı):
      İlave tahsis = max(0, I_i - B_fuel) * G_i * TRF

    Price method:
      - "Market Clearing"
      - "Average Compliance Cost"
    """
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
    benchmark_map = _compute_benchmarks(x, benchmark_top_pct=int(benchmark_top_pct))
    x["B_fuel"] = x["FuelType"].map(benchmark_map)

    # 3) Tahsis yoğunluğu (AGK)
    x["tahsis_intensity"] = x["intensity"] + float(agk) * (x["B_fuel"] - x["intensity"])

    # 4) Ücretsiz tahsis
    x["free_alloc"] = x["Generation_MWh"] * x["tahsis_intensity"]


    # Free allocation share (policy lever): 100%=full free allocation; 0%=no free allocation
    x["free_alloc"] = x["free_alloc"] * (float(free_alloc_share) / 100.0)
    
    # TRF (Geçiş Dönemi Telafi Katsayısı): benchmark nedeniyle oluşan ilave yükün pilot dönemde telafisi
    # İlave tahsis (tCO2) = max(0, intensity - B_fuel) * Generation_MWh * TRF
    trf_val = float(trf) if trf is not None else 0.0
    x["ilave_tahsis_trf"] = np.maximum(x["intensity"] - x["B_fuel"], 0.0) * x["Generation_MWh"] * trf_val

    # Toplam ücretsiz tahsis = mevcut tahsis + TRF ilavesi
    x["free_alloc_total"] = x["free_alloc"] + x["ilave_tahsis_trf"]

    # 5) Net ETS pozisyonu
    x["net_ets"] = x["Emissions_tCO2"] - x["free_alloc_total"]    # 6) BID/ASK
    x = _build_bid_ask(x, price_min, price_max, slope_bid, slope_ask, spread)

    # 7) Price
    if price_method == "Average Compliance Cost":
        clearing_price = _average_compliance_cost_price(x, price_min, price_max)
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
