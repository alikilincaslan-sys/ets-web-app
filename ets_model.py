import numpy as np
import pandas as pd


def market_clearing_price_linear(
    df_positions: pd.DataFrame,
    price_min: int,
    price_max: int,
    step: int = 1,
):
    """
    Lineer BID/ASK yaklaşımı ile piyasa clearing fiyatı.
    - Alıcılar (net_ets > 0): fiyat arttıkça talep lineer azalır ve p_bid'de sıfırlanır.
    - Satıcılar (net_ets < 0): fiyat arttıkça arz lineer artar ve price_max'ta tam kapasiteye ulaşır.
    """

    if price_max <= price_min:
        raise ValueError("price_max, price_min'den büyük olmalı.")

    prices = np.arange(price_min, price_max + step, step)

    # Alıcılar: net_ets > 0
    buyers = df_positions[df_positions["net_ets"] > 0].copy()
    # Satıcılar: net_ets < 0
    sellers = df_positions[df_positions["net_ets"] < 0].copy()

    # Toplam talep/arz fonksiyonları
    for p in prices:
        total_demand = 0.0
        total_supply = 0.0

        # ---- DEMAND (buyers) ----
        # q(p) = q0 * max(0, 1 - (p - price_min)/(p_bid - price_min))
        if not buyers.empty:
            p_bid = buyers["p_bid"].values
            q0 = buyers["net_ets"].values  # pozitif

            denom = np.maximum(p_bid - price_min, 1e-9)
            frac = 1.0 - (p - price_min) / denom
            q = q0 * np.clip(frac, 0.0, 1.0)

            # p > p_bid ise otomatik 0 olur (frac negatif)
            total_demand = float(np.sum(q))

        # ---- SUPPLY (sellers) ----
        # q(p) = q0 * max(0, (p - p_ask)/(price_max - p_ask))
        if not sellers.empty:
            p_ask = sellers["p_ask"].values
            q0 = (-sellers["net_ets"].values)  # pozitif arz kapasitesi

            denom = np.maximum(price_max - p_ask, 1e-9)
            frac = (p - p_ask) / denom
            q = q0 * np.clip(frac, 0.0, 1.0)

            total_supply = float(np.sum(q))

        # Clearing: arz >= talep olduğunda ilk fiyat
        if total_supply >= total_demand:
            return float(p)

    # Hiçbir fiyatta arz talebi karşılamazsa üst sınır
    return float(price_max)


def ets_hesapla(df: pd.DataFrame, price_min: int, price_max: int, agk: float):
    """
    1) Yakıt bazlı benchmark (B_yakıt)
    2) Tahsis yoğunluğu: T_i = B_yakıt + AGK*(I_i - B_yakıt)
    3) Net ETS: net_ets = Em - Gen*T_i
    4) Lineer BID/ASK ile birleşik piyasada clearing price

    Not: AGK tanımı senin istediğin gibi korunmuştur.
    """

    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    if price_max <= price_min:
        raise ValueError("price_max, price_min'den büyük olmalı.")
    if not (0.0 <= agk <= 1.0):
        raise ValueError("AGK 0 ile 1 arasında olmalı.")

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
    # T_i = (1-AGK)*B + AGK*I
    df["tahsis_intensity"] = df["B_fuel"] + agk * (df["intensity"] - df["B_fuel"])

    # 4) Ücretsiz tahsis
    df["free_alloc"] = df["Generation_MWh"] * df["tahsis_intensity"]

    # 5) Net ETS pozisyonu
    df["net_ets"] = df["Emissions_tCO2"] - df["free_alloc"]

    # -----------------------------
    # 6) BID/ASK fiyat parametreleri (Lineer)
    # -----------------------------
    # Basit ama gerçekçi bir "isteklilik" dönüşümü:
    # intensity - benchmark farkını €/t ölçeğine çevirip min–max aralığına kırpıyoruz.
    # Bu ölçek katsayısını ileride slider yapabiliriz.
    price_slope = 100.0  # €/tCO2 / (tCO2/MWh) ölçeği (başlangıç)

    # Temiz santral (I<B) → daha düşük fiyatlardan satmaya razı (ask düşük)
    # Kirli santral (I>B) → daha yüksek fiyata kadar almaya razı (bid yüksek)
    p_ref = price_min + price_slope * (df["intensity"] - df["B_fuel"])

    # min–max aralığına kırp
    p_ref = p_ref.clip(lower=price_min, upper=price_max)

    # Buyer için p_bid, seller için p_ask olarak kullanıyoruz
    df["p_bid"] = p_ref
    df["p_ask"] = p_ref

    # 7) Clearing price (birleşik piyasa, gerçekçi fiyata duyarlı arz-talep)
    clearing_price = market_clearing_price_linear(df[["net_ets", "p_bid", "p_ask"]], price_min, price_max, step=1)

    # 8) ETS maliyetleri (pozitif net_ets için ödeme)
    df["carbon_price"] = clearing_price
    df["ets_cost_total_€"] = df["net_ets"].clip(lower=0) * clearing_price
    df["ets_cost_€/MWh"] = df["ets_cost_total_€"] / df["Generation_MWh"]

    return df, benchmark_map, clearing_price
