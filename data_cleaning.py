import numpy as np
import pandas as pd


def clean_ets_input(df: pd.DataFrame) -> pd.DataFrame:
    """
    Temel veri temizleme:
    - Sayısal kolonları numeric'e çevirir
    - Zorunlu kolonlarda NA varsa atar
    - Generation_MWh > 0 ve Emissions_tCO2 >= 0 şartı uygular
    """
    x = df.copy()

    # Zorunlu kolonlar
    required = ["Plant", "FuelType", "Emissions_tCO2", "Generation_MWh"]
    for c in required:
        if c not in x.columns:
            raise ValueError(f"Excel kolon eksik: {c}")

    x["Emissions_tCO2"] = pd.to_numeric(x["Emissions_tCO2"], errors="coerce")
    x["Generation_MWh"] = pd.to_numeric(x["Generation_MWh"], errors="coerce")

    x = x.dropna(subset=required)
    x = x[(x["Generation_MWh"] > 0) & (x["Emissions_tCO2"] >= 0)]

    # Plant string temizliği (opsiyonel ama faydalı)
    x["Plant"] = x["Plant"].astype(str).str.strip()

    return x


def filter_intensity_outliers_by_fuel(
    df: pd.DataFrame,
    lower_pct: float,
    upper_pct: float,
):
    """
    Yakıt bazında benchmark B hesaplar (üretim ağırlıklı).
    Sonra intensity = Emissions/Generation
    intensity band dışındaysa satırı çıkarır.

    Band:
      lo = B * (1 - lower_pct)
      hi = B * (1 + upper_pct)

    Dönüş:
      cleaned_df, removed_df
    """
    x = df.copy()
    x["intensity"] = x["Emissions_tCO2"] / x["Generation_MWh"]

    keep_mask = np.ones(len(x), dtype=bool)
    removed_rows = []

    for ft, g in x.groupby("FuelType"):
        gen = g["Generation_MWh"].sum()
        em = g["Emissions_tCO2"].sum()
        if gen <= 0:
            continue

        B = em / gen
        lo = B * (1.0 - float(lower_pct))
        hi = B * (1.0 + float(upper_pct))

        idx = g.index
        ok = (g["intensity"] >= lo) & (g["intensity"] <= hi)

        keep_mask[idx] = ok.to_numpy()

        removed = g.loc[~ok, ["Plant", "FuelType", "Generation_MWh", "Emissions_tCO2", "intensity"]].copy()
        if not removed.empty:
            removed["Benchmark_B"] = B
            removed["LowerBound"] = lo
            removed["UpperBound"] = hi
            removed_rows.append(removed)

    removed_df = pd.concat(removed_rows, ignore_index=True) if removed_rows else pd.DataFrame()
    cleaned_df = x.loc[keep_mask].drop(columns=["intensity"], errors="ignore").copy()

    return cleaned_df, removed_df
