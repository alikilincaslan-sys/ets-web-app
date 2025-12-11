import pandas as pd
import numpy as np

# Excel’de olası kolon adlarını standart kolonlara map’liyoruz
COLUMN_ALIASES = {
    "Plant": ["Plant", "plant", "Santral", "Tesis", "Santral Adı", "PlantName"],
    "Generation_MWh": ["Generation_MWh", "generation_mwh", "Üretim_MWh", "Uretim_MWh", "Generation", "MWh"],
    "Emissions_tCO2": ["Emissions_tCO2", "emissions_tco2", "Emisyon_tCO2", "Emisyon", "CO2", "tCO2"],
    "FuelType": ["FuelType", "fueltype", "Yakıt", "Yakit", "Fuel", "Yakıt Türü"],
}

def _normalize_col(s: str) -> str:
    return str(s).strip().replace("\n", " ").replace("\t", " ")

def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_normalize_col(c) for c in df.columns]

    rename_map = {}
    for std, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if a in df.columns:
                rename_map[a] = std
                break

    df = df.rename(columns=rename_map)
    return df

def parse_numeric(series: pd.Series) -> pd.Series:
    """
    1.234.567,89  -> 1234567.89
    1,234,567.89  -> 1234567.89
    1234567,89    -> 1234567.89
    """
    s = series.astype(str).str.strip()

    # boş/None
    s = s.replace({"None": "", "nan": "", "NaN": ""})

    # TR formatını yakala: binlik '.' ondalık ','
    # önce binlik '.' kaldır, sonra ',' -> '.'
    tr_like = s.str.contains(r"\d+\.\d+,\d+", na=False)
    s.loc[tr_like] = (
        s.loc[tr_like]
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    # EN format: binlik ',' ondalık '.'
    en_like = s.str.contains(r"\d+,\d+\.\d+", na=False)
    s.loc[en_like] = s.loc[en_like].str.replace(",", "", regex=False)

    # kalan durumlarda: eğer sadece ',' varsa ondalık kabul et
    only_comma = (~tr_like) & (~en_like) & s.str.contains(",", na=False) & (~s.str.contains("\.", na=False))
    s.loc[only_comma] = s.loc[only_comma].str.replace(",", ".", regex=False)

    out = pd.to_numeric(s, errors="coerce")
    return out

def clean_ets_input(df: pd.DataFrame, fueltype: str | None = None) -> tuple[pd.DataFrame, dict]:
    """
    ETS modeline girecek veri temizliği.
    - kolon adlarını standardize eder
    - Plant boşsa satırı atar
    - Generation/Emissions numeric parse eder
    - Generation<=0 veya Emissions<0 satırlarını atar
    - FuelType yoksa parametreden basar
    """
    report = {
        "rows_in": int(len(df)),
        "rows_out": None,
        "dropped_missing_plant": 0,
        "dropped_invalid_generation": 0,
        "dropped_invalid_emissions": 0,
        "filled_fueltype": 0,
        "coerced_numeric_gen_na": 0,
        "coerced_numeric_emi_na": 0,
    }

    df = standardize_columns(df)

    # FuelType ekle
    if "FuelType" not in df.columns:
        if fueltype is not None:
            df["FuelType"] = str(fueltype)
            report["filled_fueltype"] = int(len(df))
        else:
            df["FuelType"] = "Unknown"
            report["filled_fueltype"] = int(len(df))

    # Plant
    if "Plant" not in df.columns:
        df["Plant"] = ""

    df["Plant"] = df["Plant"].astype(str).str.strip()
    before = len(df)
    df = df[df["Plant"] != ""]
    report["dropped_missing_plant"] = int(before - len(df))

    # Numeric parse
    if "Generation_MWh" not in df.columns:
        df["Generation_MWh"] = np.nan
    if "Emissions_tCO2" not in df.columns:
        df["Emissions_tCO2"] = np.nan

    gen = parse_numeric(df["Generation_MWh"])
    emi = parse_numeric(df["Emissions_tCO2"])

    report["coerced_numeric_gen_na"] = int(gen.isna().sum())
    report["coerced_numeric_emi_na"] = int(emi.isna().sum())

    df["Generation_MWh"] = gen
    df["Emissions_tCO2"] = emi

    # invalid generation (<=0 veya NaN)
    before = len(df)
    df = df[df["Generation_MWh"].fillna(-1) > 0]
    report["dropped_invalid_generation"] = int(before - len(df))

    # invalid emissions (<0 veya NaN)
    before = len(df)
    df = df[df["Emissions_tCO2"].fillna(-1) >= 0]
    report["dropped_invalid_emissions"] = int(before - len(df))

    report["rows_out"] = int(len(df))
    return df.reset_index(drop=True), report
   def filter_intensity_outliers_by_fuel(
    df: pd.DataFrame,
    lower_pct: float = 1.0,
    upper_pct: float = 1.0,
) -> tuple[pd.DataFrame, dict]:
    """
    Yakıt bazında üretim ağırlıklı benchmark (B_fuel) hesaplar.
    Santral yoğunluğu I = Emissions/Generation,
    eğer I, [B*(1-lower_pct), B*(1+upper_pct)] dışında ise satırı çıkarır.

    lower_pct=1.0 => alt sınır: B*(1-1.0)=0
    upper_pct=1.0 => üst sınır: B*(1+1.0)=2B  (yani %100 üst)
    """
    report = {
        "outlier_rule_applied": True,
        "outliers_dropped": 0,
        "fuel_benchmarks": {},
        "lower_pct": float(lower_pct),
        "upper_pct": float(upper_pct),
    }

    required = ["FuelType", "Generation_MWh", "Emissions_tCO2"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Outlier filtresi için kolon eksik: {c}")

    d = df.copy()
    d["intensity"] = d["Emissions_tCO2"] / d["Generation_MWh"]

    # Yakıt bazında üretim ağırlıklı benchmark
    bench = (
        d.groupby("FuelType")
        .apply(lambda g: g["Emissions_tCO2"].sum() / g["Generation_MWh"].sum())
        .to_dict()
    )
    report["fuel_benchmarks"] = {k: float(v) for k, v in bench.items()}

    # Sınırları hesapla ve filtrele
    d["B_fuel_clean"] = d["FuelType"].map(bench)
    low = d["B_fuel_clean"] * (1.0 - float(lower_pct))
    high = d["B_fuel_clean"] * (1.0 + float(upper_pct))

    before = len(d)
    d = d[(d["intensity"] >= low) & (d["intensity"] <= high)].copy()
    report["outliers_dropped"] = int(before - len(d))

    # yardımcı kolonları temizle
    d = d.drop(columns=["B_fuel_clean"], errors="ignore")
    return d.reset_index(drop=True), report
