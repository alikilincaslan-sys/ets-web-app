import numpy as np
import pandas as pd


def _compute_benchmarks(
    x: pd.DataFrame,
    benchmark_method: str = "best_plants",
    benchmark_top_pct: int = 100,
    cap_col: str = "InstalledCapacity_MW",
) -> dict:
    """
    Returns dict: FuelType -> benchmark intensity (tCO2/MWh)
    Supported: generation_weighted, capacity_weighted, best_plants
    """
    out = {}

    if benchmark_method == "generation_weighted":
        g = x.groupby("FuelType", as_index=False)[["Emissions_tCO2", "Generation_MWh"]].sum()
        g["B"] = g["Emissions_tCO2"] / g["Generation_MWh"].replace(0, np.nan)
        for _, r in g.iterrows():
            out[r["FuelType"]] = float(r["B"]) if np.isfinite(r["B"]) else np.nan
        return out

    if benchmark_method == "capacity_weighted":
        # weighted by InstalledCapacity_MW, needs cap_col
        if cap_col not in x.columns:
            raise ValueError(f"capacity_weighted benchmark requires column '{cap_col}' in input data.")
        # plant-level EI then capacity-weighted average
        plant = (
            x.groupby(["FuelType", "Plant"], as_index=False)[["Emissions_tCO2", "Generation_MWh", cap_col]]
            .sum()
        )
        plant["EI"] = plant["Emissions_tCO2"] / plant["Generation_MWh"].replace(0, np.nan)
        plant = plant.dropna(subset=["EI"])
        g = plant.groupby("FuelType", as_index=False).apply(
            lambda df: np.average(df["EI"], weights=df[cap_col].replace(0, np.nan))
        )
        g = g.reset_index().rename(columns={0: "B"})
        for _, r in g.iterrows():
            out[r["FuelType"]] = float(r["B"]) if np.isfinite(r["B"]) else np.nan
        return out

    if benchmark_method == "best_plants":
        # "best plants" defined by lowest EI covering X% of generation (benchmark_top_pct)
        pct = int(benchmark_top_pct)
        pct = max(10, min(100, pct))

        plant = x.groupby(["FuelType", "Plant"], as_index=False)[["Emissions_tCO2", "Generation_MWh"]].sum()
        plant["EI"] = plant["Emissions_tCO2"] / plant["Generation_MWh"].replace(0, np.nan)
        plant = plant.dropna(subset=["EI"])

        for ft, g in plant.groupby("FuelType"):
            gg = g.sort_values("EI", ascending=True).copy()
            gg["cum_gen"] = gg["Generation_MWh"].cumsum()
            total = float(gg["Generation_MWh"].sum())
            if total <= 0:
                out[ft] = np.nan
                continue
            threshold = total * (pct / 100.0)
            sel = gg[gg["cum_gen"] <= threshold]
            if sel.empty:
                sel = gg.head(1)
            b = float(sel["Emissions_tCO2"].sum() / sel["Generation_MWh"].sum())
            out[ft] = b if np.isfinite(b) else np.nan

        return out

    # fallback
    return _compute_benchmarks(x, benchmark_method="generation_weighted")


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
    tier_best_pct: int = 50,
    price_method: str = "Market Clearing",
    trf: float = 0.0,
    auction_supply_share: float = 1.0,
):
    """
    Returns:
      sonuc_df (plant-level),
      benchmark_map (dict),
      clearing_price (float)
    """

    x = df.copy()

    # Standardize
    for c in ["Emissions_tCO2", "Generation_MWh"]:
        if c not in x.columns:
            raise ValueError(f"Input data must include '{c}' column.")
    if "Plant" not in x.columns:
        raise ValueError("Input data must include 'Plant' column.")
    if "FuelType" not in x.columns:
        raise ValueError("Input data must include 'FuelType' column.")

    x["Emissions_tCO2"] = pd.to_numeric(x["Emissions_tCO2"], errors="coerce")
    x["Generation_MWh"] = pd.to_numeric(x["Generation_MWh"], errors="coerce")
    if cap_col in x.columns:
        x[cap_col] = pd.to_numeric(x[cap_col], errors="coerce")

    x = x.dropna(subset=["Emissions_tCO2", "Generation_MWh"])
    x = x[x["Generation_MWh"] > 0].copy()

    # 1) Plant-level EI
    x["intensity"] = x["Emissions_tCO2"] / x["Generation_MWh"].replace(0, np.nan)

    # 2) Benchmark (yakıt bazında)
    # ------------------------------------------------------------
    # New method: two-tier benchmark (best vs worst) within each fuel.
    # - Split plants by EI within fuel into two tiers (by plant count).
    # - Each tier has its own benchmark (generation-weighted EI).
    # - Allocation is calculated against the plant's tier benchmark.
    # ------------------------------------------------------------
    if benchmark_method == "two_tier":
        tier_best_pct_i = int(tier_best_pct)
        tier_best_pct_i = max(1, min(99, tier_best_pct_i))  # keep two non-empty groups

        # plant-level EI within fuel
        plant_agg = (
            x.groupby(["FuelType", "Plant"], as_index=False)[["Emissions_tCO2", "Generation_MWh"]]
            .sum()
        )
        plant_agg["EI"] = plant_agg["Emissions_tCO2"] / plant_agg["Generation_MWh"].replace(0, np.nan)
        plant_agg = plant_agg.dropna(subset=["EI"])

        tier_rows = []
        benchmark_map = {}

        for ft, g in plant_agg.groupby("FuelType"):
            gg = g.sort_values("EI", ascending=True).copy()
            n = len(gg)
            if n == 0:
                continue

            cut = int(np.ceil(n * (tier_best_pct_i / 100.0)))
            cut = max(1, min(n - 1, cut)) if n >= 2 else 1

            gg["Tier"] = "Best"
            if n >= 2:
                gg.loc[gg.index[cut:], "Tier"] = "Worst"

            # Benchmarks (generation-weighted EI) per tier
            def _tier_bench(df_tier: pd.DataFrame) -> float:
                e = float(df_tier["Emissions_tCO2"].sum())
                gen = float(df_tier["Generation_MWh"].sum())
                return float(e / gen) if gen > 0 else np.nan

            b_best = _tier_bench(gg[gg["Tier"] == "Best"])
            b_worst = _tier_bench(gg[gg["Tier"] == "Worst"])

            # Store in benchmark_map with readable keys
            benchmark_map[f"{ft} | Best tier"] = b_best
            benchmark_map[f"{ft} | Worst tier"] = b_worst

            tier_rows.append(gg[["FuelType", "Plant", "Tier"]])

        tier_df = (
            pd.concat(tier_rows, ignore_index=True)
            if len(tier_rows)
            else pd.DataFrame(columns=["FuelType", "Plant", "Tier"])
        )
        x = x.merge(tier_df, on=["FuelType", "Plant"], how="left")
        x["Tier"] = x["Tier"].fillna("Best")  # fallback

        # Map tier benchmark to rows
        def _tier_benchmark_for_row(row) -> float:
            ft = row["FuelType"]
            if row["Tier"] == "Worst":
                return benchmark_map.get(f"{ft} | Worst tier", np.nan)
            return benchmark_map.get(f"{ft} | Best tier", np.nan)

        x["B_fuel"] = x.apply(_tier_benchmark_for_row, axis=1)

    else:
        benchmark_map = _compute_benchmarks(
            x,
            benchmark_method=benchmark_method,
            benchmark_top_pct=int(benchmark_top_pct),
            cap_col=cap_col,
        )
        x["B_fuel"] = x["FuelType"].map(benchmark_map)

    if "Tier" not in x.columns:
        x["Tier"] = "All"

    # 3) AGK smoothing (effective intensity around benchmark)
    # EI_eff = B + (EI - B)*(1-AGK)
    x["EI_eff"] = x["B_fuel"] + (x["intensity"] - x["B_fuel"]) * (1 - float(agk))

    # 4) Allocation & net ETS (tCO2)
    # Allowances = B_fuel * Generation
    x["alloc"] = x["B_fuel"] * x["Generation_MWh"]
    x["net_ets"] = x["Emissions_tCO2"] - x["alloc"]

    # 5) Optional transition compensation (TRF): reduce net_ets if intensity > benchmark
    # (Only for net buyers; matches your pilot notion)
    trf = float(trf)
    if trf > 0:
        x["net_ets"] = np.where(
            x["intensity"] > x["B_fuel"],
            x["net_ets"] * (1 - trf),
            x["net_ets"],
        )

    # 6) Price formation inputs (simple bid/ask construction)
    # bids for buyers: p_bid = clamp(price_min, price_max, ...)
    # asks for sellers: p_ask = clamp(price_min, price_max, ...)
    # NOTE: This is a stylized market-curve visualization; "Auction Clearing" uses fixed supply share.
    net = x["net_ets"].astype(float)

    # Construct stylized willingness-to-pay/accept curves
    # buyers: higher net => higher WTP; sellers: higher surplus => lower ask (or vice versa)
    # We'll keep it monotone but bounded
    buyers = x[net > 0].copy()
    sellers = x[net < 0].copy()

    def _clamp(p):
        return float(np.clip(p, float(price_min), float(price_max)))

    if not buyers.empty:
        q = buyers["net_ets"].astype(float)
        qn = (q - q.min()) / (q.max() - q.min() + 1e-9)
        buyers["p_bid"] = qn.apply(lambda z: _clamp(price_min + (price_max - price_min) * (0.2 + 0.8 * z)))
        buyers["p_bid"] = buyers["p_bid"] + float(spread) / 2.0
    else:
        buyers["p_bid"] = np.nan

    if not sellers.empty:
        q = (-sellers["net_ets"].astype(float))
        qn = (q - q.min()) / (q.max() - q.min() + 1e-9)
        sellers["p_ask"] = qn.apply(lambda z: _clamp(price_min + (price_max - price_min) * (0.2 + 0.8 * z)))
        sellers["p_ask"] = sellers["p_ask"] - float(spread) / 2.0
    else:
        sellers["p_ask"] = np.nan

    x = pd.concat([buyers, sellers], ignore_index=True)

    # 7) Clearing price
    demand = float(x.loc[x["net_ets"] > 0, "net_ets"].sum())
    supply_surplus = float((-x.loc[x["net_ets"] < 0, "net_ets"]).sum())

    clearing_price = float(price_min)

    if price_method == "Average Compliance Cost":
        # Simple average cost proxy: mean of buyer bids (bounded)
        if demand > 0 and "p_bid" in x.columns:
            clearing_price = float(np.nanmean(x.loc[x["net_ets"] > 0, "p_bid"]))
        clearing_price = float(np.clip(clearing_price, float(price_min), float(price_max)))

    elif price_method == "Auction Clearing":
        # demand assumed inelastic; supply is policy-defined share of demand
        supply = float(demand * float(auction_supply_share))
        traded = min(demand, supply)
        # map scarcity to price range (stylized)
        if demand > 0:
            scarcity = 1.0 - (supply / demand)
            scarcity = float(np.clip(scarcity, -1.0, 1.0))
            # higher scarcity -> higher price
            clearing_price = float(price_min + (price_max - price_min) * np.clip(0.5 + 0.5 * scarcity, 0, 1))
        else:
            clearing_price = float(price_min)
        clearing_price = float(np.clip(clearing_price, float(price_min), float(price_max)))

    else:  # Market Clearing
        # Intersection is visualized in Streamlit. Here: a robust proxy:
        # - If supply >= demand -> price closer to lower bound
        # - If supply < demand -> price closer to upper bound
        if demand <= 0:
            clearing_price = float(price_min)
        else:
            ratio = supply_surplus / demand if demand > 0 else 0.0
            # ratio <1 => scarcity => higher price
            scarcity = 1.0 - ratio
            scarcity = float(np.clip(scarcity, -1.0, 1.0))
            clearing_price = float(price_min + (price_max - price_min) * np.clip(0.5 + 0.5 * scarcity, 0, 1))
        clearing_price = float(np.clip(clearing_price, float(price_min), float(price_max)))

    # 8) Apply price to get € impacts
    x["ets_cost_€"] = x["net_ets"] * clearing_price
    x["ets_net_cashflow_€"] = x["ets_cost_€"]
    x["ets_net_cashflow_€/MWh"] = x["ets_net_cashflow_€"] / x["Generation_MWh"].replace(0, np.nan)

    # 9) Plant-level summary output (one row per plant)
    agg_cols = {
        "FuelType": "first",
        "Tier": "first",
        "Emissions_tCO2": "sum",
        "Generation_MWh": "sum",
        "intensity": "mean",
        "B_fuel": "mean",
        "EI_eff": "mean",
        "alloc": "sum",
        "net_ets": "sum",
        "ets_cost_€": "sum",
        "ets_net_cashflow_€": "sum",
    }
    if cap_col in x.columns:
        agg_cols[cap_col] = "sum"
    if "p_bid" in x.columns:
        agg_cols["p_bid"] = "mean"
    if "p_ask" in x.columns:
        agg_cols["p_ask"] = "mean"

    sonuc_df = x.groupby("Plant", as_index=False).agg(agg_cols)

    # recompute per-MWh
    sonuc_df["ets_net_cashflow_€/MWh"] = sonuc_df["ets_net_cashflow_€"] / sonuc_df["Generation_MWh"].replace(0, np.nan)

    # keep consistent columns order
    cols = ["Plant"]
    if cap_col in sonuc_df.columns:
        cols += [cap_col]
    cols += [
        "FuelType",
        "Tier",
        "Generation_MWh",
        "Emissions_tCO2",
        "intensity",
        "B_fuel",
        "EI_eff",
        "alloc",
        "net_ets",
        "p_bid" if "p_bid" in sonuc_df.columns else None,
        "p_ask" if "p_ask" in sonuc_df.columns else None,
        "ets_net_cashflow_€",
        "ets_net_cashflow_€/MWh",
    ]
    cols = [c for c in cols if c is not None and c in sonuc_df.columns]
    sonuc_df = sonuc_df[cols]

    return sonuc_df, benchmark_map, clearing_price
