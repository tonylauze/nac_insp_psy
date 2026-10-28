"""Validate demand outcomes and construct the explicitly disjoint adult scenario."""

import numpy as np
import pandas as pd


GROUP_KEYS = ["department", "year", "demographic_scenario", "need_scenario"]
ADULT_COMPONENTS = ("episode_depressif_12_mois", "trouble_anxieux_generalise_12_mois")
ADULT_TOTAL = "besoin_total_sans_recouvrement"
NUMERIC_COLUMNS = ["year", "population_at_risk", "persons_with_need", "prevalence"]


def prepare_outcomes(data: pd.DataFrame, base_year: int = 2026) -> pd.DataFrame:
    """Reject inconsistent inputs rather than silently dropping people or scenarios."""
    required = set(GROUP_KEYS + NUMERIC_COLUMNS + ["indicator", "label"])
    missing = required.difference(data.columns)
    if missing:
        raise ValueError("Colonnes de demande absentes : " + ", ".join(sorted(missing)))
    data = data.copy()
    if data[list(required)].isna().any().any():
        raise ValueError("Les résultats de demande contiennent des valeurs manquantes.")
    for column in NUMERIC_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="raise")
    if not np.isfinite(data[NUMERIC_COLUMNS].to_numpy()).all():
        raise ValueError("Les résultats de demande contiennent des valeurs non finies.")
    if (data["year"] % 1 != 0).any():
        raise ValueError("Les années de demande doivent être entières.")
    if (data[["population_at_risk", "persons_with_need"]] < 0).any().any():
        raise ValueError("Les effectifs de demande doivent être positifs ou nuls.")
    if not data["prevalence"].between(0, 1).all():
        raise ValueError("Les prévalences doivent être comprises entre 0 et 1.")
    if not np.allclose(data["persons_with_need"], data["population_at_risk"] * data["prevalence"], atol=1, rtol=1e-5):
        raise ValueError("Les effectifs de demande ne correspondent pas aux prévalences.")
    if data.duplicated(GROUP_KEYS + ["indicator"]).any():
        raise ValueError("Doublons département × année × scénarios × indicateur de demande.")
    data["year"] = data["year"].astype(int)
    data = data.loc[(data["year"] >= base_year) & (data["indicator"] != ADULT_TOTAL)].copy()
    if data.empty:
        raise ValueError(f"Aucun résultat de demande à partir de {base_year}.")

    components = data[data["indicator"].isin(ADULT_COMPONENTS)]
    if components.empty:
        return data
    grouped = components.groupby(GROUP_KEYS, sort=False)
    if not grouped["indicator"].nunique().eq(2).all():
        raise ValueError("Le cumul adulte nécessite dépression et anxiété pour chaque territoire, année et scénario.")
    populations = grouped["population_at_risk"].agg(["min", "max"])
    if not np.allclose(populations["min"], populations["max"]):
        raise ValueError("Les composantes du cumul adulte doivent couvrir la même population.")
    if {"age_min", "age_max"}.issubset(components.columns):
        if not ((components["age_min"] == 18) & (components["age_max"] == 79)).all():
            raise ValueError("Le cumul adulte est défini uniquement sur les 18–79 ans.")
    total = grouped.agg(
        population_at_risk=("population_at_risk", "first"),
        persons_with_need=("persons_with_need", "sum"),
    ).reset_index()
    if (total["persons_with_need"] > total["population_at_risk"] + 1e-6).any():
        raise ValueError("L'hypothèse sans recouvrement dépasse la population adulte concernée.")
    total["indicator"] = ADULT_TOTAL
    total["label"] = "Besoin total (dépression + anxiété, sans recouvrement)"
    total["age_min"], total["age_max"] = 18, 79
    total["prevalence"] = total["persons_with_need"].div(total["population_at_risk"].replace(0, np.nan)).fillna(0)
    return pd.concat([data, total], ignore_index=True, sort=False)
