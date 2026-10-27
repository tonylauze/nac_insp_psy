"""Build transparent, department-level mental-health need projections.

The program uses only:
* population projections by age and sex from Insee Omphale workbooks;
* epidemiological parameters maintained in CSV files next to the workbooks;
* optional, explicit territorial multipliers.

It deliberately does not infer need from health-care utilisation. Different
indicators are kept separate because their prevalences overlap and must not be
summed without a validated comorbidity model.
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd


HEADER_ROW = 5  # zero-based row containing ZONE, SEXE, TRAGE and POP_YYYY
YEAR_PATTERN = re.compile(r"^POP_(\d{4})$")
AGE_PATTERN = re.compile(r"\[(\d+);(\d+)(?:\[|\+\])")
SCENARIO_FILES = {
    "Central": "donnees_det_Central.xlsx",
    "Esperance_de_vie_basse": "donnees_det_EDV_basse.xlsx",
    "Esperance_de_vie_haute": "donnees_det_EDV_haute.xlsx",
    "Fecondite_basse": "donnees_det_Fec_basse.xlsx",
    "Fecondite_haute": "donnees_det_Fec_haute.xlsx",
    "Migrations_basses": "donnees_det_Mig_basse.xlsx",
    "Migrations_hautes": "donnees_det_Mig_haute.xlsx",
    "Population_agee": "donnees_det_Pop_agee.xlsx",
    "Population_basse": "donnees_det_Pop_basse.xlsx",
    "Population_haute": "donnees_det_Pop_haute.xlsx",
    "Population_jeune": "donnees_det_Pop_jeune.xlsx",
}
SEX_CODES = {"hommes": 1, "femmes": 2}
REQUIRED_PARAMETER_COLUMNS = {
    "indicator",
    "label",
    "source",
    "sex",
    "age_min",
    "age_max",
    "baseline_rate",
}


@dataclass(frozen=True)
class AgeBand:
    lower: int
    upper: int


def parse_age_band(label: object) -> AgeBand:
    """Convert Insee labels such as ``[15;20[`` and ``[95;99+]`` to ranges."""
    match = AGE_PATTERN.fullmatch(str(label).strip())
    if not match:
        raise ValueError(f"Tranche d'âge Insee non reconnue : {label!r}")
    lower, upper = (int(value) for value in match.groups())
    return AgeBand(lower, upper)


def overlap_share(band: AgeBand, age_min: int, age_max: int) -> float:
    """Share of a five-year Insee band belonging to an inclusive age interval."""
    requested_upper = age_max + 1
    overlap = max(0, min(band.upper, requested_upper) - max(band.lower, age_min))
    return overlap / (band.upper - band.lower)


def read_population_workbook(path: Path) -> tuple[pd.DataFrame, list[int]]:
    frame = pd.read_excel(path, sheet_name="Population", header=HEADER_ROW)
    year_columns = [column for column in frame.columns if YEAR_PATTERN.fullmatch(str(column))]
    if not year_columns:
        raise ValueError(f"Aucune colonne POP_YYYY dans {path.name}")
    years = [int(str(column).split("_")[1]) for column in year_columns]
    required = ["ZONE", "SEXE", "TRAGE", *year_columns]
    frame = frame.loc[:, required].dropna(subset=["ZONE", "SEXE", "TRAGE"])
    frame["SEXE"] = pd.to_numeric(frame["SEXE"], errors="raise").astype(int)
    frame["_band"] = frame["TRAGE"].map(parse_age_band)
    return frame, years


def validate_parameters(parameters: pd.DataFrame) -> pd.DataFrame:
    missing = REQUIRED_PARAMETER_COLUMNS.difference(parameters.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans prevalence_parameters.csv : {sorted(missing)}")
    parameters = parameters.copy()
    parameters["sex"] = parameters["sex"].str.strip().str.lower()
    allowed_sexes = {*SEX_CODES, "tous"}
    invalid = sorted(set(parameters["sex"]) - allowed_sexes)
    if invalid:
        raise ValueError(f"Valeurs sex invalides : {invalid}; valeurs admises : {sorted(allowed_sexes)}")
    for column in ("age_min", "age_max", "baseline_rate"):
        parameters[column] = pd.to_numeric(parameters[column], errors="raise")
    if (parameters["age_min"] > parameters["age_max"]).any():
        raise ValueError("Chaque age_min doit être inférieur ou égal à age_max.")
    if ((parameters["baseline_rate"] < 0) | (parameters["baseline_rate"] > 1)).any():
        raise ValueError("Chaque baseline_rate doit être compris entre 0 et 1.")
    return parameters


def read_need_scenarios(path: Path) -> pd.DataFrame:
    scenarios = pd.read_csv(path)
    expected = {"need_scenario", "annual_relative_change", "min_multiplier", "max_multiplier"}
    missing = expected.difference(scenarios.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans need_scenarios.csv : {sorted(missing)}")
    for column in expected - {"need_scenario"}:
        scenarios[column] = pd.to_numeric(scenarios[column], errors="raise")
    if (scenarios["min_multiplier"] <= 0).any() or (scenarios["max_multiplier"] <= 0).any():
        raise ValueError("Les multiplicateurs de scénario doivent être strictement positifs.")
    if (scenarios["min_multiplier"] > scenarios["max_multiplier"]).any():
        raise ValueError("min_multiplier doit être inférieur ou égal à max_multiplier.")
    return scenarios


def read_territorial_modifiers(path: Path | None) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame(columns=["department", "indicator", "territorial_multiplier"])
    modifiers = pd.read_csv(path)
    expected = {"department", "indicator", "territorial_multiplier"}
    missing = expected.difference(modifiers.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans territorial_modifiers.csv : {sorted(missing)}")
    modifiers = modifiers.loc[:, ["department", "indicator", "territorial_multiplier"]].copy()
    modifiers["territorial_multiplier"] = pd.to_numeric(
        modifiers["territorial_multiplier"], errors="raise"
    )
    if (modifiers["territorial_multiplier"] <= 0).any():
        raise ValueError("Les multiplicateurs territoriaux doivent être strictement positifs.")
    if modifiers.duplicated(["department", "indicator"]).any():
        raise ValueError("territorial_modifiers.csv contient des doublons département × indicateur.")
    return modifiers


def at_risk_population(population: pd.DataFrame, parameter: pd.Series, years: list[int]) -> pd.DataFrame:
    """Aggregate the appropriate sex and fractional age bands for one indicator."""
    selected = population.copy()
    if parameter.sex != "tous":
        selected = selected.loc[selected["SEXE"] == SEX_CODES[parameter.sex]]
    selected = selected.copy()
    selected["_share"] = selected["_band"].map(
        lambda band: overlap_share(band, int(parameter.age_min), int(parameter.age_max))
    )
    selected = selected.loc[selected["_share"] > 0]
    columns = [f"POP_{year}" for year in years]
    selected.loc[:, columns] = selected.loc[:, columns].mul(selected["_share"], axis=0)
    grouped = selected.groupby("ZONE", as_index=False)[columns].sum()
    return grouped.melt(id_vars="ZONE", var_name="_year", value_name="population_at_risk").assign(
        year=lambda data: data["_year"].str.extract(r"(\d{4})").astype(int),
        department=lambda data: data["ZONE"].astype(str),
    )[["department", "year", "population_at_risk"]]


def project_workbook(
    population: pd.DataFrame,
    years: list[int],
    demographic_scenario: str,
    parameters: pd.DataFrame,
    need_scenarios: pd.DataFrame,
    modifiers: pd.DataFrame,
    base_year: int,
) -> pd.DataFrame:
    records: list[pd.DataFrame] = []
    for parameter in parameters.itertuples(index=False):
        at_risk = at_risk_population(population, parameter, years)
        at_risk["indicator"] = parameter.indicator
        at_risk["label"] = parameter.label
        at_risk["source"] = parameter.source
        at_risk["sex"] = parameter.sex
        at_risk["age_min"] = int(parameter.age_min)
        at_risk["age_max"] = int(parameter.age_max)
        at_risk["baseline_rate"] = float(parameter.baseline_rate)
        at_risk = at_risk.merge(modifiers, how="left", on=["department", "indicator"])
        at_risk["territorial_multiplier"] = pd.to_numeric(
            at_risk["territorial_multiplier"], errors="coerce"
        ).fillna(1.0)
        for scenario in need_scenarios.itertuples(index=False):
            result = at_risk.copy()
            elapsed = result["year"] - base_year
            multiplier = (1.0 + scenario.annual_relative_change) ** elapsed
            result["need_multiplier"] = np.clip(
                multiplier, scenario.min_multiplier, scenario.max_multiplier
            )
            result["prevalence"] = (
                result["baseline_rate"]
                * result["territorial_multiplier"]
                * result["need_multiplier"]
            ).clip(lower=0, upper=1)
            result["persons_with_need"] = result["population_at_risk"] * result["prevalence"]
            result["demographic_scenario"] = demographic_scenario
            result["need_scenario"] = scenario.need_scenario
            records.append(result)
    return pd.concat(records, ignore_index=True)


def build_projection(data_dir: Path, output_dir: Path, base_year: int) -> pd.DataFrame:
    parameters = validate_parameters(pd.read_csv(data_dir / "prevalence_parameters.csv"))
    need_scenarios = read_need_scenarios(data_dir / "need_scenarios.csv")
    modifiers = read_territorial_modifiers(data_dir / "territorial_modifiers.csv")
    results: list[pd.DataFrame] = []
    for scenario, filename in SCENARIO_FILES.items():
        path = data_dir / filename
        if not path.exists():
            raise FileNotFoundError(f"Fichier Insee absent : {path}")
        population, years = read_population_workbook(path)
        results.append(
            project_workbook(
                population, years, scenario, parameters, need_scenarios, modifiers, base_year
            )
        )
    projection = pd.concat(results, ignore_index=True)
    projection = projection.loc[
        :,
        [
            "department",
            "year",
            "demographic_scenario",
            "need_scenario",
            "indicator",
            "label",
            "source",
            "sex",
            "age_min",
            "age_max",
            "population_at_risk",
            "baseline_rate",
            "territorial_multiplier",
            "need_multiplier",
            "prevalence",
            "persons_with_need",
        ],
    ].sort_values(["demographic_scenario", "need_scenario", "indicator", "department", "year"])
    output_dir.mkdir(parents=True, exist_ok=True)
    projection.to_csv(output_dir / "demand_projection.csv", index=False, float_format="%.6f")
    summary_keys = [
        "department",
        "year",
        "demographic_scenario",
        "need_scenario",
        "indicator",
        "label",
        "source",
        "age_min",
        "age_max",
    ]
    summary = (
        projection.groupby(summary_keys, as_index=False)
        .agg(
            population_at_risk=("population_at_risk", "sum"),
            persons_with_need=("persons_with_need", "sum"),
        )
        .assign(prevalence=lambda data: data["persons_with_need"] / data["population_at_risk"])
    )
    summary.to_csv(output_dir / "demand_projection_summary.csv", index=False, float_format="%.6f")
    metadata = pd.DataFrame(
        {
            "field": ["base_year", "method", "health_data"],
            "value": [
                str(base_year),
                "Population Omphale × prevalence epidemiologique × multiplicateurs explicites",
                "Aucune donnee SNDS, PMSI, ATIH ou d'activite de soins n'est utilisee.",
            ],
        }
    )
    metadata.to_csv(output_dir / "metadata.csv", index=False)
    return projection


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=Path("dataDemand"))
    parser.add_argument("--output-dir", type=Path, default=Path("dataDemand/output"))
    parser.add_argument("--base-year", type=int, default=2026)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    projection = build_projection(args.data_dir, args.output_dir, args.base_year)
    print(
        f"Projection créée : {len(projection):,} lignes détaillées dans "
        f"{args.output_dir / 'demand_projection.csv'}"
    )


if __name__ == "__main__":
    main()
