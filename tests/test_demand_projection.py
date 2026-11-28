import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from demand_projection.run import (
    AgeBand, at_risk_population, overlap_share, project_workbook,
    validate_parameters, read_need_scenarios, read_territorial_modifiers,
    read_population_workbook, summarize_projection,
)


class DemandProjectionTests(unittest.TestCase):
    def test_invalid_prevalence_parameters_are_rejected(self):
        valid = dict(indicator="test", label="Test", source="test", sex="tous",
                     age_min=18, age_max=79, baseline_rate=0.1)
        for column, value in [("baseline_rate", float("nan")), ("age_min", float("inf")),
                              ("age_min", 18.5), ("age_min", -1)]:
            with self.subTest(column=column, value=value), self.assertRaises(ValueError):
                validate_parameters(pd.DataFrame([{**valid, column: value}]))
        with self.assertRaises(ValueError):
            validate_parameters(pd.DataFrame([valid, valid]))

    def test_invalid_need_scenarios_are_rejected(self):
        valid = dict(need_scenario="stable", annual_relative_change=0.0,
                     min_multiplier=1.0, max_multiplier=1.0)
        for column, value in [("annual_relative_change", -1), ("annual_relative_change", -2),
                              ("min_multiplier", float("nan")), ("max_multiplier", float("inf"))]:
            with self.subTest(column=column, value=value):
                with patch("demand_projection.run.pd.read_csv", return_value=pd.DataFrame([{**valid, column: value}])):
                    with self.assertRaises(ValueError):
                        read_need_scenarios(Path("scenarios.csv"))

    def test_missing_modifier_is_rejected_instead_of_replaced_by_one(self):
        for value in [float("nan"), float("inf")]:
            with self.subTest(value=value), patch.object(Path, "exists", return_value=True):
                frame = pd.DataFrame([dict(department="Test", indicator="test", territorial_multiplier=value)])
                with patch("demand_projection.run.pd.read_csv", return_value=frame):
                    with self.assertRaises(ValueError):
                        read_territorial_modifiers(Path("modifiers.csv"))

    def test_missing_or_duplicate_population_is_rejected(self):
        valid = dict(ZONE="Test", SEXE=1, TRAGE="[20;25[", POP_2026=100)
        for rows in [[{**valid, "POP_2026": float("nan")}], [valid, valid], [{**valid, "SEXE": 1.5}]]:
            with self.subTest(rows=rows), patch("demand_projection.run.pd.read_excel", return_value=pd.DataFrame(rows)):
                with self.assertRaises(ValueError):
                    read_population_workbook(Path("population.xlsx"))

    def test_zero_population_summary_has_finite_zero_prevalence(self):
        row = dict(department="Test", year=2026, demographic_scenario="Central", need_scenario="stable",
                   indicator="test", label="Test", source="test", age_min=18, age_max=79,
                   population_at_risk=0, persons_with_need=0)
        summary = summarize_projection(pd.DataFrame([row]))
        self.assertEqual(summary.prevalence.iloc[0], 0.0)

    def test_overlap_share_handles_partial_five_year_band(self):
        self.assertEqual(overlap_share(AgeBand(15, 20), 18, 79), 0.4)
        self.assertEqual(overlap_share(AgeBand(75, 80), 18, 79), 1.0)
        self.assertEqual(overlap_share(AgeBand(80, 85), 18, 79), 0.0)

    def test_at_risk_population_aggregates_both_sexes(self):
        population = pd.DataFrame(
            {
                "ZONE": ["Test", "Test", "Test", "Test"],
                "SEXE": [1, 2, 1, 2],
                "TRAGE": ["[15;20[", "[15;20[", "[20;25[", "[20;25["],
                "POP_2026": [100, 200, 300, 400],
                "_band": [AgeBand(15, 20), AgeBand(15, 20), AgeBand(20, 25), AgeBand(20, 25)],
            }
        )
        parameter = pd.Series({"sex": "tous", "age_min": 18, "age_max": 24})
        result = at_risk_population(population, parameter, [2026])
        # 18-19 represents 2/5 of the first band; 20-24 includes the second one.
        self.assertAlmostEqual(result.loc[0, "population_at_risk"], 820.0)

    def test_project_workbook_keeps_indicators_separate(self):
        population = pd.DataFrame(
            {
                "ZONE": ["Test", "Test"],
                "SEXE": [1, 2],
                "TRAGE": ["[20;25[", "[20;25["],
                "POP_2026": [100, 200],
                "_band": [AgeBand(20, 25), AgeBand(20, 25)],
            }
        )
        parameters = pd.DataFrame(
            [
                {
                    "indicator": "a",
                    "label": "A",
                    "source": "test",
                    "sex": "tous",
                    "age_min": 18,
                    "age_max": 79,
                    "baseline_rate": 0.1,
                },
                {
                    "indicator": "b",
                    "label": "B",
                    "source": "test",
                    "sex": "tous",
                    "age_min": 18,
                    "age_max": 79,
                    "baseline_rate": 0.2,
                },
            ]
        )
        scenarios = pd.DataFrame(
            [{"need_scenario": "stable", "annual_relative_change": 0.0, "min_multiplier": 1.0, "max_multiplier": 1.0}]
        )
        output = project_workbook(
            population,
            [2026],
            "Central",
            parameters,
            scenarios,
            pd.DataFrame(columns=["department", "indicator", "territorial_multiplier"]),
            2026,
        )
        self.assertEqual(set(output["indicator"]), {"a", "b"})
        self.assertEqual(len(output), 2)
        self.assertAlmostEqual(output.loc[output["indicator"] == "a", "persons_with_need"].iloc[0], 30.0)
        self.assertAlmostEqual(output.loc[output["indicator"] == "b", "persons_with_need"].iloc[0], 60.0)


if __name__ == "__main__":
    unittest.main()
