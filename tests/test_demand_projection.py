import unittest

import pandas as pd

from demand_projection.run import AgeBand, at_risk_population, overlap_share, project_workbook


class DemandProjectionTests(unittest.TestCase):
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
