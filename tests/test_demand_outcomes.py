import unittest

import pandas as pd

from demand_projection.outcomes import ADULT_COMPONENTS, ADULT_TOTAL, prepare_outcomes
from demand_projection.run import AgeBand, at_risk_population


class DemandOutcomesTests(unittest.TestCase):
    def setUp(self):
        self.data = pd.DataFrame([
            dict(department="Paris", year=2026, demographic_scenario="Central",
                 need_scenario="stable", indicator=indicator, label=indicator,
                 population_at_risk=population, persons_with_need=need,
                 prevalence=need / population, age_min=age_min, age_max=age_max)
            for indicator, population, need, age_min, age_max in [
                (ADULT_COMPONENTS[0], 1000, 150, 18, 79),
                (ADULT_COMPONENTS[1], 1000, 60, 18, 79),
                ("trouble_mental_probable_6_11", 100, 13, 6, 11),
            ]
        ])

    def test_total_adds_adults_only_and_counts_population_once(self):
        result = prepare_outcomes(self.data)
        total = result.loc[result.indicator == ADULT_TOTAL].iloc[0]
        self.assertEqual(total.persons_with_need, 210)
        self.assertEqual(total.population_at_risk, 1000)
        self.assertAlmostEqual(total.prevalence, 0.21)
        self.assertEqual(len(result), 4)

    def test_missing_component_in_one_department_is_rejected(self):
        extra = self.data.iloc[[0]].assign(department="Ain")
        with self.assertRaisesRegex(ValueError, "cumul adulte"):
            prepare_outcomes(pd.concat([self.data, extra], ignore_index=True))

    def test_different_adult_populations_are_rejected(self):
        self.data.loc[1, ["population_at_risk", "persons_with_need"]] = [2000, 120]
        with self.assertRaisesRegex(ValueError, "même population"):
            prepare_outcomes(self.data)

    def test_duplicates_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "Doublons"):
            prepare_outcomes(pd.concat([self.data, self.data.iloc[[0]]]))

    def test_empty_future_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Aucun résultat"):
            prepare_outcomes(self.data.assign(year=2020))

    def test_invalid_numbers_are_rejected(self):
        for value in [float("nan"), float("inf"), -1, "invalid"]:
            with self.subTest(value=value), self.assertRaises(ValueError):
                data = self.data.copy()
                data["persons_with_need"] = value
                prepare_outcomes(data)

    def test_children_age_band_includes_only_six_to_eleven(self):
        population = pd.DataFrame({
            "ZONE": ["Test"] * 3, "SEXE": [1] * 3,
            "_band": [AgeBand(0, 5), AgeBand(5, 10), AgeBand(10, 15)],
            "POP_2026": [500, 500, 500],
        })
        parameter = pd.Series({"sex": "tous", "age_min": 6, "age_max": 11})
        result = at_risk_population(population, parameter, [2026])
        self.assertAlmostEqual(result.population_at_risk.iloc[0], 600)


if __name__ == "__main__":
    unittest.main()
