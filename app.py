"""Carte des psychiatres : calcul en mémoire, propre à chaque session utilisateur.

Installation : pip install streamlit plotly pandas numpy openpyxl
Lancement   : streamlit run app.py

Placer les deux classeurs et requirements.txt à côté de ce script. Sur Streamlit
Cloud, requirements.txt installe notamment openpyxl, nécessaire pour lire Excel.

Les résultats sont un DataFrame de cinq colonnes :
    dep, annee, scenario, indicateur, valeur
Ils restent dans st.session_state, sans aucun fichier de résultats ou paramètres.
Chaque utilisateur dispose de ses propres paramètres et projections. Changer
seulement un filtre d'affichage ne relance pas la simulation. Les paramètres sont
appliqués après Valider et reviennent aux valeurs initiales dans une nouvelle session.

Si les projections régionales de psychiatrie manquent, repli AUTOMATIQUE :
    N_dep(t, s) = N_RPPS_dep(b) * N_DREES_France(t, s) / N_DREES_France(b, s)
où b est la dernière année observée, 2026 dans les fichiers fournis. La croissance
nationale est appliquée aux régions puis répartie à parts départementales fixes.
L'interface indique la méthode utilisée. Si des projections régionales détaillées
existent, elles restent prioritaires ; des données régionales invalides déclenchent
une erreur explicite et ne sont pas silencieusement remplacées.

Les résultats portent sur les effectifs de psychiatres et leur répartition par
statut. La part salariée publique est en %. Les nombres fractionnaires restent
des effectifs attendus. Les taux nationaux de postes non pourvus sont intégrés
au code avec leurs sources ; aucun fichier supplémentaire n’est requis.

Extension : ajouter une fonction et une entrée à modulation_registry, dans ce fichier.
Dynamiques départementales et places de formation : voir README.md pour les formules.
"""
from dataclasses import dataclass
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from typing import Callable
import json
import re
import unicodedata
import urllib.request

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import openpyxl
except ModuleNotFoundError as exc:
    if exc.name != "openpyxl":
        raise
    openpyxl = None


def require_excel_reader():
    if openpyxl is None:
        raise RuntimeError(
            "La dépendance openpyxl manque. Ajouter openpyxl>=3.1,<4 dans "
            "requirements.txt, à côté de app.py sur GitHub, puis faire un commit "
            "et un push pour que Streamlit Cloud installe les dépendances."
        )


# Chemins à adapter seulement si les fichiers sont dans un autre dossier.
ROOT = Path(__file__).resolve().parent
RPPS_PATH = ROOT / "Médecins RPPS 2012-2026.xlsx"
DREES_PATH = ROOT / "DREES_projections_medecins_2026.xlsx"
GEOJSON_URL = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements-version-simplifiee.geojson"
COLUMNS = ["dep", "annee", "scenario", "indicateur", "valeur"]
COUNT = "Effectif de psychiatres"
SALARIED_SHARE = "Part de salariés"
LIBERAL_COUNT = "Nombre brut de libéraux"
SALARIED_COUNT = "Nombre brut de salariés"
INDICATORS = {"effectif": COUNT, "part_salaries": SALARIED_SHARE,
              "liberaux": LIBERAL_COUNT, "salaries": SALARIED_COUNT}
# Groupes exclusifs : salariés hospitaliers + autres salariés ; libéraux + mixtes.
EXERCISE_GROUPS = {"ensemble": "total", "liberal": "liberal_exclusif",
                   "mixte": "mixte", "salariehospitalier": "hospitalier",
                   "autresalarie": "autre_salarie"}
# Série nationale publiée, pas une série contenue dans les classeurs DREES.
# Sources et champs détaillés dans README.md ; les taux sont des fractions.
VACANCY_SOURCES = {
    "cnup_2024": "https://csa.eu/wp-content/uploads/2024/01/DossierPresse_ChoisirPsychiatrie_CNUP_Final_Janv_24.pdf",
    "an_2024": "https://www.assemblee-nationale.fr/dyn/17/rapports/cion-soc/l17b0714_rapport-information.pdf",
    "cnup_2025": "https://www.santementale.fr/2025/09/internes-en-psychiatrie-15-des-postes-restent-vacants-en-2025/",
    "cnup_2026": "https://fr.linkedin.com/posts/olivier-bonnot-84804584_choisirpsychiatrie-nuitsdelapsychiatrie-activity-7510984553386209281-GRtp",
}
VACANCY_OBSERVATIONS = (
    (2013, 0.023, "cnup_2024"), (2014, 0.043, "cnup_2024"),
    (2015, 0.029, "cnup_2024"), (2016, 0.016, "cnup_2024"),
    (2017, 0.014, "cnup_2024"), (2018, 0.036, "cnup_2024"),
    (2019, 0.175, "cnup_2024"), (2020, 0.105, "cnup_2024"),
    (2021, 0.132, "cnup_2024"), (2022, 0.060, "cnup_2024"),
    (2023, 67 / 547, "cnup_2024"), (2024, 65 / 489, "an_2024"),
    (2025, 86 / 554, "cnup_2025"), (2026, 37 / 653, "cnup_2026"),
)
METRO = {f"{n:02d}" for n in range(1, 96) if n != 20} | {"2A", "2B"}
DES_YEARS = 5
STUDY_YEARS = 6 + DES_YEARS  # Durée nominale ; aucune moyenne empirique fournie.
SMOOTH_FLOOR = 1.0  # Psychiatre ; raccord C1 à la fonction identité.
REGIONS = {
    "11": "Île-de-France", "24": "Centre-Val de Loire",
    "27": "Bourgogne-Franche-Comté", "28": "Normandie",
    "32": "Hauts-de-France", "44": "Grand Est", "52": "Pays de la Loire",
    "53": "Bretagne", "75": "Nouvelle-Aquitaine", "76": "Occitanie",
    "84": "Auvergne-Rhône-Alpes", "93": "Provence-Alpes-Côte d’Azur", "94": "Corse",
}


def normalize(value):
    value = re.sub(r"^\s*\d+\s*-\s*", "", str(value))
    value = unicodedata.normalize("NFKD", value.lower())
    value = "".join(c for c in value if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", value)


# Correspondance département -> anciennes régions lorsque le classeur DREES
# emploie le découpage antérieur à 2016. Les codes régionaux DREES ne sont pas
# assimilés aux codes régionaux Insee : on rapproche les libellés.
OLD_REGIONS = {
    "Alsace": "67 68", "Aquitaine": "24 33 40 47 64",
    "Auvergne": "03 15 43 63", "Basse-Normandie": "14 50 61",
    "Bourgogne": "21 58 71 89", "Bretagne": "22 29 35 56",
    "Centre": "18 28 36 37 41 45", "Champagne-Ardenne": "08 10 51 52",
    "Corse": "2A 2B", "Franche-Comté": "25 39 70 90",
    "Haute-Normandie": "27 76", "Ile-de-France": "75 77 78 91 92 93 94 95",
    "Languedoc-Roussillon": "11 30 34 48 66", "Limousin": "19 23 87",
    "Lorraine": "54 55 57 88", "Midi-Pyrénées": "09 12 31 32 46 65 81 82",
    "Nord-Pas-de-Calais": "59 62", "Pays de la Loire": "44 49 53 72 85",
    "Picardie": "02 60 80", "Poitou-Charentes": "16 17 79 86",
    "Provence-Alpes-Côte d'Azur": "04 05 06 13 83 84",
    "Rhône-Alpes": "01 07 26 38 42 69 73 74",
}


def regional_psychiatry(path):
    """Trouver l'onglet par son contenu, même si son nom varie."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    candidates, diagnostics = [], []
    try:
        for sheet in book:
            rows = sheet.iter_rows(values_only=True)
            positions = None
            # La source actuelle a ses en-têtes ligne 7.
            for _ in range(30):
                row = next(rows, None)
                if row is None:
                    break
                p = {normalize(v): i for i, v in enumerate(row) if v is not None}
                if {"annee", "scenario", "region", "effectif"} <= p.keys():
                    positions = p
                    break
            if positions is None:
                continue
            spec = next((k for k in ("specialite", "specialites", "specialiteagregee", "specialitesagregees") if k in positions), None)
            if spec is None or "statut" not in positions:
                continue
            labels, records = set(), []
            for row in rows:
                if not isinstance(row[positions["annee"]], (int, float)):
                    continue
                speciality = row[positions[spec]]
                labels.add(str(speciality))
                if normalize(speciality) != "psychiatrie":
                    continue
                if normalize(row[positions["statut"]]) != "medecinsenactivite":
                    continue
                scenario, region = row[positions["scenario"]], row[positions["region"]]
                if normalize(scenario) == "observe" or normalize(region) == "ensemble":
                    continue
                records.append({"region_source": normalize(region),
                                "annee": int(row[positions["annee"]]),
                                "scenario": re.sub(r"^\d+-", "", str(scenario)),
                                "effectif_region": float(row[positions["effectif"]])})
            diagnostics.append(f"{sheet.title.strip()} : {', '.join(sorted(labels))}")
            if records:
                candidates.append((sheet.title, pd.DataFrame(records)))
    finally:
        book.close()
    if len(candidates) != 1:
        detail = " ; ".join(diagnostics) or "Aucun onglet effectif × région × spécialité trouvé."
        if not candidates:
            return None, detail
        raise ValueError("Plusieurs onglets régionaux de psychiatrie trouvés : " + ", ".join(n for n, _ in candidates))
    sheet_name, result = candidates[0]
    if result.duplicated(["region_source", "annee", "scenario"]).any():
        raise ValueError("L'onglet régional contient plusieurs valeurs pour une même région, année et scénario.")
    if not np.isfinite(result.effectif_region).all() or (result.effectif_region < 0).any():
        raise ValueError("Effectif régional invalide.")
    return result, sheet_name


def national_psychiatry(path):
    """Lire les effectifs nationaux de psychiatres actifs, sans les étudiants."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    candidates = []
    try:
        for sheet in book:
            rows = sheet.iter_rows(values_only=True)
            positions = None
            for _ in range(30):
                row = next(rows, None)
                if row is None:
                    break
                p = {normalize(v): i for i, v in enumerate(row) if v is not None}
                if {"annee", "scenario", "statut", "effectif"} <= p.keys() and "region" not in p:
                    positions = p
                    break
            if positions is None:
                continue
            spec = next((k for k in ("specialite", "specialites") if k in positions), None)
            if spec is None:
                continue
            records = []
            for row in rows:
                if not isinstance(row[positions["annee"]], (int, float)):
                    continue
                if normalize(row[positions[spec]]) != "psychiatrie":
                    continue
                if normalize(row[positions["statut"]]) != "medecinsenactivite":
                    continue
                scenario = str(row[positions["scenario"]])
                if normalize(scenario) == "observe":
                    continue
                records.append({"annee": int(row[positions["annee"]]),
                                "scenario": re.sub(r"^\d+-", "", scenario),
                                "effectif_national": float(row[positions["effectif"]])})
            if records:
                candidates.append((sheet.title, pd.DataFrame(records)))
    finally:
        book.close()
    if len(candidates) != 1:
        raise ValueError("Un seul onglet national de projections de psychiatrie doit être disponible pour le repli.")
    sheet, result = candidates[0]
    if result.duplicated(["annee", "scenario"]).any():
        raise ValueError("Projections nationales de psychiatrie dupliquées.")
    if not np.isfinite(result.effectif_national).all() or (result.effectif_national < 0).any():
        raise ValueError("Effectif national invalide.")
    return result, sheet


def national_to_regions(history, names, national):
    """Rétablir le raccord à l'année observée et la répartition régionale fixe."""
    anchor = int(history.annee.max())
    last = history[history.annee == anchor].merge(
        names[["dep", "region_rpps"]], on="dep", validate="one_to_one")
    regional_base = last.groupby("region_rpps", as_index=False).effectif.sum()
    regional_base = regional_base.rename(columns={"region_rpps": "region_source", "effectif": "effectif_reference"})
    result = []
    for scenario, series in national.groupby("scenario"):
        initial = series.loc[series.annee == anchor, "effectif_national"]
        if len(initial) != 1 or initial.iloc[0] <= 0:
            raise ValueError(f"Ancrage national manquant ou nul en {anchor} pour {scenario}.")
        future = series[series.annee > anchor].copy()
        if future.empty:
            raise ValueError(f"Aucune projection nationale après {anchor} pour {scenario}.")
        future["indice_national"] = future.effectif_national / initial.iloc[0]
        regional = regional_base.merge(future[["annee", "scenario", "indice_national"]], how="cross")
        regional["effectif_region"] = regional.effectif_reference * regional.indice_national
        result.append(regional[["region_source", "annee", "scenario", "effectif_region"]])
    if not result:
        raise ValueError("Aucun scénario national disponible.")
    return pd.concat(result, ignore_index=True)


def historical_psychiatry(path):
    """Lire l'onglet Effectifs : total et quatre modes d'exercice exclusifs."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    records, names = [], []
    try:
        rows = book["Effectifs"].iter_rows(values_only=True)
        pos = {v: i for i, v in enumerate(next(rows)) if v is not None}
        for row in rows:
            if any(row[pos[k]] != v for k, v in [("tranche_age", "00-Ensemble"), ("sexe", "0-Ensemble")]):
                continue
            if normalize(row[pos["territoire"]]) != "francemetropolitaine":
                continue
            if normalize(row[pos["specialites"]]) != "psychiatrie":
                continue
            raw_dep = str(row[pos["departement"]]).split("-", 1)[0].strip()
            dep = raw_dep[-2:] if len(raw_dep) == 3 and raw_dep.startswith("0") else raw_dep
            if dep not in METRO:
                continue
            label = normalize(row[pos["exercice"]])
            if label not in EXERCISE_GROUPS:
                raise ValueError(f"Mode d'exercice inconnu : {row[pos['exercice']]!r}.")
            kind = EXERCISE_GROUPS[label]
            if kind == "total":
                names.append({"dep": dep, "nom": str(row[pos["departement"]]).split("-", 1)[1].strip(),
                              "region_code": str(row[pos["region"]]).split("-", 1)[0].strip().zfill(2),
                              "region_rpps": normalize(row[pos["region"]])})
            for column, index in pos.items():
                if re.fullmatch(r"effectif_\d{4}", column):
                    value = row[index]
                    if not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
                        raise ValueError(f"Effectif indisponible : {dep}, {column}, {kind}, {value!r}.")
                    records.append({"dep": dep, "annee": int(column[-4:]), "statut": kind, "valeur": float(value)})
    finally:
        book.close()
    raw = pd.DataFrame(records)
    if raw.empty or raw.duplicated(["dep", "annee", "statut"]).any():
        raise ValueError("Effectifs par statut absents ou dupliqués.")
    wide = raw.pivot(index=["dep", "annee"], columns="statut", values="valeur")
    if "total" not in wide or wide.total.isna().any():
        raise ValueError("Total d'effectifs manquant pour certaines cellules département/année.")
    kinds = ["liberal_exclusif", "mixte", "hospitalier", "autre_salarie"]
    # Les catégories absentes ne valent zéro que si les autres reconstituent
    # exactement le total publié. Une cellule présente mais vide a déjà été rejetée.
    detail = wide.reindex(columns=kinds).fillna(0.0)
    if not np.allclose(detail.sum(axis=1), wide.total, atol=1e-8, rtol=0):
        raise ValueError("Les quatre modes d'exercice ne reconstituent pas le total publié.")
    history = pd.DataFrame({"effectif": wide.total,
                            "salaries": detail.hospitalier + detail.autre_salarie,
                            "liberaux": detail.liberal_exclusif + detail.mixte}).reset_index()
    history["part_salaries"] = np.divide(history.salaries, history.effectif,
                                          out=np.zeros(len(history)), where=history.effectif > 0)
    years = sorted(history.annee.unique())
    if (years != list(range(min(years), max(years) + 1))
            or set(history.dep) != METRO or len(history) != len(METRO) * len(years)):
        raise ValueError("Historique départemental incomplet.")
    return history, pd.DataFrame(names)


def assemble_reference(history, names, regional):
    """Garder le RPPS jusqu'à sa dernière année ; ensuite ventiler les totaux régionaux."""
    anchor = int(history.annee.max())
    available = set(regional.region_source)
    names = names.copy()
    mapping = {}
    for row in names.itertuples():
        possible = {row.region_rpps} & available
        possible |= {normalize(r) for r, deps in OLD_REGIONS.items() if row.dep in deps.split()} & available
        # Le libellé Centre-Val de Loire est parfois abrégé en Centre.
        if row.region_rpps == "centrevaldeloire" and "centre" in available:
            possible.add("centre")
        if len(possible) != 1:
            raise ValueError(f"Région DREES absente ou ambiguë pour le département {row.dep}.")
        mapping[row.dep] = possible.pop()
    names["region_source"] = names.dep.map(mapping)
    baseline = history[history.annee == anchor].merge(names[["dep", "region_source"]], on="dep", validate="one_to_one")
    totals = baseline.groupby("region_source").effectif.transform("sum")
    if (totals <= 0).any():
        raise ValueError("Impossible de répartir une région dont l'effectif RPPS est nul.")
    baseline["part_departement"] = baseline.effectif / totals
    future_reg = regional[(regional.annee > anchor) & regional.region_source.isin(mapping.values())].copy()
    if future_reg.empty:
        raise ValueError(f"Aucune projection après {anchor}.")
    expected_years = tuple(range(anchor + 1, int(future_reg.annee.max()) + 1))
    coverage = future_reg.groupby(["scenario", "region_source"]).annee.apply(lambda s: tuple(sorted(s)))
    if len(coverage) != future_reg.scenario.nunique() * len(set(mapping.values())) or not coverage.map(lambda x: x == expected_years).all():
        raise ValueError("Couverture région × année × scénario incomplète.")
    future = future_reg.merge(baseline[["dep", "region_source", "part_departement"]], on="region_source", validate="many_to_many")
    # Le tableau régional peut provenir des projections détaillées DREES ou du
    # repli national_to_regions. Chaque département reçoit seulement sa part.
    future["effectif"] = future.effectif_region * future.part_departement
    future["projete"] = True
    models = fit_status_models(history, anchor)
    future["part_salaries"] = project_status_shares(future, models, anchor)
    future["salaries"] = future.effectif * future.part_salaries
    future["liberaux"] = future.effectif - future.salaries
    observed = history.merge(names[["dep", "region_source"]], on="dep", validate="many_to_one")
    observed = observed.merge(future_reg[["scenario"]].drop_duplicates(), how="cross")
    observed["projete"] = False
    columns = ["dep", "region_source", "annee", "scenario", "effectif", "projete", "part_salaries", "salaries", "liberaux"]
    return pd.concat([observed[columns], future[columns]], ignore_index=True), anchor


# ---------- Parts par statut et dynamique nationale d'internat ----------
def expit(values):
    """Logistique stable sans dépendance supplémentaire."""
    return np.exp(-np.logaddexp(0.0, -np.asarray(values, dtype=float)))


def logistic_fit(years, successes, totals, pseudocount=0.0):
    """Logit binomial agrégé (constante + année), Newton avec recherche de pas.

    La demi-observation optionnelle évite les séparations sur petits effectifs.
    Pour les taux nationaux publiés, totals=1 donne le même poids à chaque année.
    La très faible pénalisation numérique ne concerne que la pente.
    """
    years, successes, totals = [np.asarray(v, dtype=float) for v in (years, successes, totals)]
    if (not (years.shape == successes.shape == totals.shape) or len(years) < 2
            or len(np.unique(years)) < 2 or not np.isfinite([years, successes, totals]).all()
            or (totals <= 0).any() or (successes < 0).any() or (successes > totals).any()):
        raise ValueError("Observations invalides pour l'extrapolation logistique.")
    center = float(years.mean())
    x = np.column_stack([np.ones(len(years)), (years - center) / 10.0])
    y, n = successes + pseudocount, totals + 2 * pseudocount
    initial = np.clip(y.sum() / n.sum(), 1e-12, 1 - 1e-12)
    beta = np.array([np.log(initial / (1 - initial)), 0.0])
    ridge = np.diag([0.0, 1e-8])

    def loss(b):
        z = x @ b
        return float(np.sum(n * np.logaddexp(0.0, z) - y * z) + 0.5 * b @ ridge @ b)

    for _ in range(100):
        prob = expit(x @ beta)
        gradient = x.T @ (n * prob - y) + ridge @ beta
        information = x.T @ ((n * prob * (1 - prob))[:, None] * x) + ridge
        step = np.linalg.solve(information, gradient)
        if np.max(np.abs(step)) < 1e-9:
            break
        scale, current = 1.0, loss(beta)
        while scale > 1e-9 and loss(beta - scale * step) > current + 1e-10:
            scale *= 0.5
        if scale <= 1e-9:
            if np.max(np.abs(gradient)) < 1e-7:
                break
            raise ValueError("L'ajustement logistique n'a pas convergé.")
        beta -= scale * step
    else:
        raise ValueError("Nombre d'itérations logistiques dépassé.")
    return {"intercept": float(beta[0]), "slope": float(beta[1] / 10.0), "center": center}


def anchored_logistic(share, slope, elapsed):
    """Raccord exact ; les parts exactement nulles ou égales à un restent fixes."""
    share, slope, elapsed = np.broadcast_arrays(share, slope, elapsed)
    clipped = np.clip(share, 1e-12, 1 - 1e-12)
    result = expit(np.log(clipped / (1 - clipped)) + slope * elapsed)
    return np.where(share <= 0, 0.0, np.where(share >= 1, 1.0, result))


def fit_status_models(history, anchor):
    """Une pente logistique par département, estimée sur tous les comptes observés."""
    records = []
    for dep, group in history.groupby("dep"):
        valid = group[(group.annee <= anchor) & (group.effectif > 0)].sort_values("annee")
        if valid.empty:
            raise ValueError(f"Aucun effectif positif pour estimer le statut dans {dep}.")
        last = valid.iloc[-1]
        constant = len(valid) < 2 or (valid.salaries == 0).all() or (valid.salaries == valid.effectif).all()
        fit = {"slope": 0.0, "intercept": 0.0, "center": float(anchor)} if constant else logistic_fit(
            valid.annee, valid.salaries, valid.effectif, pseudocount=0.5)
        records.append({"dep": dep, "pente_logit": fit["slope"],
                        "part_ancrage": float(last.part_salaries),
                        "annee_part_ancrage": int(last.annee), "nb_annees": len(valid)})
    return pd.DataFrame(records)


def project_status_shares(frame, models, anchor):
    by_dep = models.set_index("dep")
    if not frame.dep.isin(by_dep.index).all():
        raise ValueError("Modèle de statut manquant pour certains départements.")
    return anchored_logistic(frame.dep.map(by_dep.part_ancrage).to_numpy(),
                             frame.dep.map(by_dep.pente_logit).to_numpy(),
                             frame.annee.to_numpy() - anchor)


def vacancy_history():
    return pd.DataFrame(VACANCY_OBSERVATIONS, columns=["annee", "taux", "source"])


def vacancy_projection(start, anchor, end):
    """Tendance nationale, incluant 2026 ; raccord au dernier taux réellement publié."""
    observed = vacancy_history()
    observed = observed[observed.annee <= anchor]
    estimation = observed[observed.annee >= start]
    if len(estimation) < 5:
        raise ValueError("Au moins cinq années sont nécessaires pour la tendance d'internat.")
    fit = logistic_fit(estimation.annee, estimation.taux, np.ones(len(estimation)))
    last = observed.sort_values("annee").iloc[-1]
    years = np.arange(int(observed.annee.min()), end + 1)
    curve = pd.DataFrame({"annee": years})
    curve["taux_projete"] = anchored_logistic(float(last.taux), fit["slope"], years - last.annee)
    curve["taux_observe"] = curve.annee.map(observed.set_index("annee").taux)
    curve["taux"] = curve.taux_observe.fillna(curve.taux_projete)
    # On compare une dégradation future à un comportement constant à l'ancrage.
    # Le niveau déjà observé n'est pas soustrait une seconde fois aux stocks DREES.
    baseline = float(curve.loc[curve.annee == anchor, "taux"].iloc[0])
    curve["coefficient"] = np.where(curve.annee > anchor,
                                     np.minimum(1.0, (1 - curve.taux) / (1 - baseline)), 1.0)
    return curve, {"pente_logit": fit["slope"], "taux_reference": baseline,
                   "derniere_observation": int(last.annee), "debut": start}


def prepare_national_flows(reference, parameters, context):
    """Calcul commun aux départements ; chaque coefficient porte sur une cohorte."""
    anchor = context["anchor"]
    curve, calibration = vacancy_projection(parameters["debut"], anchor, int(reference.annee.max()))
    stocks = context["stocks"].sort_values(["scenario", "annee"])
    flows = stocks[stocks.annee > anchor][["annee", "scenario", "psychiatrie"]].copy()
    flows["annee_entree_internat"] = flows.annee - DES_YEARS
    coefficients = curve.set_index("annee").coefficient
    flows["coefficient_entrees"] = flows.annee_entree_internat.map(coefficients)
    # Les cohortes entrées au plus tard à l'ancrage sont déjà dans la référence.
    flows.loc[flows.annee_entree_internat <= anchor, "coefficient_entrees"] = 1.0
    if flows.coefficient_entrees.isna().any():
        raise ValueError("Coefficient de non-remplissage manquant pour une cohorte.")
    if not parameters["active"]:
        flows["coefficient_entrees"] = 1.0
    flows["flux_fr_reference"] = flows.psychiatrie / DES_YEARS
    flows["flux_fr_corrige"] = flows.flux_fr_reference * flows.coefficient_entrees
    # Conversion France -> métropole par la calibration des formations et leur mobilité.
    retention = context["routing"].groupby("origine").part.sum()
    weights = context["weights"].copy()
    weights["contribution_metro"] = weights.poids * weights.origine.map(retention).fillna(0.0)
    shares = weights.groupby("scenario").contribution_metro.sum()
    flows["part_flux_metropole"] = flows.scenario.map(shares)
    if flows.part_flux_metropole.isna().any() or not flows.part_flux_metropole.between(0, 1 + 1e-8).all():
        raise ValueError("Calibration France/métropole invalide pour les flux d'internes.")
    flows["manque_annuel_metro"] = (flows.flux_fr_reference - flows.flux_fr_corrige) * flows.part_flux_metropole
    flows["manque_cumule_metro"] = flows.groupby("scenario").manque_annuel_metro.cumsum()
    totals = reference[reference.projete].groupby(["annee", "scenario"], as_index=False).effectif.sum()
    flows = flows.merge(totals.rename(columns={"effectif": "total_reference"}), on=["annee", "scenario"], validate="one_to_one")
    flows["retrait_applique"] = np.minimum(flows.manque_cumule_metro, flows.total_reference)
    flows["plafond_atteint"] = flows.manque_cumule_metro > flows.total_reference
    flows["coefficient_stock"] = np.divide(flows.total_reference - flows.retrait_applique,
                                           flows.total_reference, out=np.ones(len(flows)), where=flows.total_reference > 0)
    return flows, curve, calibration


def internat_correction(projected, parameters, context):
    """Le déficit de nouvelles cohortes réduit uniformément les niveaux départementaux."""
    result = projected.copy()
    stats = context["national_flows"].set_index(["scenario", "annee"])
    keys = pd.MultiIndex.from_frame(result[["scenario", "annee"]])
    if not keys.isin(stats.index).all():
        raise ValueError("Correction nationale absente pour certaines années ou scénarios.")
    result["effectif"] *= stats.coefficient_stock.reindex(keys).to_numpy()
    return result


# ---------- Calibration : étudiants, mobilité, écarts départementaux ----------
def training_students(path, anchor, names):
    """Stocks DREES du troisième cycle ; jamais assimilés directement à des flux."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    national, regional = [], []
    try:
        for sheet in book:
            rows = sheet.iter_rows(values_only=True)
            positions = None
            for _ in range(30):
                row = next(rows, None)
                if row is None:
                    break
                p = {normalize(v): i for i, v in enumerate(row) if v is not None}
                if {"annee", "scenario", "statut", "effectif"} <= p.keys():
                    positions = p
                    break
            if positions is None:
                continue
            p = positions
            spec = next((k for k in ("specialite", "specialites") if k in p), None)
            is_regional = "region" in p and not any(k.startswith("specialite") for k in p)
            is_national = spec is not None and "region" not in p
            if not (is_regional or is_national):
                continue
            for row in rows:
                if not isinstance(row[p["annee"]], (int, float)):
                    continue
                year = int(row[p["annee"]])
                if year < anchor or normalize(row[p["scenario"]]) == "observe":
                    continue
                if normalize(row[p["statut"]]) != "etudiantsen3emecycle":
                    continue
                label = normalize(row[p["region"]] if is_regional else row[p[spec]])
                if is_regional and year != anchor:
                    continue
                if is_national and label not in {"ensemble", "psychiatrie"}:
                    continue
                value = row[p["effectif"]]
                if not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
                    raise ValueError(f"Stock d'étudiants invalide : {sheet.title}, {year}, {label}.")
                record = {"annee": year, "scenario": re.sub(r"^\d+-", "", str(row[p["scenario"]])),
                          "libelle": label, "stock": float(value)}
                (regional if is_regional else national).append(record)
    finally:
        book.close()
    stocks, regions = pd.DataFrame(national), pd.DataFrame(regional)
    for frame in (stocks, regions):
        if frame.empty or frame.duplicated(["annee", "scenario", "libelle"]).any():
            raise ValueError("Stocks d'étudiants absents ou ambigus dans les projections DREES.")
    stocks = stocks.pivot(index=["annee", "scenario"], columns="libelle", values="stock").reset_index()
    if not {"ensemble", "psychiatrie"} <= set(stocks) or stocks.isna().any().any():
        raise ValueError("Les stocks nationaux doivent contenir Ensemble et Psychiatrie pour chaque année/scénario.")
    if (stocks.ensemble <= 0).any() or (stocks.psychiatrie > stocks.ensemble).any():
        raise ValueError("Stocks nationaux d'étudiants incohérents.")
    mapping = {normalize(REGIONS[c]): c for c in REGIONS}
    mapping.update(dict(zip(names.region_rpps, names.region_code)))
    for old, departments in OLD_REGIONS.items():
        codes = names.loc[names.dep.isin(departments.split()), "region_code"].unique()
        if len(codes) != 1:
            raise ValueError(f"Correspondance régionale impossible : {old}.")
        mapping[normalize(old)] = codes[0]
    overseas = {"guyanemartiniqueguadeloupe", "lareunionmayotte"}
    weights = []
    for scenario, group in regions.groupby("scenario"):
        total = group.loc[group.libelle == "ensemble", "stock"]
        detail = group[group.libelle != "ensemble"].copy()
        if len(total) != 1 or total.iloc[0] <= 0:
            raise ValueError(f"Total régional des étudiants absent : {scenario}.")
        if not np.isclose(detail.stock.sum(), total.iloc[0], atol=len(detail), rtol=0):
            raise ValueError(f"Les stocks régionaux d'étudiants ne reconstituent pas le total : {scenario}.")
        unknown = set(detail.libelle) - set(mapping) - overseas
        if unknown:
            raise ValueError(f"Régions de formation inconnues : {sorted(unknown)}.")
        national_total = stocks.loc[(stocks.annee == anchor) & (stocks.scenario == scenario), "ensemble"]
        if len(national_total) != 1 or not np.isclose(national_total.iloc[0], total.iloc[0], atol=1, rtol=0):
            raise ValueError("Totaux national et régional des étudiants incompatibles.")
        detail["origine"] = detail.libelle.map(mapping)
        for origin, value in detail.dropna(subset=["origine"]).groupby("origine").stock.sum().items():
            weights.append({"scenario": scenario, "origine": origin, "poids": value / total.iloc[0]})
    weights = pd.DataFrame(weights)
    coverage = weights.groupby("scenario").origine.apply(lambda values: frozenset(values))
    if coverage.nunique() != 1:
        raise ValueError("Régions de formation manquantes dans certains scénarios.")
    return stocks, weights


def diploma_mobility(path, anchor):
    """P(lieu d'exercice | lieu de diplôme), psychiatres RPPS à l'ancrage.

    Garder les destinations hors métropole au dénominateur. Les lignes marginales
    ne sont pas ajoutées aux cellules. Les zéros omis sont autorisés uniquement
    lorsque les cellules reconstituent les marges de chaque origine.
    """
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    records = []
    try:
        sheets = [s for s in book if normalize(s.title) == "effectifsparlieudiplome"]
        if len(sheets) != 1:
            raise ValueError("Onglet Effectifs par lieu diplome absent ou ambigu.")
        rows = sheets[0].iter_rows(values_only=True)
        p = {v: i for i, v in enumerate(next(rows)) if v is not None}
        for row in rows:
            if normalize(row[p["specialites"]]) != "psychiatrie":
                continue
            value = row[p[f"effectif_{anchor}"]]
            if not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
                raise ValueError("Cellule de mobilité RPPS manquante ou invalide.")
            records.append({"origine": str(row[p["region_diplome"]]).split("-", 1)[0].strip().zfill(2),
                            "destination": str(row[p["region"]]).split("-", 1)[0].strip().zfill(2),
                            "effectif": float(value)})
    finally:
        book.close()
    cells = pd.DataFrame(records)
    if cells.empty or cells.duplicated(["origine", "destination"]).any():
        raise ValueError("Matrice diplôme/exercice absente ou dupliquée.")
    margins = cells[(cells.destination == "00") & cells.origine.isin(REGIONS)].set_index("origine").effectif
    detail = cells[(cells.destination != "00") & cells.origine.isin(REGIONS)].copy()
    sums = detail.groupby("origine").effectif.sum().reindex(margins.index, fill_value=0)
    if not np.allclose(sums, margins, atol=0.01, rtol=0):
        raise ValueError("La matrice diplôme/exercice ne reconstitue pas ses marges.")
    detail = detail[detail.origine.isin(margins[margins > 0].index)].copy()
    detail["part"] = detail.effectif / detail.origine.map(margins)
    return detail[["origine", "destination", "part"]]


def count_reference(frame, baseline, count_column="effectif"):
    """Effectif national ventilé selon les parts départementales à l'ancrage."""
    shares = baseline.set_index("dep").part_nationale
    groups = ["annee", "scenario"] if "scenario" in frame else ["annee"]
    totals = frame.groupby(groups)[count_column].transform("sum")
    return frame.dep.map(shares) * totals


def estimate_department_trends(history, baseline, start, anchor):
    """OLS de l'écart d'effectif à la référence nationale, en psychiatres/an."""
    h = history[history.annee.between(start, anchor)].copy()
    years = sorted(h.annee.unique())
    if (len(years) < 2 or years != list(range(start, anchor + 1))
            or set(h.dep) != set(baseline.dep)
            or h.duplicated(["dep", "annee"]).any()
            or not h.groupby("dep").size().eq(len(years)).all()):
        raise ValueError("Fenêtre d'estimation incomplète ou trop courte.")
    h["ecart"] = h.effectif - count_reference(h, baseline)
    x = h.annee - h.annee.mean()
    h["produit"], h["carre"] = x * h.ecart, x * x
    sums = h.groupby("dep")[["produit", "carre", "ecart"]].agg({"produit": "sum", "carre": "sum", "ecart": "mean"})
    slopes = sums.produit / sums.carre
    # Les écarts somment à zéro chaque année, donc les pentes aussi. Ce centrage
    # enlève uniquement le résidu numérique des opérations en virgule flottante.
    slopes -= slopes.mean()
    return pd.DataFrame({"dep": slopes.index, "pente": slopes.to_numpy(),
                         "niveau_ajuste_ancrage": (sums.ecart + slopes * (anchor - h.annee.mean())).to_numpy()})


def build_context(history, names, stocks, weights, mobility):
    anchor = int(history.annee.max())
    baseline = history[history.annee == anchor].merge(names, on="dep", validate="one_to_one")
    baseline["part_nationale"] = baseline.effectif / baseline.effectif.sum()
    baseline["part_regionale"] = baseline.effectif / baseline.groupby("region_code").effectif.transform("sum")
    trends = estimate_department_trends(history, baseline, int(history.annee.min()), anchor)
    baseline = baseline.merge(trends, on="dep", validate="one_to_one")
    routing = mobility.merge(baseline[["dep", "region_code", "part_regionale"]],
                             left_on="destination", right_on="region_code", validate="many_to_many")
    routing["part"] *= routing.part_regionale
    available = set(mobility.origine) & set(weights.loc[weights.poids > 0, "origine"])
    return {"anchor": anchor, "baseline": baseline, "stocks": stocks,
            "history": history.copy(deep=True),
            "status_models": fit_status_models(history, anchor),
            "weights": weights, "mobility": mobility, "routing": routing,
            "available": available, "observed_start": int(history.annee.min())}


def department_dynamics(projected, parameters, context):
    if not parameters["active"]:
        return projected
    result = projected.copy()
    base = context["baseline"].set_index("dep")
    slopes = estimate_department_trends(context["history"], context["baseline"],
                                        parameters["debut"], context["anchor"]).set_index("dep").pente
    total = result.groupby(["annee", "scenario"]).effectif.transform("sum")
    annual_effect = result.dep.map(slopes)
    raw = total * result.dep.map(base.part_nationale) + annual_effect * (result.annee - context["anchor"])
    # Identité au-dessus du seuil, approche exponentielle de zéro en dessous.
    result["effectif"] = np.where(raw >= SMOOTH_FLOOR, raw,
                                  SMOOTH_FLOOR * np.exp(np.minimum(raw / SMOOTH_FLOOR, 1) - 1))
    positive_total = result.groupby(["annee", "scenario"]).effectif.transform("sum")
    result["effectif"] *= np.divide(total, positive_total, out=np.zeros(len(result)), where=positive_total > 0)
    return result


def training_additions(projected, parameters, context):
    """Hausse annuelle permanente ; cumul des cohortes après bac+11, sans sorties."""
    result = projected.copy()
    weights = context["weights"].copy()
    weights["hausse"] = weights.origine.map(lambda code: parameters.get(f"region_{code}", 0) / 100)
    if (weights.hausse == 0).all():
        return result
    first_graduation = parameters["rentree"] + STUDY_YEARS
    stocks = context["stocks"].copy().sort_values(["scenario", "annee"])
    flow_keys = pd.MultiIndex.from_frame(stocks[["scenario", "annee"]])
    corrected = context["national_flows"].set_index(["scenario", "annee"]).flux_fr_corrige
    stocks["flux_psy"] = np.where(stocks.annee >= first_graduation,
                                   corrected.reindex(flow_keys, fill_value=0.0).to_numpy(), 0.0)
    stocks["cohortes_cumulees"] = stocks.groupby("scenario").flux_psy.cumsum()
    allocation = weights.merge(context["routing"][["origine", "dep", "part"]], on="origine", validate="many_to_many")
    allocation["coefficient"] = allocation.poids * allocation.hausse * allocation.part
    coefficients = allocation.groupby(["scenario", "dep"]).coefficient.sum()
    flux = stocks.set_index(["scenario", "annee"]).cohortes_cumulees
    keys = pd.MultiIndex.from_frame(result[["scenario", "annee"]])
    if not keys.isin(flux.index).all():
        raise ValueError("Stocks d'internes manquants pour certaines années projetées.")
    dep_keys = pd.MultiIndex.from_frame(result[["scenario", "dep"]])
    result["effectif"] += flux.reindex(keys).to_numpy() * coefficients.reindex(dep_keys, fill_value=0).to_numpy()
    return result


# ---------- Modulations indépendantes, interface générée depuis le registre ----------
@dataclass(frozen=True)
class Control:
    key: str
    label: str
    kind: str  # slider, checkbox, number, select
    default: object
    minimum: object = None
    maximum: object = None
    step: object = None
    options: tuple = ()
    disabled: bool = False
    help: str = ""


@dataclass(frozen=True)
class Modulation:
    key: str
    label: str
    description: str
    controls: tuple
    apply: Callable


def modulation_registry(context):
    """Ajouter une entrée ici suffit pour enregistrer une nouvelle modulation.

    La fonction apply(projected, parameters, context) reçoit des copies. Seule
    la colonne effectif est modifiable ; les autres restent réservées au moteur.
    L'ordre du tuple est l'ordre d'application ; tous les contrôles sont dans
    le même formulaire. Le contexte dépend des sources, jamais de l'utilisateur.
    """
    regional_controls = tuple(
        Control(f"region_{code}", f"{label} (%)", "slider", 0, 0, 100, 1,
                disabled=code not in context["available"],
                help="Hausse annuelle permanente des places, à spécialisation et mobilité constantes."
                if code in context["available"] else "Pas de calibration régionale disponible dans les deux sources ; curseur désactivé.")
        for code, label in sorted(REGIONS.items(), key=lambda item: item[1])
    )
    return (
        Modulation("dynamique", "Dynamiques départementales",
                   f"Tendances des effectifs (psychiatres par an), estimées jusqu'à {context['anchor']}. Effet dès {context['anchor'] + 1}, indépendant du délai de formation.",
                   (Control("active", "Extrapoler les dynamiques départementales", "checkbox", True),
                    Control("debut", "Première année pour estimer la tendance", "number", context["observed_start"],
                            context["observed_start"], context["anchor"] - 4, 1,
                            help="Tout l'historique par défaut. Choisir 2017 pour dix années, ou 2022 pour cinq années avec les sources 2026. Une fenêtre courte peut modifier fortement la pente.")), department_dynamics),
        Modulation("internat", "Dynamique nationale de l'internat",
                   "Correction uniforme des futures cohortes. La tendance inclut l'amélioration de 2026 ; les cohortes déjà entrées ne sont pas corrigées.",
                   (Control("active", "Appliquer la correction des postes non pourvus", "checkbox", True),
                    Control("debut", "Première année pour la tendance des postes non pourvus", "number",
                            min(v[0] for v in VACANCY_OBSERVATIONS), min(v[0] for v in VACANCY_OBSERVATIONS),
                            max(v[0] for v in VACANCY_OBSERVATIONS) - 4, 1,
                            help="Logistique nationale, avec au moins cinq années, jusqu'à la dernière observation. Une tendance à l'amélioration ne crée pas de hausse supplémentaire : le coefficient est plafonné à 1.")),
                   internat_correction),
        Modulation("formation", "Places de formation par région",
                   "Première rentrée au niveau du bac ; hausse reconduite pour chaque cohorte. Premiers psychiatres 11 ans plus tard.",
                   (Control("rentree", "Première rentrée augmentée (année du bac)", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()) - STUDY_YEARS, 1),) + regional_controls,
                   training_additions),
    )


def default_parameters(registry):
    return {m.key: {c.key: c.default for c in m.controls} for m in registry}


def validate_parameters(parameters, registry):
    for modulation in registry:
        values = parameters[modulation.key]
        for control in modulation.controls:
            value = values[control.key]
            if control.disabled and value != control.default:
                raise ValueError(f"Calibration indisponible : {control.label}.")
            if control.kind in {"slider", "number"}:
                if isinstance(value, bool) or not isinstance(value, (int, float)) or not np.isfinite(value):
                    raise ValueError(f"Valeur numérique invalide : {control.label}.")
                if (control.minimum is not None and value < control.minimum) or (control.maximum is not None and value > control.maximum):
                    raise ValueError(f"Valeur hors limites : {control.label}.")
            elif control.kind == "checkbox" and not isinstance(value, bool):
                raise ValueError(f"Valeur booléenne attendue : {control.label}.")
            elif control.kind == "select" and value not in control.options:
                raise ValueError(f"Option invalide : {control.label}.")


def apply_modulations(reference, parameters, context, return_contributions=False, return_details=False):
    registry = modulation_registry(context)
    validate_parameters(parameters, registry)
    flows, curve, calibration = prepare_national_flows(reference, parameters["internat"], context)
    run_context = deepcopy(context)
    run_context["national_flows"] = flows
    result = reference.copy(deep=True)
    projected = result[result.projete].copy()
    immutable = [c for c in projected if c != "effectif"]
    contributions = pd.DataFrame({"reference": result.effectif.copy()}, index=result.index)
    for modulation in registry:
        values = parameters[modulation.key]
        changed = modulation.apply(projected.copy(deep=True), dict(values), deepcopy(run_context))
        if not isinstance(changed, pd.DataFrame) or list(changed.columns) != list(projected.columns) or not changed.index.equals(projected.index):
            raise ValueError(f"{modulation.key} a modifié la structure des données.")
        if not changed[immutable].equals(projected[immutable]):
            raise ValueError(f"{modulation.key} a modifié une clé ou une colonne réservée.")
        if not np.isfinite(changed.effectif).all() or (changed.effectif < 0).any():
            raise ValueError(f"Résultat invalide : {modulation.key}.")
        contributions[modulation.key] = 0.0
        contributions.loc[projected.index, modulation.key] = changed.effectif - projected.effectif
        projected = changed
    result.loc[projected.index, "effectif"] = projected.effectif
    # Répartition finale du total simulé ; les comptes historiques restent exacts.
    result.loc[projected.index, "salaries"] = projected.effectif * projected.part_salaries
    result.loc[projected.index, "liberaux"] = projected.effectif - result.loc[projected.index, "salaries"]
    result.loc[result.effectif == 0, "part_salaries"] = 0.0
    if not np.allclose(result.salaries + result.liberaux, result.effectif, atol=1e-8, rtol=1e-12):
        raise ValueError("Les comptes par statut ne reconstituent pas le total.")
    if return_details:
        return result, contributions, {"national_flows": flows, "vacancy_curve": curve, "vacancy_calibration": calibration}
    return (result, contributions) if return_contributions else result


def to_projections_frame(panel):
    values = panel[["dep", "annee", "scenario"] + list(INDICATORS)].copy()
    # Convention publique explicite : la part est exprimée de 0 à 100, en %.
    values["part_salaries"] *= 100.0
    result = values.melt(id_vars=["dep", "annee", "scenario"], value_vars=list(INDICATORS),
                         var_name="indicateur", value_name="valeur")
    result.indicateur = result.indicateur.map(INDICATORS)
    result = result[COLUMNS].sort_values(COLUMNS[:-1]).reset_index(drop=True)
    if result.empty or result.isna().any().any() or result.duplicated(COLUMNS[:-1]).any():
        raise ValueError("Table de projections vide, incomplète ou dupliquée.")
    if not np.isfinite(result.valeur).all() or (result.valeur < 0).any():
        raise ValueError("Valeur projetée invalide.")
    if not result.loc[result.indicateur == SALARIED_SHARE, "valeur"].between(0, 100).all():
        raise ValueError("Part salariée hors de l'intervalle 0–100 %.")
    return result


# ---------- Référence commune en mémoire et projections propres à la session ----------
def file_hash(path):
    h = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def source_signature():
    paths = [RPPS_PATH, DREES_PATH, Path(__file__)]
    return sha256(json.dumps([(str(p), file_hash(p)) for p in paths]).encode()).hexdigest()


@st.cache_data(show_spinner=False, persist=None, max_entries=2)
def load_reference(signature):
    # Seule la référence SANS paramètres utilisateur est mise en cache commun.
    # st.cache_data restitue une copie indépendante à chaque appel. Aucun cache disque.
    # signature inclut les sources et le code pour invalider les anciennes références.
    regional, sheet = regional_psychiatry(DREES_PATH)
    history, names = historical_psychiatry(RPPS_PATH)
    mode = "regional"
    if regional is None:
        national, sheet = national_psychiatry(DREES_PATH)
        regional = national_to_regions(history, names, national)
        mode = "national"
    reference, anchor = assemble_reference(history, names, regional)
    stocks, weights = training_students(DREES_PATH, anchor, names)
    mobility = diploma_mobility(RPPS_PATH, anchor)
    context = build_context(history, names, stocks, weights, mobility)
    expected = pd.MultiIndex.from_frame(reference[["annee", "scenario"]].drop_duplicates())
    future = expected[expected.get_level_values("annee") >= anchor]
    actual = pd.MultiIndex.from_frame(stocks[["annee", "scenario"]])
    if not future.isin(actual).all() or set(reference.scenario) != set(weights.scenario):
        raise ValueError("Calibration des étudiants incomplète sur les années/scénarios projetés.")
    return reference, {"annee_reference": anchor, "onglet_drees": sheet,
                       "mode_projection": mode}, context


def ensure_projections(signature, parameters, reference, metadata, context):
    """Recalcul seulement au premier accès, changement de sources ou validation.

    Les projections modulées et leurs paramètres appartiennent exclusivement à
    cette session. Aucun fichier, cache global de résultats ou cache de ressources.
    """
    key = sha256(json.dumps({"sources": signature, "parameters": parameters}, sort_keys=True).encode()).hexdigest()
    state = st.session_state.get("psy_projection", {})
    if state.get("key") == key:
        return state["projections"], state["metadata"], False
    # Toujours repartir de la référence non modulée : aucun cumul entre clics.
    panel, contributions, details = apply_modulations(reference, parameters, context, return_details=True)
    output = to_projections_frame(panel)
    # Remplacement de l'état seulement lorsque le calcul a entièrement réussi.
    st.session_state.psy_projection = {
        "key": key, "parameters": deepcopy(parameters),
        "metadata": deepcopy(metadata), "projections": output, "panel": panel,
        "contributions": contributions, **details,
    }
    return output, metadata, True


def render_parameters(applied, registry):
    draft = {}
    for modulation in registry:
        st.subheader(modulation.label)
        st.caption(modulation.description)
        values = {}
        for control in modulation.controls:
            value = applied.get(modulation.key, {}).get(control.key, control.default)
            common = {"label": control.label, "key": f"param_{modulation.key}_{control.key}",
                      "disabled": control.disabled, "help": control.help or None}
            if control.kind == "slider":
                value = st.slider(**common, min_value=control.minimum, max_value=control.maximum, value=value, step=control.step)
            elif control.kind == "number":
                value = st.number_input(**common, min_value=control.minimum, max_value=control.maximum, value=value, step=control.step)
            elif control.kind == "checkbox":
                value = st.checkbox(**common, value=value)
            elif control.kind == "select":
                value = st.selectbox(**common, options=control.options, index=control.options.index(value))
            else:
                raise ValueError(f"Contrôle inconnu : {control.kind}.")
            values[control.key] = value
        draft[modulation.key] = values
    return draft


@st.cache_data(persist=None)
def load_geojson():
    with urllib.request.urlopen(GEOJSON_URL, timeout=20) as handle:
        gj = json.load(handle)
    gj["features"] = [f for f in gj["features"] if f["properties"]["code"] in METRO]
    return gj


def dynamics_diagnostic(panel, contributions, context, scenario):
    """Écarts d'effectif à une référence nationale commune aux deux vues."""
    if not panel.index.equals(contributions.index):
        raise ValueError("Diagnostic et projection désalignés.")
    diagnostic = panel[panel.scenario == scenario].copy()
    pieces = contributions.loc[diagnostic.index]
    diagnostic["effectif_reference"] = pieces.reference
    diagnostic["effet_dynamique"] = pieces.dynamique
    diagnostic["effet_internat"] = pieces.internat
    diagnostic["ajout_formation"] = pieces.formation
    diagnostic["hors_formation"] = pieces.reference + pieces.dynamique + pieces.internat
    diagnostic["reference_nationale"] = count_reference(diagnostic, context["baseline"], "hors_formation")
    diagnostic["ecart_hors_formation"] = diagnostic.hors_formation - diagnostic.reference_nationale
    diagnostic["ecart_complet"] = diagnostic.effectif - diagnostic.reference_nationale
    return diagnostic


def dynamics_figure(diagnostic, context, departments, full=False, start_year=None, show_fit=False):
    """Observé puis extrapolation ; les nouvelles places sont une vue distincte."""
    anchor = context["anchor"]
    start_year = context["observed_start"] if start_year is None else start_year
    trends = estimate_department_trends(context["history"], context["baseline"], start_year, anchor).set_index("dep")
    base = context["baseline"].set_index("dep")
    column = "ecart_complet" if full else "ecart_hors_formation"
    figure = go.Figure()
    colors = px.colors.qualitative.Dark24
    for index, dep in enumerate(departments):
        series = diagnostic[diagnostic.dep == dep].sort_values("annee")
        label = f"{dep} · {base.loc[dep, 'nom']}"
        for projected, label_period, dash in ((False, "observé", "solid"), (True, "projeté", "dash")):
            part = series[series.annee >= anchor] if projected else series[~series.projete]
            figure.add_trace(go.Scatter(x=part.annee, y=part[column], mode="lines",
                                       name=f"{label} — {label_period}", legendgroup=dep,
                                       line={"color": colors[index % len(colors)], "dash": dash},
                                       hovertemplate="%{x} : %{y:.2f} psychiatres<extra>%{fullData.name}</extra>"))
        if show_fit:
            years = np.arange(start_year, anchor + 1)
            fitted = trends.loc[dep, "niveau_ajuste_ancrage"] + trends.loc[dep, "pente"] * (years - anchor)
            figure.add_trace(go.Scatter(x=years, y=fitted, mode="lines",
                                       name=f"{label} — droite estimée", legendgroup=dep, showlegend=False,
                                       line={"color": colors[index % len(colors)], "dash": "dot", "width": 1}, opacity=0.6,
                                       hovertemplate="%{x} : %{y:.2f} psychiatres<extra>%{fullData.name}</extra>"))
    figure.add_hline(y=0, line_color="#64748b", line_width=1,
                    annotation_text=f"Part de l'effectif national fixée en {anchor}, sans nouvelles places")
    figure.add_vline(x=anchor, line_color="#94a3b8", line_dash="dot")
    figure.update_layout(height=570, xaxis_title="Année", yaxis_title="Écart à la référence nationale (nombre de psychiatres)",
                         margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified")
    return figure


def calibration_table(reference, context, scenario, parameters, flows):
    first = parameters["formation"]["rentree"] + STUDY_YEARS
    row = flows[(flows.scenario == scenario) & (flows.annee == first)].iloc[0]
    flow = float(row.flux_fr_corrige)
    national = float(row.total_reference * row.coefficient_stock)
    weights = context["weights"]
    weights = weights[weights.scenario == scenario].set_index("origine").poids
    retention = context["routing"].groupby("origine").part.sum()
    records = []
    for code, label in REGIONS.items():
        available = code in context["available"]
        additions = 0.01 * weights.get(code, 0) * flow * retention.get(code, 0)
        records.append({"Région de formation": label, "Calibration": "Disponible" if available else "Indisponible",
                        "Hausse validée (%)": parameters["formation"][f"region_{code}"],
                        "Poids formation France (%)": 100 * weights.get(code, 0) if available else np.nan,
                        "Exercice métropolitain (%)": 100 * retention.get(code, 0) if available else np.nan,
                        f"Psychiatres métro. en plus en {first} pour +1 %": additions if available else np.nan,
                        f"X (%) du total métro. en {first} pour +1 %": 100 * additions / national if available and national > 0 else np.nan})
    return pd.DataFrame(records)


def status_figure(panel, context, scenario, departments):
    figure = go.Figure()
    names = context["baseline"].set_index("dep").nom
    for i, dep in enumerate(departments):
        series = panel[(panel.scenario == scenario) & (panel.dep == dep)].sort_values("annee")
        for projected, dash, period in ((False, "solid", "observé"), (True, "dash", "projeté")):
            part = series[series.annee >= context["anchor"]] if projected else series[~series.projete]
            figure.add_trace(go.Scatter(x=part.annee, y=100 * part.part_salaries, mode="lines",
                                       name=f"{dep} · {names[dep]} — {period}", legendgroup=dep,
                                       line={"color": px.colors.qualitative.Dark24[i % 24], "dash": dash},
                                       hovertemplate="%{x} : %{y:.2f} %<extra>%{fullData.name}</extra>"))
    figure.add_vline(x=context["anchor"], line_dash="dot", line_color="#94a3b8")
    figure.update_layout(height=500, xaxis_title="Année", yaxis_title="Part de salariés (%)",
                         yaxis_range=[0, 100], hovermode="x unified")
    return figure


def vacancy_figure(curve, calibration):
    last = calibration["derniere_observation"]
    observed = curve[curve.taux_observe.notna()]
    future = curve[curve.annee >= last]
    figure = go.Figure()
    figure.add_trace(go.Scatter(x=observed.annee, y=100 * observed.taux_observe, mode="lines+markers", name="Observé"))
    figure.add_trace(go.Scatter(x=future.annee, y=100 * future.taux_projete, mode="lines", name="Extrapolation logistique", line_dash="dash"))
    figure.add_hline(y=100 * calibration["taux_reference"], line_dash="dot", annotation_text="Taux de référence constant")
    figure.update_layout(height=420, xaxis_title="Année d'entrée en internat", yaxis_title="Postes non pourvus (%)",
                         yaxis_range=[0, 100], hovermode="x unified")
    return figure


def main():
    st.set_page_config(page_title="Projections psychiatres", layout="wide")
    st.title("Projections des effectifs de psychiatres par département")
    try:
        require_excel_reader()
        signature = source_signature()
        with st.spinner("Lecture et calibration des données…"):
            reference, metadata, context = load_reference(signature)
    except FileNotFoundError as exc:
        st.error(f"Fichier absent : {exc.filename}. Placer les deux classeurs à côté de app.py.")
        st.stop()
    except (ValueError, KeyError, OSError, RuntimeError) as exc:
        st.error(str(exc))
        st.stop()
    registry = modulation_registry(context)
    # Un changement de sources/code réinitialise les anciens contrôles, dont
    # l'ancien curseur national. Un simple changement de filtre ne fait rien.
    if st.session_state.get("parameter_source") != signature:
        for key in list(st.session_state):
            if key.startswith("param_"):
                del st.session_state[key]
        st.session_state.applied_parameters = default_parameters(registry)
        st.session_state.parameter_source = signature
    with st.sidebar:
        with st.form("modulations"):
            draft = render_parameters(st.session_state.applied_parameters, registry)
            submitted = st.form_submit_button("Valider", type="primary")
        st.caption("Les modifications prennent effet après Valider. Vos paramètres et résultats sont propres à cette session.")
    proposed = draft if submitted else st.session_state.applied_parameters
    try:
        with st.spinner("Calcul du scénario…"):
            df, metadata, recalculated = ensure_projections(signature, proposed, reference, metadata, context)
        st.session_state.applied_parameters = deepcopy(proposed)
    except (ValueError, KeyError, RuntimeError) as exc:
        st.error(str(exc))
        st.stop()
    if recalculated:
        st.success("Projections recalculées pour votre session.")
    anchor = context["anchor"]
    st.caption(f"Observé : {df.annee.min()}–{anchor}. Projeté : {anchor + 1}–{df.annee.max()}.")
    if metadata["mode_projection"] == "national":
        st.info(f"Psychiatrie absente des projections régionales : croissance nationale DREES raccordée aux effectifs RPPS de {anchor}.")
    else:
        st.caption("Référence issue des projections régionales détaillées de psychiatrie.")
    if proposed["dynamique"]["active"]:
        st.caption(f"Paramètres validés : dynamiques des effectifs actives dès {anchor + 1}, estimées sur {proposed['dynamique']['debut']}–{anchor}. Les pentes sont en psychiatres par an ; le total national est conservé à cette étape, puis la correction de l’internat et les ajouts de places sont appliqués.")
    else:
        st.warning(f"Paramètres validés : dynamiques départementales désactivées. Leur pente n'est pas appliquée. Pour les activer dès {anchor + 1}, cocher la case dans le formulaire puis cliquer sur Valider.")
    state = st.session_state.psy_projection
    if proposed["internat"]["active"]:
        st.caption(f"Correction nationale des cohortes activée : premiers effets possibles sur les effectifs en {anchor + 1 + DES_YEARS}. Les observations restent inchangées.")
    else:
        st.caption("Correction nationale des postes non pourvus désactivée.")
    with st.sidebar:
        indicator = st.selectbox("Indicateur", list(INDICATORS.values()))
        scenario = st.selectbox("Scénario", sorted(df.scenario.unique()))
        year = st.slider("Année", int(df.annee.min()), int(df.annee.max()), anchor)
    tab_map, tab_trends, tab_status, tab_internat, tab_method = st.tabs(["Carte et effectifs", "Dynamiques départementales", "Statuts d’exercice", "Internat national", "Calibration et hypothèses"])
    names = context["baseline"].set_index("dep").nom.to_dict()
    with tab_map:
        selected = df[(df.indicateur == indicator) & (df.scenario == scenario) & (df.annee == year)].copy()
        selected["nom"] = selected.dep.map(names)
        try:
            gj = load_geojson()
        except (OSError, ValueError) as exc:
            gj = None
            st.warning(f"Carte indisponible ; les tableaux et les courbes restent accessibles : {exc}")
        is_share = indicator == SALARIED_SHARE
        value_label = indicator + " (%)" if is_share else indicator
        vmin, vmax = (0.0, 100.0) if is_share else df.loc[df.indicateur == indicator, "valeur"].quantile([0.02, 0.98])
        if vmin == vmax:
            vmax = vmin + 1
        map_col, table_col = st.columns([2, 1])
        if gj is not None:
            fig = px.choropleth(selected, geojson=gj, locations="dep", featureidkey="properties.code", color="valeur", color_continuous_scale="Blues", range_color=(vmin, vmax), hover_name="nom", labels={"valeur": value_label})
            fig.update_geos(fitbounds="locations", visible=False)
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=650)
            map_col.plotly_chart(fig, width="stretch")
        table_col.dataframe(selected.sort_values("valeur")[["dep", "nom", "valeur"]].rename(columns={"valeur": value_label}), hide_index=True, width="stretch")
        st.caption("Salariés = salariés hospitaliers + autres salariés. Libéraux = libéraux exclusifs + exercice mixte. Chaque personne est comptée une seule fois ; la part salariée est en %.")
    with tab_trends:
        active = proposed["dynamique"]["active"]
        start = proposed["dynamique"]["debut"]
        if active:
            st.info(f"Dynamique réellement appliquée : OUI, dès {anchor + 1}. Pente d'effectif estimée sur {start}–{anchor} puis raccordée à l'observation {anchor}.")
        else:
            st.warning("Dynamique réellement appliquée : NON. La pente d'effectif n'est pas appliquée. Avec le repli national, l'écart reste nul sans nouvelles places. Des projections régionales détaillées peuvent avoir leur propre évolution.")
        st.write(f"Les courbes montrent un écart en nombre de psychiatres. Chaque effectif est comparé au total national après correction de l’internat et avant nouvelles places, multiplié par la part du département observée en {anchor}. Cette référence est commune aux deux vues.")
        view = st.selectbox("Projection affichée", ["Sans ouverture de places", "Avec les ouvertures de places"], key="trend_view")
        full = view == "Avec les ouvertures de places"
        show_fit = st.checkbox("Afficher la droite estimée sur les observations", value=False, key="show_trend_fit")
        departments = st.multiselect("Départements à comparer", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]],
                                     format_func=lambda dep: f"{dep} · {names[dep]}")
        if departments:
            state = st.session_state.psy_projection
            diagnostic = dynamics_diagnostic(state["panel"], state["contributions"], context, scenario)
            st.plotly_chart(dynamics_figure(diagnostic, context, departments, full, start, show_fit), width="stretch")
            if full:
                has_places = any(proposed["formation"][f"region_{code}"] for code in REGIONS)
                if has_places:
                    st.caption(f"Ouvertures validées : premiers effets en {proposed['formation']['rentree'] + STUDY_YEARS}. Ce délai concerne seulement les places de formation, jamais la tendance départementale.")
                else:
                    st.caption("Aucune hausse de places validée : les deux vues coïncident.")
            else:
                st.caption("Cette vue isole la trajectoire sans les ajouts de places, même si des curseurs régionaux sont positifs. La carte conserve le scénario complet validé.")
            st.caption(f"Trait plein : observations. Tirets : projection. Toutes les courbes valent zéro à l'ancrage {anchor}. Un plat à zéro signifie que l'effectif suit la référence nationale, pas qu'il est constant. Un écart négatif correspond à moins de psychiatres que cette référence. Le filtre Année de la carte ne coupe pas les courbes.")
            if show_fit:
                st.caption("Les fines lignes pointillées sont les régressions des écarts d'effectif, avec constante. Le futur repart de l'observation réelle et prolonge la pente estimée. La saturation et la renormalisation conservent le total à cette étape ; la correction nationale de l’internat peut ensuite réduire les niveaux et les écarts départementaux.")
            slopes = estimate_department_trends(context["history"], context["baseline"], start, anchor)
            slopes = slopes[slopes.dep.isin(departments)][["dep", "pente"]].copy()
            slopes.insert(1, "nom", slopes.dep.map(names))
            slopes["Pente avant saturation et correction nationale"] = slopes.pente if active else 0.0
            st.dataframe(slopes.rename(columns={"pente": "Pente estimée (psychiatres/an)"}), hide_index=True, width="stretch")
            st.caption("Pentes en psychiatres par an, relativement à la référence nationale. La pente appliquée est nulle lorsque les dynamiques sont désactivées. Après saturation, renormalisation et correction nationale, l’évolution finale peut différer de cette pente.")
            with st.expander(f"Décomposition des effectifs en {year}"):
                detail = diagnostic[(diagnostic.annee == year) & diagnostic.dep.isin(departments)].copy()
                detail.insert(1, "nom", detail.dep.map(names))
                st.dataframe(detail[["dep", "nom", "effectif_reference", "effet_dynamique", "effet_internat", "ajout_formation", "effectif"]].rename(
                    columns={"effectif_reference": "Référence sans modulation", "effet_dynamique": "Effet départemental", "effet_internat": "Correction internat",
                             "ajout_formation": "Ajout des places", "effectif": "Effectif final"}), hide_index=True, width="stretch")
                st.caption("L'effet départemental inclut la saturation et la renormalisation. L'ajout des places est un nombre de psychiatres supplémentaires, toujours positif ou nul.")
        else:
            st.info("Sélectionner au moins un département pour afficher les courbes.")
    with tab_status:
        st.write(f"La part salariée est estimée séparément pour chaque département par une tendance logistique sur les comptes {context['observed_start']}–{anchor} de l’onglet Effectifs. La projection repart exactement de la part observée en {anchor}.")
        st.caption("Salariés : hospitaliers et autres salariés, hors exercice mixte. Libéraux : libéraux exclusifs et mixtes. Les deux groupes reconstituent toujours l’effectif total.")
        status_departments = st.multiselect("Départements pour les statuts", sorted(names),
                                            default=["75"] if "75" in names else [sorted(names)[0]],
                                            format_func=lambda dep: f"{dep} · {names[dep]}", key="status_departments")
        if status_departments:
            st.plotly_chart(status_figure(state["panel"], context, scenario, status_departments), width="stretch")
            selected_status = state["panel"][(state["panel"].scenario == scenario) & (state["panel"].annee == year)
                                               & state["panel"].dep.isin(status_departments)].copy()
            selected_status["part_salaries"] *= 100
            selected_status.insert(1, "nom", selected_status.dep.map(names))
            st.dataframe(selected_status[["dep", "nom", "effectif", "salaries", "liberaux", "part_salaries"]].rename(
                columns={"effectif": COUNT, "salaries": SALARIED_COUNT, "liberaux": LIBERAL_COUNT,
                         "part_salaries": SALARIED_SHARE + " (%)"}), hide_index=True, width="stretch")
        st.caption("Les nouvelles cohortes reçoivent la même répartition par statut que le total de leur département et de leur année. Cette extrapolation ne modélise pas les transitions individuelles entre modes d’exercice.")
    with tab_internat:
        calibration = state["vacancy_calibration"]
        st.write("Le coefficient corrige les arrivées futures, en comparant le taux de remplissage extrapolé à un taux constant de référence. Il s’applique à la cohorte d’entrée en internat, puis produit un effet sur les effectifs cinq ans plus tard.")
        st.info("La série inclut le recul récent des postes non pourvus : 86/554 en 2025 (15,52 %), puis 37/653 en 2026 (5,67 %), hors CESP pour ces deux derniers millésimes. L’évolution n’est donc pas une hausse continue.")
        st.caption(f"Estimation validée : {calibration['debut']}–{calibration['derniere_observation']}. Taux de référence : {100 * calibration['taux_reference']:.2f} %. Correction {'active' if proposed['internat']['active'] else 'désactivée'}. La courbe ci-dessous montre l’extrapolation estimée dans les deux cas.")
        st.plotly_chart(vacancy_figure(state["vacancy_curve"], calibration), width="stretch")
        st.caption(f"Coefficient de flux = min(1, (1 − taux futur) / (1 − taux de référence)). Si la tendance estimée s’améliore, aucune réduction supplémentaire n’est appliquée. Les cohortes entrées au plus tard en {anchor} conservent un coefficient de 1.")
        flows = state["national_flows"]
        shown = flows[flows.scenario == scenario]
        if shown.plafond_atteint.any():
            st.warning("L’extrapolation retirerait plus de psychiatres que le total de référence dans certaines années. Le retrait est limité à ce total ; ces horizons demandent de revoir les hypothèses.")
        st.dataframe(shown[["annee", "annee_entree_internat", "psychiatrie", "coefficient_entrees", "flux_fr_reference", "flux_fr_corrige", "manque_cumule_metro", "coefficient_stock"]].rename(
            columns={"annee": "Année d’arrivée", "annee_entree_internat": "Cohorte d’internat",
                     "psychiatrie": "Stock d’étudiants DREES", "coefficient_entrees": "Coefficient du flux",
                     "flux_fr_reference": "Flux France avant correction", "flux_fr_corrige": "Flux France corrigé",
                     "manque_cumule_metro": "Déficit cumulé métropole", "coefficient_stock": "Facteur uniforme des effectifs"}),
                     hide_index=True, width="stretch")
        st.write("Le stock d’étudiants est divisé par cinq pour approcher un flux annuel. Le déficit des futures cohortes est cumulé sans sorties ultérieures, puis retiré du total métropolitain selon la calibration formation/mobilité. Le même facteur annuel s’applique à chaque département. Les places supplémentaires utilisent le flux corrigé une seule fois.")
        st.warning("Scénario additionnel : la documentation publique ne fournit pas le taux exact de non-remplissage déjà inclus dans ProMédI. Le taux de référence constant est une convention, pas un paramètre DREES identifié. La division du stock par cinq et l’absence de sorties des cohortes manquantes restent des approximations.")
        with st.expander("Série historique et sources des postes non pourvus"):
            history_table = vacancy_history()
            history_table["Postes non pourvus (%)"] = history_table.taux * 100
            history_table["Source"] = history_table.source.map(VACANCY_SOURCES)
            st.dataframe(history_table[["annee", "Postes non pourvus (%)", "Source"]].rename(columns={"annee": "Année"}), hide_index=True, width="stretch")
            st.caption("2013–2022 : pourcentages publiés arrondis ; 2023–2026 : rapports calculés à partir des comptes publiés. Les anciennes publications ne précisent pas toutes le traitement des CESP. Cette limite de comparabilité est conservée dans le README.")
    with tab_method:
        first = proposed["formation"]["rentree"] + STUDY_YEARS
        st.write(f"Places : hausse annuelle permanente à partir de {proposed['formation']['rentree']}, avec premiers effets en {first}. Le délai retenu est de {STUDY_YEARS} ans après le bac : 6 ans avant le DES et 5 ans de DES. C'est une durée nominale, pas une moyenne mesurée.")
        st.write("Le stock national d'internes de psychiatrie de chaque année de sortie est divisé par 5 pour approcher un flux annuel. La géographie des places est approchée par les stocks régionaux d'internes toutes spécialités de la dernière année observée. Le flux est ensuite corrigé par le coefficient des postes non pourvus à l’entrée en internat, cinq ans plus tôt, si cette modulation est active.")
        st.write("Chaque cohorte supplémentaire est répartie selon la matrice diplôme → exercice des psychiatres RPPS, puis selon les parts départementales observées dans chaque région d'exercice. Les destinations hors métropole restent au dénominateur : elles ne sont pas redistribuées aux départements affichés. Les cohortes sont cumulées sans attrition ni retraite supplémentaire jusqu'en 2070.")
        st.caption("X dépend de la région, du scénario et de l'année. Le tableau tient compte de la correction de l’internat et donne l’effet de +1 % dans une seule région pour la première cohorte diplômée ; ensuite les cohortes s'accumulent. Les nombres fractionnaires sont des effectifs attendus.")
        st.dataframe(calibration_table(reference, context, scenario, proposed, state["national_flows"]), hide_index=True, width="stretch")
        missing = [REGIONS[code] for code in REGIONS if code not in context["available"]]
        if missing:
            st.warning("Calibration de formation absente pour : " + ", ".join(missing) + ". Une ouverture ex nihilo nécessite des hypothèses supplémentaires ; le curseur reste désactivé. Ces régions peuvent recevoir des diplômés formés ailleurs.")
        st.write(f"Dynamiques : régression des écarts d'effectif départemental à la référence nationale sur {proposed['dynamique']['debut']}–{anchor}. La référence vaut le total métropolitain de chaque année multiplié par la part d'effectif du département observée en {anchor}. Chaque pente est exprimée en psychiatres par an.")
        st.write(f"Projection avant saturation = total national de référence × part départementale de {anchor} + pente départementale × années depuis {anchor}. Les pentes somment à zéro. L'effet commence dès {anchor + 1}. Le lissage sous 1 psychiatre et la renormalisation préservent des effectifs non négatifs et le total national.")
        if metadata["mode_projection"] == "regional":
            st.write(f"Lorsque les dynamiques sont activées, le total des projections régionales sert de trajectoire nationale : leur ventilation régionale est remplacée par les parts {anchor} et les écarts départementaux. Désactiver les dynamiques rétablit la ventilation régionale de référence.")
        with st.expander("Sources et champ des effectifs"):
            st.write(f"Onglet de projection des actifs : {metadata['onglet_drees']}.")
            if metadata["mode_projection"] == "national":
                st.write(f"Référence départementale = effectif RPPS en {anchor} × effectif national DREES futur / effectif national DREES en {anchor}. La croissance du champ DREES (70 ans ou moins) est appliquée au RPPS tous âges ; les plus de 70 ans ne font pas l'objet d'une projection distincte.")
            else:
                st.write("Les effectifs régionaux DREES de référence sont utilisés directement. Le RPPS comprend tous les âges, contre 70 ans ou moins pour les actifs DREES : une rupture de champ peut apparaître au raccord.")
            st.write("Les résultats sont des nombres de psychiatres et une part salariée en %. Les observations proviennent directement de l’onglet Effectifs. Les valeurs projetées fractionnaires représentent des effectifs attendus.")
            st.write("Aucun fichier de projections n'est écrit. Le calcul de référence est partagé en mémoire ; les paramètres et résultats simulés appartiennent à chaque session.")
        st.caption("Les formules, limites et règles d'extension sont détaillées dans README.md.")


if __name__ == "__main__":
    main()
