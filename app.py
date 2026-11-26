"""Carte des psychiatres : calcul en mémoire, propre à chaque session utilisateur.

Installation : pip install streamlit plotly pandas numpy openpyxl
Lancement   : streamlit run app.py

Version : France (101 départements), correspondances fixes, périmètre adulte,
plateau de postes vacants et politiques nationales de prévalence par âge.
Les politiques sont des hypothèses de scénario, appliquées avant l'agrégation.

Placer les deux classeurs et requirements.txt à côté de ce script. Sur Streamlit
Cloud, requirements.txt installe notamment openpyxl, nécessaire pour lire Excel.

Les résultats sont un DataFrame de cinq colonnes :
    dep, annee, scenario, indicateur, valeur
Ils restent dans st.session_state, sans aucun fichier de résultats ou paramètres.
Chaque utilisateur dispose de ses propres paramètres et projections. Changer
seulement un filtre d'affichage ne relance pas la simulation.

Interface : deux pages réunies dans le « Tableau de bord » de la barre latérale.
  - Scénarios : choisir un scénario prédéfini, en créer un, l'exporter ou l'importer
    (CSV, une ligne par paramètre). Les paramètres ne s'appliquent qu'au clic sur
    « Appliquer » et reviennent au scénario Base dans une nouvelle session.
  - Projections : carte et sélecteurs (thème, indicateur, projection DREES, année),
    puis les onglets de diagnostic du scénario actif.

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
Les formules utilisées sont décrites dans l'onglet Calibration et hypothèses.
La demande nécessite demand_projection/run.py, outcomes.py et les CSV de dataDemand/.
Seuls les classeurs Insee présents sont utilisés ; un scénario démographique suffit.
"""
from dataclasses import dataclass
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from typing import Callable
import csv
import json
import re
import unicodedata
import urllib.request
from io import BytesIO, StringIO
from html import escape

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
DEMAND_SUMMARY_PATH = ROOT / "dataDemand" / "output" / "demand_projection_summary.csv"
TOTAL_POPULATION_PATH = ROOT / "dataDemand" / "output" / "population_totale.csv"
ACTIVITY_PROJECTION_PATH = ROOT / "dataSupply" / "actes_par_psychiatre.csv"
GEOJSON_URL = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements-avec-outre-mer.geojson"
COLUMNS = ["dep", "annee", "scenario", "indicateur", "valeur"]
COUNT = "Effectif de psychiatres"
SALARIED_SHARE = "Part de salariés"
LIBERAL_COUNT = "Nombre brut de libéraux"
SALARIED_COUNT = "Nombre brut de salariés"
INDICATORS = {"effectif": COUNT, "part_salaries": SALARIED_SHARE,
              "liberaux": LIBERAL_COUNT, "salaries": SALARIED_COUNT}
ACTS_PER_PSY = "Actes annuels par psychiatre"
AVAILABLE_ACTS = "Actes disponibles pour le besoin sélectionné"
REQUIRED_ACTS = "Actes nécessaires pour le besoin sélectionné"
ACTS_BALANCE = "Solde d’actes par habitant et par an"
PSY_TARGET = "Effectif cible estimé de psychiatres"
PSY_BALANCE = "Solde de psychiatres par rapport à la cible"
PSY_SHORTFALL = "Psychiatres manquants pour atteindre la cible"
PSY_SURPLUS = "Psychiatres au-dessus de la cible"
NEED_COVERAGE = "Couverture théorique des actes nécessaires (%)"
RESIDENTS_PER_PSY = "Habitants par psychiatre"
MATCH_INDICATORS = {"actes_par_psychiatre": ACTS_PER_PSY, "actes_disponibles": AVAILABLE_ACTS,
                    "actes_necessaires": REQUIRED_ACTS, "solde_actes_habitant": ACTS_BALANCE,
                    "psychiatres_cibles": PSY_TARGET, "solde_psychiatres": PSY_BALANCE,
                    "psychiatres_manquants": PSY_SHORTFALL, "psychiatres_surplus": PSY_SURPLUS,
                    "couverture_actes_pct": NEED_COVERAGE, "habitants_par_psychiatre": RESIDENTS_PER_PSY}
ACTIVITY_INDICATORS = ("actes_par_psychiatre", "actes_disponibles", "actes_necessaires", "solde_actes_habitant")
# 15,865 millions de consultations/téléconsultations / 6 308 psychiatres libéraux (CNAM 2024).
# Les autres valeurs sont des hypothèses de scénario, pas des normes de soins.
MATCH_DEFAULTS = {"actes_liberal": 2515.0, "ratio_salarie": 0.65, "croissance_activite": 0.0,
                  "actes_par_besoin": 8.0, "couverture": 25.0, "part_offre": 50.0,
                  "marge_activite": 10.0}
MATCH_SOURCES = {
    "Activité libérale CNAM 2024": "https://www.assurance-maladie.ameli.fr/sites/default/files/2024_fiche_psychiatres-et-neuro-psychiatres-liberaux.pdf",
    "Suivi individualisé HAS": "https://www.has-sante.fr/jcms/pprd_2974237/fr/depression-de-l-adulte-reperage-et-prise-en-charge-initiale",
}
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
DROM = {"971": "Guadeloupe", "972": "Martinique", "973": "Guyane", "974": "La Réunion", "976": "Mayotte"}
DEPARTMENTS = METRO | set(DROM)
DES_YEARS = 5
STUDY_YEARS = 5 + DES_YEARS  # Depuis l'entrée en deuxième année ; réglable.
VACANCY_LINEAR = "Extrapolation linéaire"
VACANCY_PLATEAU = "Plateau au maximum observé"
SMOOTH_FLOOR = 1.0  # Psychiatre ; raccord C1 à la fonction identité.
REGIONS = {
    "01": "Guadeloupe", "02": "Martinique", "03": "Guyane", "04": "La Réunion", "06": "Mayotte",
    "11": "Île-de-France", "24": "Centre-Val de Loire",
    "27": "Bourgogne-Franche-Comté", "28": "Normandie",
    "32": "Hauts-de-France", "44": "Grand Est", "52": "Pays de la Loire",
    "53": "Bretagne", "75": "Nouvelle-Aquitaine", "76": "Occitanie",
    "84": "Auvergne-Rhône-Alpes", "93": "Provence-Alpes-Côte d’Azur", "94": "Corse",
}
# Les lieux de formation ultramarins sont groupés dans les sources DREES/RPPS.
TRAINING_REGIONS = {code: label for code, label in REGIONS.items() if code not in {"02", "03", "06"}}
TRAINING_REGIONS.update({"01": "Antilles-Guyane", "04": "La Réunion-Mayotte"})



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
            if normalize(row[pos["territoire"]]) not in {"francemetropolitaine", "drom"}:
                continue
            if normalize(row[pos["specialites"]]) != "psychiatrie":
                continue
            raw_dep = str(row[pos["departement"]]).split("-", 1)[0].strip()
            dep = raw_dep[-2:] if len(raw_dep) == 3 and raw_dep.startswith("0") else raw_dep
            if dep not in DEPARTMENTS:
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
            or set(history.dep) != DEPARTMENTS or len(history) != len(DEPARTMENTS) * len(years)):
        raise ValueError("Historique départemental incomplet.")
    return history, pd.DataFrame(names)


def pediatric_shares(path):
    """Part nationale des qualifications enfant/adolescent, sans double comptage."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = book["Effectifs par spécialités"].iter_rows(values_only=True)
        header = next(rows)
        p = {value: i for i, value in enumerate(header) if value is not None}
        selected = [row for row in rows if normalize(row[p["specialites"]]) == "psychiatrie"]
        total = [row for row in selected if normalize(row[p["specialites_detaillees"]]) == "ensemble"]
        children = [row for row in selected if "enfant" in normalize(row[p["specialites_detaillees"]])
                    and "adolescent" in normalize(row[p["specialites_detaillees"]])]
        if len(total) != 1 or not children:
            raise ValueError("Qualifications de pédopsychiatrie ou total national absents du RPPS.")
        records = []
        for column, index in p.items():
            if re.fullmatch(r"effectif_\d{4}", str(column)):
                n, child = float(total[0][index]), sum(float(row[index]) for row in children)
                if not np.isfinite([n, child]).all() or not 0 <= child < n:
                    raise ValueError("Part nationale de pédopsychiatrie invalide.")
                records.append({"annee": int(column[-4:]), "total": n, "pedopsychiatres": child, "part": child / n})
        return pd.DataFrame(records).sort_values("annee")
    finally:
        book.close()


def scope_factors(years, context, excluded):
    if not excluded:
        return np.ones(len(years))
    shares = context["pediatric_shares"].set_index("annee").part
    values = pd.Series(np.minimum(np.asarray(years), context["anchor"])).map(shares)
    if values.isna().any():
        raise ValueError("Part nationale de pédopsychiatrie manquante pour certaines années.")
    return 1 - values.to_numpy()


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


def vacancy_projection(start, anchor, end, method=VACANCY_LINEAR):
    """Tendance linéaire des taux ; raccord au dernier taux réellement publié.

    MCO avec constante et poids égaux par année. Seule la pente est prolongée,
    en fractions par an. Les proportions projetées sont bornées entre 0 et 1.
    """
    observed = vacancy_history()
    observed = observed[observed.annee <= anchor]
    estimation = observed[observed.annee >= start]
    if len(estimation) < 5:
        raise ValueError("Au moins cinq années sont nécessaires pour la tendance d'internat.")
    x = estimation.annee.to_numpy(dtype=float)
    y = estimation.taux.to_numpy(dtype=float)
    centered = x - x.mean()
    slope = float(np.dot(centered, y - y.mean()) / np.dot(centered, centered))
    last = observed.sort_values("annee").iloc[-1]
    years = np.arange(int(observed.annee.min()), end + 1)
    curve = pd.DataFrame({"annee": years})
    raw = float(last.taux) + slope * (years - last.annee)
    if method == VACANCY_PLATEAU:
        raw = np.where(years > anchor, float(estimation.taux.max()), raw)
    elif method != VACANCY_LINEAR:
        raise ValueError("Méthode de projection des postes non pourvus inconnue.")
    curve["taux_projete"] = np.clip(raw, 0.0, 1.0)
    curve["borne_atteinte"] = (years > last.annee) & ((raw < 0.0) | (raw > 1.0))
    curve["taux_observe"] = curve.annee.map(observed.set_index("annee").taux)
    curve["taux"] = curve.taux_observe.fillna(curve.taux_projete)
    # On compare une dégradation future à un comportement constant à l'ancrage.
    # Le niveau déjà observé n'est pas soustrait une seconde fois aux stocks DREES.
    baseline = float(curve.loc[curve.annee == anchor, "taux"].iloc[0])
    if baseline >= 1.0:
        raise ValueError("Le taux de référence est de 100 % : le coefficient relatif de remplissage n'est pas défini.")
    curve["coefficient"] = np.where(curve.annee > anchor,
                                     np.minimum(1.0, (1 - curve.taux) / (1 - baseline)), 1.0)
    return curve, {"pente_annuelle": slope, "taux_reference": baseline,
                   "derniere_observation": int(last.annee), "debut": start,
                   "methode": method, "plateau": float(estimation.taux.max())}


def prepare_national_flows(reference, parameters, context):
    """Calcul commun aux départements ; chaque coefficient porte sur une cohorte."""
    anchor = context["anchor"]
    curve, calibration = vacancy_projection(parameters["debut"], anchor, int(reference.annee.max()), parameters["methode"])
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
    # Le champ est désormais la France : le flux national ne subit plus de retrait DROM.
    flows["part_flux_france"] = 1.0
    flows["manque_annuel_france"] = (flows.flux_fr_reference - flows.flux_fr_corrige) * flows.part_flux_france
    flows["manque_cumule_france"] = flows.groupby("scenario").manque_annuel_france.cumsum()
    totals = reference[reference.projete].groupby(["annee", "scenario"], as_index=False).effectif.sum()
    flows = flows.merge(totals.rename(columns={"effectif": "total_reference"}), on=["annee", "scenario"], validate="one_to_one")
    flows["retrait_applique"] = np.minimum(flows.manque_cumule_france, flows.total_reference)
    flows["plafond_atteint"] = flows.manque_cumule_france > flows.total_reference
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
def training_students(path, anchor, names, mobility):
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
        for label, origin in {"guyanemartiniqueguadeloupe": "01", "lareunionmayotte": "04"}.items():
            value = float(detail.loc[detail.libelle == label, "stock"].sum())
            if value:
                weights.append({"scenario": scenario, "origine": origin, "poids": value / total.iloc[0]})
        for origin, value in detail.dropna(subset=["origine"]).groupby("origine").stock.sum().items():
            weights.append({"scenario": scenario, "origine": origin, "poids": value / total.iloc[0]})
    weights = pd.DataFrame(weights)
    coverage = weights.groupby("scenario").origine.apply(lambda values: frozenset(values))
    if coverage.nunique() != 1:
        raise ValueError("Régions de formation manquantes dans certains scénarios.")
    return stocks, weights


def diploma_history(path, anchor):
    """Comptes annuels par lieu de diplôme et région d'exercice, dont le code 99."""
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
            for column, index in p.items():
                if not re.fullmatch(r"effectif_\d{4}", column) or int(column[-4:]) > anchor:
                    continue
                value = row[index]
                if not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
                    raise ValueError(f"Cellule de mobilité manquante ou invalide : {column}.")
                records.append({"origine": str(row[p["region_diplome"]]).split("-", 1)[0].strip().zfill(2),
                                "destination": str(row[p["region"]]).split("-", 1)[0].strip().zfill(2),
                                "annee": int(column[-4:]), "effectif": float(value)})
    finally:
        book.close()
    cells = pd.DataFrame(records)
    if cells.empty or cells.duplicated(["origine", "destination", "annee"]).any():
        raise ValueError("Matrice diplôme/exercice absente ou dupliquée.")
    years = sorted(cells.annee.unique())
    if years != list(range(min(years), anchor + 1)):
        raise ValueError("Historique des lieux de diplôme incomplet.")
    margins = cells[(cells.destination == "00") & (cells.origine != "00")].set_index(["origine", "annee"]).effectif
    detail = cells[(cells.destination != "00") & (cells.origine != "00")].copy()
    origins = sorted(margins.index.get_level_values("origine").unique())
    complete_margins = pd.MultiIndex.from_product([origins, years], names=["origine", "annee"])
    if not margins.index.sort_values().equals(complete_margins) or "99" not in origins:
        raise ValueError("Marges annuelles ou origine 99 absentes de la matrice de diplôme.")
    sums = detail.groupby(["origine", "annee"]).effectif.sum().reindex(margins.index, fill_value=0)
    if not np.allclose(sums, margins, atol=0.01, rtol=0):
        raise ValueError("La matrice diplôme/exercice ne reconstitue pas ses marges.")
    # Les cellules absentes ne sont des zéros qu'après vérification des marges.
    destinations = sorted(cells.loc[cells.destination != "00", "destination"].unique())
    grid = pd.MultiIndex.from_product([origins, destinations, years], names=["origine", "destination", "annee"])
    result = detail.set_index(["origine", "destination", "annee"]).effectif.reindex(grid, fill_value=0).reset_index()
    return result


def project_diploma_distribution(history, parameters, anchor, end):
    """Matrice diplôme -> exercice fixe ; tendance des stocks nationaux par origine."""
    selected = history[history.annee.between(parameters["debut"], anchor)]
    base = history[history.annee == anchor][["origine", "destination", "effectif"]].rename(columns={"effectif": "ancrage"})
    totals = base.groupby("origine").ancrage.transform("sum")
    base["part"] = np.divide(base.ancrage, totals, out=np.zeros(len(base)), where=totals > 0)
    national = selected.groupby(["origine", "annee"], as_index=False).effectif.sum()
    slopes = {}
    for origin, group in national.groupby("origine"):
        if len(group) < 5:
            raise ValueError("Au moins cinq années pour estimer les stocks par lieu de diplôme.")
        slopes[origin] = float(np.polyfit(group.annee - anchor, group.effectif, 1)[0])
    base["stock_origine"] = totals
    base["pente_origine"] = base.origine.map(slopes)
    panel = base.merge(pd.DataFrame({"annee": range(anchor, end + 1)}), how="cross")
    panel["effectif"] = np.maximum(0, panel.stock_origine + panel.pente_origine * (panel.annee - anchor)) * panel.part
    panel["projete"] = panel.annee > anchor
    panel["repli"] = "Correspondance observée fixe"
    return panel


def annual_routing(panel, baseline):
    routing = panel[["origine", "destination", "annee", "part"]].merge(
        baseline[["dep", "region_code", "part_regionale"]],
        left_on="destination", right_on="region_code", validate="many_to_many")
    routing["part"] *= routing.part_regionale
    return routing[["origine", "annee", "dep", "part"]]


def accumulate_arrivals(reference, cohorts, routing):
    """Router chaque cohorte avec les parts de son année, puis cumuler sans sorties."""
    distributed = cohorts.merge(routing, on=["origine", "annee"], validate="many_to_many")
    distributed["flux_ajoute"] = distributed.flux_ajoute * distributed.part
    annual = distributed.groupby(["dep", "annee", "scenario"], as_index=False).flux_ajoute.sum()
    result = reference.loc[reference.projete, ["dep", "annee", "scenario"]].merge(
        annual, on=["dep", "annee", "scenario"], how="left", validate="one_to_one")
    result["flux_ajoute"] = result.flux_ajoute.fillna(0.0)
    result = result.sort_values(["scenario", "dep", "annee"])
    result["cumul"] = result.groupby(["scenario", "dep"]).flux_ajoute.cumsum()
    return result


def foreign_reference_flows(panel, anchor):
    """Variation nette du stock étranger projeté, sans hypothèse de retraite."""
    foreign = panel[(panel.origine == "99") & (panel.destination != "99")]
    totals = foreign.groupby("annee", as_index=False).effectif.sum().sort_values("annee")
    totals["ecart_flux"] = totals.effectif.diff().fillna(0.0)
    totals["flux_reference"] = totals.ecart_flux.clip(lower=0.0)
    totals["origine"] = "99"
    return totals


def foreign_france_distribution(panel):
    """Correspondance fixe, conditionnelle à un exercice en France, DROM inclus."""
    france = panel[(panel.origine == "99") & (panel.destination != "99")].copy()
    total = france.groupby("annee").part.transform("sum")
    if (total <= 0).any():
        raise ValueError("Aucune correspondance française pour les diplômés à l'étranger.")
    france["part"] /= total
    return france


def prepare_arrivals(reference, parameters, context):
    flows = context["national_flows"]
    training = flows[["annee", "scenario", "flux_fr_reference", "coefficient_entrees"]].merge(
        context["weights"], on="scenario", validate="many_to_many")
    training["flux_reference"] = training.flux_fr_reference * training.poids
    training["flux_corrige"] = training.flux_reference * training.coefficient_entrees
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    increases = training.origine.map(lambda code: parameters["formation"].get(f"region_{code}", 0) / 100.0)
    training["flux_ajoute"] = np.where(training.annee >= first, training.flux_corrige * increases, 0.0)
    panel = context["mobility_panel"]
    foreign = foreign_reference_flows(panel, context["anchor"])
    foreign = foreign.merge(pd.DataFrame({"scenario": sorted(reference.scenario.unique())}), how="cross")
    foreign["flux_hausse"] = np.where(foreign.annee >= parameters["immigration"]["debut"],
        foreign.flux_reference * parameters["immigration"]["hausse"] / 100.0, 0.0)
    foreign["flux_ajoute"] = foreign.ecart_flux + foreign.flux_hausse
    routing = annual_routing(foreign_france_distribution(panel), context["baseline"])
    return {"training_cohorts": training, "foreign_cohorts": foreign,
            "training_arrivals": accumulate_arrivals(reference, training, context["routing"]),
            "foreign_routing": routing, "foreign_intervention_routing": routing,
            "foreign_trend_arrivals": accumulate_arrivals(reference, foreign.assign(flux_ajoute=foreign.ecart_flux), routing),
            "foreign_arrivals": accumulate_arrivals(reference, foreign.assign(flux_ajoute=foreign.flux_hausse), routing)}


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
    origin_totals = mobility[mobility.annee == anchor].groupby("origine").effectif.sum()
    available = set(origin_totals[origin_totals > 0].index) & set(REGIONS) & set(weights.loc[weights.poids > 0, "origine"])
    return {"anchor": anchor, "baseline": baseline, "stocks": stocks,
            "history": history.copy(deep=True),
            "status_models": fit_status_models(history, anchor),
            "weights": weights, "mobility_history": mobility,
            "mobility_start": int(mobility.annee.min()),
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
    """Hausse permanente des cohortes de deuxième année, routage annuel puis cumul."""
    return add_arrival_stock(projected, context["training_arrivals"])


def immigration_additions(projected, parameters, context):
    """Hausse des arrivées de diplômés à l'étranger, sans délai de formation."""
    return add_arrival_stock(projected, context["foreign_arrivals"])


def foreign_trend_adjustment(projected, parameters, context):
    """Écart cumulé au flux étranger de b, actif même à 0 % de hausse."""
    result = add_arrival_stock(projected, context["foreign_trend_arrivals"])
    result["effectif"] = result.effectif.clip(lower=0.0)
    return result


def add_arrival_stock(projected, arrivals):
    result = projected.copy()
    keys = pd.MultiIndex.from_frame(result[["dep", "annee", "scenario"]])
    added = arrivals.set_index(["dep", "annee", "scenario"]).cumul.reindex(keys)
    if added.isna().any():
        raise ValueError("Cumul des installations absent pour certaines lignes.")
    result["effectif"] += added.to_numpy()
    return result


def calibration_only(projected, parameters, context):
    return projected


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
                if code in context["available"] else "Pas d'effectif de référence disponible pour ce lieu de diplôme ; curseur désactivé.")
        for code, label in sorted(TRAINING_REGIONS.items(), key=lambda item: item[1])
    )
    return (
        Modulation("perimetre", "Périmètre de soins",
                   "Le retrait des pédopsychiatres applique la même part nationale à chaque département et statut. La demande des moins de 18 ans est également exclue.",
                   (Control("exclure_pedopsychiatres", "Exclure les pédopsychiatres et la demande des moins de 18 ans", "checkbox", False),), calibration_only),
        Modulation("mobilite", "Correspondance diplôme → exercice",
                   f"Parts régionales fixes observées en {context['anchor']}, puis parts départementales fixes. Les nouveaux diplômés français sont calibrés par les stocks d'internes DREES divisés par cinq. Aucun taux de retraite n'intervient.",
                   (Control("debut", "Première année pour la tendance des effectifs diplômés à l'étranger", "number",
                            context["mobility_start"], context["mobility_start"], context["anchor"] - 4, 1),), calibration_only),
        Modulation("dynamique", "Dynamiques départementales",
                   f"Tendances des effectifs (psychiatres par an), estimées jusqu'à {context['anchor']}. Effet dès {context['anchor'] + 1}, indépendant du délai de formation.",
                   (Control("active", "Extrapoler les dynamiques départementales", "checkbox", True),
                    Control("debut", "Première année pour estimer la tendance", "number", context["observed_start"],
                            context["observed_start"], context["anchor"] - 4, 1,
                            help="Tout l'historique par défaut. Choisir 2017 pour dix années, ou 2022 pour cinq années avec les sources 2026. Une fenêtre courte peut modifier fortement la pente.")), department_dynamics),
        Modulation("internat", "Dynamique nationale de l'internat",
                   "Correction des futures cohortes : tendance linéaire ou plateau immédiat au maximum observé sur la période sélectionnée, dès la première année projetée.",
                   (Control("active", "Appliquer la correction des postes non pourvus", "checkbox", True),
                    Control("methode", "Projection des postes non pourvus", "select", VACANCY_LINEAR,
                            options=(VACANCY_LINEAR, VACANCY_PLATEAU)),
                    Control("debut", "Première année pour la tendance des postes non pourvus", "number",
                            min(v[0] for v in VACANCY_OBSERVATIONS), min(v[0] for v in VACANCY_OBSERVATIONS),
                            max(v[0] for v in VACANCY_OBSERVATIONS) - 4, 1,
                            help="Régression linéaire des taux, avec au moins cinq années, puis raccord au dernier taux observé. Les taux restent entre 0 et 100 %. Une tendance à l'amélioration ne crée pas de hausse supplémentaire : le coefficient est plafonné à 1.")),
                   internat_correction),
        Modulation("etranger_tendance", "Tendance nationale des arrivées étrangères", "", (), foreign_trend_adjustment),
        Modulation("immigration", "Diplômés à l'étranger — code 99",
                   "La tendance du stock national reste active à 0 %. Le curseur majore son accroissement net positif annuel, réparti avec la correspondance observée. Ce flux net ne mesure pas les entrées brutes. Le code 99 désigne le lieu du diplôme.",
                   (Control("hausse", "Majoration de l'accroissement net des diplômés à l'étranger (%)", "slider", 0, 0, 100, 1),
                    Control("debut", "Première année de hausse des arrivées étrangères", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()), 1)), immigration_additions),
        Modulation("formation", "Places de formation par région",
                   "Hausse des places en deuxième année, reconduite pour chaque cohorte. Volume : stock national d'internes en psychiatrie / 5 × poids régional de formation DREES. Répartition selon la correspondance diplôme → exercice observée.",
                   (Control("rentree", "Première rentrée augmentée en deuxième année", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()) - DES_YEARS, 1),
                    Control("delai", "Délai deuxième année → installation (années)", "number", STUDY_YEARS,
                            DES_YEARS, 20, 1,
                            help="Dix ans par défaut : cinq ans avant le DES, puis cinq ans de DES. Le délai du DES utilisé pour les postes non pourvus reste de cinq ans.")) + regional_controls,
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
    formation = parameters["formation"]
    end = next(m for m in registry if m.key == "immigration").controls[1].maximum
    if formation["rentree"] + formation["delai"] > end:
        raise ValueError(f"La première installation doit intervenir au plus tard en {end}. Réduire la rentrée ou le délai.")


def apply_modulations(reference, parameters, context, return_contributions=False, return_details=False):
    registry = modulation_registry(context)
    validate_parameters(parameters, registry)
    run_context = deepcopy(context)
    mobility_panel = project_diploma_distribution(context["mobility_history"], parameters["mobilite"],
                                                  context["anchor"], int(reference.annee.max()))
    run_context["mobility_panel"] = mobility_panel
    run_context["routing"] = annual_routing(mobility_panel, context["baseline"])
    run_context["intervention_routing"] = run_context["routing"]
    flows, curve, calibration = prepare_national_flows(reference, parameters["internat"], run_context)
    run_context["national_flows"] = flows
    arrival_details = prepare_arrivals(reference, parameters, run_context)
    run_context.update(arrival_details)
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
    factors = scope_factors(result.annee, context, parameters["perimetre"]["exclure_pedopsychiatres"])
    result[["effectif", "salaries", "liberaux"]] = result[["effectif", "salaries", "liberaux"]].mul(factors, axis=0)
    contributions = contributions.mul(factors, axis=0)
    if return_details:
        keys = pd.MultiIndex.from_frame(projected[["dep", "annee", "scenario"]])
        requested = run_context["foreign_trend_arrivals"].set_index(["dep", "annee", "scenario"]).cumul.reindex(keys).to_numpy() * factors[projected.index]
        limited = int((~np.isclose(contributions.loc[projected.index, "etranger_tendance"], requested, atol=1e-8, rtol=1e-10)).sum())
        return result, contributions, {"national_flows": flows, "vacancy_curve": curve,
            "vacancy_calibration": calibration, "mobility_panel": mobility_panel,
            "routing": run_context["routing"], "intervention_routing": run_context["intervention_routing"],
            "foreign_trend_limited_rows": limited, **arrival_details}
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
    mobility = diploma_history(RPPS_PATH, anchor)
    stocks, weights = training_students(DREES_PATH, anchor, names, mobility)
    context = build_context(history, names, stocks, weights, mobility)
    context["pediatric_shares"] = pediatric_shares(RPPS_PATH)
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


# ---------- Politiques nationales : règles explicites par trouble et tranche d'âge ----------
POLICY_COLUMNS = ["id", "nom", "active", "indicateur", "age_min", "age_max", "baisse_pct", "debut", "fin", "montee"]
POLICY_INDICATORS = {
    "tous": "Tous les troubles couverts",
    "episode_depressif_12_mois": "Dépression (18–79 ans)",
    "trouble_anxieux_generalise_12_mois": "Anxiété généralisée (18–79 ans)",
    "trouble_mental_probable_6_11": "Trouble probable chez l'enfant (6–11 ans)",
}
POLICY_COVERAGE = {"episode_depressif_12_mois": (18, 79),
                   "trouble_anxieux_generalise_12_mois": (18, 79),
                   "trouble_mental_probable_6_11": (6, 11)}


def policy_templates(anchor=2026):
    """Hypothèses illustratives, sans efficacité empirique revendiquée."""
    def rule(identifier, name, indicator, lo, hi, reduction):
        return {"id": identifier, "nom": name, "active": True, "indicateur": indicator,
                "age_min": lo, "age_max": hi, "baisse_pct": float(reduction),
                "debut": anchor + 1, "fin": 2070, "montee": 1}
    return [rule("scolaire", "Prévention de l'anxiété à l'école — illustration", "trouble_anxieux_generalise_12_mois", 6, 17, 15),
            rule("jeunes", "Prévention de la dépression des jeunes adultes — illustration", "episode_depressif_12_mois", 18, 25, 10),
            rule("enfance", "Prévention des troubles de l'enfant — illustration", "trouble_mental_probable_6_11", 6, 11, 10)]


def validate_policies(policies):
    if not isinstance(policies, list) or len(policies) > 100:
        raise ValueError("Les politiques doivent être une liste d'au plus 100 règles.")
    ids = set()
    for policy in policies:
        if not isinstance(policy, dict) or set(policy) != set(POLICY_COLUMNS):
            raise ValueError("Format de politique invalide ; utiliser le modèle exporté.")
        if not isinstance(policy["id"], str) or not re.fullmatch(r"[a-zA-Z0-9_-]{1,80}", policy["id"]) or policy["id"] in ids:
            raise ValueError("Chaque règle doit avoir un identifiant unique (lettres, chiffres, - ou _).")
        ids.add(policy["id"])
        if not isinstance(policy["nom"], str) or not policy["nom"].strip() or len(policy["nom"]) > 160:
            raise ValueError("Le nom d'une politique doit contenir de 1 à 160 caractères.")
        if not isinstance(policy["active"], bool) or policy["indicateur"] not in POLICY_INDICATORS:
            raise ValueError("Activation ou trouble ciblé invalide.")
        for key, low, high in [("age_min", 0, 120), ("age_max", 0, 120),
                               ("debut", 2026, 2070), ("fin", 2026, 2070), ("montee", 1, 45)]:
            value = policy[key]
            if isinstance(value, bool) or not isinstance(value, int) or not low <= value <= high:
                raise ValueError(f"Valeur entière hors limites pour {key} : {value!r}.")
        reduction = policy["baisse_pct"]
        if isinstance(reduction, bool) or not isinstance(reduction, (int, float)) or not np.isfinite(reduction) or not 0 <= reduction <= 100:
            raise ValueError("La baisse de prévalence doit être comprise entre 0 et 100 %.")
        if policy["age_min"] > policy["age_max"] or policy["debut"] > policy["fin"]:
            raise ValueError("L'âge ou l'année de début doit précéder la fin.")
    # Une même politique ne peut appliquer deux fois son effet à la même cellule.
    for index, a in enumerate(policies):
        for b in policies[index + 1:]:
            if (a["active"] and b["active"] and a["nom"].strip().casefold() == b["nom"].strip().casefold()
                    and (a["indicateur"] == b["indicateur"] or "tous" in {a["indicateur"], b["indicateur"]})
                    and max(a["age_min"], b["age_min"]) <= min(a["age_max"], b["age_max"])
                    and max(a["debut"], b["debut"]) <= min(a["fin"], b["fin"])):
                raise ValueError(f"Deux règles de « {a['nom']} » se recouvrent : séparer les âges ou les années.")


def policy_warnings(policies, exclude_children=False):
    messages = []
    for p in policies:
        if not p["active"]:
            continue
        codes = list(POLICY_COVERAGE) if p["indicateur"] == "tous" else [p["indicateur"]]
        covered = set()
        for code in codes:
            lo, hi = POLICY_COVERAGE[code]
            lo = max(lo, 18) if exclude_children else lo
            covered.update(range(max(lo, p["age_min"]), min(hi, p["age_max"]) + 1))
        requested = set(range(p["age_min"], p["age_max"] + 1))
        if not covered:
            messages.append(f"{p['nom']} : aucun effet calculable dans le périmètre actuel ; aucun taux de prévalence n'est disponible pour ce trouble à ces âges.")
        elif requested - covered:
            messages.append(f"{p['nom']} : effet limité aux âges couverts ({min(covered)}–{max(covered)} ans) ; les autres âges n'ont pas de prévalence dans ce modèle.")
    return messages


def policies_csv(policies):
    return pd.DataFrame(policies, columns=POLICY_COLUMNS).to_csv(index=False, sep=";").encode("utf-8-sig")


def parse_policies_csv(data):
    if len(data) > 1_000_000:
        raise ValueError("Fichier de politiques trop volumineux (maximum 1 Mo).")
    table = pd.read_csv(BytesIO(data), sep=None, engine="python", dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if set(table.columns) != set(POLICY_COLUMNS):
        raise ValueError("Colonnes de politiques incorrectes ; télécharger le modèle CSV.")
    result = []
    for row in table.to_dict("records"):
        active = row["active"].strip().lower()
        if active not in TRUE_WORDS | FALSE_WORDS:
            raise ValueError("La colonne active doit valoir true/false ou oui/non.")
        row["active"] = active in TRUE_WORDS
        for key in ["age_min", "age_max", "debut", "fin", "montee"]:
            row[key] = parse_number(row[key], True)
        row["baisse_pct"] = parse_number(row["baisse_pct"], False)
        result.append(row)
    validate_policies(result)
    return result


def policy_factor(years, policies):
    factor = np.ones(len(years))
    values = np.asarray(years)
    for p in policies:
        progress = np.clip((values - p["debut"] + 1) / p["montee"], 0, 1)
        progress = np.where(values <= p["fin"], progress, 0)
        factor *= 1 - (p["baisse_pct"] / 100) * progress
    return factor


def project_policy_demand(population, years, scenario, parameters, scenarios, modifiers, policies, exclude_children):
    """Découper aux frontières d'âge des politiques avant d'agréger les effectifs."""
    from demand_projection.run import project_workbook, summarize_projection
    validate_policies(policies)
    records = []
    for _, parameter in parameters.iterrows():
        lo, hi = int(parameter.age_min), int(parameter.age_max)
        if exclude_children:
            lo = max(lo, 18)
        if lo > hi:
            continue
        applicable = [p for p in policies if p["active"] and p["indicateur"] in ("tous", parameter.indicator)
                      and p["age_min"] <= hi and p["age_max"] >= lo]
        cuts = sorted({lo, hi + 1} | {max(lo, p["age_min"]) for p in applicable}
                      | {min(hi + 1, p["age_max"] + 1) for p in applicable})
        for left, right in zip(cuts, cuts[1:]):
            row = parameter.copy()
            row["age_min"], row["age_max"] = left, right - 1
            part = project_workbook(population, years, scenario, pd.DataFrame([row]), scenarios, modifiers, 2026)
            rules = [p for p in applicable if p["age_min"] <= left and p["age_max"] >= right - 1]
            part["persons_without_policy"] = part.persons_with_need
            part["persons_with_need"] *= policy_factor(part.year, rules)
            part["prevalence"] = part.persons_with_need / part.population_at_risk.replace(0, np.nan)
            part["prevalence"] = part.prevalence.fillna(0)
            # Rétablir le périmètre de l'indicateur avant l'agrégation entre segments.
            part["age_min"], part["age_max"] = lo, hi
            records.append(part)
    if not records:
        raise ValueError("Aucun besoin couvert dans le périmètre sélectionné.")
    detail = pd.concat(records, ignore_index=True)
    summary = summarize_projection(detail)
    keys = ["department", "year", "demographic_scenario", "need_scenario", "indicator", "age_min", "age_max"]
    before = detail.groupby(keys, as_index=False).persons_without_policy.sum()
    summary = summary.merge(before, on=keys, validate="one_to_one")
    summary["prevalence"] = summary.prevalence.fillna(0)
    return summary


# ---------- Scénarios : définition, prédéfinis, export et import ----------
# Un scénario regroupe tous les paramètres structurels et de politiques publiques :
#   {"nom", "origine", "offre": paramètres des modulations, "rapprochement": paramètres offre/demande}
# Il ne contient aucun résultat : les projections sont recalculées à son application.
MATCH_BLOCK = "rapprochement"
MATCH_BLOCK_LABEL = "Rapprochement offre / demande"
MATCH_CHOICES = (("demographic_scenario", "Scénario démographique Insee"),
                 ("need_scenario", "Scénario de besoin"),
                 ("need_indicator", "Besoin de référence"))
# (clé, libellé, minimum, maximum, pas, aide) ; bornes identiques à validate_match_parameters.
MATCH_CONTROLS = (
    ("actes_liberal", "Consultations annuelles par psychiatre libéral ou mixte", 0.0, 20000.0, 50.0, "Valeur initiale approximative CNAM 2024 : 2 515 consultations et téléconsultations par praticien. Portée en 2026 par convention."),
    ("ratio_salarie", "Activité salariée / activité libérale", 0.0, 2.0, 0.05, "Hypothèse initiale de 0,65, conversion en actes équivalents ; ajustable, ce n'est pas une mesure actuelle de toute l'activité hospitalière."),
    ("croissance_activite", "Évolution annuelle de l'activité par psychiatre (%)", -5.0, 5.0, 0.1, "0 % maintient la productivité de référence. Une projection d'activité fournie remplace ces trois paramètres."),
    ("actes_par_besoin", "Actes annuels par personne ciblée", 0.0, 100.0, 1.0, "Hypothèse initiale : 8 actes équivalents par an. Ordre de grandeur descriptif, pas une recommandation clinique par trouble."),
    ("couverture", "Couverture cible du besoin par un psychiatre (%)", 0.0, 100.0, 5.0, "Part des personnes avec besoin potentiel ciblée par un suivi psychiatrique. 25 % est une hypothèse de scénario, pas un taux observé."),
    ("part_offre", "Part des actes mobilisable pour ce besoin (%)", 0.0, 100.0, 5.0, "50 % est une hypothèse de scénario pour réserver une part de l'activité aux troubles et âges hors référentiel. Ce paramètre évite d'affecter implicitement toute l'offre à un seul besoin."),
    ("marge_activite", "Variation d'activité pour la fourchette d'effectif cible (± %) ", 0.0, 50.0, 1.0,
     "Hypothèse de sensibilité : ±10 % par défaut autour de l'activité mobilisable d'un psychiatre. Une activité plus élevée donne la borne basse d'effectif ; une activité plus faible donne la borne haute. Ce n'est pas un intervalle de confiance. Mettre 0 pour une cible unique."),
)
# Regroupement des modulations sur la page Scénarios.
SCENARIO_GROUPS = (("Paramètres structurels", ("perimetre", "dynamique", "internat", "mobilite")),
                   ("Politiques publiques", ("formation", "immigration")),
                   ("Politiques de demande", ("politiques",)),
                   ("Offre / demande", (MATCH_BLOCK,)))
SCENARIO_COLUMNS = ["scenario", "bloc", "parametre", "libelle", "valeur"]
MODE_PREDEFINED, MODE_CREATE, MODE_IMPORT = "Scénarios prédéfinis", "Créer un scénario", "Importer un scénario"
TRUE_WORDS, FALSE_WORDS = {"oui", "vrai", "true", "1"}, {"non", "faux", "false", "0"}


def matching_choices(demand):
    """Options des listes du rapprochement, lues dans la demande (None si indisponible)."""
    if demand is None or demand.empty:
        return None, {}
    required = {"indicator", "label", "demographic_scenario", "need_scenario"}
    if not required <= set(demand):
        raise ValueError("Colonnes de scénarios ou de référentiel absentes de la demande.")
    indicators = sorted(demand.indicator.unique())
    labels = demand.drop_duplicates("indicator").set_index("indicator").label.to_dict()
    return {"demographic_scenario": sorted(demand.demographic_scenario.unique()),
            "need_scenario": sorted(demand.need_scenario.unique()), "need_indicator": indicators}, labels


def matching_defaults(choices, labels):
    values = dict(MATCH_DEFAULTS)
    if choices is None:
        values.update(dict.fromkeys(key for key, _ in MATCH_CHOICES))
        return values
    demographic, need, indicators = (choices[key] for key, _ in MATCH_CHOICES)
    values["demographic_scenario"] = "Central" if "Central" in demographic else demographic[0]
    values["need_scenario"] = "stable" if "stable" in need else need[0]
    values["need_indicator"] = next((i for i in indicators if "total" in normalize(labels[i])), indicators[0])
    return values


def make_scenario(name, origin, offre, rapprochement, description="", politiques=None):
    return {"nom": name, "origine": origin, "description": description,
            "offre": deepcopy(offre), "rapprochement": dict(rapprochement), "politiques": deepcopy(politiques or [])}


def predefined_scenarios(registry, context, match_values):
    """Scénarios illustratifs : seuls les paramètres cités diffèrent du scénario Base."""
    base = default_parameters(registry)
    first = context["anchor"] + 1
    places = {f"region_{code}": 20 for code in TRAINING_REGIONS if code in context["available"]}

    def variant(name, description, **changes):
        offre = deepcopy(base)
        for block, values in changes.items():
            offre[block].update(values)
        return make_scenario(name, "Prédéfini", offre, match_values, description)

    scenarios = (
        variant("Base", "Réglages par défaut du modèle : dynamiques départementales et correction de "
                "l'internat actives, aucune hausse de places ni d'arrivées étrangères au-delà de la tendance."),
        variant("Hausse des places de formation",
                f"+20 % de places en deuxième année dans chaque région de formation calibrée, à partir de la "
                f"rentrée {first} ; premières installations en {first + STUDY_YEARS}.",
                formation={**places, "rentree": first}),
        variant("Attractivité internationale",
                f"+30 % d'arrivées de psychiatres diplômés à l'étranger chaque année dès {first}, sans délai.",
                immigration={"hausse": 30, "debut": first}),
        variant("Sans correction de l'internat",
                "Aucune dégradation future du remplissage de l'internat : toutes les cohortes conservent "
                f"le taux observé en {context['anchor']}.", internat={"active": False}),
    )
    result = {scenario["nom"]: scenario for scenario in scenarios}
    prevention = make_scenario("Prévention nationale — illustration", "Prédéfini", base, match_values,
        "Hypothèses : −10 % de dépression chez les 18–25 ans et −10 % de troubles probables chez les 6–11 ans dès 2027. Effets illustratifs, non estimés.",
        politiques=policy_templates(context["anchor"])[1:])
    result[prevention["nom"]] = prevention
    return result


def complete_matching(scenario, defaults):
    """Compléter les listes offre/demande d'un scénario créé sans données de demande."""
    result = deepcopy(scenario)
    for key, value in MATCH_DEFAULTS.items():
        result["rapprochement"].setdefault(key, value)
    for key, _ in MATCH_CHOICES:
        if result["rapprochement"].get(key) is None:
            result["rapprochement"][key] = defaults[key]
    return result


def check_scenario(scenario, registry, choices):
    validate_parameters(scenario["offre"], registry)
    validate_policies(scenario.get("politiques", []))
    match = scenario["rapprochement"]
    validate_match_parameters(match)
    if scenario["offre"]["perimetre"]["exclure_pedopsychiatres"] and match.get("need_indicator") == "trouble_mental_probable_6_11":
        raise ValueError("Le besoin enfant ne peut pas être sélectionné lorsque les pédopsychiatres sont exclus ; choisir un besoin adulte.")
    if choices is not None:
        for key, label in MATCH_CHOICES:
            if match.get(key) not in choices[key]:
                raise ValueError(f"Valeur inconnue pour « {label} » : {match.get(key)!r}.")


def scenario_rows(scenario, registry):
    """Une ligne par paramètre, dans l'ordre de la page Scénarios."""
    rows = []
    for modulation in registry:
        for control in modulation.controls:
            rows.append({"bloc": modulation.key, "parametre": control.key,
                         "libelle": f"{modulation.label} — {control.label}",
                         "valeur": scenario["offre"][modulation.key][control.key]})
    match = scenario["rapprochement"]
    for key, label in MATCH_CHOICES:
        rows.append({"bloc": MATCH_BLOCK, "parametre": key, "libelle": f"{MATCH_BLOCK_LABEL} — {label}", "valeur": match.get(key)})
    for key, label, *_ in MATCH_CONTROLS:
        rows.append({"bloc": MATCH_BLOCK, "parametre": key, "libelle": f"{MATCH_BLOCK_LABEL} — {label}", "valeur": match[key]})
    rows.append({"bloc": "demande", "parametre": "politiques_json", "libelle": "Politiques nationales de demande",
                 "valeur": json.dumps(scenario.get("politiques", []), ensure_ascii=False, sort_keys=True)})
    return rows


def export_value(value):
    if isinstance(value, bool):
        return "oui" if value else "non"
    if isinstance(value, float):
        return f"{value:.12g}"
    return "" if value is None else str(value)


def display_value(value, labels=None):
    if isinstance(value, bool):
        return "oui" if value else "non"
    if isinstance(value, (int, float)):
        return f"{value:g}".replace(".", ",")
    return (labels or {}).get(value, "—" if value is None else str(value))


def scenario_table(scenario, registry):
    table = pd.DataFrame(scenario_rows(scenario, registry))
    table["valeur"] = table.valeur.map(export_value)
    table.insert(0, "scenario", scenario["nom"])
    return table[SCENARIO_COLUMNS]


def scenario_csv(scenario, registry):
    """Point-virgule et BOM UTF-8 : le fichier s'ouvre directement dans Excel en français."""
    return scenario_table(scenario, registry).to_csv(sep=";", index=False).encode("utf-8-sig")


def scenario_file_name(scenario):
    stem = re.sub(r"[^a-z0-9]+", "_", unicodedata.normalize("NFKD", scenario["nom"].lower())
                  .encode("ascii", "ignore").decode()).strip("_") or "scenario"
    return f"scenario_{stem}.csv"


def scenario_differences(scenario, reference, registry, labels=None):
    """Libellés et valeurs des paramètres qui diffèrent du scénario de référence."""
    changed, regions = [], []
    for row, base in zip(scenario_rows(scenario, registry), scenario_rows(reference, registry)):
        if row["valeur"] == base["valeur"]:
            continue
        if row["bloc"] == "demande":
            changed.append(f"Politiques nationales : {sum(p['active'] for p in scenario.get('politiques', []))} règle(s) active(s)")
            continue
        if row["bloc"] == "formation" and row["parametre"].startswith("region_"):
            regions.append(row)
            continue
        changed.append(f"{row['libelle'].rsplit(' — ', 1)[1]} : {display_value(row['valeur'], labels)}")
    # Les curseurs régionaux identiques sont résumés en une seule ligne.
    values = {row["valeur"] for row in regions}
    if len(regions) > 3 and len(values) == 1:
        changed.insert(0, f"Hausse des places de formation : +{values.pop()} % dans {len(regions)} régions")
    else:
        changed[:0] = [f"Places — {row['libelle'].rsplit(' — ', 1)[1]} : {display_value(row['valeur'])}" for row in regions]
    return changed


def parse_number(text, integer):
    cleaned = re.sub(r"[\s  ]", "", str(text)).replace(",", ".")
    value = float(cleaned)
    if not np.isfinite(value):
        raise ValueError
    if integer:
        if value != int(value):
            raise ValueError
        return int(value)
    return value


def convert_control(control, text):
    if control.kind == "checkbox":
        word = normalize(text)
        if word in TRUE_WORDS or word in FALSE_WORDS:
            return word in TRUE_WORDS
        raise ValueError
    if control.kind == "select":
        if text not in control.options:
            raise ValueError
        return text
    integer = isinstance(control.default, int) and not isinstance(control.default, bool)
    return parse_number(text, integer)


def parse_scenario_file(data, registry, reference, choices):
    """Lire un CSV exporté (séparateur , ou ; et décimales à virgule acceptés)."""
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1")
    try:
        table = pd.read_csv(StringIO(text), sep=None, engine="python", dtype=str, keep_default_na=False)
    except (pd.errors.ParserError, csv.Error) as exc:
        raise ValueError(f"Fichier illisible : {exc}") from exc
    table.columns = [normalize(column) for column in table.columns]
    missing = {"bloc", "parametre", "valeur"} - set(table.columns)
    if missing:
        raise ValueError("Colonnes absentes : " + ", ".join(sorted(missing)) + ". Utiliser le format exporté par l'application.")
    table = table.apply(lambda column: column.str.strip())
    table = table[(table.bloc != "") | (table.parametre != "")]
    if table.empty:
        raise ValueError("Le fichier ne contient aucun paramètre.")
    duplicates = table[table.duplicated(["bloc", "parametre"], keep=False)]
    if not duplicates.empty:
        raise ValueError("Paramètres répétés : " + ", ".join(sorted(set(duplicates.bloc + "/" + duplicates.parametre))) + ".")
    controls = {(m.key, c.key): c for m in registry for c in m.controls}
    choice_keys = {key for key, _ in MATCH_CHOICES}
    number_keys = {key: (label, low, high) for key, label, low, high, *_ in MATCH_CONTROLS}
    known = set(controls) | {(MATCH_BLOCK, key) for key in choice_keys | set(number_keys)} | {("demande", "politiques_json")}
    obsolete = (table.bloc == "mobilite") & table.parametre.isin(["methode", "retraite"])
    migrated = bool(obsolete.any())
    table = table[~obsolete].copy()
    unknown = [f"{b}/{p}" for b, p in zip(table.bloc, table.parametre) if (b, p) not in known]
    if unknown:
        raise ValueError("Paramètres inconnus : " + ", ".join(unknown[:8]) + ".")
    names = [value for value in table.get("scenario", pd.Series(dtype=str)) if value]
    scenario = make_scenario(names[0] if names else "Scénario importé", "Importé",
                             reference["offre"], reference["rapprochement"])
    for row in table.itertuples(index=False):
        try:
            if (row.bloc, row.parametre) == ("demande", "politiques_json"):
                scenario["politiques"] = json.loads(row.valeur)
                validate_policies(scenario["politiques"])
            elif (row.bloc, row.parametre) in controls:
                scenario["offre"][row.bloc][row.parametre] = convert_control(controls[(row.bloc, row.parametre)], row.valeur)
            elif row.parametre in choice_keys:
                scenario["rapprochement"][row.parametre] = row.valeur
            else:
                scenario["rapprochement"][row.parametre] = parse_number(row.valeur, False)
        except (ValueError, TypeError):
            raise ValueError(f"Valeur invalide pour {row.bloc}/{row.parametre} : {row.valeur!r}.") from None
    absent = sorted(known - set(zip(table.bloc, table.parametre)))
    warnings = ["Ancien scénario : méthode delta + retraites supprimée ; correspondance diplôme → exercice fixe utilisée."] if migrated else []
    if absent:
        warnings.append(f"{len(absent)} paramètre(s) absent(s) du fichier : valeurs du scénario Base conservées ("
                        + ", ".join(f"{b}/{p}" for b, p in absent[:6]) + ("…" if len(absent) > 6 else "") + ").")
    check_scenario(scenario, registry, choices)
    return scenario, warnings


def scenario_highlights(scenario, context, labels):
    """Résumé lisible du scénario, pour la barre latérale et l'en-tête des projections."""
    offre, anchor = scenario["offre"], context["anchor"]
    lines = [("Dynamiques départementales : actives, estimées sur "
              f"{offre['dynamique']['debut']}–{anchor}") if offre["dynamique"]["active"] else "Dynamiques départementales : désactivées",
             (f"Correction de l'internat : active, tendance depuis {offre['internat']['debut']}")
             if offre["internat"]["active"] else "Correction de l'internat : désactivée"]
    places = [offre["formation"][f"region_{code}"] for code in TRAINING_REGIONS if offre["formation"].get(f"region_{code}")]
    if places:
        size = f"+{places[0]} %" if len(set(places)) == 1 else f"de +{min(places)} à +{max(places)} %"
        first = offre["formation"]["rentree"] + offre["formation"]["delai"]
        lines.append(f"Places de formation : {size} dans {len(places)} région(s), rentrée "
                     f"{offre['formation']['rentree']}, installations dès {first}")
    else:
        lines.append("Places de formation : aucune hausse")
    hausse = offre["immigration"]["hausse"]
    lines.append(f"Diplômés à l'étranger : +{hausse} % dès {offre['immigration']['debut']}"
                 if hausse else "Diplômés à l'étranger : tendance seule")
    lines.append("Périmètre : " + ("adultes, pédopsychiatres retirés" if offre["perimetre"]["exclure_pedopsychiatres"] else "adultes et enfants couverts"))
    lines.append("Internat : " + offre["internat"]["methode"])
    lines.append(f"Politiques de demande actives : {sum(p['active'] for p in scenario.get('politiques', []))}")
    match = scenario["rapprochement"]
    if match.get("need_indicator"):
        lines.append(f"Offre / demande : {labels.get(match['need_indicator'], match['need_indicator'])}, Insee "
                     f"{match['demographic_scenario']}, besoin {match['need_scenario']}, couverture "
                     f"{display_value(match['couverture'])} %")
    return lines


# ---------- Éditeur de scénario : widgets alimentés par un brouillon de session ----------
# Streamlit efface l'état d'un widget absent de la page affichée. Le brouillon
# (st.session_state.scenario_draft) conserve donc les valeurs entre les pages.
def widget_key(block, parameter):
    return f"edit__{block}__{parameter}"


def write_widgets(scenario, registry):
    """Recopier un scénario dans les widgets de l'éditeur (appelé avant leur affichage)."""
    st.session_state.scenario_draft = make_scenario(scenario["nom"], "Personnalisé",
                                                    scenario["offre"], scenario["rapprochement"], politiques=scenario.get("politiques", []))
    for key in [k for k in st.session_state if str(k).startswith("policy__")]:
        del st.session_state[key]
    st.session_state[widget_key("meta", "nom")] = scenario["nom"]
    for modulation in registry:
        for control in modulation.controls:
            st.session_state[widget_key(modulation.key, control.key)] = scenario["offre"][modulation.key][control.key]
    for key, _ in MATCH_CHOICES:
        st.session_state[widget_key(MATCH_BLOCK, key)] = scenario["rapprochement"].get(key)
    for key, *_ in MATCH_CONTROLS:
        st.session_state[widget_key(MATCH_BLOCK, key)] = float(scenario["rapprochement"][key])


def scenario_from_widgets(registry):
    draft = st.session_state.scenario_draft
    value = lambda block, key, fallback: st.session_state.get(widget_key(block, key), fallback)
    offre = {m.key: {c.key: value(m.key, c.key, draft["offre"][m.key][c.key]) for c in m.controls} for m in registry}
    match = {key: value(MATCH_BLOCK, key, draft["rapprochement"].get(key))
             for key in [k for k, _ in MATCH_CHOICES] + [k for k, *_ in MATCH_CONTROLS]}
    name = (value("meta", "nom", draft["nom"]) or "").strip() or "Scénario personnalisé"
    return make_scenario(name, "Personnalisé", offre, match, politiques=policies_from_widgets())


def seed_widget(key, value):
    if key not in st.session_state:
        st.session_state[key] = value


def render_control(control, block):
    key = widget_key(block, control.key)
    seed_widget(key, st.session_state.scenario_draft["offre"][block][control.key])
    common = {"label": control.label, "key": key, "disabled": control.disabled, "help": control.help or None}
    if control.kind == "slider":
        st.slider(**common, min_value=control.minimum, max_value=control.maximum, step=control.step)
    elif control.kind == "number":
        st.number_input(**common, min_value=control.minimum, max_value=control.maximum, step=control.step)
    elif control.kind == "checkbox":
        st.checkbox(**common)
    elif control.kind == "select":
        st.selectbox(**common, options=control.options)
    else:
        raise ValueError(f"Contrôle inconnu : {control.kind}.")


def set_all_regions(registry):
    value = st.session_state["edit_all_regions"]
    formation = next(m for m in registry if m.key == "formation")
    for control in formation.controls:
        if control.key.startswith("region_") and not control.disabled:
            st.session_state[widget_key("formation", control.key)] = value


def render_modulation(modulation, registry):
    with st.container(border=True):
        st.markdown(f"**{modulation.label}**")
        if modulation.description:
            st.caption(modulation.description)
        if modulation.key != "formation":
            columns = st.columns(2)
            for index, control in enumerate(modulation.controls):
                with columns[index % 2]:
                    render_control(control, modulation.key)
            return
        general = [c for c in modulation.controls if not c.key.startswith("region_")]
        regional = [c for c in modulation.controls if c.key.startswith("region_")]
        columns = st.columns(2)
        for column, control in zip(columns, general):
            with column:
                render_control(control, modulation.key)
        st.markdown("Hausse annuelle permanente des places par région de formation (%)")
        left, right = st.columns([3, 2], vertical_alignment="bottom")
        left.slider("Appliquer la même hausse à toutes les régions calibrées (%)", 0, 100, 0, 1, key="edit_all_regions")
        right.button("Appliquer à toutes les régions", on_click=set_all_regions, args=(registry,), width="stretch")
        columns = st.columns(3)
        for index, control in enumerate(regional):
            with columns[index % 3]:
                render_control(control, modulation.key)


def policies_from_widgets():
    policies = deepcopy(st.session_state.scenario_draft.get("politiques", []))
    for p in policies:
        for field in POLICY_COLUMNS:
            if field != "id":
                p[field] = st.session_state.get(f"policy__{p['id']}__{field}", p[field])
    return policies


def add_policy(template, registry):
    from uuid import uuid4
    draft = scenario_from_widgets(registry)
    policy = deepcopy(template)
    policy["id"] = uuid4().hex[:16]
    used = {p["nom"] for p in draft["politiques"]}
    if policy["nom"] in used:
        policy["nom"] += f" ({len(draft['politiques']) + 1})"
    draft["politiques"].append(policy)
    st.session_state.scenario_draft = draft


def remove_policy(identifier, registry):
    draft = scenario_from_widgets(registry)
    draft["politiques"] = [p for p in draft["politiques"] if p["id"] != identifier]
    st.session_state.scenario_draft = draft


def replace_policies(policies, registry):
    draft = scenario_from_widgets(registry)
    draft["politiques"] = deepcopy(policies)
    st.session_state.scenario_draft = draft
    for key in [k for k in st.session_state if str(k).startswith("policy__")]:
        del st.session_state[key]


def render_policy_editor(context):
    registry = modulation_registry(context)
    st.markdown("**Politiques nationales agissant sur la prévalence**")
    st.caption("Une règle décrit un trouble, une tranche d'âge et une baisse relative de prévalence identique dans tous les départements. Pour plusieurs tranches, créer plusieurs règles portant le même nom, sans recouvrement. Les exemples sont des hypothèses illustratives, pas des effets évalués.")
    templates = policy_templates(context["anchor"])
    custom = {**templates[1], "id": "custom", "nom": "Ma politique nationale", "indicateur": "tous", "age_min": 18, "age_max": 79}
    options = {"Politique personnalisée": custom, **{p["nom"]: p for p in templates}}
    col, button = st.columns([4, 1], vertical_alignment="bottom")
    choice = col.selectbox("Modèle de politique", list(options), key="policy_template")
    button.button("Ajouter", key="add_policy", on_click=add_policy, args=(options[choice], registry),
                  disabled=len(st.session_state.scenario_draft.get("politiques", [])) >= 100)
    for p in st.session_state.scenario_draft.get("politiques", []):
        with st.expander(p["nom"], expanded=True):
            for field in POLICY_COLUMNS:
                if field != "id":
                    seed_widget(f"policy__{p['id']}__{field}", p[field])
            key = lambda field: f"policy__{p['id']}__{field}"
            name, active, delete = st.columns([4, 1, 1], vertical_alignment="bottom")
            name.text_input("Nom de la politique", key=key("nom"), max_chars=160)
            active.checkbox("Active", key=key("active"))
            delete.button("Supprimer", key=f"remove_{p['id']}", on_click=remove_policy, args=(p["id"], registry))
            target, lo, hi, reduction = st.columns([3, 1, 1, 2])
            target.selectbox("Trouble ciblé", list(POLICY_INDICATORS), format_func=POLICY_INDICATORS.get, key=key("indicateur"))
            lo.number_input("Âge minimum", 0, 120, step=1, key=key("age_min"))
            hi.number_input("Âge maximum inclus", 0, 120, step=1, key=key("age_max"))
            reduction.number_input("Baisse de prévalence (%)", 0.0, 100.0, step=1.0, key=key("baisse_pct"))
            start, stop, ramp = st.columns(3)
            start.number_input("Première année d'effet", 2026, 2070, step=1, key=key("debut"))
            stop.number_input("Dernière année d'effet", 2026, 2070, step=1, key=key("fin"))
            ramp.number_input("Montée en charge (années)", 1, 45, step=1, key=key("montee"),
                              help="1 : effet complet dès la première année. 3 : un tiers, deux tiers, puis effet complet. Après la dernière année, l'effet s'arrête.")
    policies = policies_from_widgets()
    excluded = st.session_state.get(widget_key("perimetre", "exclure_pedopsychiatres"), st.session_state.scenario_draft["offre"]["perimetre"]["exclure_pedopsychiatres"])
    try:
        validate_policies(policies)
        for message in policy_warnings(policies, excluded):
            st.warning(message)
        invalid = False
    except ValueError as exc:
        st.error(str(exc))
        invalid = True
    export, model = st.columns(2)
    export.download_button("Exporter les politiques (.csv)", policies_csv(policies), "politiques_nationales.csv", "text/csv",
                           key="export_policies", disabled=invalid)
    model.download_button("Télécharger les trois exemples", policies_csv(templates), "exemples_politiques.csv", "text/csv", key="policy_examples")
    upload = st.file_uploader("Importer des politiques (.csv)", type="csv", key="policies_upload")
    if upload is not None:
        try:
            imported = parse_policies_csv(upload.getvalue())
            st.caption(f"{len(imported)} règle(s) valide(s). L'import remplace les politiques du brouillon ; cliquer ensuite sur Appliquer ce scénario.")
            st.button("Charger ces politiques dans le brouillon", key="import_policies", on_click=replace_policies, args=(imported, registry))
        except (ValueError, TypeError, pd.errors.ParserError, UnicodeError) as exc:
            st.error(f"Import impossible : {exc}")
    st.caption("Les politiques sont également incluses dans l'export CSV du scénario complet. Sans données pour un trouble à un âge donné, aucun effet n'est inventé. Les effets de politiques différentes se multiplient sur les âges communs.")


def render_matching_editor(choices, labels, has_activity):
    if choices is None:
        st.info("Les données de demande sont indisponibles : seuls les paramètres d'activité sont enregistrés. "
                "Le rapprochement sera calculé lorsque demand_projection/ et dataDemand/ seront présents.")
    draft = st.session_state.scenario_draft["rapprochement"]
    with st.container(border=True):
        st.markdown(f"**{MATCH_BLOCK_LABEL}**")
        st.caption("Un seul indicateur de besoin est utilisé : le total et ses composantes ne sont jamais additionnés. "
                   "Ces hypothèses sont distinctes du référentiel de prévalence.")
        if choices is not None:
            columns = st.columns(3)
            for column, (key, label) in zip(columns, MATCH_CHOICES):
                widget = widget_key(MATCH_BLOCK, key)
                if st.session_state.get(widget, draft.get(key)) not in choices[key]:
                    st.session_state.pop(widget, None)
                    draft[key] = matching_defaults(choices, labels)[key]
                seed_widget(widget, draft.get(key))
                column.selectbox(label, choices[key], key=widget,
                                 format_func=(lambda value: labels.get(value, value)) if key == "need_indicator" else str)
        columns = st.columns(3)
        for index, (key, label, low, high, step, help_text) in enumerate(MATCH_CONTROLS):
            widget = widget_key(MATCH_BLOCK, key)
            seed_widget(widget, float(draft[key]))
            columns[index % 3].number_input(label, min_value=low, max_value=high, step=step, key=widget, help=help_text)
        if has_activity:
            st.caption("Projection d'activité fournie : les trois paramètres de productivité sont remplacés par ses valeurs.")


@st.cache_data(persist=None)
def load_geojson():
    with urllib.request.urlopen(GEOJSON_URL, timeout=20) as handle:
        gj = json.load(handle)
    gj["features"] = [f for f in gj["features"] if f["properties"]["code"] in DEPARTMENTS]
    missing = DEPARTMENTS - {f["properties"]["code"] for f in gj["features"]}
    if missing:
        raise ValueError("Contours départementaux manquants : " + ", ".join(sorted(missing)))
    # Plotly (d3) attend l'anneau extérieur dans le sens horaire, contrairement à
    # la norme GeoJSON. Sans ce retournement, la projection Lambert remplit le globe.
    def area(ring):
        return sum(x0 * y1 - x1 * y0 for (x0, y0), (x1, y1) in zip(ring, ring[1:] + ring[:1]))

    def rewind(polygon):
        return [ring if (area(ring) < 0) == (i == 0) else ring[::-1] for i, ring in enumerate(polygon)]

    for feature in gj["features"]:
        geometry = feature["geometry"]
        if geometry["type"] == "Polygon":
            geometry["coordinates"] = rewind(geometry["coordinates"])
        else:
            geometry["coordinates"] = [rewind(p) for p in geometry["coordinates"]]
    return gj


def dynamics_diagnostic(panel, contributions, context, scenario):
    """Écarts territoriaux, à périmètre identique entre effectif et référence."""
    if not panel.index.equals(contributions.index):
        raise ValueError("Diagnostic et projection désalignés.")
    diagnostic = panel[panel.scenario == scenario].copy()
    pieces = contributions.loc[diagnostic.index]
    diagnostic["effectif_reference"] = pieces.reference
    diagnostic["effet_dynamique"] = pieces.dynamique
    diagnostic["effet_internat"] = pieces.internat
    diagnostic["tendance_etranger"] = pieces.etranger_tendance
    diagnostic["ajout_immigration"] = pieces.immigration
    diagnostic["ajout_formation"] = pieces.formation
    diagnostic["avant_ajouts"] = pieces.reference + pieces.dynamique + pieces.internat
    diagnostic["hors_formation"] = diagnostic.avant_ajouts + pieces.etranger_tendance + pieces.immigration
    # Les diplômés à l'étranger sont déjà inclus des deux côtés dans l'observé.
    # En projection, leur croissance doit aussi entrer dans le total national
    # comparé au département. Même règle pour les ajouts de politique publique :
    # une hausse répartie selon les parts à l'ancrage ne crée aucun écart territorial.
    diagnostic["reference_nationale_hors_formation"] = count_reference(
        diagnostic, context["baseline"], "hors_formation")
    diagnostic["reference_nationale"] = count_reference(diagnostic, context["baseline"], "effectif")
    diagnostic["ecart_hors_formation"] = diagnostic.hors_formation - diagnostic.reference_nationale_hors_formation
    diagnostic["ecart_complet"] = diagnostic.effectif - diagnostic.reference_nationale
    return diagnostic


def dynamics_figure(diagnostic, context, departments, full=False, start_year=None, show_fit=False):
    """Observé puis projeté, comparés au total national du même périmètre."""
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
    scope_label = "tous effets inclus" if full else "hors ouvertures de places"
    figure.add_hline(y=0, line_color="#64748b", line_width=1,
                    annotation_text=f"Part nationale de {anchor} × total national du scénario ({scope_label})")
    figure.add_vline(x=anchor, line_color="#94a3b8", line_dash="dot")
    figure.update_layout(height=570, xaxis_title="Année", yaxis_title="Écart à la référence nationale (nombre de psychiatres)",
                         margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified")
    return figure


def calibration_table(reference, context, scenario, parameters, flows, routing, cohorts):
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    row = flows[(flows.scenario == scenario) & (flows.annee == first)].iloc[0]
    scope = scope_factors([first], context, parameters["perimetre"]["exclure_pedopsychiatres"])[0]
    national = float(row.total_reference * row.coefficient_stock) * scope
    origin_flows = cohorts[(cohorts.scenario == scenario) & (cohorts.annee == first)].set_index("origine")
    retention = routing[routing.annee == first].groupby("origine").part.sum()
    records = []
    for code, label in TRAINING_REGIONS.items():
        available = code in context["available"]
        flow = float(origin_flows.loc[code, "flux_corrige"]) if available else 0.0
        additions = 0.01 * flow * retention.get(code, 0) * scope
        records.append({"Région de formation": label, "Calibration": "Disponible" if available else "Indisponible",
                        "Hausse validée (%)": parameters["formation"][f"region_{code}"],
                        "Diplômés estimés France (internes / 5)": float(origin_flows.loc[code, "flux_reference"]) if available else np.nan,
                        "Après correction et périmètre": flow * scope if available else np.nan,
                        "Exercice en France (%)": 100 * retention.get(code, 0) if available else np.nan,
                        f"Psychiatres France en plus en {first} pour +1 %": additions if available else np.nan,
                        f"X (%) du total France en {first} pour +1 %": 100 * additions / national if available and national > 0 else np.nan})
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
    figure.add_trace(go.Scatter(x=future.annee, y=100 * future.taux_projete, mode="lines", name=calibration["methode"], line_dash="dash"))
    figure.add_hline(y=100 * calibration["taux_reference"], line_dash="dot", annotation_text="Taux de référence constant")
    figure.update_layout(height=420, xaxis_title="Année d'entrée en internat", yaxis_title="Postes non pourvus (%)",
                         yaxis_range=[0, 100], hovermode="x unified")
    return figure


def insee_total_population(population, years, scenario):
    """Total tous âges, sans modifier les modules de demande ni sommer leurs indicateurs."""
    from demand_projection.run import parse_age_band

    keys = ["ZONE", "SEXE", "TRAGE"]
    columns = [f"POP_{year}" for year in years]
    if population.empty or population[keys].isna().any().any():
        raise ValueError("Population Insee vide ou clés âge/sexe/département manquantes.")
    if population.duplicated(keys).any():
        raise ValueError("Population Insee dupliquée par département, sexe et tranche d'âge.")
    if set(population.SEXE.unique()) != {1, 2}:
        raise ValueError("Le total Insee nécessite les deux sexes, sans ligne de total ajoutée.")
    bands = sorted(population.TRAGE.unique(), key=lambda label: parse_age_band(label).lower)
    ranges = [parse_age_band(label) for label in bands]
    if ranges[0].lower != 0 or not str(bands[-1]).endswith("+]") or any(
        left.upper != right.lower for left, right in zip(ranges, ranges[1:])
    ):
        raise ValueError("Les tranches Insee doivent couvrir tous les âges sans trou ni recouvrement.")
    if (not population.groupby(["ZONE", "SEXE"]).TRAGE.nunique().eq(len(bands)).all()
            or not population.groupby("ZONE").SEXE.nunique().eq(2).all()):
        raise ValueError("Tranches d'âge ou sexe manquants pour un département Insee.")
    values = population[columns].apply(pd.to_numeric, errors="raise")
    if not np.isfinite(values.to_numpy()).all() or (values < 0).any().any():
        raise ValueError("Population Insee manquante, négative ou non finie.")
    grouped = values.groupby(population.ZONE).sum().reset_index()
    totals = grouped.melt(id_vars="ZONE", var_name="_year", value_name="population_total")
    totals["department"] = totals.ZONE.astype(str)
    totals["year"] = totals._year.str.extract(r"(\d{4})").astype(int)
    totals["demographic_scenario"] = scenario
    return totals[["department", "year", "demographic_scenario", "population_total"]]


def join_demand_population(demand, population):
    """Joindre après les agrégations de besoin : chaque total est utilisé une seule fois."""
    keys = ["department", "year", "demographic_scenario"]
    if population.duplicated(keys).any():
        raise ValueError("Population totale dupliquée par département, année et scénario Insee.")
    # Les anciens modules ne transmettent pas population_total au cumul adulte.
    # Reconstituer cette colonne pour tous les indicateurs depuis la même source.
    result = demand.drop(columns=["population_total"], errors="ignore").merge(
        population[keys + ["population_total"]], on=keys, how="left", validate="many_to_one")
    invalid = ~np.isfinite(result.population_total) | (result.population_total <= 0)
    if invalid.any():
        examples = result.loc[invalid, keys].drop_duplicates().head(4).to_dict("records")
        raise ValueError(f"Population totale Insee absente ou invalide pour la demande : {examples}.")
    return result


def available_population_files():
    from demand_projection.run import SCENARIO_FILES
    available = {scenario: ROOT / "dataDemand" / filename for scenario, filename in SCENARIO_FILES.items()
                 if (ROOT / "dataDemand" / filename).is_file()}
    if not available:
        raise FileNotFoundError("Aucun classeur Insee déclaré dans SCENARIO_FILES n'est présent dans dataDemand/.")
    return available


@st.cache_data(show_spinner=False, max_entries=8)
def compute_demand_summary(signature, policies_json="[]", exclude_children=False):
    """Recalcul depuis les sources, pour les seuls scénarios démographiques présents."""
    del signature
    from demand_projection.run import (read_population_workbook, validate_parameters as validate_prevalence,
        read_need_scenarios, read_territorial_modifiers)
    from demand_projection.outcomes import prepare_outcomes
    data_dir = ROOT / "dataDemand"
    policies = json.loads(policies_json)
    validate_policies(policies)
    parameters = validate_prevalence(pd.read_csv(data_dir / "prevalence_parameters.csv"))
    scenarios = read_need_scenarios(data_dir / "need_scenarios.csv")
    modifiers = read_territorial_modifiers(data_dir / "territorial_modifiers.csv")
    summaries, totals = [], []
    for scenario, path in available_population_files().items():
        population, years = read_population_workbook(path)
        totals.append(insee_total_population(population, years, scenario))
        columns = [f"POP_{year}" for year in years]
        population = population.astype({column: float for column in columns})
        summaries.append(project_policy_demand(population, years, scenario, parameters, scenarios, modifiers, policies, exclude_children))
    summary = pd.concat(summaries, ignore_index=True)
    demand = prepare_outcomes(summary)
    # Le cumul adulte garde la même hypothèse de non-recouvrement avant/après politique.
    keys = ["department", "year", "demographic_scenario", "need_scenario", "indicator"]
    adult = summary[summary.indicator.isin(["episode_depressif_12_mois", "trouble_anxieux_generalise_12_mois"])]
    totals_before = adult.groupby(keys[:-1], as_index=False).persons_without_policy.sum()
    totals_before["indicator"] = "besoin_total_sans_recouvrement"
    before = pd.concat([summary[keys + ["persons_without_policy"]], totals_before], ignore_index=True)
    demand = demand.drop(columns=["persons_without_policy"], errors="ignore").merge(before, on=keys, validate="one_to_one")
    demand["persons_avoided"] = (demand.persons_without_policy - demand.persons_with_need).clip(lower=0)
    return join_demand_population(demand, pd.concat(totals, ignore_index=True))


def demand_outcomes(policies=None, exclude_children=False):
    policies = [] if policies is None else policies
    validate_policies(policies)
    data_dir = ROOT / "dataDemand"
    inputs = list(available_population_files().values())
    inputs += [data_dir / "prevalence_parameters.csv", data_dir / "need_scenarios.csv",
               ROOT / "demand_projection" / "run.py", ROOT / "demand_projection" / "outcomes.py", Path(__file__)]
    modifiers = data_dir / "territorial_modifiers.csv"
    if modifiers.exists():
        inputs.append(modifiers)
    signature = tuple((str(path), path.stat().st_mtime_ns, path.stat().st_size) for path in inputs)
    return compute_demand_summary(signature, json.dumps(policies, sort_keys=True), exclude_children)


def render_demand_tab(demand=None, error=None, policies=None, exclude_children=False):
    """Display demand outcomes with filters independent of the supply scenario."""
    st.write("Projection démographique-épidémiologique indépendante de la projection des effectifs de psychiatres.")
    if error:
        st.error(f"Demande de soins indisponible : {error}")
        return
    if demand is None:
        try:
            with st.spinner("Chargement des projections de demande…"):
                demand = demand_outcomes()
        except (OSError, ValueError, KeyError, ImportError) as exc:
            st.error(f"Demande de soins indisponible : {exc}")
            return
    st.caption("Périmètre : adultes uniquement." if exclude_children else "Périmètre : adultes 18–79 ans et enfants 6–11 ans, selon l'indicateur.")
    if any(p["active"] for p in policies or []):
        st.caption("Les politiques actives du scénario sont appliquées à chaque département. Les effets simultanés se multiplient sur les âges communs : −10 % puis −15 % donnent −23,5 %. La réduction est un niveau, pas une nouvelle baisse cumulée chaque année.")
    demographic_scenarios = sorted(demand["demographic_scenario"].unique())
    need_scenarios = sorted(demand["need_scenario"].unique())
    indicators = sorted(demand["indicator"].unique())
    controls_a, controls_b, controls_c = st.columns(3)
    with controls_a:
        demand_demographic_scenario = st.selectbox(
            "Scénario démographique Insee", demographic_scenarios,
            index=demographic_scenarios.index("Central") if "Central" in demographic_scenarios else 0,
            key="demand_demographic_scenario",
        )
    with controls_b:
        demand_need_scenario = st.selectbox(
            "Scénario de besoin", need_scenarios,
            index=need_scenarios.index("stable") if "stable" in need_scenarios else 0,
            key="demand_need_scenario",
        )
    with controls_c:
        demand_indicator = st.selectbox(
            "Indicateur de besoin", indicators,
            format_func=lambda value: demand.loc[demand["indicator"] == value, "label"].iloc[0],
            key="demand_indicator",
        )
    filtered_demand = demand.loc[
        (demand["demographic_scenario"] == demand_demographic_scenario)
        & (demand["need_scenario"] == demand_need_scenario)
        & (demand["indicator"] == demand_indicator)
    ].copy()
    if filtered_demand.empty:
        st.info("Aucun résultat pour cette combinaison de scénarios et d’indicateur.")
        return
    departments_demand = sorted(filtered_demand["department"].unique())
    default_demand_departments = ["Paris"] if "Paris" in departments_demand else departments_demand[:1]
    selected_demand_departments = st.multiselect(
        "Départements à comparer", departments_demand, default=default_demand_departments,
        key="demand_departments",
    )
    selected_demand = filtered_demand[filtered_demand["department"].isin(selected_demand_departments)].copy()
    national_demand = filtered_demand.groupby("year", as_index=False)["persons_with_need"].sum()
    first_demand_year = int(national_demand["year"].min())
    last_demand_year = int(national_demand["year"].max())
    first_demand_value = national_demand.loc[national_demand["year"] == first_demand_year, "persons_with_need"].iloc[0]
    last_demand_value = national_demand.loc[national_demand["year"] == last_demand_year, "persons_with_need"].iloc[0]
    demand_change = last_demand_value / first_demand_value - 1 if first_demand_value else 0.0
    metric_a, metric_b, metric_c = st.columns(3)
    metric_a.metric(f"Personnes avec besoin projeté — France, {first_demand_year}", f"{first_demand_value:,.0f}".replace(",", " "))
    metric_b.metric(f"Personnes avec besoin projeté — France, {last_demand_year}", f"{last_demand_value:,.0f}".replace(",", " "))
    metric_c.metric(f"Évolution {first_demand_year}–{last_demand_year}", f"{demand_change:+.1%}" if first_demand_value else "Non calculable")
    if "persons_without_policy" in filtered_demand:
        last = filtered_demand[filtered_demand.year == last_demand_year]
        st.caption(f"En {last_demand_year}, avant politiques : {last.persons_without_policy.sum():,.0f} personnes ; réduction simulée : {last.persons_avoided.sum():,.0f} personnes.")
        st.download_button("Exporter la demande annuelle de ce scénario", filtered_demand.to_csv(index=False).encode("utf-8-sig"),
                           "demande_annuelle.csv", "text/csv", key="export_demand_results")
    if selected_demand.empty:
        st.info("Sélectionnez au moins un département pour afficher les résultats.")
    else:
        figure = px.line(
            selected_demand, x="year", y="persons_with_need", color="department",
            labels={"year": "Année", "persons_with_need": "Personnes avec besoin projeté", "department": "Département"},
            title=filtered_demand["label"].iloc[0],
        )
        figure.update_layout(legend_title_text="Département", hovermode="x unified")
        st.plotly_chart(figure, width="stretch")
        display_years = [2026, 2030, 2040, 2050, 2060, 2070]
        display = selected_demand.loc[selected_demand["year"].isin(display_years), ["department", "year", "population_at_risk", "prevalence", "persons_with_need"]].copy()
        display["population_at_risk"] = display["population_at_risk"].round().astype(int)
        display["prevalence"] = (100 * display["prevalence"]).round(2)
        display["persons_with_need"] = display["persons_with_need"].round().astype(int)
        st.dataframe(display.rename(columns={"department": "Département", "year": "Année", "population_at_risk": "Population concernée", "prevalence": "Prévalence (%)", "persons_with_need": "Personnes avec besoin"}), hide_index=True, width="stretch")
    with st.expander("Périmètre et limites"):
        st.markdown(
            "Les résultats estiment des personnes avec besoin potentiel : population projetée × prévalence. "
            "Ils ne mesurent ni des recours observés, ni le nombre de psychiatres requis. "
            "L'indicateur « Besoin total » additionne dépression et anxiété selon l'hypothèse explicite qu'aucune personne n'est concernée par les deux ; il peut donc surestimer le besoin si cette hypothèse est fausse. "
            "Le modèle n'utilise aucune donnée SNDS, PMSI, ATIH ou d'activité de soins. "
            "Les 6–11 ans sont couverts par Enabee 2022 : 13 % présentent au moins un trouble probable parmi les troubles étudiés. "
            "Ce taux, mesuré chez les enfants scolarisés en élémentaire en France hexagonale, est extrapolé à la population projetée des 6–11 ans. "
            "Les moins de 6 ans, les 12–17 ans, les 80 ans et plus, les addictions et les troubles sévères adultes restent hors périmètre. "
            "Le cumul dépression + anxiété concerne uniquement les adultes de 18–79 ans ; il n’inclut pas les enfants."
        )


def match_department_codes(frame, baseline):
    """Rapprocher codes ou noms sans confondre la Corse avec un code numérique."""
    column = next((c for c in ("dep", "department_code", "code_departement", "department", "departement") if c in frame), None)
    if column is None:
        raise ValueError("Colonne département absente (dep ou department).")
    names = {normalize(row.nom): row.dep for row in baseline.itertuples()}
    names.update({normalize(name): code for code, name in {
        "971": "Guadeloupe", "972": "Martinique", "973": "Guyane", "974": "La Réunion", "976": "Mayotte"}.items()})
    names[normalize("Réunion")] = "974"  # Libellé Omphale, reconnu avant le filtre France.
    # Le classeur d'effectifs nomme le département 62 « Nord-Pas-de-Calais ».
    names[normalize("Pas-de-Calais")] = "62"
    def code(value):
        raw = str(value).strip().upper()
        if re.fullmatch(r"\d+\.0", raw):
            raw = raw[:-2]
        match = re.match(r"^(0?2[AB]|\d{1,3})(?:\s*[-–].*)?$", raw)
        if match:
            token = match.group(1)
            return token[1:] if len(token) == 3 and token.startswith("0") else token.zfill(2)
        return names.get(normalize(raw))
    result = frame[column].map(code)
    if result.isna().any():
        unknown = ", ".join(map(str, frame.loc[result.isna(), column].drop_duplicates().head(5)))
        raise ValueError(f"Départements non reconnus pour le rapprochement : {unknown}.")
    return result


def match_years(frame):
    column = "annee" if "annee" in frame else "year"
    if column not in frame:
        raise ValueError("Colonne année absente (annee ou year).")
    values = pd.to_numeric(frame[column], errors="coerce")
    if not np.isfinite(values).all() or not (values % 1 == 0).all():
        raise ValueError("Année non entière ou manquante dans les données de rapprochement.")
    return values.astype(int)


def prepare_total_population(frame, baseline, demographic_scenario, explicit_total=False):
    """Jamais de substitution de population_at_risk à la population tous âges."""
    candidates = ("population_total", "total_population", "population_totale")
    if explicit_total:
        candidates += ("population",)
    column = next((c for c in candidates if c in frame), None)
    if column is None:
        raise ValueError("Population totale tous âges absente. Ajouter population_total aux résultats de demande, ou fournir population_totale.csv. La population à risque ne peut pas servir de dénominateur par habitant.")
    selected = frame.copy()
    if "demographic_scenario" not in selected:
        raise ValueError("Le fichier de population totale doit préciser demographic_scenario pour éviter un mélange de scénarios Insee.")
    selected = selected[selected.demographic_scenario == demographic_scenario].copy()
    if selected.empty:
        raise ValueError(f"Population totale indisponible pour le scénario Insee {demographic_scenario}.")
    selected["dep"] = match_department_codes(selected, baseline)
    selected["annee"] = match_years(selected)
    selected["population_total"] = pd.to_numeric(selected[column], errors="coerce")
    selected = selected[selected.dep.isin(baseline.dep)]
    if not np.isfinite(selected.population_total).all() or (selected.population_total <= 0).any():
        raise ValueError("La population totale doit être finie et strictement positive.")
    if (selected.groupby(["dep", "annee"]).population_total.nunique() > 1).any():
        raise ValueError("Plusieurs populations totales différentes pour le même département et la même année.")
    return selected[["dep", "annee", "population_total"]].drop_duplicates()


def validate_match_parameters(parameters):
    limits = {"actes_liberal": (0, 20000), "ratio_salarie": (0, 2), "croissance_activite": (-5, 5),
              "actes_par_besoin": (0, 100), "couverture": (0, 100), "part_offre": (0, 100),
              "marge_activite": (0, 50)}
    for name, (lo, hi) in limits.items():
        value = parameters.get(name, MATCH_DEFAULTS[name])
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not np.isfinite(value) or not lo <= value <= hi:
            raise ValueError(f"Paramètre offre/demande invalide : {name}.")


def compute_supply_demand(panel, demand, population, parameters, baseline, anchor, activity=None):
    """Solde annuel signé ; jointures strictes sur département, année et scénarios."""
    validate_match_parameters(parameters)
    required = {"demographic_scenario", "need_scenario", "indicator", "persons_with_need"}
    if not required <= set(demand.columns):
        raise ValueError("Résultats de demande incomplets : " + ", ".join(sorted(required - set(demand.columns))))
    need = demand[(demand.demographic_scenario == parameters["demographic_scenario"])
                  & (demand.need_scenario == parameters["need_scenario"])
                  & (demand.indicator == parameters["need_indicator"])].copy()
    if need.empty:
        raise ValueError("Aucune demande pour le référentiel et les scénarios validés.")
    need["dep"], need["annee"] = match_department_codes(need, baseline), match_years(need)
    need = need[need.dep.isin(baseline.dep)]
    if need.empty or need.duplicated(["dep", "annee"]).any():
        raise ValueError("Demande départementale absente ou dupliquée ; sélectionner un seul indicateur de besoin.")
    need["persons_with_need"] = pd.to_numeric(need.persons_with_need, errors="coerce")
    if not np.isfinite(need.persons_with_need).all() or (need.persons_with_need < 0).any():
        raise ValueError("Nombre de personnes avec besoin négatif ou manquant.")
    populations = prepare_total_population(population if population is not None else demand,
        baseline, parameters["demographic_scenario"], explicit_total=population is not None)
    start = max(anchor, int(need.annee.min()))
    end = min(int(panel.annee.max()), int(need.annee.max()))
    result = panel[panel.annee.between(start, end)].copy()
    if result.empty:
        raise ValueError("Aucune année commune entre l'offre et la demande à partir de l'ancrage.")
    if result.duplicated(["dep", "annee", "scenario"]).any():
        raise ValueError("Offre dupliquée avant le rapprochement.")
    result = result.merge(need[["dep", "annee", "persons_with_need"]], on=["dep", "annee"], how="left", validate="many_to_one")
    result = result.merge(populations, on=["dep", "annee"], how="left", validate="many_to_one")
    missing = result[["persons_with_need", "population_total"]].isna().any(axis=1)
    if missing.any():
        examples = result.loc[missing, ["dep", "annee"]].drop_duplicates().head(4).to_dict("records")
        raise ValueError(f"Demande ou population manquante sur la période commune : {examples}. Aucun zéro ni report d'année n'est substitué.")
    if (result.persons_with_need > result.population_total + 1e-8).any():
        raise ValueError("Le besoin sélectionné dépasse la population totale ; vérifier le champ ou les doubles comptes.")
    if activity is None:
        result["actes_liberal"] = parameters["actes_liberal"] * (1 + parameters["croissance_activite"] / 100.0) ** (result.annee - anchor)
        result["actes_salarie"] = result.actes_liberal * parameters["ratio_salarie"]
    else:
        activity = activity.copy()
        activity["annee"] = match_years(activity)
        keys = ["annee"]
        if any(c in activity for c in ("dep", "department_code", "code_departement", "department", "departement")):
            activity["dep"] = match_department_codes(activity, baseline)
            keys.append("dep")
        if "scenario" in activity:
            keys.append("scenario")
        if {"actes_liberal", "actes_salarie"} <= set(activity.columns):
            pass
        elif "actes_par_psychiatre" in activity:
            activity["actes_liberal"] = activity["actes_salarie"] = activity.actes_par_psychiatre
        else:
            raise ValueError("La projection d'activité doit contenir actes_liberal et actes_salarie, ou actes_par_psychiatre.")
        if activity.duplicated(keys).any():
            raise ValueError("Projection d'actes par psychiatre dupliquée.")
        for column in ("actes_liberal", "actes_salarie"):
            activity[column] = pd.to_numeric(activity[column], errors="coerce")
            if not np.isfinite(activity[column]).all() or (activity[column] < 0).any():
                raise ValueError("Actes par psychiatre manquants, négatifs ou non finis.")
        result = result.merge(activity[keys + ["actes_liberal", "actes_salarie"]], on=keys, how="left", validate="many_to_one")
        if result[["actes_liberal", "actes_salarie"]].isna().any().any():
            raise ValueError("La projection d'activité ne couvre pas toutes les années, départements et scénarios du rapprochement.")
    result["actes_totaux"] = result.liberaux * result.actes_liberal + result.salaries * result.actes_salarie
    result["actes_par_psychiatre"] = np.divide(result.actes_totaux, result.effectif, out=np.zeros(len(result)), where=result.effectif > 0)
    result["actes_disponibles"] = result.actes_totaux * parameters["part_offre"] / 100.0
    result["personnes_ciblees"] = result.persons_with_need * parameters["couverture"] / 100.0
    result["actes_necessaires"] = result.personnes_ciblees * parameters["actes_par_besoin"]
    result["solde_actes"] = result.actes_disponibles - result.actes_necessaires
    result["solde_actes_habitant"] = result.solde_actes / result.population_total
    numeric = list(ACTIVITY_INDICATORS) + ["actes_totaux", "solde_actes"]
    if not np.isfinite(result[numeric]).all().all():
        raise ValueError("Résultat offre/demande non fini.")
    return add_psychiatrist_targets(result, parameters, baseline)


def add_psychiatrist_targets(result, parameters, baseline):
    """Effectif requis pour le besoin sélectionné, à activité et mix de statuts fixés."""
    result = result.copy()
    count = result.effectif.to_numpy(dtype=float)
    # Même mix salariés/libéraux que l'offre simulée. En cas d'effectif nul,
    # garder le mix observé du département pour valoriser les recrutements possibles.
    base = baseline.set_index("dep")
    total = float(baseline.effectif.sum())
    national_share = float(baseline.salaries.sum() / total) if total > 0 else np.nan
    base_shares = base.part_salaries.where(base.effectif > 0, national_share)
    fallback = result.dep.map(base_shares).to_numpy(dtype=float)
    share = np.divide(result.salaries.to_numpy(dtype=float), count, out=fallback.copy(), where=count > 0)
    if not np.isfinite(share).all() or ((share < 0) | (share > 1)).any():
        raise ValueError("Mix salariés/libéraux indisponible pour estimer l'effectif cible.")
    result["mix_cible_salaries"] = share
    result["mix_cible_repli"] = count <= 0
    activity = (1 - share) * result.actes_liberal.to_numpy() + share * result.actes_salarie.to_numpy()
    result["capacite_mobilisable_par_psychiatre"] = activity * parameters["part_offre"] / 100.0
    capacity = result.capacite_mobilisable_par_psychiatre.to_numpy()
    required = result.actes_necessaires.to_numpy(dtype=float)
    target = np.divide(required, capacity, out=np.full(len(result), np.nan), where=capacity > 0)
    target[required == 0] = 0.0
    margin = parameters.get("marge_activite", MATCH_DEFAULTS["marge_activite"]) / 100.0
    result["psychiatres_cibles"] = target
    result["psychiatres_cibles_bas"] = target / (1 + margin)
    result["psychiatres_cibles_haut"] = target / (1 - margin)
    result["solde_psychiatres"] = count - target
    result["psychiatres_manquants"] = (target - count).clip(min=0)
    result["psychiatres_surplus"] = (count - target).clip(min=0)
    result["manque_minimal_fourchette"] = (result.psychiatres_cibles_bas - count).clip(lower=0)
    result["surplus_fourchette"] = (count - result.psychiatres_cibles_haut).clip(lower=0)
    result["couverture_actes_pct"] = np.divide(100 * result.actes_disponibles.to_numpy(), required,
        out=np.full(len(result), np.nan), where=required > 0)
    result["habitants_par_psychiatre"] = np.divide(result.population_total.to_numpy(), count,
        out=np.full(len(result), np.nan), where=count > 0)
    result["personnes_ciblees_par_psychiatre"] = np.divide(result.personnes_ciblees.to_numpy(), count,
        out=np.full(len(result), np.nan), where=count > 0)
    lower, upper = result.psychiatres_cibles_bas, result.psychiatres_cibles_haut
    tolerance = 1e-9 * np.maximum(1, np.nan_to_num(target))
    result["statut_couverture"] = np.select(
        [required == 0, ~np.isfinite(target), count < lower - tolerance, count > upper + tolerance],
        ["Aucun besoin ciblé", "Capacité mobilisable nulle", "Sous la fourchette cible", "Au-dessus de la fourchette cible"],
        default="Dans la fourchette cible")
    return result


def format_quantity(value, decimals=0, suffix="", signed=False):
    if value is None or not np.isfinite(value):
        return "Non calculable"
    spec = f"{'+' if signed else ''},.{decimals}f"
    return format(float(value), spec).replace(",", "\u202f").replace(".", ",") + suffix


def target_card_values(row, parameters):
    """Libellés identiques sur la carte et dans la fiche départementale."""
    finite = np.isfinite(row.psychiatres_cibles)
    if not finite:
        gap_label, gap = "Manque / surplus", "Non calculable"
    elif row.manque_minimal_fourchette > 1e-7:
        gap_label, gap = "Manque minimal", format_quantity(row.manque_minimal_fourchette, 1)
    elif row.surplus_fourchette > 1e-7:
        gap_label, gap = "Au-delà de la fourchette", "+" + format_quantity(row.surplus_fourchette, 1)
    else:
        gap_label, gap = "Écart à la cible centrale", format_quantity(row.solde_psychiatres, 1, signed=True)
    margin = parameters.get("marge_activite", MATCH_DEFAULTS["marge_activite"])
    band = (format_quantity(row.psychiatres_cibles_bas, 1) + " – " + format_quantity(row.psychiatres_cibles_haut, 1)
            if finite and margin > 0 else format_quantity(row.psychiatres_cibles, 1))
    status = str(row.statut_couverture)
    tone = ("red" if status == "Sous la fourchette cible" else "green" if status == "Au-dessus de la fourchette cible"
            else "amber" if status == "Dans la fourchette cible" else "neutral")
    return {"nom": str(row.nom), "annee": str(int(row.annee)), "population": format_quantity(row.population_total),
            "effectif": format_quantity(row.effectif, 1), "habitants_par_psy": format_quantity(row.habitants_par_psychiatre),
            "personnes_ciblees": format_quantity(row.personnes_ciblees), "cible": format_quantity(row.psychiatres_cibles, 1),
            "fourchette": band, "ecart_label": gap_label, "ecart": gap,
            "couverture": format_quantity(row.couverture_actes_pct, 1, " %"), "statut": status, "ton": tone,
            "capacite": format_quantity(row.capacite_mobilisable_par_psychiatre, 1),
            "besoin": format_quantity(row.actes_necessaires), "marge": format_quantity(margin),
            "mix_repli": bool(row.mix_cible_repli)}


def territory_hover(row, parameters, label):
    v = target_card_values(row, parameters)
    color = {"red": "#b93855", "green": "#397d52", "amber": "#936a25", "neutral": "#65707c"}[v["ton"]]
    lines = [f"<b>{escape(v['nom'])}</b> · {v['annee']}",
             f"Population : <b>{v['population']}</b>", f"Psychiatres : <b>{v['effectif']}</b>",
             f"Habitants / psychiatre : <b>{v['habitants_par_psy']}</b>",
             f"Personnes ciblées : <b>{v['personnes_ciblees']}</b>",
             f"Effectif cible estimé : <b>{v['cible']}</b>", f"Fourchette : <b>{v['fourchette']}</b>",
             f"{v['ecart_label']} : <b>{v['ecart']}</b>", f"Couverture des actes : <b>{v['couverture']}</b>",
             f"<span style='color:{color}'><b>{escape(v['statut'])}</b></span>",
             f"<span style='font-size:11px'>{escape(label)}</span>"]
    return "<br>".join(lines)


def territory_card_html(row, parameters, label):
    v = target_card_values(row, parameters)
    def item(label, value, emphasis=False):
        css = " psy-card-focus" if emphasis else ""
        return f'<div class="psy-card-row{css}"><span>{escape(label)}</span><strong>{escape(value)}</strong></div>'
    items = "".join([item("Population", v["population"]), item("Psychiatres", v["effectif"]),
        item("Habitants / psychiatre", v["habitants_par_psy"]), item("Personnes ciblées", v["personnes_ciblees"]),
        item("Effectif cible estimé", v["cible"], True), item("Fourchette d'effectif cible", v["fourchette"]),
        item(v["ecart_label"], v["ecart"]), item("Couverture des actes nécessaires", v["couverture"])])
    note = "Mix de statuts de la dernière année observée utilisé pour estimer les recrutements." if v["mix_repli"] else "Mix salariés/libéraux du département conservé pour estimer la cible."
    return f'''<div class="psy-territory-card">
<div class="psy-card-heading"><strong>{escape(v['nom'])}</strong><span>{v['annee']}</span></div>
<div class="psy-card-subtitle">{escape(label)}</div>{items}
<div class="psy-card-status psy-card-{v['ton']}"><small>COUVERTURE THÉORIQUE</small><strong>{escape(v['statut'])}</strong></div>
<div class="psy-card-footnote">Fourchette : activité par psychiatre ±{v['marge']} %. Hypothèse de sensibilité, pas un intervalle de confiance. {escape(note)}</div>
</div>'''


def target_evolution_figure(detail, department):
    rows = detail[detail.dep == department].sort_values("annee")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=rows.annee, y=rows.psychiatres_cibles_bas, mode="lines", line=dict(width=0),
                            showlegend=False, hoverinfo="skip", connectgaps=False))
    fig.add_trace(go.Scatter(x=rows.annee, y=rows.psychiatres_cibles_haut, mode="lines", line=dict(width=0),
                            fill="tonexty", fillcolor="rgba(90,145,120,.16)", name="Fourchette cible", hoverinfo="skip", connectgaps=False))
    fig.add_trace(go.Scatter(x=rows.annee, y=rows.psychiatres_cibles, mode="lines", name="Effectif cible estimé",
                            line=dict(color="#3e7954", width=2, dash="dash"), connectgaps=False))
    fig.add_trace(go.Scatter(x=rows.annee, y=rows.effectif, mode="lines", name="Psychiatres disponibles",
                            line=dict(color="#354f70", width=3)))
    fig.update_layout(height=430, margin=dict(l=10, r=10, t=15, b=10), hovermode="x unified",
                      xaxis_title="Année", yaxis_title="Nombre de psychiatres", legend=dict(orientation="h", y=-.2))
    return fig


def match_projections_frame(detail):
    result = detail[["dep", "annee", "scenario"] + list(MATCH_INDICATORS)].melt(
        id_vars=["dep", "annee", "scenario"], var_name="indicateur", value_name="valeur")
    result.indicateur = result.indicateur.map(MATCH_INDICATORS)
    return result[COLUMNS].sort_values(COLUMNS[:-1]).reset_index(drop=True)


OPTIONAL_FILES = {
    "population": ("Population totale tous âges (CSV, si absente des données de demande)", TOTAL_POPULATION_PATH),
    "activity": ("Projection d’actes par psychiatre (CSV, facultatif)", ACTIVITY_PROJECTION_PATH),
}


def read_optional_match_csv(data, path):
    """Fichier chargé dans la session, sinon fichier local, sinon None."""
    if data is not None:
        return pd.read_csv(BytesIO(data), dtype=str)
    if path.exists():
        return pd.read_csv(path, dtype=str)
    return None


def optional_match_inputs():
    """Le contenu des fichiers chargés est gardé en session, même hors de la page Scénarios."""
    stored = st.session_state.get("optional_files", {})
    return tuple(read_optional_match_csv(stored.get(name, (None, None))[1], path)
                 for name, (_, path) in OPTIONAL_FILES.items())


def store_optional_file(name):
    upload = st.session_state.get(f"upload_{name}")
    files = st.session_state.setdefault("optional_files", {})
    if upload is None:
        files.pop(name, None)
    else:
        files[name] = (upload.name, upload.getvalue())


def forget_optional_file(name):
    st.session_state.setdefault("optional_files", {}).pop(name, None)


def ensure_matching(state, demand, population, parameters, context, activity=None):
    hasher = sha256((state["key"] + json.dumps(parameters, sort_keys=True)).encode())
    for frame in (demand, population, activity):
        if frame is not None:
            hasher.update(json.dumps(list(frame.columns)).encode())
            hasher.update(pd.util.hash_pandas_object(frame, index=True).to_numpy().tobytes())
    key = hasher.hexdigest()
    saved = st.session_state.get("psy_matching", {})
    if saved.get("key") == key:
        return saved
    detail = compute_supply_demand(state["panel"], demand, population, parameters, context["baseline"], context["anchor"], activity)
    result = {"key": key, "parameters": deepcopy(parameters), "detail": detail,
              "projections": match_projections_frame(detail), "external_activity": activity is not None}
    st.session_state.psy_matching = result
    return result


def render_match_tab(matching, error, scenario, year, names):
    st.write("L'effectif cible estime le nombre de psychiatres nécessaire pour produire les actes requis par le besoin sélectionné. Il conserve le mix salariés/libéraux et la part d'activité mobilisable du scénario.")
    st.caption("Un solde négatif indique un manque ; un solde positif indique une marge théorique dans le périmètre choisi. Il ne mesure pas des rendez-vous réellement libres ni un excédent pour l'ensemble de la psychiatrie.")
    if matching is None:
        st.info(error or "Renseigner les données et paramètres de rapprochement sur la page Scénarios.")
        return
    params = matching["parameters"]
    st.caption(f"Paramètres validés : scénario Insee {params['demographic_scenario']} ; besoin {params['need_scenario']} / {params['need_indicator']} ; {params['actes_par_besoin']:g} actes par personne ciblée, couverture {params['couverture']:g} %, offre mobilisable {params['part_offre']:g} %.")
    detail = matching["detail"].query("scenario == @scenario").copy()
    available_years = sorted(detail.annee.unique())
    selected_year = year if year in available_years else available_years[0]
    if selected_year != year:
        st.info(f"Le rapprochement couvre {available_years[0]}–{available_years[-1]}. Les mesures ci-dessous portent sur {selected_year}.")
    selected = detail[detail.annee == selected_year].copy()
    a, b, c = st.columns(3)
    a.metric("Actes disponibles — France", f"{selected.actes_disponibles.sum():,.0f}".replace(",", " "))
    b.metric("Actes nécessaires — France", f"{selected.actes_necessaires.sum():,.0f}".replace(",", " "))
    c.metric("Solde annuel par habitant — France", f"{selected.solde_actes.sum() / selected.population_total.sum():+.3f}")
    staffing_a, staffing_b, staffing_c = st.columns(3)
    staffing_a.metric("Psychiatres disponibles — France", format_quantity(selected.effectif.sum(), 1))
    staffing_b.metric("Effectif cible estimé — France", format_quantity(selected.psychiatres_cibles.sum(skipna=False), 1))
    staffing_c.metric("Manques départementaux cumulés", format_quantity(selected.psychiatres_manquants.sum(skipna=False), 1),
                      help="Somme des manques par rapport à la cible centrale. Un surplus dans un département ne compense pas automatiquement un manque ailleurs.")
    unavailable = int(selected.psychiatres_cibles.isna().sum())
    if unavailable:
        st.info(f"Effectif cible non calculable dans {unavailable} département(s) : besoin positif mais capacité mobilisable par psychiatre nulle. Les totaux de cible incomplets ne sont pas affichés comme des totaux nationaux.")
    st.caption(f"Fourchette : activité par psychiatre ±{params.get('marge_activite', MATCH_DEFAULTS['marge_activite']):g} %. Hypothèse de sensibilité, sans valeur de norme clinique.")
    departments = st.multiselect("Départements du rapprochement", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]], format_func=lambda dep: f"{dep} · {names[dep]}", key="match_departments")
    shown = detail[detail.dep.isin(departments)].copy()
    shown["Département"] = shown.dep.map(names)
    if not shown.empty:
        graph = st.selectbox("Évolution à afficher", [PSY_TARGET, PSY_BALANCE, NEED_COVERAGE, ACTS_BALANCE], key="matching_evolution")
        if graph == PSY_TARGET and len(departments) == 1:
            figure = target_evolution_figure(detail, departments[0])
        elif graph == PSY_TARGET:
            figure = go.Figure()
            for index, department in enumerate(departments):
                series = shown[shown.dep == department].sort_values("annee")
                color = px.colors.qualitative.Dark24[index % len(px.colors.qualitative.Dark24)]
                figure.add_trace(go.Scatter(x=series.annee, y=series.effectif, name=f"{names[department]} · offre", line=dict(color=color)))
                figure.add_trace(go.Scatter(x=series.annee, y=series.psychiatres_cibles, name=f"{names[department]} · cible", line=dict(color=color, dash="dash")))
            figure.update_layout(xaxis_title="Année", yaxis_title="Nombre de psychiatres", hovermode="x unified")
        else:
            column = next(key for key, label in MATCH_INDICATORS.items() if label == graph)
            figure = px.line(shown, x="annee", y=column, color="Département", labels={"annee": "Année", column: graph})
            figure.add_hline(y=100 if graph == NEED_COVERAGE else 0, line_dash="dot", line_color="#475569")
        st.plotly_chart(figure, width="stretch", key="matching_evolution_chart")
    selected.insert(1, "nom", selected.dep.map(names))
    st.dataframe(selected[["dep", "nom", "annee", "population_total", "effectif", "persons_with_need", "personnes_ciblees", "psychiatres_cibles", "psychiatres_cibles_bas", "psychiatres_cibles_haut", "solde_psychiatres", "psychiatres_manquants", "psychiatres_surplus", "couverture_actes_pct", "statut_couverture", "actes_par_psychiatre", "actes_totaux", "actes_disponibles", "actes_necessaires", "solde_actes", "solde_actes_habitant"]].rename(columns={
        "effectif": "Psychiatres disponibles", "psychiatres_cibles": PSY_TARGET, "psychiatres_cibles_bas": "Cible basse", "psychiatres_cibles_haut": "Cible haute",
        "solde_psychiatres": PSY_BALANCE, "psychiatres_manquants": PSY_SHORTFALL, "psychiatres_surplus": PSY_SURPLUS,
        "couverture_actes_pct": NEED_COVERAGE, "statut_couverture": "Situation théorique",
        "annee": "Année", "population_total": "Population totale", "persons_with_need": "Personnes avec besoin potentiel", "personnes_ciblees": "Personnes ciblées", "actes_par_psychiatre": ACTS_PER_PSY,
        "actes_totaux": "Actes totaux avant affectation", "actes_disponibles": AVAILABLE_ACTS, "actes_necessaires": REQUIRED_ACTS, "solde_actes": "Solde annuel d’actes", "solde_actes_habitant": ACTS_BALANCE}), hide_index=True, width="stretch")
    exported = detail.assign(nom=detail.dep.map(names))
    st.download_button("Exporter les effectifs cibles et le rapprochement annuel", exported.to_csv(index=False).encode("utf-8-sig"),
                       "effectifs_cibles_psychiatres.csv", "text/csv", key="export_targets")
    with st.expander("Hypothèses et sources du rapprochement"):
        st.latex(r"N_{\mathrm{cible}} = \frac{\mathrm{Personnes\ avec\ besoin} \times \mathrm{Couverture\ cible} \times \mathrm{Actes/personne}}{\mathrm{Actes/psychiatre} \times \mathrm{Part\ mobilisable}}")
        st.write("La cible centrale est un effectif attendu, éventuellement fractionnaire. La borne basse divise les actes nécessaires par une capacité accrue de la marge choisie ; la borne haute utilise une capacité diminuée de cette marge. Manques et surplus des indicateurs sont calculés par rapport à la cible centrale ; la fiche indique aussi le manque minimal ou le dépassement de la fourchette.")
        st.write("Si aucun psychiatre n'est disponible, la cible reste calculable avec le mix de statuts observé à l'ancrage. Si la capacité mobilisable est nulle alors que le besoin est positif, la cible est signalée comme non calculable. Un besoin ciblé nul donne une cible nulle et aucun taux de couverture n'est calculé.")
        st.write("Productivité par statut × effectifs projetés = actes totaux. La part mobilisable limite l'offre au périmètre choisi. Personnes avec besoin × couverture cible × actes par personne = actes nécessaires. Les actes annuels ne sont pas cumulés entre années.")
        st.write("Les valeurs initiales de 8 actes, 25 % de couverture et 50 % d'offre mobilisable sont des hypothèses ajustables, pas un référentiel clinique validé. Les actes libéraux et salariés sont exprimés en équivalents de consultation ; l'hospitalisation et le travail en équipe ne sont pas directement comparables aux consultations facturées.")
        st.write("La population totale tous âges est requise. La population à risque des 18–79 ans ou des 6–11 ans n'est jamais utilisée comme substitut. Le cumul adulte dépression/anxiété conserve l'hypothèse d'absence de recouvrement du modèle de demande.")
        for label, url in MATCH_SOURCES.items():
            st.markdown(f"[{label}]({url})")


# ---------- Interface : navigation, page Scénarios, page Projections ----------
FAMILY_SUPPLY, FAMILY_DEMAND, FAMILY_MATCH = "Offre", "Demande", "Offre / demande"
DEMAND_MEASURES = {"count": "personnes avec besoin potentiel", "rate": "pour 1 000 habitants"}
SIDEBAR_WIDTH = 220  # Pixels.
# La carte suit la hauteur de la fenêtre : espace libre sous l'en-tête et les onglets,
# multiplié par MAP_SCALE. Avec 1, titre, carte et sélecteurs tiennent sans défilement.
MAP_RESERVED_HEIGHT = 305  # Pixels hors carte, mesurés sur la page Projections.
MAP_MIN_HEIGHT = 340
MAP_SCALE = 1.2  # Carte 20 % plus haute que l'espace libre : les sélecteurs passent un peu plus bas.
PAGE_STYLE = f"""<style>
[data-testid="stMainBlockContainer"] {{padding-top: 2.3rem; padding-bottom: 2rem;}}
[data-testid="stSidebar"] {{width: {SIDEBAR_WIDTH}px !important; min-width: {SIDEBAR_WIDTH}px !important;
                           max-width: {SIDEBAR_WIDTH}px !important;}}
[data-testid="stSidebarContent"] {{padding-top: 0.4rem;}}
[data-testid="stNavSectionHeader"] p {{font-size: 0.78rem; font-weight: 600; letter-spacing: 0.06em;
                                       text-transform: uppercase; opacity: 0.75;}}
[data-testid="stSidebarNavLink"] {{border: 1px solid rgba(49, 51, 63, 0.2); border-radius: 0.5rem;
                                   padding: 0.55rem 0.75rem; margin-bottom: 0.4rem;}}
[data-testid="stSidebarNavLink"] p {{font-size: 1.02rem; font-weight: 600;}}
[data-testid="stSidebarNavLink"][aria-current="page"] {{background: #c53948; border-color: #c53948;}}
[data-testid="stSidebarNavLink"][aria-current="page"] * {{color: #ffffff !important;}}
.st-key-map_box {{height: calc({MAP_SCALE} * (100vh - {MAP_RESERVED_HEIGHT}px)) !important;
                  min-height: {round(MAP_SCALE * MAP_MIN_HEIGHT)}px;
                  flex: 0 0 auto !important;}}

.psy-territory-card {{max-width: 520px; background: #fff; color: #25282c; border: 1px solid #e0e3e6;
 border-radius: 18px; padding: 24px 26px; box-shadow: 0 10px 30px rgba(20,34,46,.09); margin: 8px 0 18px;}}
.psy-card-heading {{display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:8px;}}
.psy-card-heading strong {{font-size:27px; font-weight:750;}}.psy-card-heading span {{font-size:18px; color:#777;}}
.psy-card-subtitle {{font-size:12px; color:#73808a; margin-bottom:14px; line-height:1.4;}}
.psy-card-row {{display:flex; align-items:baseline; justify-content:space-between; gap:20px; border-bottom:1px solid #eceef0; padding:10px 0; font-size:16px;}}
.psy-card-row span {{color:#71777d;}}.psy-card-row strong {{font-size:18px; font-weight:650; text-align:right; white-space:nowrap;}}
.psy-card-focus {{border-top:1px dashed #cdd4da; margin-top:8px; padding-top:15px;}}
.psy-card-focus strong {{font-size:23px; color:#315f48;}}.psy-card-status {{margin-top:18px; padding:12px 16px; border-left:5px solid; border-radius:0 8px 8px 0;}}
.psy-card-status small {{display:block; font-size:12px; color:#68716e; letter-spacing:.06em;}}
.psy-card-status strong {{display:block; margin-top:3px; font-size:19px; line-height:1.35;}}
.psy-card-red {{background:#fbf0f3; border-color:#b93855; color:#a82d4a;}}
.psy-card-green {{background:#edf5ef; border-color:#50895e; color:#397447;}}
.psy-card-amber {{background:#faf5e9; border-color:#b18a34; color:#876421;}}
.psy-card-neutral {{background:#f2f4f6; border-color:#82909b; color:#5b6770;}}
.psy-card-footnote {{font-size:11px; color:#7d858c; line-height:1.5; margin-top:14px;}}
</style>"""


def keep_selection(key):
    """Un contrôle segmenté ne peut pas être vidé : on rétablit le dernier choix."""
    if st.session_state.get(key) is None:
        st.session_state[key] = st.session_state.get(f"{key}__last")
    st.session_state[f"{key}__last"] = st.session_state[key]


def segmented(label, options, key, **kwargs):
    if st.session_state.get(key) not in options:
        st.session_state[key] = options[0]
    st.session_state.setdefault(f"{key}__last", st.session_state[key])
    return st.segmented_control(label, options, key=key, on_change=keep_selection, args=(key,), **kwargs)


def apply_scenario(scenario, registry, choices, from_widgets=False):
    """Bouton « Appliquer » : vérifier, activer, puis afficher les projections."""
    if from_widgets:
        scenario = scenario_from_widgets(registry)
        st.session_state.scenario_draft = deepcopy(scenario)
    try:
        check_scenario(scenario, registry, choices)
    except (ValueError, KeyError) as exc:
        st.session_state.scenario_error = f"Scénario non appliqué : {exc}"
        return
    st.session_state.scenario_previous = st.session_state.scenario_active
    st.session_state.scenario_active = deepcopy(scenario)
    st.session_state.goto_projections = True


def edit_scenario(scenario, registry, name=None):
    """Charger un scénario dans l'éditeur et ouvrir la création."""
    copy = deepcopy(scenario)
    if name:
        copy["nom"] = name
    write_widgets(copy, registry)
    st.session_state.scenario_mode = MODE_CREATE
    st.session_state.scenario_mode__last = MODE_CREATE


def scenario_short_lines(scenario):
    """Trois repères courts pour la barre latérale ; le détail est sur les pages."""
    offre = scenario["offre"]
    places = [offre["formation"][f"region_{code}"] for code in TRAINING_REGIONS if offre["formation"].get(f"region_{code}")]
    first = offre["formation"]["rentree"] + offre["formation"]["delai"]
    size = (f"+{places[0]} %" if len(set(places)) == 1 else f"+{min(places)} à +{max(places)} %") if places else ""
    hausse = offre["immigration"]["hausse"]
    return [f"Places : {size} ({len(places)} rég.) dès {first}" if places else "Places : inchangées",
            f"Étranger : +{hausse} % dès {offre['immigration']['debut']}" if hausse else "Étranger : tendance seule",
            "Internat : correction " + ("active" if offre["internat"]["active"] else "désactivée")]


def render_sidebar_summary(scenario):
    with st.sidebar:
        st.caption("SCÉNARIO ACTIF")
        st.markdown(f"**{scenario['nom']}**  \n:gray[{scenario['origine']}]")
        st.caption("  \n".join(scenario_short_lines(scenario)))


def render_scenarios_page(app, projections_page):
    if st.session_state.pop("goto_projections", False):
        st.switch_page(projections_page)
    registry, context, choices, labels = app["registry"], app["context"], app["choices"], app["labels"]
    predefined, active = app["predefined"], st.session_state.scenario_active
    st.title("Scénarios")
    st.write("Choisissez les paramètres structurels et les politiques publiques utilisés pour les projections. "
             "Un scénario appliqué est recalculé pour votre session, puis affiché dans Projections.")
    if "scenario_error" in st.session_state:
        st.error(st.session_state.pop("scenario_error"))
    with st.container(border=True):
        left, right = st.columns([4, 1], vertical_alignment="center")
        left.markdown(f"Scénario actif : **{active['nom']}** · :gray[{active['origine']}]")
        left.caption(" · ".join(scenario_highlights(active, context, labels)))
        right.download_button("Exporter (.csv)", scenario_csv(active, registry), scenario_file_name(active),
                              "text/csv", key="export_active", width="stretch")
    mode = segmented("Mode", [MODE_PREDEFINED, MODE_CREATE, MODE_IMPORT], "scenario_mode", label_visibility="collapsed")
    if mode == MODE_PREDEFINED:
        render_predefined(predefined, active, registry, choices, labels)
    elif mode == MODE_CREATE:
        render_creation(app)
    else:
        render_import(app)
    with st.expander("Données facultatives pour le rapprochement offre / demande"):
        st.caption("Ces fichiers s'appliquent à tous les scénarios de la session et ne sont pas inclus dans l'export. "
                   "Une projection d'activité remplace les trois paramètres de productivité.")
        stored = st.session_state.get("optional_files", {})
        for name, (label, path) in OPTIONAL_FILES.items():
            st.file_uploader(label, type="csv", key=f"upload_{name}", on_change=store_optional_file, args=(name,))
            if name in stored and st.session_state.get(f"upload_{name}") is None:
                info, remove = st.columns([4, 1], vertical_alignment="center")
                info.caption(f"Fichier utilisé : {stored[name][0]}")
                remove.button("Retirer", key=f"forget_{name}", on_click=forget_optional_file, args=(name,))
            elif name not in stored and path.exists():
                st.caption(f"Fichier local utilisé : {path.relative_to(ROOT)}")


def render_predefined(predefined, active, registry, choices, labels):
    st.caption("Scénarios illustratifs, construits à partir des leviers du modèle. "
               "« Personnaliser » les copie dans l'éditeur pour les ajuster.")
    base = predefined["Base"]
    columns = st.columns(2)
    for index, scenario in enumerate(predefined.values()):
        with columns[index % 2].container(border=True):
            is_active = active["origine"] == "Prédéfini" and active["nom"] == scenario["nom"]
            st.markdown(f"#### {scenario['nom']}" + ("  :green-badge[Actif]" if is_active else ""))
            st.write(scenario["description"])
            changes = scenario_differences(scenario, base, registry, labels)
            st.caption("  \n".join(changes) if changes else "Valeurs par défaut du modèle.")
            apply, edit, export = st.columns(3)
            apply.button("Appliquer", key=f"apply_{index}", type="primary", width="stretch",
                         on_click=apply_scenario, args=(scenario, registry, choices))
            edit.button("Personnaliser", key=f"edit_{index}", width="stretch",
                        on_click=edit_scenario, args=(scenario, registry, f"{scenario['nom']} (personnalisé)"))
            export.download_button("Exporter", scenario_csv(scenario, registry), scenario_file_name(scenario),
                                   "text/csv", key=f"export_{index}", width="stretch")


def render_creation(app):
    registry, choices, labels, predefined = app["registry"], app["choices"], app["labels"], app["predefined"]
    st.caption("Les valeurs ne sont prises en compte qu'au clic sur « Appliquer ». Le brouillon est conservé "
               "pendant la session, même en passant par la page Projections.")
    starts = {**predefined, "Scénario actif": st.session_state.scenario_active}
    name_col, start_col, load_col = st.columns([2, 2, 1], vertical_alignment="bottom")
    with name_col:
        seed_widget(widget_key("meta", "nom"), st.session_state.scenario_draft["nom"])
        st.text_input("Nom du scénario", key=widget_key("meta", "nom"), max_chars=80)
    start = start_col.selectbox("Partir des valeurs de", list(starts), key="edit_start")
    load_col.button("Charger", width="stretch", on_click=edit_scenario,
                    args=(starts[start], registry, st.session_state.get(widget_key("meta", "nom"))),
                    help="Remplace les valeurs de l'éditeur par celles du scénario choisi, en gardant le nom.")
    by_key = {m.key: m for m in registry}
    for tab, (_, blocks) in zip(st.tabs([label for label, _ in SCENARIO_GROUPS]), SCENARIO_GROUPS):
        with tab:
            for block in blocks:
                if block == "politiques":
                    render_policy_editor(app["context"])
                elif block == MATCH_BLOCK:
                    effective_choices = deepcopy(choices)
                    if effective_choices is not None and st.session_state.get(widget_key("perimetre", "exclure_pedopsychiatres"), False):
                        effective_choices["need_indicator"] = [i for i in effective_choices["need_indicator"] if i != "trouble_mental_probable_6_11"]
                    render_matching_editor(effective_choices, labels, optional_match_inputs()[1] is not None)
                elif by_key[block].controls:
                    render_modulation(by_key[block], registry)
    draft = scenario_from_widgets(registry)
    st.session_state.scenario_draft = draft
    try:
        check_scenario(draft, registry, choices)
        problem = None
    except (ValueError, KeyError) as exc:
        problem = str(exc)
        st.error(problem)
    changes = scenario_differences(draft, predefined["Base"], registry, labels)
    st.caption(("Écarts au scénario Base : " + " · ".join(changes)) if changes else "Identique au scénario Base.")
    apply, export, _ = st.columns([1, 1, 2])
    apply.button("Appliquer ce scénario", type="primary", width="stretch", disabled=problem is not None,
                 on_click=apply_scenario, args=(draft, registry, choices, True))
    export.download_button("Exporter (.csv)", scenario_csv(draft, registry), scenario_file_name(draft),
                           "text/csv", key="export_draft", width="stretch", disabled=problem is not None)


def render_import(app):
    registry, choices, labels, base = app["registry"], app["choices"], app["labels"], app["predefined"]["Base"]
    st.caption("Charger un fichier exporté par l'application. Les paramètres absents prennent la valeur du "
               "scénario Base ; toute valeur hors limites est refusée.")
    upload = st.file_uploader("Fichier de scénario (.csv)", type="csv", key="scenario_upload")
    with st.expander("Format attendu"):
        st.write("Une ligne par paramètre, séparateur point-virgule ou virgule. Colonnes : "
                 "`scenario` (nom), `bloc`, `parametre`, `libelle` (indicatif) et `valeur`. "
                 "Les décimales peuvent utiliser la virgule ; les cases à cocher valent oui ou non.")
        st.dataframe(scenario_table(base, registry).head(8), hide_index=True, width="stretch")
        st.download_button("Télécharger un modèle (scénario Base)", scenario_csv(base, registry),
                           "modele_scenario.csv", "text/csv", key="export_template")
    if upload is None:
        return
    try:
        scenario, warnings = parse_scenario_file(upload.getvalue(), registry, base, choices)
    except (ValueError, KeyError) as exc:
        st.error(f"Import impossible : {exc}")
        return
    for message in warnings:
        st.warning(message)
    with st.container(border=True):
        st.markdown(f"#### {scenario['nom']}")
        st.caption("  \n".join(scenario_highlights(scenario, app["context"], labels)))
        changes = scenario_differences(scenario, base, registry, labels)
        st.caption(("Écarts au scénario Base : " + " · ".join(changes)) if changes else "Identique au scénario Base.")
        apply, edit, _ = st.columns([1, 1, 2])
        apply.button("Appliquer ce scénario", type="primary", width="stretch", key="apply_import",
                     on_click=apply_scenario, args=(scenario, registry, choices))
        edit.button("Modifier avant d'appliquer", width="stretch", key="edit_import",
                    on_click=edit_scenario, args=(scenario, registry))
    with st.expander("Valeurs importées"):
        st.dataframe(scenario_table(scenario, registry), hide_index=True, width="stretch")


def demand_map_frame(demand, baseline, match):
    """Demande du scénario actif, rapprochée des codes départementaux français."""
    selected = demand[(demand.demographic_scenario == match["demographic_scenario"])
                      & (demand.need_scenario == match["need_scenario"])].copy()
    names = pd.DataFrame({"department": selected.department.unique()})
    names["dep"] = match_department_codes(names, baseline)
    selected["dep"] = selected.department.map(names.set_index("department").dep)
    selected["annee"] = match_years(selected)
    return selected[selected.dep.isin(set(baseline.dep))]


def render_map_tab(app, df, demand, demand_error, matching, matching_error, scenario):
    """Carte pleine largeur, sélecteurs dessous ; renvoie la projection DREES et l'année choisies."""
    context, names = app["context"], app["names"]
    anchor, labels = context["anchor"], app["labels"]
    map_area = st.container(key="map_box")
    controls = st.container()
    with controls:
        family_col, indicator_col, drees_col, year_col = st.columns([2.35, 1.75, 1.35, 1.75], vertical_alignment="bottom")
        with family_col:
            family = segmented("Thème", [FAMILY_SUPPLY, FAMILY_DEMAND, FAMILY_MATCH], "map_family")
        demand_rows = None
        if family == FAMILY_SUPPLY:
            options = list(INDICATORS.values())
            indicator = indicator_col.selectbox("Indicateur", options, key="map_indicator_supply")
            rows = df[df.indicateur == indicator]
        elif family == FAMILY_MATCH:
            options = list(MATCH_INDICATORS.values())
            indicator = indicator_col.selectbox("Indicateur", options, index=options.index(PSY_TARGET), key="map_indicator_match")
            rows = matching["projections"] if matching is not None else df.iloc[0:0]
            rows = rows[rows.indicateur == indicator]
        else:
            options = [(code, measure) for code in (app["choices"] or {}).get("need_indicator", []) for measure in DEMAND_MEASURES]
            indicator = indicator_col.selectbox(
                "Indicateur", options, key="map_indicator_demand",
                format_func=lambda item: f"{labels.get(item[0], item[0])} — {DEMAND_MEASURES[item[1]]}") if options else None
            rows = df.iloc[0:0]
            if indicator is not None and demand is not None:
                code, measure = indicator
                demand_rows = demand_map_frame(demand, context["baseline"], scenario["rapprochement"])
                demand_rows = demand_rows[demand_rows.indicator == code].copy()
                demand_rows["valeur"] = (demand_rows.persons_with_need if measure == "count"
                                         else 1000 * demand_rows.persons_with_need / demand_rows.population_total)
                rows = demand_rows
        drees_options = sorted(df.scenario.unique())
        drees = drees_col.selectbox("Projection DREES", drees_options, key="map_drees", disabled=family == FAMILY_DEMAND,
                                    help="Scénario des projections DREES d'effectifs. Sans effet sur la demande.")
        years = rows.annee if not rows.empty else df.annee
        lower, upper = int(years.min()), int(years.max())
        st.session_state.map_year = max(lower, min(upper, st.session_state.get("map_year", anchor)))
        with year_col:
            year = (st.slider("Année", lower, upper, key="map_year") if lower < upper
                    else st.selectbox("Année", [lower], key="map_year"))
    if family == FAMILY_DEMAND:
        unit = DEMAND_MEASURES[indicator[1]] if indicator else ""
        value_label = f"Besoin potentiel ({unit})"
        selected = rows[rows.annee == year] if not rows.empty else rows
        scale_values = rows.valeur if not rows.empty else pd.Series(dtype=float)
    else:
        value_label = indicator + " (%)" if indicator == SALARIED_SHARE else indicator
        selected = rows[(rows.scenario == drees) & (rows.annee == year)]
        scale_values = rows.valeur
    selected = selected[["dep", "valeur"]].copy() if not selected.empty else pd.DataFrame(columns=["dep", "valeur"])
    selected["nom"] = selected.dep.map(names)
    map_matching = matching
    if family == FAMILY_DEMAND and indicator is not None and demand is not None and indicator[0] != scenario["rapprochement"].get("need_indicator"):
        try:
            map_parameters = {**scenario["rapprochement"], "need_indicator": indicator[0]}
            population, activity = optional_match_inputs()
            map_matching = ensure_matching(st.session_state.psy_projection, demand, population, map_parameters, context, activity)
        except (OSError, ValueError, KeyError, TypeError) as exc:
            map_matching = None
            st.info(f"Effectif cible indisponible pour ce besoin : {exc}")
    card_detail = None
    if map_matching is not None:
        card_detail = map_matching["detail"]
        card_detail = card_detail[(card_detail.scenario == drees) & (card_detail.annee == year)].copy()
        if not card_detail.empty:
            card_detail["nom"] = card_detail.dep.map(names)
            label = labels.get(map_matching["parameters"]["need_indicator"], map_matching["parameters"]["need_indicator"])
            card_detail["info_bulle"] = card_detail.apply(lambda row: territory_hover(row, map_matching["parameters"], label), axis=1)
            selected = selected.merge(card_detail[["dep", "info_bulle"]], on="dep", how="left", validate="one_to_one")
    with map_area:
        if family == FAMILY_DEMAND and demand is None:
            st.warning(f"Demande de soins indisponible : {demand_error}" if demand_error else "Demande de soins indisponible.")
        elif family == FAMILY_MATCH and selected.empty:
            st.warning(matching_error or "Indicateur disponible uniquement sur la période commune à l'offre, à la demande et à la population totale.")
        render_choropleth(selected, scale_values, value_label, indicator)
    with st.expander(f"Valeurs par département en {year}"):
        if not selected.empty:
            st.dataframe(selected.sort_values("valeur")[["dep", "nom", "valeur"]].rename(
                columns={"dep": "Code", "nom": "Département", "valeur": value_label}), hide_index=True, width="stretch")
    if family == FAMILY_MATCH:
        st.caption("L'effectif cible est estimé à partir du besoin sélectionné et de la capacité mobilisable par psychiatre. Un solde de psychiatres négatif indique un manque par rapport à la cible centrale ; les manques et surplus sont également disponibles séparément. Les cartes de soldes sont centrées sur zéro.")
        if matching is not None:
            mp = matching["parameters"]
            st.caption(f"Référence du rapprochement : Insee {mp['demographic_scenario']} ; besoin {mp['need_scenario']} / {mp['need_indicator']} ; couverture {mp['couverture']:g} %, {mp['actes_par_besoin']:g} actes par personne ciblée, {mp['part_offre']:g} % de l'offre mobilisable.")
    elif family == FAMILY_DEMAND:
        match = scenario["rapprochement"]
        st.caption(f"Personnes avec besoin potentiel = population projetée × prévalence, selon le scénario Insee {match['demographic_scenario']} "
                   f"et le scénario de besoin {match['need_scenario']} du scénario actif. Le taux rapporte ce nombre à la population totale tous âges. "
                   "Ce n'est ni un recours observé, ni un nombre de psychiatres requis ; voir l'onglet Demande de soins.")
    else:
        st.caption("Salariés = salariés hospitaliers + autres salariés. Libéraux = libéraux exclusifs + exercice mixte. Chaque personne est comptée une seule fois ; la part salariée est en %.")
    if card_detail is not None and not card_detail.empty:
        with st.expander("Explorer un département : effectif disponible et cible", expanded=True):
            departments = sorted(card_detail.dep.unique())
            if st.session_state.get("target_department") not in departments:
                st.session_state.target_department = "75" if "75" in departments else departments[0]
            department = st.selectbox("Département à explorer", departments, key="target_department", format_func=lambda dep: f"{dep} · {names[dep]}")
            row = card_detail[card_detail.dep == department].iloc[0]
            label = labels.get(map_matching["parameters"]["need_indicator"], map_matching["parameters"]["need_indicator"])
            left, right = st.columns([1, 1.55], vertical_alignment="top")
            with left:
                st.markdown(territory_card_html(row, map_matching["parameters"], label), unsafe_allow_html=True)
            with right:
                trend = map_matching["detail"].query("scenario == @drees")
                st.plotly_chart(target_evolution_figure(trend, department), width="stretch", key="target_department_evolution")
                st.caption("L'effectif cible concerne uniquement le besoin sélectionné, avec la couverture cible et la part d'activité mobilisable du scénario. La couverture théorique ne mesure pas les délais ni l'accès réel aux soins.")
    return drees, year


def france_figure(selected, scale_values, value_label, indicator, gj):
    """Métropole et cinq encarts DROM, même échelle de couleur partout."""
    is_share, is_balance = indicator == SALARIED_SHARE, indicator in (ACTS_BALANCE, PSY_BALANCE)
    scale_values = pd.to_numeric(scale_values, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if is_balance:
        bound = max(float(scale_values.abs().quantile(0.98)), 0.001) if not scale_values.empty else 1.0
        vmin, vmax = -bound, bound
    elif indicator == NEED_COVERAGE:
        vmin, vmax = 0.0, max(200.0, float(scale_values.quantile(.98))) if not scale_values.empty else 200.0
    else:
        vmin, vmax = (0.0, 100.0) if is_share else scale_values.quantile([0.02, 0.98]) if not scale_values.empty else (0.0, 1.0)
    if vmin == vmax:
        vmax = vmin + 1
    unit = " %" if is_share or indicator == NEED_COVERAGE else ""
    signed_scale = [[0, "#ad536f"], [.5, "#f6f4f3"], [1, "#5c9d77"]]
    colorscale = signed_scale if is_balance else px.colors.sequential.Blues[1:]
    if indicator == NEED_COVERAGE:
        midpoint = 100 / vmax
        colorscale = [[0, "#ad536f"], [midpoint, "#f6f4f3"], [min(1.0, midpoint + .12), "#81b28f"], [1, "#39754d"]]
    fig = go.Figure()
    panels = [("Métropole", METRO, [0.20, 0.88], [0, 1])]
    for index, (dep, label) in enumerate(DROM.items()):
        top = 1 - index * 0.20
        panels.append((label, {dep}, [0.0, 0.18], [top - 0.17, top - 0.025]))
    layout = {}
    for index, (label, codes, xdomain, ydomain) in enumerate(panels):
        geo = "geo" if index == 0 else f"geo{index + 1}"
        subset = selected[selected.dep.isin(codes)]
        contours = {"type": "FeatureCollection", "features": [f for f in gj["features"] if f["properties"]["code"] in codes]}
        subset = subset.copy()
        simple = pd.Series([
            f"<b>{escape(str(row.nom))}</b> ({row.dep})<br>{escape(value_label)} : {format_quantity(row.valeur, 1)}"
            for row in subset.itertuples()
        ], index=subset.index, dtype="str")
        subset["hover_card"] = subset.info_bulle.fillna(simple) if "info_bulle" in subset else simple
        if "info_bulle" in subset and not subset.empty:
            subset["hover_card"] += subset.valeur.map(lambda value: f"<br><i>{escape(value_label)} : {format_quantity(value, 1)}{unit}</i>")
        valid = subset[np.isfinite(subset.valeur)]
        fig.add_trace(go.Choropleth(geo=geo, geojson=contours, locations=valid.dep, z=valid.valeur,
            featureidkey="properties.code", coloraxis="coloraxis", text=valid.nom,
            customdata=valid[["hover_card"]].to_numpy(), marker_line_color="white", marker_line_width=0.5,
            hovertemplate="%{customdata[0]}<extra></extra>"))
        missing = subset[~np.isfinite(subset.valeur)]
        if not missing.empty:
            fig.add_trace(go.Choropleth(geo=geo, geojson=contours, locations=missing.dep, z=np.zeros(len(missing)),
                featureidkey="properties.code", colorscale=[[0, "#d6dade"], [1, "#d6dade"]], showscale=False,
                customdata=missing[["hover_card"]].to_numpy(), marker_line_color="white", marker_line_width=.5,
                hovertemplate="%{customdata[0]}<extra></extra>"))
        layout[geo] = dict(domain=dict(x=xdomain, y=ydomain), visible=False, bgcolor="rgba(0,0,0,0)",
                           projection=dict(type="mercator"), fitbounds="locations")
        if index == 0:
            layout[geo].update(projection=dict(type="conic conformal", parallels=[44, 49], rotation=dict(lon=3)),
                center=dict(lon=2.3, lat=46.3), lonaxis_range=[-5.3, 9.8], lataxis_range=[41.2, 51.2], fitbounds=False)
        else:
            fig.add_annotation(x=sum(xdomain)/2, y=ydomain[1]+0.02, xref="paper", yref="paper", text=label,
                               showarrow=False, font=dict(size=11), xanchor="center", yanchor="middle")
    fig.update_layout(**layout, margin=dict(l=0, r=0, t=10, b=0), dragmode=False,
        paper_bgcolor="rgba(0,0,0,0)", separators=", ",
        hoverlabel=dict(bgcolor="white", bordercolor="#dfe4e7", font=dict(size=14, color="#28313a", family="Arial, sans-serif"), align="left"),
        coloraxis=dict(cmin=float(vmin), cmax=float(vmax), colorscale=colorscale,
                       colorbar=dict(title=dict(text=value_label, side="right"), thickness=14, len=0.7, outlinewidth=0)))
    return fig


def render_choropleth(selected, scale_values, value_label, indicator):
    try:
        gj = load_geojson()
    except (OSError, ValueError) as exc:
        st.warning(f"Carte indisponible ; les tableaux et les courbes restent accessibles : {exc}")
        return
    if selected.empty:
        return
    fig = france_figure(selected, scale_values, value_label, indicator, gj)
    st.plotly_chart(fig, width="stretch", height="stretch")
    st.caption("France : 101 départements, métropole et cinq DROM. Encarts ultramarins à des échelles géographiques différentes ; couleurs comparables. Les collectivités d'outre-mer hors départements ne figurent pas dans ces données.")


def render_projections_page(app, scenarios_page):
    reference, metadata, context = app["reference"], app["metadata"], app["context"]
    signature, names, labels = app["signature"], app["names"], app["labels"]
    scenario = complete_matching(st.session_state.scenario_active, app["match_defaults"])
    proposed = scenario["offre"]
    # En-tête sur une ligne : titre, scénario actif (détail au clic), lien vers Scénarios.
    title, detail, change = st.columns([3.3, 1.6, 1.25], vertical_alignment="center")
    title.markdown("### Psychiatres : offre, besoins et effectifs cibles")
    try:
        with st.spinner(f"Calcul du scénario « {scenario['nom']} »…"):
            df, metadata, recalculated = ensure_projections(signature, proposed, reference, metadata, context)
    except (ValueError, KeyError, RuntimeError) as exc:
        failed = scenario["nom"]
        previous = st.session_state.get("scenario_previous")
        if previous is not None and previous is not st.session_state.scenario_active:
            st.session_state.scenario_active = previous
        st.error(f"Le scénario « {failed} » n'a pas pu être calculé : {exc}")
        st.page_link(scenarios_page, label="Revenir aux scénarios", icon=":material/tune:")
        st.stop()
    if recalculated:
        st.toast(f"Projections recalculées : {scenario['nom']}.")
    anchor = context["anchor"]
    state = st.session_state.psy_projection
    if proposed["perimetre"]["exclure_pedopsychiatres"]:
        context = deepcopy(context)
        for name in ("history", "baseline"):
            frame = context[name]
            factors = scope_factors(frame.annee, context, True)
            frame[["effectif", "salaries", "liberaux"]] = frame[["effectif", "salaries", "liberaux"]].mul(factors, axis=0)
    mode_text = (f"croissance nationale DREES raccordée aux effectifs RPPS de {anchor} (psychiatrie absente des projections régionales)"
                 if metadata["mode_projection"] == "national" else "projections régionales détaillées de psychiatrie")
    with detail.popover(f"Scénario : {scenario['nom']}", icon=":material/info:", width="stretch",
                        help="Paramètres du scénario actif"):
        st.markdown(f"**{scenario['nom']}** · :gray[{scenario['origine']}]")
        st.markdown("  \n".join(scenario_highlights(scenario, context, labels)))
        st.caption(f"Observé : {df.annee.min()}–{anchor}. Projeté : {anchor + 1}–{df.annee.max()}. Référence : {mode_text}.")
    change.page_link(scenarios_page, label="Changer de scénario", icon=":material/tune:")
    demand, demand_error = app["demand"], app["demand_error"]
    excluded = proposed["perimetre"]["exclure_pedopsychiatres"]
    if excluded or scenario.get("politiques"):
        try:
            with st.spinner("Application des politiques nationales et du périmètre de soins…"):
                demand = demand_outcomes(scenario.get("politiques", []), excluded)
            demand_error = None
        except (OSError, ValueError, KeyError, ImportError) as exc:
            demand, demand_error = None, str(exc)
    local_choices, local_labels = matching_choices(demand)
    app = {**app, "demand": demand, "choices": local_choices, "labels": local_labels}
    if excluded:
        share = context["pediatric_shares"].set_index("annee").loc[anchor]
        st.caption(f"Périmètre adulte : retrait uniforme de {100 * share.part:.2f} % de l'offre projetée ({share.pedopsychiatres:,.0f} / {share.total:,.0f} en {anchor}). Demande des moins de 18 ans exclue. Dans l'historique, le taux national de chaque année est utilisé.")
    for message in policy_warnings(scenario.get("politiques", []), excluded):
        st.warning(message)
    matching, matching_error = None, None
    try:
        if demand is not None:
            population, activity = optional_match_inputs()
            matching = ensure_matching(state, demand, population, scenario["rapprochement"], context, activity)
        else:
            matching_error = "Les dossiers demand_projection/ et dataDemand/ doivent être présents pour calculer la demande."
    except (OSError, ValueError, KeyError, TypeError) as exc:
        matching_error = str(exc)
    tab_map, tab_trends, tab_status, tab_internat, tab_demand, tab_match, tab_method = st.tabs(
        ["Carte et indicateurs", "Dynamiques départementales", "Statuts d’exercice", "Internat national",
         "Demande de soins", "Offre / demande", "Calibration et hypothèses"])
    with tab_map:
        scenario_drees, year = render_map_tab(app, df, demand, demand_error, matching, matching_error, scenario)
    shown_choice = f"Projection DREES : {scenario_drees} · année {year}, réglables sous la carte."
    with tab_trends:
        st.caption(shown_choice)
        active = proposed["dynamique"]["active"]
        start = proposed["dynamique"]["debut"]
        if active:
            st.info(f"Dynamique réellement appliquée : OUI, dès {anchor + 1}. Pente d'effectif estimée sur {start}–{anchor} puis raccordée à l'observation {anchor}.")
        else:
            st.warning("Dynamique réellement appliquée : NON. La tendance nationale des arrivées étrangères reste active, même avec le curseur à 0 %. Elle peut produire des écarts territoriaux. Des projections régionales détaillées peuvent aussi avoir leur propre évolution.")
        st.write(f"Les courbes montrent un écart en nombre de psychiatres. Chaque effectif départemental est comparé au total national de la même année et du même scénario, multiplié par la part du département observée en {anchor}. La correction de l’internat, la tendance des diplômés à l’étranger et les hausses de places ou d’arrivées sont incluses dans les deux termes de la comparaison.")
        st.caption("Une hausse nationale répartie selon les parts départementales de l’année d’ancrage ne crée aucun écart. La somme des écarts de tous les départements vaut zéro chaque année. Les effets des politiques sur les effectifs restent visibles dans les indicateurs et la décomposition ci-dessous.")
        show_fit = st.checkbox("Afficher la droite estimée sur les observations", value=False, key="show_trend_fit")
        departments = st.multiselect("Départements à comparer", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]],
                                     format_func=lambda dep: f"{dep} · {names[dep]}", key="trend_departments")
        if departments:
            diagnostic = dynamics_diagnostic(state["panel"], state["contributions"], context, scenario_drees)
            st.plotly_chart(dynamics_figure(diagnostic, context, departments, True, start, show_fit), width="stretch")
            if any(proposed["formation"][f"region_{code}"] for code in TRAINING_REGIONS):
                st.caption(f"Ouvertures de places du scénario : premiers effets en {proposed['formation']['rentree'] + proposed['formation']['delai']}. Ce délai concerne les places de formation ; la hausse des arrivées étrangères intervient sans délai.")
            else:
                st.caption("Aucune hausse de places dans ce scénario.")
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
                st.dataframe(detail[["dep", "nom", "effectif_reference", "effet_dynamique", "effet_internat", "tendance_etranger", "ajout_immigration", "ajout_formation", "effectif", "reference_nationale", "ecart_complet"]].rename(
                    columns={"effectif_reference": "Référence sans modulation", "effet_dynamique": "Effet départemental", "effet_internat": "Correction internat",
                             "tendance_etranger": "Tendance étrangère / 2026", "ajout_immigration": "Hausse du curseur étranger", "ajout_formation": "Ajout des places", "effectif": "Effectif final",
                             "reference_nationale": "Référence nationale du graphique", "ecart_complet": "Écart affiché"}), hide_index=True, width="stretch")
                st.caption("L'effet départemental inclut la saturation et la renormalisation. L'ajout des places est un nombre de psychiatres supplémentaires, toujours positif ou nul.")
                st.caption(f"Référence nationale du graphique = total national final du scénario × part du département en {anchor}. Écart affiché = effectif final − cette référence. Cette référence sert à la comparaison ; elle ne s’ajoute pas aux composantes de l’effectif.")
            if state["foreign_trend_limited_rows"]:
                st.warning("L’écart cumulé des arrivées étrangères aurait rendu certains effectifs départementaux négatifs. Leur correction a été limitée à un effectif nul ; la décomposition montre le retrait effectivement appliqué.")
        else:
            st.info("Sélectionner au moins un département pour afficher les courbes.")
    with tab_status:
        st.caption(shown_choice)
        st.write(f"La part salariée est estimée séparément pour chaque département par une tendance logistique sur les comptes {context['observed_start']}–{anchor} de l’onglet Effectifs. La projection repart exactement de la part observée en {anchor}.")
        st.caption("Salariés : hospitaliers et autres salariés, hors exercice mixte. Libéraux : libéraux exclusifs et mixtes. Les deux groupes reconstituent toujours l’effectif total.")
        status_departments = st.multiselect("Départements pour les statuts", sorted(names),
                                            default=["75"] if "75" in names else [sorted(names)[0]],
                                            format_func=lambda dep: f"{dep} · {names[dep]}", key="status_departments")
        if status_departments:
            st.plotly_chart(status_figure(state["panel"], context, scenario_drees, status_departments), width="stretch")
            selected_status = state["panel"][(state["panel"].scenario == scenario_drees) & (state["panel"].annee == year)
                                               & state["panel"].dep.isin(status_departments)].copy()
            selected_status["part_salaries"] *= 100
            selected_status.insert(1, "nom", selected_status.dep.map(names))
            st.dataframe(selected_status[["dep", "nom", "effectif", "salaries", "liberaux", "part_salaries"]].rename(
                columns={"effectif": COUNT, "salaries": SALARIED_COUNT, "liberaux": LIBERAL_COUNT,
                         "part_salaries": SALARIED_SHARE + " (%)"}), hide_index=True, width="stretch")
        st.caption("Les nouvelles cohortes reçoivent la même répartition par statut que le total de leur département et de leur année. Cette extrapolation ne modélise pas les transitions individuelles entre modes d’exercice.")
    with tab_internat:
        st.caption(shown_choice)
        calibration = state["vacancy_calibration"]
        st.write("Le coefficient corrige les arrivées futures, en comparant le taux de remplissage extrapolé à un taux constant de référence. Il s’applique à la cohorte d’entrée en internat, puis produit un effet sur les effectifs cinq ans plus tard.")
        st.info("La série inclut le recul récent des postes non pourvus : 86/554 en 2025 (15,52 %), puis 37/653 en 2026 (5,67 %), hors CESP pour ces deux derniers millésimes. L’évolution n’est donc pas une hausse continue.")
        st.caption(f"Estimation : {calibration['debut']}–{calibration['derniere_observation']}. Taux de référence : {100 * calibration['taux_reference']:.2f} %. Correction {'active' if proposed['internat']['active'] else 'désactivée'}. La courbe ci-dessous montre l’extrapolation estimée dans les deux cas.")
        if calibration["methode"] == VACANCY_PLATEAU:
            st.caption(f"Plateau immédiat dès {anchor + 1} : {100 * calibration['plateau']:.2f} %, maximum observé sur {calibration['debut']}–{anchor}. Les années observées sont conservées.")
        else:
            st.caption(f"Projection linéaire : {100 * calibration['pente_annuelle']:+.3f} point par an, raccordée au dernier taux observé, bornée entre 0 et 100 %.")
        st.plotly_chart(vacancy_figure(state["vacancy_curve"], calibration), width="stretch")
        if state["vacancy_curve"].borne_atteinte.any():
            st.warning("La droite extrapolée sort de l’intervalle 0–100 % à certains horizons. Les taux y sont maintenus à la borne correspondante.")
        st.caption(f"Coefficient de flux = min(1, (1 − taux futur) / (1 − taux de référence)). Si la tendance estimée s’améliore, aucune réduction supplémentaire n’est appliquée. Les cohortes entrées au plus tard en {anchor} conservent un coefficient de 1.")
        flows = state["national_flows"]
        shown = flows[flows.scenario == scenario_drees]
        if shown.plafond_atteint.any():
            st.warning("L’extrapolation retirerait plus de psychiatres que le total de référence dans certaines années. Le retrait est limité à ce total ; ces horizons demandent de revoir les hypothèses.")
        st.dataframe(shown[["annee", "annee_entree_internat", "psychiatrie", "coefficient_entrees", "flux_fr_reference", "flux_fr_corrige", "manque_cumule_france", "coefficient_stock"]].rename(
            columns={"annee": "Année d’arrivée", "annee_entree_internat": "Cohorte d’internat",
                     "psychiatrie": "Stock d’étudiants DREES", "coefficient_entrees": "Coefficient du flux",
                     "flux_fr_reference": "Flux France avant correction", "flux_fr_corrige": "Flux France corrigé",
                     "manque_cumule_france": "Déficit cumulé France", "coefficient_stock": "Facteur uniforme des effectifs"}),
                     hide_index=True, width="stretch")
        st.write("Le stock d’étudiants est divisé par cinq pour approcher un flux annuel. Le déficit des futures cohortes est cumulé sans sorties ultérieures, puis retiré du total français selon la calibration formation/mobilité. Le même facteur annuel s’applique à chaque département. Les places supplémentaires utilisent le flux corrigé une seule fois.")
        st.warning("Scénario additionnel : la documentation publique ne fournit pas le taux exact de non-remplissage déjà inclus dans ProMédI. Le taux de référence constant est une convention, pas un paramètre DREES identifié. La division du stock par cinq et l’absence de sorties des cohortes manquantes restent des approximations.")
        with st.expander("Série historique et sources des postes non pourvus"):
            history_table = vacancy_history()
            history_table["Postes non pourvus (%)"] = history_table.taux * 100
            history_table["Source"] = history_table.source.map(VACANCY_SOURCES)
            st.dataframe(history_table[["annee", "Postes non pourvus (%)", "Source"]].rename(columns={"annee": "Année"}), hide_index=True, width="stretch")
            st.caption("2013–2022 : pourcentages publiés arrondis ; 2023–2026 : rapports calculés à partir des comptes publiés. Les anciennes publications ne précisent pas toutes le traitement des CESP. Cette limite de comparabilité est conservée dans le README.")
    with tab_demand:
        render_demand_tab(demand, demand_error, scenario.get("politiques", []), excluded)
    with tab_match:
        st.caption(shown_choice)
        render_match_tab(matching, matching_error, scenario_drees, year, names)
    with tab_method:
        st.caption(shown_choice)
        first = proposed["formation"]["rentree"] + proposed["formation"]["delai"]
        st.write(f"Places en deuxième année : hausse annuelle permanente à partir de {proposed['formation']['rentree']}, avec premiers effets en {first}. Délai retenu : {proposed['formation']['delai']} ans. Le défaut de dix ans correspond à cinq ans avant le DES et cinq ans de DES.")
        st.write("Pour chaque lieu de diplôme français, le volume annuel est approché par le stock national d'internes DREES en psychiatrie divisé par cinq, pondéré par la répartition régionale des étudiants à l'ancrage. La part de psychiatrie dans les étudiants est supposée identique entre régions. Une hausse de places de +u % ajoute u % de ce flux à chaque cohorte après le délai prévu. Le coefficient de postes non pourvus à l’entrée en internat, cinq ans plus tôt, est appliqué une seule fois aux cohortes françaises si la correction est activée.")
        st.write("Chaque cohorte supplémentaire est répartie selon la matrice diplôme → exercice observée à l'année d'ancrage, puis selon les parts départementales observées dans chaque région d'exercice. Les destinations hors France restent au dénominateur. Les cohortes sont cumulées sans attrition ni retraite supplémentaire jusqu'en 2070 ; les cohortes déjà arrivées ne sont pas redistribuées à chaque nouvelle année.")
        st.caption("Le tableau donne l’effet de +1 % pour la première cohorte supplémentaire ; les cohortes suivantes s’accumulent. Les volumes dépendent du scénario DREES. Les lieux de diplôme ultramarins sont groupés : Antilles-Guyane et La Réunion-Mayotte. Une région sans correspondance observée conserve un curseur désactivé ; ses habitants peuvent recevoir des psychiatres formés ailleurs.")
        st.dataframe(calibration_table(reference, context, scenario_drees, proposed, state["national_flows"], state["intervention_routing"], state["training_cohorts"]), hide_index=True, width="stretch")
        missing = [TRAINING_REGIONS[code] for code in TRAINING_REGIONS if code not in context["available"]]
        if missing:
            st.warning("Calibration de formation absente pour : " + ", ".join(missing) + ". Une ouverture ex nihilo nécessite des hypothèses supplémentaires ; le curseur reste désactivé. Ces régions peuvent recevoir des diplômés formés ailleurs.")
        st.write(f"Dynamiques : régression des écarts d'effectif départemental à la référence nationale sur {proposed['dynamique']['debut']}–{anchor}. La référence vaut le total français de chaque année multiplié par la part d'effectif du département observée en {anchor}. Chaque pente est exprimée en psychiatres par an.")
        st.write(f"Projection avant saturation = total national de référence × part départementale de {anchor} + pente départementale × années depuis {anchor}. Les pentes somment à zéro. L'effet commence dès {anchor + 1}. Le lissage sous 1 psychiatre et la renormalisation préservent des effectifs non négatifs et le total national.")
        if metadata["mode_projection"] == "regional":
            st.write(f"Lorsque les dynamiques sont activées, le total des projections régionales sert de trajectoire nationale : leur ventilation régionale est remplacée par les parts {anchor} et les écarts départementaux. Désactiver les dynamiques rétablit la ventilation régionale de référence.")
        st.write(f"Correspondance diplôme → exercice : parts fixes en {anchor}. La tendance nationale du stock de diplômés étrangers est estimée depuis {proposed['mobilite']['debut']}, sans taux de retraite. Le curseur majore uniquement son accroissement net positif, qui ne mesure pas les entrées brutes. Le code 99 désigne le lieu du diplôme, pas la nationalité.")
        st.caption("Le retrait des pédopsychiatres repose sur les qualifications enfant/adolescent du RPPS (SM43 et SM92), pas sur toute l'activité de soins réalisée auprès des enfants. Le taux national observé de chaque année s'applique uniformément ; le dernier taux est maintenu en projection.")
        with st.expander("Sources et champ des effectifs"):
            st.write(f"Onglet de projection des actifs : {metadata['onglet_drees']}.")
            if metadata["mode_projection"] == "national":
                st.write(f"Référence départementale = effectif RPPS en {anchor} × effectif national DREES futur / effectif national DREES en {anchor}. La croissance du champ DREES (70 ans ou moins) est appliquée au RPPS tous âges ; les plus de 70 ans ne font pas l'objet d'une projection distincte.")
            else:
                st.write("Les effectifs régionaux DREES de référence sont utilisés directement. Le RPPS comprend tous les âges, contre 70 ans ou moins pour les actifs DREES : une rupture de champ peut apparaître au raccord.")
            st.write("Les résultats sont des nombres de psychiatres et une part salariée en %. Les observations proviennent directement de l’onglet Effectifs. Les valeurs projetées fractionnaires représentent des effectifs attendus.")
            st.write("Aucun fichier de projections n'est écrit. Le calcul de référence est partagé en mémoire ; les paramètres et résultats simulés appartiennent à chaque session.")
        st.caption("Politiques nationales : prévalence × produit des coefficients (1 − baisse × montée en charge), uniquement dans les âges et années ciblés. Les populations des tranches Insee sont réparties au prorata des âges ; aucun taux propre aux 12–17 ans n'est déduit des taux adultes.")


def main():
    st.set_page_config(page_title="Projections psychiatres", layout="wide")
    st.markdown(PAGE_STYLE, unsafe_allow_html=True)
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
    demand, demand_error = None, None
    try:
        demand = demand_outcomes()
        choices, labels = matching_choices(demand)
    except (OSError, ValueError, KeyError, ImportError) as exc:
        demand, demand_error, choices, labels = None, str(exc), None, {}
    match_defaults = matching_defaults(choices, labels)
    predefined = predefined_scenarios(registry, context, match_defaults)
    # Un changement de sources/code réinitialise le scénario et l'éditeur.
    # Une nouvelle session démarre aussi sur le scénario Base.
    if st.session_state.get("parameter_source") != signature:
        for key in [k for k in st.session_state if str(k).startswith(("edit__", "edit_all"))]:
            del st.session_state[key]
        st.session_state.scenario_active = deepcopy(predefined["Base"])
        st.session_state.scenario_previous = None
        st.session_state.scenario_draft = make_scenario("Mon scénario", "Personnalisé",
                                                        predefined["Base"]["offre"], match_defaults)
        st.session_state.parameter_source = signature
    app = {"reference": reference, "metadata": metadata, "context": context, "registry": registry,
           "signature": signature, "demand": demand, "demand_error": demand_error, "choices": choices,
           "labels": labels, "match_defaults": match_defaults, "predefined": predefined,
           "names": context["baseline"].set_index("dep").nom.to_dict()}
    pages = {}

    def projections():
        render_projections_page(app, pages["scenarios"])

    def scenarios():
        render_scenarios_page(app, pages["projections"])

    pages["projections"] = st.Page(projections, title="Projections", icon=":material/map:", url_path="projections", default=True)
    pages["scenarios"] = st.Page(scenarios, title="Scénarios", icon=":material/tune:", url_path="scenarios")
    navigation = st.navigation({"Tableau de bord": [pages["projections"], pages["scenarios"]]})
    render_sidebar_summary(st.session_state.scenario_active)
    navigation.run()


if __name__ == "__main__":
    main()
