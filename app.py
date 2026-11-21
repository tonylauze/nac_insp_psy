"""Carte des psychiatres : calcul en mémoire, propre à chaque session utilisateur.

Installation : pip install streamlit plotly pandas numpy openpyxl
Lancement   : streamlit run app.py

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
Dynamiques départementales et places de formation : voir README.md pour les formules.
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
GEOJSON_URL = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements-version-simplifiee.geojson"
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
MATCH_INDICATORS = {"actes_par_psychiatre": ACTS_PER_PSY, "actes_disponibles": AVAILABLE_ACTS,
                    "actes_necessaires": REQUIRED_ACTS, "solde_actes_habitant": ACTS_BALANCE}
# 15,865 millions de consultations/téléconsultations / 6 308 psychiatres libéraux (CNAM 2024).
# Les autres valeurs sont des hypothèses de scénario, pas des normes de soins.
MATCH_DEFAULTS = {"actes_liberal": 2515.0, "ratio_salarie": 0.65, "croissance_activite": 0.0,
                  "actes_par_besoin": 8.0, "couverture": 25.0, "part_offre": 50.0}
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
DES_YEARS = 5
STUDY_YEARS = 5 + DES_YEARS  # Depuis l'entrée en deuxième année ; réglable.
FLOW_ALLOCATION = "Installations estimées (delta + retraites)"
STOCK_ALLOCATION = "Répartition des effectifs projetés"
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
    retention = context["routing"].groupby(["origine", "annee"], as_index=False).part.sum()
    weights = context["weights"].merge(retention, on="origine", validate="many_to_many")
    weights["contribution_metro"] = weights.poids * weights.part
    shares = weights.groupby(["scenario", "annee"]).contribution_metro.sum()
    keys = pd.MultiIndex.from_frame(flows[["scenario", "annee"]])
    flows["part_flux_metropole"] = shares.reindex(keys).to_numpy()
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
    """Stocks linéaires raccordés, puis I(t)=max(0, N(t)-(1-r)N(t-1))."""
    start = parameters["debut"]
    retirement = parameters["retraite"] / 100.0
    selected = history[history.annee.between(start, anchor)].copy()
    expected_years = list(range(start, anchor + 1))
    if (len(expected_years) < 5 or sorted(selected.annee.unique()) != expected_years
            or not selected.groupby(["origine", "destination"]).size().eq(len(expected_years)).all()):
        raise ValueError("Fenêtre d'estimation diplôme/exercice incomplète.")
    x = selected.annee - np.mean(expected_years)
    selected["produit"], selected["carre"] = x * selected.effectif, x * x
    sums = selected.groupby(["origine", "destination"])[["produit", "carre"]].sum()
    slopes = (sums.produit / sums.carre).rename("pente")
    base = history[history.annee == anchor].set_index(["origine", "destination"]).effectif.rename("ancrage")
    previous = history[history.annee == anchor - 1].set_index(["origine", "destination"]).effectif
    models = pd.concat([base, slopes], axis=1).reset_index()
    panel = models.merge(pd.DataFrame({"annee": range(anchor, end + 1)}), how="cross")
    panel["effectif"] = np.maximum(0.0, panel.ancrage + panel.pente * (panel.annee - anchor))
    panel = panel.sort_values(["origine", "destination", "annee"]).reset_index(drop=True)
    panel["precedent"] = panel.groupby(["origine", "destination"]).effectif.shift()
    at_anchor = panel.annee == anchor
    keys = pd.MultiIndex.from_frame(panel.loc[at_anchor, ["origine", "destination"]])
    panel.loc[at_anchor, "precedent"] = previous.reindex(keys).to_numpy()
    if panel.precedent.isna().any():
        raise ValueError("Effectif précédent absent pour le calcul des installations.")
    panel["installations_brutes"] = panel.effectif - (1 - retirement) * panel.precedent
    panel["installations"] = panel.installations_brutes.clip(lower=0.0)
    group = ["origine", "annee"]
    stock_total = panel.groupby(group).effectif.transform("sum")
    base_total = panel.groupby(group).ancrage.transform("sum")
    base_share = np.divide(panel.ancrage.to_numpy(), base_total.to_numpy(), out=np.zeros(len(panel)), where=base_total.to_numpy() > 0)
    stock_share = np.divide(panel.effectif.to_numpy(), stock_total.to_numpy(), out=base_share.copy(), where=stock_total.to_numpy() > 0)
    if parameters["methode"] == FLOW_ALLOCATION:
        flow_total = panel.groupby(group).installations.transform("sum")
        panel["part"] = np.divide(panel.installations.to_numpy(), flow_total.to_numpy(), out=stock_share.copy(), where=flow_total.to_numpy() > 0)
        panel["repli"] = np.where(flow_total > 0, "Aucun", np.where(stock_total > 0, "Effectifs projetés", "Répartition à l'ancrage"))
    else:
        panel["part"] = stock_share
        panel["repli"] = np.where(stock_total > 0, "Aucun", "Répartition à l'ancrage")
    # Les origines intégralement nulles sont indisponibles, pas renormalisées arbitrairement.
    nonzero = base_total > 0
    totals = panel[nonzero].groupby(group).part.sum()
    if not np.allclose(totals, 1.0, atol=1e-10) or not panel.part.between(0, 1).all():
        raise ValueError("Répartition d'installations invalide.")
    panel["projete"] = panel.annee > anchor
    return panel


def installation_history(history, retirement):
    """Reconstituer les flux sur les stocks OBSERVÉS avant toute régression."""
    result = history.sort_values(["origine", "destination", "annee"]).copy()
    if result.duplicated(["origine", "destination", "annee"]).any():
        raise ValueError("Historique d'installations dupliqué.")
    grouped = result.groupby(["origine", "destination"])
    result["precedent"] = grouped.effectif.shift()
    previous_year = grouped.annee.shift()
    valid = result.precedent.notna()
    if not (result.loc[valid, "annee"] - previous_year[valid]).eq(1).all():
        raise ValueError("Année manquante dans l'historique d'installations.")
    result = result.loc[valid].copy()
    result["installations_brutes"] = result.effectif - (1 - retirement) * result.precedent
    result["installations"] = result.installations_brutes.clip(lower=0.0)
    return result


def extrapolate_installations(observed, start, anchor, end):
    """MCO sur les flux historiques, puis pente raccordée au flux reconstitué en b."""
    start = max(start, int(observed.annee.min()))
    selected = observed[observed.annee.between(start, anchor)].copy()
    years = list(range(start, anchor + 1))
    keys = ["origine", "destination"]
    if (len(years) < 5 or sorted(selected.annee.unique()) != years
            or not selected.groupby(keys).size().eq(len(years)).all()):
        raise ValueError("Au moins cinq années complètes d'installations sont nécessaires.")
    x = selected.annee - np.mean(years)
    selected["produit"], selected["carre"] = x * selected.installations, x * x
    sums = selected.groupby(keys)[["produit", "carre"]].sum()
    slopes = (sums.produit / sums.carre).rename("pente_installations")
    base = observed[observed.annee == anchor].set_index(keys).installations.rename("installations_ancrage")
    models = pd.concat([base, slopes], axis=1).reset_index()
    if models[["installations_ancrage", "pente_installations"]].isna().any().any():
        raise ValueError("Modèle d'installations incomplet à l'ancrage.")
    result = models.merge(pd.DataFrame({"annee": range(anchor, end + 1)}), how="cross")
    result["installations_brutes"] = result.installations_ancrage + result.pente_installations * (result.annee - anchor)
    result["installations"] = result.installations_brutes.clip(lower=0.0)
    result["debut_estimation"] = start
    result["projete"] = result.annee > anchor
    return result


def installation_distribution(panel, method):
    """Parts pour les curseurs, avec repli sur les stocks si tous les flux sont nuls."""
    result = panel.copy()
    groups = ["origine", "annee"]
    totals = result.groupby(groups)
    stock_total = totals.effectif.transform("sum").to_numpy()
    base_total = totals.ancrage.transform("sum").to_numpy()
    base_share = np.divide(result.ancrage.to_numpy(), base_total, out=np.zeros(len(result)), where=base_total > 0)
    stock_share = np.divide(result.effectif.to_numpy(), stock_total, out=base_share.copy(), where=stock_total > 0)
    if method == FLOW_ALLOCATION:
        flow_total = totals.installations.transform("sum").to_numpy()
        result["part"] = np.divide(result.installations.to_numpy(), flow_total, out=stock_share.copy(), where=flow_total > 0)
        result["repli"] = np.where(flow_total > 0, "Aucun", np.where(stock_total > 0, "Effectifs projetés", "Répartition à l'ancrage"))
    else:
        result["part"] = stock_share
        result["repli"] = np.where(stock_total > 0, "Aucun", "Répartition à l'ancrage")
    nonzero = base_total > 0
    if (not np.allclose(result[nonzero].groupby(groups).part.sum(), 1.0, atol=1e-10)
            or not result.part.between(0, 1).all()):
        raise ValueError("Répartition des cohortes supplémentaires invalide.")
    return result


def project_installation_distribution(history, parameters, stock_panel, anchor, end):
    """I historiques -> proportions historiques -> extrapolation bornée des proportions."""
    france = history[history.destination != "99"]
    observed = installation_history(france, parameters["retraite"] / 100.0)
    projected = stock_panel[stock_panel.destination != "99"].drop(
        columns=["installations", "installations_brutes", "part", "repli"]).copy()
    base = projected[projected.annee == anchor][["origine", "destination", "ancrage"]]
    observed = observed.merge(base, on=["origine", "destination"], validate="many_to_one")
    observed["projete"] = False
    observed = installation_distribution(observed, FLOW_ALLOCATION)
    observed["part_installations"] = observed.part
    start = max(parameters["debut"], int(observed.annee.min()))
    selected = observed[observed.annee.between(start, anchor)].copy()
    years = list(range(start, anchor + 1))
    keys = ["origine", "destination"]
    if (len(years) < 5 or sorted(selected.annee.unique()) != years
            or not selected.groupby(keys).size().eq(len(years)).all()):
        raise ValueError("Fenêtre des proportions d'installations incomplète.")
    x = selected.annee - np.mean(years)
    selected["produit"], selected["carre"] = x * selected.part, x * x
    sums = selected.groupby(keys)[["produit", "carre"]].sum()
    slopes = (sums.produit / sums.carre).rename("pente_part")
    initial = observed[observed.annee == anchor].set_index(keys).part.rename("part_ancrage")
    models = pd.concat([initial, slopes], axis=1).reset_index()
    projected = projected.merge(models, on=keys, validate="many_to_one")
    projected["part_brute"] = projected.part_ancrage + projected.pente_part * (projected.annee - anchor)
    projected["part_bornee"] = projected.part_brute.clip(0.0, 1.0)
    projected["borne_part"] = (projected.part_brute < 0) | (projected.part_brute > 1)
    projected["debut_estimation"] = start
    fallback = installation_distribution(projected, STOCK_ALLOCATION)
    if parameters["methode"] == FLOW_ALLOCATION:
        total = projected.groupby(["origine", "annee"]).part_bornee.transform("sum").to_numpy()
        projected["part"] = np.divide(projected.part_bornee.to_numpy(), total,
                                      out=fallback.part.to_numpy().copy(), where=total > 0)
        projected["repli"] = np.where(total > 0, "Aucun", np.where(fallback.repli == "Aucun", "Effectifs projetés", fallback.repli))
    else:
        projected["part"], projected["repli"] = fallback.part, fallback.repli
        observed = installation_distribution(observed, STOCK_ALLOCATION)
    if not projected.part.between(0, 1).all():
        raise ValueError("Proportion d'installation hors de 0–100 %.")
    return projected, observed


def origin_installation_flows(history, parameters, anchor, end):
    """Volumes des deux curseurs : stocks observés agrégés par origine -> I -> projection."""
    france = history[history.destination != "99"]
    national = france.groupby(["origine", "annee"], as_index=False).effectif.sum()
    national["destination"] = "France"
    observed = installation_history(national, parameters["retraite"] / 100.0)
    projected = extrapolate_installations(observed, parameters["debut"], anchor, end)
    return projected, observed


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


def foreign_reference_flows(panel, retirement, anchor):
    """Flux national de diplômés étrangers, puis écart annuel au flux observé en b.

    Le max est appliqué APRÈS agrégation. La destination 99 (COM ou étranger)
    n'est pas une région d'exercice française identifiable et reste exclue.
    Les écarts au flux de référence peuvent être signés, jamais les flux eux-mêmes.
    """
    foreign = panel[(panel.origine == "99") & (panel.destination != "99")]
    totals = foreign.groupby("annee", as_index=False)[["effectif", "precedent"]].sum()
    totals["flux_reference"] = (totals.effectif - (1 - retirement) * totals.precedent).clip(lower=0.0)
    baseline = totals.loc[totals.annee == anchor, "flux_reference"]
    if len(baseline) != 1:
        raise ValueError("Flux étranger de l'année de référence absent ou ambigu.")
    totals["flux_ancrage"] = float(baseline.iloc[0])
    totals["ecart_flux"] = np.where(totals.annee > anchor, totals.flux_reference - totals.flux_ancrage, 0.0)
    totals["origine"] = "99"
    return totals


def foreign_france_distribution(panel):
    """Répartition conditionnelle à une installation en France, DROM inclus.

    Le flux calculé porte sur la France : les parts sont renormalisées sur ce
    même champ avant de retenir les départements métropolitains de l'application.
    """
    france = panel[(panel.origine == "99") & (panel.destination != "99")].copy()
    total = france.groupby("annee").part.transform("sum").to_numpy()
    current = france.groupby("annee").effectif.transform("sum").to_numpy()
    initial = france.groupby("annee").ancrage.transform("sum").to_numpy()
    initial_parts = np.divide(france.ancrage.to_numpy(), initial, out=np.zeros(len(france)), where=initial > 0)
    fallback = np.divide(france.effectif.to_numpy(), current, out=initial_parts.copy(), where=current > 0)
    france["part"] = np.divide(france.part.to_numpy(), total, out=fallback.copy(), where=total > 0)
    if not np.allclose(france.groupby("annee").part.sum(), 1.0):
        raise ValueError("Aucune répartition française disponible pour l'origine 99.")
    return france


def prepare_arrivals(reference, parameters, context):
    flows = context["national_flows"]
    origin_flows, historical_origins = origin_installation_flows(context["mobility_history"],
        parameters["mobilite"], context["anchor"], int(reference.annee.max()))
    training = origin_flows[origin_flows.origine.isin(context["available"])].rename(columns={"installations": "flux_reference"})
    training = training.merge(flows[["annee", "scenario", "coefficient_entrees"]], on="annee", validate="many_to_many")
    training["flux_corrige"] = training.flux_reference * training.coefficient_entrees
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    increases = training.origine.map(lambda code: parameters["formation"].get(f"region_{code}", 0) / 100.0)
    training["flux_ajoute"] = np.where(training.annee >= first, training.flux_corrige * increases, 0.0)
    panel = context["mobility_panel"]
    foreign = foreign_reference_flows(panel, parameters["mobilite"]["retraite"] / 100.0, context["anchor"])
    intervention = origin_flows[origin_flows.origine == "99"]
    foreign = foreign.merge(intervention[["annee", "installations", "pente_installations"]].rename(
        columns={"installations": "flux_intervention", "pente_installations": "pente_flux_intervention"}),
        on="annee", validate="one_to_one")
    foreign = foreign.merge(pd.DataFrame({"scenario": sorted(reference.scenario.unique())}), how="cross")
    foreign["flux_hausse"] = np.where(foreign.annee >= parameters["immigration"]["debut"],
        foreign.flux_intervention * parameters["immigration"]["hausse"] / 100.0, 0.0)
    foreign["flux_ajoute"] = foreign.ecart_flux + foreign.flux_hausse
    routing = annual_routing(foreign_france_distribution(panel), context["baseline"])
    intervention_routing = annual_routing(foreign_france_distribution(context["installation_panel"]), context["baseline"])
    return {"training_cohorts": training, "foreign_cohorts": foreign,
            "training_arrivals": accumulate_arrivals(reference, training, context["intervention_routing"]),
            "foreign_routing": routing, "foreign_intervention_routing": intervention_routing,
            "origin_installations": origin_flows, "origin_installation_history": historical_origins,
            "foreign_trend_arrivals": accumulate_arrivals(reference, foreign.assign(flux_ajoute=foreign.ecart_flux), routing),
            "foreign_arrivals": accumulate_arrivals(reference, foreign.assign(flux_ajoute=foreign.flux_hausse), intervention_routing)}


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
    available = set(origin_totals[origin_totals > 0].index) & set(REGIONS)
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
        for code, label in sorted(REGIONS.items(), key=lambda item: item[1])
    )
    return (
        Modulation("mobilite", "Installations selon le lieu de diplôme",
                   "La trajectoire démographique conserve les extrapolations de stocks. Pour les politiques publiques, les proportions et les volumes sont estimés séparément à partir des installations historiques, puis extrapolés.",
                   (Control("methode", "Répartition des nouveaux arrivants pour les politiques publiques", "select", FLOW_ALLOCATION,
                            options=(FLOW_ALLOCATION, STOCK_ALLOCATION),
                            help="Ce choix modifie uniquement les proportions des cohortes supplémentaires. Il ne change ni leurs volumes par lieu de diplôme, ni la trajectoire démographique sans hausse."),
                    Control("retraite", "Taux annuel uniforme de départ à la retraite (%)", "number", 3.0, 0.0, 20.0, 0.1,
                            help="Hypothèse réglable, non estimée : installations = max(0, effectif courant − (1 − taux) × effectif précédent)."),
                    Control("debut", "Première année pour les dynamiques diplôme/exercice", "number",
                            context["mobility_start"], context["mobility_start"], context["anchor"] - 4, 1,
                            help="Effectifs à partir de cette année. Pour les installations, première année disponible : 2013, calculée avec 2012. Une fenêtre débutant plus tard utilise aussi le compte précédent pour son premier delta.")), calibration_only),
        Modulation("dynamique", "Dynamiques départementales",
                   f"Tendances des effectifs (psychiatres par an), estimées jusqu'à {context['anchor']}. Effet dès {context['anchor'] + 1}, indépendant du délai de formation.",
                   (Control("active", "Extrapoler les dynamiques départementales", "checkbox", True),
                    Control("debut", "Première année pour estimer la tendance", "number", context["observed_start"],
                            context["observed_start"], context["anchor"] - 4, 1,
                            help="Tout l'historique par défaut. Choisir 2017 pour dix années, ou 2022 pour cinq années avec les sources 2026. Une fenêtre courte peut modifier fortement la pente.")), department_dynamics),
        Modulation("internat", "Dynamique nationale de l'internat",
                   "Correction uniforme des futures cohortes par une tendance linéaire du taux de postes non pourvus. L'amélioration de 2026 est incluse.",
                   (Control("active", "Appliquer la correction des postes non pourvus", "checkbox", True),
                    Control("debut", "Première année pour la tendance des postes non pourvus", "number",
                            min(v[0] for v in VACANCY_OBSERVATIONS), min(v[0] for v in VACANCY_OBSERVATIONS),
                            max(v[0] for v in VACANCY_OBSERVATIONS) - 4, 1,
                            help="Régression linéaire des taux, avec au moins cinq années, puis raccord au dernier taux observé. Les taux restent entre 0 et 100 %. Une tendance à l'amélioration ne crée pas de hausse supplémentaire : le coefficient est plafonné à 1.")),
                   internat_correction),
        Modulation("etranger_tendance", "Tendance nationale des arrivées étrangères", "", (), foreign_trend_adjustment),
        Modulation("immigration", "Diplômés à l'étranger — code 99",
                   "La tendance démographique issue des effectifs reste active à 0 %. Le curseur ajoute un pourcentage du flux national reconstitué sur l'historique puis extrapolé directement, sans délai de formation. Le code 99 désigne le lieu du diplôme.",
                   (Control("hausse", "Hausse des arrivées de diplômés à l'étranger (%)", "slider", 0, 0, 100, 1),
                    Control("debut", "Première année de hausse des arrivées étrangères", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()), 1)), immigration_additions),
        Modulation("formation", "Places de formation par région",
                   "Hausse des places en deuxième année, reconduite pour chaque cohorte. Volume prévu à partir des installations historiques du lieu de diplôme ; répartition selon les pourcentages extrapolés ou les stocks, au choix.",
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
    demographic_parameters = {**parameters["mobilite"], "methode": FLOW_ALLOCATION}
    mobility_panel = project_diploma_distribution(context["mobility_history"], demographic_parameters,
                                                  context["anchor"], int(reference.annee.max()))
    run_context["mobility_panel"] = mobility_panel
    run_context["routing"] = annual_routing(mobility_panel, context["baseline"])
    installation_panel, historical_installations = project_installation_distribution(
        context["mobility_history"], parameters["mobilite"], mobility_panel, context["anchor"], int(reference.annee.max()))
    run_context["installation_panel"] = installation_panel
    run_context["intervention_routing"] = annual_routing(installation_panel, context["baseline"])
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
    if return_details:
        keys = pd.MultiIndex.from_frame(projected[["dep", "annee", "scenario"]])
        requested = run_context["foreign_trend_arrivals"].set_index(["dep", "annee", "scenario"]).cumul.reindex(keys).to_numpy()
        limited = int((~np.isclose(contributions.loc[projected.index, "etranger_tendance"], requested, atol=1e-8, rtol=1e-10)).sum())
        return result, contributions, {"national_flows": flows, "vacancy_curve": curve,
            "vacancy_calibration": calibration, "mobility_panel": mobility_panel,
            "routing": run_context["routing"], "installation_panel": installation_panel,
            "installation_history": historical_installations, "intervention_routing": run_context["intervention_routing"],
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
    stocks, weights = training_students(DREES_PATH, anchor, names)
    mobility = diploma_history(RPPS_PATH, anchor)
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
)
# Regroupement des modulations sur la page Scénarios.
SCENARIO_GROUPS = (("Paramètres structurels", ("dynamique", "internat", "mobilite")),
                   ("Politiques publiques", ("formation", "immigration")),
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


def make_scenario(name, origin, offre, rapprochement, description=""):
    return {"nom": name, "origine": origin, "description": description,
            "offre": deepcopy(offre), "rapprochement": dict(rapprochement)}


def predefined_scenarios(registry, context, match_values):
    """Scénarios illustratifs : seuls les paramètres cités diffèrent du scénario Base."""
    base = default_parameters(registry)
    first = context["anchor"] + 1
    places = {f"region_{code}": 20 for code in REGIONS if code in context["available"]}

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
    return {scenario["nom"]: scenario for scenario in scenarios}


def complete_matching(scenario, defaults):
    """Compléter les listes offre/demande d'un scénario créé sans données de demande."""
    result = deepcopy(scenario)
    for key, _ in MATCH_CHOICES:
        if result["rapprochement"].get(key) is None:
            result["rapprochement"][key] = defaults[key]
    return result


def check_scenario(scenario, registry, choices):
    validate_parameters(scenario["offre"], registry)
    match = scenario["rapprochement"]
    validate_match_parameters(match)
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
    known = set(controls) | {(MATCH_BLOCK, key) for key in choice_keys | set(number_keys)}
    unknown = [f"{b}/{p}" for b, p in zip(table.bloc, table.parametre) if (b, p) not in known]
    if unknown:
        raise ValueError("Paramètres inconnus : " + ", ".join(unknown[:8]) + ".")
    names = [value for value in table.get("scenario", pd.Series(dtype=str)) if value]
    scenario = make_scenario(names[0] if names else "Scénario importé", "Importé",
                             reference["offre"], reference["rapprochement"])
    for row in table.itertuples(index=False):
        try:
            if (row.bloc, row.parametre) in controls:
                scenario["offre"][row.bloc][row.parametre] = convert_control(controls[(row.bloc, row.parametre)], row.valeur)
            elif row.parametre in choice_keys:
                scenario["rapprochement"][row.parametre] = row.valeur
            else:
                scenario["rapprochement"][row.parametre] = parse_number(row.valeur, False)
        except (ValueError, TypeError):
            raise ValueError(f"Valeur invalide pour {row.bloc}/{row.parametre} : {row.valeur!r}.") from None
    absent = sorted(known - set(zip(table.bloc, table.parametre)))
    warnings = []
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
    places = [offre["formation"][f"region_{code}"] for code in REGIONS if offre["formation"].get(f"region_{code}")]
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
    lines.append(f"Départs à la retraite : {display_value(offre['mobilite']['retraite'])} % par an")
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
                                                    scenario["offre"], scenario["rapprochement"])
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
    return make_scenario(name, "Personnalisé", offre, match)


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
    gj["features"] = [f for f in gj["features"] if f["properties"]["code"] in METRO]
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
    """Écarts d'effectif à une référence nationale commune aux deux vues."""
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
    diagnostic["reference_nationale"] = count_reference(diagnostic, context["baseline"], "avant_ajouts")
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
                    annotation_text=f"Part nationale fixée en {anchor}, avant évolution étrangère et nouvelles places")
    figure.add_vline(x=anchor, line_color="#94a3b8", line_dash="dot")
    figure.update_layout(height=570, xaxis_title="Année", yaxis_title="Écart à la référence nationale (nombre de psychiatres)",
                         margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified")
    return figure


def calibration_table(reference, context, scenario, parameters, flows, routing, cohorts):
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    row = flows[(flows.scenario == scenario) & (flows.annee == first)].iloc[0]
    national = float(row.total_reference * row.coefficient_stock)
    origin_flows = cohorts[(cohorts.scenario == scenario) & (cohorts.annee == first)].set_index("origine")
    retention = routing[routing.annee == first].groupby("origine").part.sum()
    records = []
    for code, label in REGIONS.items():
        available = code in context["available"]
        flow = float(origin_flows.loc[code, "flux_corrige"]) if available else 0.0
        additions = 0.01 * flow * retention.get(code, 0)
        records.append({"Région de formation": label, "Calibration": "Disponible" if available else "Indisponible",
                        "Hausse validée (%)": parameters["formation"][f"region_{code}"],
                        "Installations extrapolées France": float(origin_flows.loc[code, "flux_reference"]) if available else np.nan,
                        "Après correction des postes vacants": flow if available else np.nan,
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
    figure.add_trace(go.Scatter(x=future.annee, y=100 * future.taux_projete, mode="lines", name="Extrapolation linéaire", line_dash="dash"))
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


@st.cache_data(show_spinner=False, max_entries=2)
def load_demand_summary(path: str, modified_at_ns: int, signature):
    """Compléter aussi les exports existants, sans exiger un nouveau CSV."""
    del modified_at_ns, signature
    from demand_projection.run import SCENARIO_FILES, read_population_workbook
    from demand_projection.outcomes import prepare_outcomes

    demand = prepare_outcomes(pd.read_csv(path).drop(columns=["population_total"], errors="ignore"))
    totals = []
    for scenario, filename in SCENARIO_FILES.items():
        population, years = read_population_workbook(ROOT / "dataDemand" / filename)
        totals.append(insee_total_population(population, years, scenario))
    return join_demand_population(demand, pd.concat(totals, ignore_index=True))


@st.cache_data(show_spinner=False, max_entries=2)
def compute_demand_summary(signature):
    """Calcul en mémoire avec les modules d'origine ; adaptation limitée à app.py."""
    del signature
    from demand_projection.run import (SCENARIO_FILES, read_population_workbook,
        validate_parameters, read_need_scenarios, read_territorial_modifiers,
        project_workbook, summarize_projection)
    from demand_projection.outcomes import prepare_outcomes

    data_dir = ROOT / "dataDemand"
    parameters = validate_parameters(pd.read_csv(data_dir / "prevalence_parameters.csv"))
    scenarios = read_need_scenarios(data_dir / "need_scenarios.csv")
    modifiers = read_territorial_modifiers(data_dir / "territorial_modifiers.csv")
    projections, totals = [], []
    for scenario, filename in SCENARIO_FILES.items():
        population, years = read_population_workbook(data_dir / filename)
        totals.append(insee_total_population(population, years, scenario))
        # Le module d'origine affecte des proratas aux colonnes POP_YYYY.
        # Préparer des flottants localement évite les erreurs de dtype sous pandas 3.
        columns = [f"POP_{year}" for year in years]
        population = population.astype({column: float for column in columns})
        projections.append(project_workbook(population, years, scenario, parameters, scenarios, modifiers, 2026))
    summary = summarize_projection(pd.concat(projections, ignore_index=True))
    demand = prepare_outcomes(summary.drop(columns=["population_total"], errors="ignore"))
    return join_demand_population(demand, pd.concat(totals, ignore_index=True))


def demand_outcomes():
    """Use a current local export, otherwise calculate from versioned inputs."""
    from demand_projection.run import SCENARIO_FILES

    data_dir = ROOT / "dataDemand"
    inputs = [data_dir / filename for filename in SCENARIO_FILES.values()]
    inputs += [data_dir / "prevalence_parameters.csv", data_dir / "need_scenarios.csv",
               ROOT / "demand_projection" / "run.py", ROOT / "demand_projection" / "outcomes.py"]
    modifiers = data_dir / "territorial_modifiers.csv"
    if modifiers.exists():
        inputs.append(modifiers)
    signature = tuple((str(path), path.stat().st_mtime_ns, path.stat().st_size) for path in inputs)
    if DEMAND_SUMMARY_PATH.exists():
        modified = DEMAND_SUMMARY_PATH.stat().st_mtime_ns
        if modified >= max(item[1] for item in signature):
            return load_demand_summary(str(DEMAND_SUMMARY_PATH), modified, signature)
    return compute_demand_summary(signature)


def render_demand_tab(demand=None, error=None):
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
    names[normalize("Réunion")] = "974"  # Libellé Omphale, reconnu avant le filtre métropole.
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
              "actes_par_besoin": (0, 100), "couverture": (0, 100), "part_offre": (0, 100)}
    for name, (lo, hi) in limits.items():
        value = parameters[name]
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
    numeric = list(MATCH_INDICATORS) + ["actes_totaux", "solde_actes"]
    if not np.isfinite(result[numeric]).all().all():
        raise ValueError("Résultat offre/demande non fini.")
    return result


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
    st.write("Le solde compare des actes annuels équivalents de consultation pour le besoin sélectionné : (actes disponibles − actes nécessaires) / population totale du département.")
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
    a.metric("Actes disponibles — métropole", f"{selected.actes_disponibles.sum():,.0f}".replace(",", " "))
    b.metric("Actes nécessaires — métropole", f"{selected.actes_necessaires.sum():,.0f}".replace(",", " "))
    c.metric("Solde annuel par habitant — métropole", f"{selected.solde_actes.sum() / selected.population_total.sum():+.3f}")
    departments = st.multiselect("Départements du rapprochement", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]], format_func=lambda dep: f"{dep} · {names[dep]}", key="match_departments")
    shown = detail[detail.dep.isin(departments)].copy()
    shown["Département"] = shown.dep.map(names)
    if not shown.empty:
        figure = px.line(shown, x="annee", y="solde_actes_habitant", color="Département", labels={"annee": "Année", "solde_actes_habitant": "Solde annuel d’actes par habitant"})
        figure.add_hline(y=0, line_dash="dot", line_color="#475569")
        st.plotly_chart(figure, width="stretch")
    selected.insert(1, "nom", selected.dep.map(names))
    st.dataframe(selected[["dep", "nom", "annee", "population_total", "persons_with_need", "personnes_ciblees", "actes_par_psychiatre", "actes_totaux", "actes_disponibles", "actes_necessaires", "solde_actes", "solde_actes_habitant"]].rename(columns={
        "annee": "Année", "population_total": "Population totale", "persons_with_need": "Personnes avec besoin potentiel", "personnes_ciblees": "Personnes ciblées", "actes_par_psychiatre": ACTS_PER_PSY,
        "actes_totaux": "Actes totaux avant affectation", "actes_disponibles": AVAILABLE_ACTS, "actes_necessaires": REQUIRED_ACTS, "solde_actes": "Solde annuel d’actes", "solde_actes_habitant": ACTS_BALANCE}), hide_index=True, width="stretch")
    with st.expander("Hypothèses et sources du rapprochement"):
        st.write("Productivité par statut × effectifs projetés = actes totaux. La part mobilisable limite l'offre au périmètre choisi. Personnes avec besoin × couverture cible × actes par personne = actes nécessaires. Les actes annuels ne sont pas cumulés entre années.")
        st.write("Les valeurs initiales de 8 actes, 25 % de couverture et 50 % d'offre mobilisable sont des hypothèses ajustables, pas un référentiel clinique validé. Les actes libéraux et salariés sont exprimés en équivalents de consultation ; l'hospitalisation et le travail en équipe ne sont pas directement comparables aux consultations facturées.")
        st.write("La population totale tous âges est requise. La population à risque des 18–79 ans ou des 6–11 ans n'est jamais utilisée comme substitut. Le cumul adulte dépression/anxiété conserve l'hypothèse d'absence de recouvrement du modèle de demande.")
        for label, url in MATCH_SOURCES.items():
            st.markdown(f"[{label}]({url})")


# ---------- Interface : navigation, page Scénarios, page Projections ----------
FAMILY_SUPPLY, FAMILY_DEMAND, FAMILY_MATCH = "Offre", "Demande", "Offre / demande"
DEMAND_MEASURES = {"count": "personnes avec besoin potentiel", "rate": "pour 1 000 habitants"}
SIDEBAR_WIDTH = 220  # Pixels.
# La carte occupe la hauteur de la fenêtre moins l'en-tête, les onglets et la ligne
# de sélecteurs : titre, carte et options restent visibles ensemble, sans défilement.
MAP_RESERVED_HEIGHT = 305  # Pixels hors carte, mesurés sur la page Projections.
MAP_MIN_HEIGHT = 340
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
.st-key-map_box {{height: calc(100vh - {MAP_RESERVED_HEIGHT}px) !important; min-height: {MAP_MIN_HEIGHT}px;
                  flex: 0 0 auto !important;}}
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
    places = [offre["formation"][f"region_{code}"] for code in REGIONS if offre["formation"].get(f"region_{code}")]
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
                if block == MATCH_BLOCK:
                    render_matching_editor(choices, labels, optional_match_inputs()[1] is not None)
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
    """Demande du scénario actif, rapprochée des codes départementaux métropolitains."""
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
            indicator = indicator_col.selectbox("Indicateur", options, key="map_indicator_match")
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
        st.caption("Rapprochement annuel selon les paramètres offre / demande du scénario actif. Solde négatif : manque d'actes ; positif : marge théorique pour le besoin choisi. La carte est centrée sur zéro pour le solde, sans supprimer les valeurs négatives.")
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
    return drees, year


def render_choropleth(selected, scale_values, value_label, indicator):
    try:
        gj = load_geojson()
    except (OSError, ValueError) as exc:
        st.warning(f"Carte indisponible ; les tableaux et les courbes restent accessibles : {exc}")
        return
    if selected.empty:
        return
    is_share, is_balance = indicator == SALARIED_SHARE, indicator == ACTS_BALANCE
    if is_balance:
        bound = max(float(scale_values.abs().quantile(0.98)), 0.001) if not scale_values.empty else 1.0
        vmin, vmax = -bound, bound
    else:
        vmin, vmax = (0.0, 100.0) if is_share else scale_values.quantile([0.02, 0.98]) if not scale_values.empty else (0.0, 1.0)
    if vmin == vmax:
        vmax = vmin + 1
    # Le premier bleu, presque blanc, se confondait avec le fond.
    fig = px.choropleth(selected, geojson=gj, locations="dep", featureidkey="properties.code", color="valeur",
                        color_continuous_scale="RdBu" if is_balance else px.colors.sequential.Blues[1:], range_color=(vmin, vmax),
                        hover_name="nom", custom_data=["dep"], labels={"valeur": value_label})
    unit = " %" if is_share else ""
    number_format = "%{z:+.3f} acte/habitant/an" if is_balance else "%{z:,.1f}" + unit
    fig.update_traces(marker_line_color="white", marker_line_width=0.6,
                      hovertemplate="<b>%{hovertext}</b> (%{customdata[0]})<br>" + number_format + "<extra></extra>")
    # Projection conique conforme de Lambert, comme le Lambert 93 officiel :
    # la France garde ses proportions, au lieu d'être étirée en largeur.
    fig.update_geos(projection=dict(type="conic conformal", parallels=[44, 49], rotation=dict(lon=3)),
                    center=dict(lon=2.3, lat=46.3), lonaxis_range=[-5.3, 9.8], lataxis_range=[41.2, 51.2],
                    visible=False, bgcolor="rgba(0,0,0,0)")
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), dragmode=False,
                      paper_bgcolor="rgba(0,0,0,0)", separators=", ",
                      coloraxis_colorbar=dict(title=dict(text=value_label, side="right"),
                                              thickness=14, len=0.7, outlinewidth=0))
    st.plotly_chart(fig, width="stretch", height="stretch")


def render_projections_page(app, scenarios_page):
    reference, metadata, context = app["reference"], app["metadata"], app["context"]
    signature, names, labels = app["signature"], app["names"], app["labels"]
    scenario = complete_matching(st.session_state.scenario_active, app["match_defaults"])
    proposed = scenario["offre"]
    # En-tête sur une ligne : titre, scénario actif (détail au clic), lien vers Scénarios.
    title, detail, change = st.columns([3.3, 1.6, 1.25], vertical_alignment="center")
    title.markdown("### Projections des effectifs de psychiatres par département")
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
    mode_text = (f"croissance nationale DREES raccordée aux effectifs RPPS de {anchor} (psychiatrie absente des projections régionales)"
                 if metadata["mode_projection"] == "national" else "projections régionales détaillées de psychiatrie")
    with detail.popover(f"Scénario : {scenario['nom']}", icon=":material/info:", width="stretch",
                        help="Paramètres du scénario actif"):
        st.markdown(f"**{scenario['nom']}** · :gray[{scenario['origine']}]")
        st.markdown("  \n".join(scenario_highlights(scenario, context, labels)))
        st.caption(f"Observé : {df.annee.min()}–{anchor}. Projeté : {anchor + 1}–{df.annee.max()}. Référence : {mode_text}.")
    change.page_link(scenarios_page, label="Changer de scénario", icon=":material/tune:")
    demand, demand_error = app["demand"], app["demand_error"]
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
        st.write(f"Les courbes montrent un écart en nombre de psychiatres. Chaque effectif est comparé au total national après correction de l’internat, avant tendance étrangère et hausses des places ou des arrivées, multiplié par la part du département observée en {anchor}. Les courbes incluent tous les effets du scénario actif, ouvertures de places comprises.")
        show_fit = st.checkbox("Afficher la droite estimée sur les observations", value=False, key="show_trend_fit")
        departments = st.multiselect("Départements à comparer", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]],
                                     format_func=lambda dep: f"{dep} · {names[dep]}", key="trend_departments")
        if departments:
            diagnostic = dynamics_diagnostic(state["panel"], state["contributions"], context, scenario_drees)
            st.plotly_chart(dynamics_figure(diagnostic, context, departments, True, start, show_fit), width="stretch")
            if any(proposed["formation"][f"region_{code}"] for code in REGIONS):
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
                st.dataframe(detail[["dep", "nom", "effectif_reference", "effet_dynamique", "effet_internat", "tendance_etranger", "ajout_immigration", "ajout_formation", "effectif"]].rename(
                    columns={"effectif_reference": "Référence sans modulation", "effet_dynamique": "Effet départemental", "effet_internat": "Correction internat",
                             "tendance_etranger": "Tendance étrangère / 2026", "ajout_immigration": "Hausse du curseur étranger", "ajout_formation": "Ajout des places", "effectif": "Effectif final"}), hide_index=True, width="stretch")
                st.caption("L'effet départemental inclut la saturation et la renormalisation. L'ajout des places est un nombre de psychiatres supplémentaires, toujours positif ou nul.")
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
        st.caption(f"Projection linéaire : {100 * calibration['pente_annuelle']:+.3f} point de pourcentage par an, avec raccord au dernier taux observé. Les taux sont limités à l’intervalle 0–100 %.")
        st.plotly_chart(vacancy_figure(state["vacancy_curve"], calibration), width="stretch")
        if state["vacancy_curve"].borne_atteinte.any():
            st.warning("La droite extrapolée sort de l’intervalle 0–100 % à certains horizons. Les taux y sont maintenus à la borne correspondante.")
        st.caption(f"Coefficient de flux = min(1, (1 − taux futur) / (1 − taux de référence)). Si la tendance estimée s’améliore, aucune réduction supplémentaire n’est appliquée. Les cohortes entrées au plus tard en {anchor} conservent un coefficient de 1.")
        flows = state["national_flows"]
        shown = flows[flows.scenario == scenario_drees]
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
    with tab_demand:
        render_demand_tab(demand, demand_error)
    with tab_match:
        st.caption(shown_choice)
        render_match_tab(matching, matching_error, scenario_drees, year, names)
    with tab_method:
        st.caption(shown_choice)
        first = proposed["formation"]["rentree"] + proposed["formation"]["delai"]
        st.write(f"Places en deuxième année : hausse annuelle permanente à partir de {proposed['formation']['rentree']}, avec premiers effets en {first}. Délai retenu : {proposed['formation']['delai']} ans. Le défaut de dix ans correspond à cinq ans avant le DES et cinq ans de DES.")
        st.write("Pour chaque lieu de diplôme, le volume de référence des politiques publiques vient des installations historiques nationales, extrapolées directement. Une hausse de places de +u % ajoute u % de ce flux à chaque cohorte après le délai prévu. Le coefficient de postes non pourvus à l’entrée en internat, cinq ans plus tôt, est appliqué une seule fois aux cohortes françaises si la correction est activée.")
        st.write("Chaque cohorte supplémentaire est répartie selon la matrice diplôme → exercice prédite pour son année d’installation, puis selon les parts départementales observées dans chaque région d'exercice. Les destinations hors métropole restent au dénominateur. Les cohortes sont cumulées sans attrition ni retraite supplémentaire jusqu'en 2070 ; les cohortes déjà arrivées ne sont pas redistribuées à chaque nouvelle année.")
        st.caption("Le tableau donne l’effet de +1 % pour la première cohorte supplémentaire ; les cohortes suivantes s’accumulent. Le volume par origine ne dépend pas du scénario DREES ; X, exprimé relativement au total, en dépend. Les stocks d’internes DREES restent utilisés uniquement pour la correction nationale de la trajectoire démographique.")
        st.dataframe(calibration_table(reference, context, scenario_drees, proposed, state["national_flows"], state["intervention_routing"], state["training_cohorts"]), hide_index=True, width="stretch")
        missing = [REGIONS[code] for code in REGIONS if code not in context["available"]]
        if missing:
            st.warning("Calibration de formation absente pour : " + ", ".join(missing) + ". Une ouverture ex nihilo nécessite des hypothèses supplémentaires ; le curseur reste désactivé. Ces régions peuvent recevoir des diplômés formés ailleurs.")
        st.write(f"Dynamiques : régression des écarts d'effectif départemental à la référence nationale sur {proposed['dynamique']['debut']}–{anchor}. La référence vaut le total métropolitain de chaque année multiplié par la part d'effectif du département observée en {anchor}. Chaque pente est exprimée en psychiatres par an.")
        st.write(f"Projection avant saturation = total national de référence × part départementale de {anchor} + pente départementale × années depuis {anchor}. Les pentes somment à zéro. L'effet commence dès {anchor + 1}. Le lissage sous 1 psychiatre et la renormalisation préservent des effectifs non négatifs et le total national.")
        if metadata["mode_projection"] == "regional":
            st.write(f"Lorsque les dynamiques sont activées, le total des projections régionales sert de trajectoire nationale : leur ventilation régionale est remplacée par les parts {anchor} et les écarts départementaux. Désactiver les dynamiques rétablit la ventilation régionale de référence.")
        st.write(f"Installations selon le lieu de diplôme : {proposed['mobilite']['methode']} ; départs à la retraite {proposed['mobilite']['retraite']:.1f} % par an ; estimation à partir de {proposed['mobilite']['debut']}. Le code 99 désigne le lieu du diplôme (étranger), pas la nationalité ; sa tendance démographique reste active à 0 % de hausse.")
        with st.expander("Sources et champ des effectifs"):
            st.write(f"Onglet de projection des actifs : {metadata['onglet_drees']}.")
            if metadata["mode_projection"] == "national":
                st.write(f"Référence départementale = effectif RPPS en {anchor} × effectif national DREES futur / effectif national DREES en {anchor}. La croissance du champ DREES (70 ans ou moins) est appliquée au RPPS tous âges ; les plus de 70 ans ne font pas l'objet d'une projection distincte.")
            else:
                st.write("Les effectifs régionaux DREES de référence sont utilisés directement. Le RPPS comprend tous les âges, contre 70 ans ou moins pour les actifs DREES : une rupture de champ peut apparaître au raccord.")
            st.write("Les résultats sont des nombres de psychiatres et une part salariée en %. Les observations proviennent directement de l’onglet Effectifs. Les valeurs projetées fractionnaires représentent des effectifs attendus.")
            st.write("Aucun fichier de projections n'est écrit. Le calcul de référence est partagé en mémoire ; les paramètres et résultats simulés appartiennent à chaque session.")
        st.caption("Les formules, limites et règles d'extension sont détaillées dans README.md.")


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
