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
from demand_projection.outcomes import prepare_outcomes

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
    training = flows[["annee", "scenario", "flux_fr_corrige"]].merge(context["weights"], on="scenario", validate="many_to_many")
    training["flux_reference"] = training.flux_fr_corrige * training.poids
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    increases = training.origine.map(lambda code: parameters["formation"].get(f"region_{code}", 0) / 100.0)
    training["flux_ajoute"] = np.where(training.annee >= first, training.flux_reference * increases, 0.0)
    panel = context["mobility_panel"]
    foreign = foreign_reference_flows(panel, parameters["mobilite"]["retraite"] / 100.0, context["anchor"])
    foreign = foreign.merge(pd.DataFrame({"scenario": sorted(reference.scenario.unique())}), how="cross")
    foreign["flux_hausse"] = np.where(foreign.annee >= parameters["immigration"]["debut"],
        foreign.flux_reference * parameters["immigration"]["hausse"] / 100.0, 0.0)
    foreign["flux_ajoute"] = foreign.ecart_flux + foreign.flux_hausse
    routing = annual_routing(foreign_france_distribution(panel), context["baseline"])
    return {"training_cohorts": training, "foreign_cohorts": foreign,
            "training_arrivals": accumulate_arrivals(reference, training, context["routing"]),
            "foreign_routing": routing,
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
    available = set(origin_totals[origin_totals > 0].index) & set(weights.loc[weights.poids > 0, "origine"]) & set(REGIONS)
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
                if code in context["available"] else "Pas de calibration régionale disponible dans les deux sources ; curseur désactivé.")
        for code, label in sorted(REGIONS.items(), key=lambda item: item[1])
    )
    return (
        Modulation("mobilite", "Installations selon le lieu de diplôme",
                   "Extrapolation annuelle des effectifs par diplôme et région d'exercice. Le taux de retraite sert à reconstituer les flux ; il ne réduit pas une seconde fois la référence DREES.",
                   (Control("methode", "Méthode de répartition des installations", "select", FLOW_ALLOCATION,
                            options=(FLOW_ALLOCATION, STOCK_ALLOCATION)),
                    Control("retraite", "Taux annuel uniforme de départ à la retraite (%)", "number", 3.0, 0.0, 20.0, 0.1,
                            help="Hypothèse réglable, non estimée : installations = max(0, effectif courant − (1 − taux) × effectif précédent)."),
                    Control("debut", "Première année pour les dynamiques diplôme/exercice", "number",
                            context["mobility_start"], context["mobility_start"], context["anchor"] - 4, 1)), calibration_only),
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
                   "L'écart annuel au flux de 2026 est cumulé dès 2027, même à 0 %. Le curseur ajoute une hausse permanente de ce flux, sans délai de formation. Le code 99 désigne le lieu du diplôme.",
                   (Control("hausse", "Hausse des arrivées de diplômés à l'étranger (%)", "slider", 0, 0, 100, 1),
                    Control("debut", "Première année de hausse des arrivées étrangères", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()), 1)), immigration_additions),
        Modulation("formation", "Places de formation par région",
                   "Hausse des places en deuxième année, reconduite pour chaque cohorte. Répartition prévue pour l'année d'installation.",
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
            "routing": run_context["routing"], "foreign_trend_limited_rows": limited, **arrival_details}
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


def render_parameters(applied, registry):
    draft = {}
    for modulation in registry:
        if not modulation.controls:
            draft[modulation.key] = {}
            continue
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


def calibration_table(reference, context, scenario, parameters, flows, routing):
    first = parameters["formation"]["rentree"] + parameters["formation"]["delai"]
    row = flows[(flows.scenario == scenario) & (flows.annee == first)].iloc[0]
    flow = float(row.flux_fr_corrige)
    national = float(row.total_reference * row.coefficient_stock)
    weights = context["weights"]
    weights = weights[weights.scenario == scenario].set_index("origine").poids
    retention = routing[routing.annee == first].groupby("origine").part.sum()
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
    figure.add_trace(go.Scatter(x=future.annee, y=100 * future.taux_projete, mode="lines", name="Extrapolation linéaire", line_dash="dash"))
    figure.add_hline(y=100 * calibration["taux_reference"], line_dash="dot", annotation_text="Taux de référence constant")
    figure.update_layout(height=420, xaxis_title="Année d'entrée en internat", yaxis_title="Postes non pourvus (%)",
                         yaxis_range=[0, 100], hovermode="x unified")
    return figure


@st.cache_data(show_spinner=False, max_entries=2)
def load_demand_summary(path: str, modified_at_ns: int):
    """Read generated outcomes; the timestamp invalidates the cache."""
    del modified_at_ns
    return prepare_outcomes(pd.read_csv(path))


@st.cache_data(show_spinner=False, max_entries=2)
def compute_demand_summary(signature):
    """Support fresh checkouts without committing generated CSV files."""
    del signature
    from demand_projection.run import build_projection, summarize_projection

    return prepare_outcomes(summarize_projection(build_projection(ROOT / "dataDemand", None, 2026)))


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
            return load_demand_summary(str(DEMAND_SUMMARY_PATH), modified)
    return compute_demand_summary(signature)


def render_demand_tab():
    """Display demand outcomes with filters independent of the supply scenario."""
    st.write("Projection démographique-épidémiologique indépendante de la projection des effectifs de psychiatres.")
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
        st.caption(f"Paramètres validés : dynamiques des effectifs actives dès {anchor + 1}, estimées sur {proposed['dynamique']['debut']}–{anchor}. Le total national est conservé à cette étape, puis la correction de l’internat et les hausses d'arrivées et de places sont appliquées.")
    else:
        st.warning(f"Paramètres validés : dynamiques départementales désactivées. Leur pente n'est pas appliquée. Pour les activer dès {anchor + 1}, cocher la case dans le formulaire puis cliquer sur Valider.")
    state = st.session_state.psy_projection
    st.caption(f"Répartition validée : {proposed['mobilite']['methode']} ; retraites {proposed['mobilite']['retraite']:.1f} % par an. Places en deuxième année : premières installations en {proposed['formation']['rentree'] + proposed['formation']['delai']}. Code 99 : tendance nationale active dès {anchor + 1}, plus hausse de {proposed['immigration']['hausse']} % dès {proposed['immigration']['debut']}.")
    if proposed["internat"]["active"]:
        st.caption(f"Correction nationale des cohortes activée : premiers effets possibles sur les effectifs en {anchor + 1 + DES_YEARS}. Les observations restent inchangées.")
    else:
        st.caption("Correction nationale des postes non pourvus désactivée.")
    with st.sidebar:
        indicator = st.selectbox("Indicateur", list(INDICATORS.values()))
        scenario = st.selectbox("Scénario", sorted(df.scenario.unique()))
        year = st.slider("Année", int(df.annee.min()), int(df.annee.max()), anchor)
    tab_map, tab_trends, tab_status, tab_internat, tab_mobility, tab_demand, tab_method = st.tabs(["Carte et effectifs", "Dynamiques départementales", "Statuts d’exercice", "Internat national", "Diplômes et installations", "Demande de soins", "Calibration et hypothèses"])
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
            # Le premier bleu, presque blanc, se confondait avec le fond.
            fig = px.choropleth(selected, geojson=gj, locations="dep", featureidkey="properties.code", color="valeur",
                                color_continuous_scale=px.colors.sequential.Blues[1:], range_color=(vmin, vmax),
                                hover_name="nom", custom_data=["dep"], labels={"valeur": value_label})
            unit = " %" if is_share else ""
            fig.update_traces(marker_line_color="white", marker_line_width=0.6,
                              hovertemplate="<b>%{hovertext}</b> (%{customdata[0]})<br>%{z:,.1f}" + unit + "<extra></extra>")
            # Projection conique conforme de Lambert, comme le Lambert 93 officiel :
            # la France garde ses proportions, au lieu d'être étirée en largeur.
            fig.update_geos(projection=dict(type="conic conformal", parallels=[44, 49], rotation=dict(lon=3)),
                            center=dict(lon=2.3, lat=46.3), lonaxis_range=[-5.3, 9.8], lataxis_range=[41.2, 51.2],
                            visible=False, bgcolor="rgba(0,0,0,0)")
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=650, dragmode=False,
                              paper_bgcolor="rgba(0,0,0,0)", separators=", ",
                              coloraxis_colorbar=dict(title=dict(text=value_label, side="right"),
                                                      thickness=12, len=0.6, outlinewidth=0))
            map_col.plotly_chart(fig, width="stretch")
        table_col.dataframe(selected.sort_values("valeur")[["dep", "nom", "valeur"]].rename(columns={"valeur": value_label}), hide_index=True, width="stretch")
        st.caption("Salariés = salariés hospitaliers + autres salariés. Libéraux = libéraux exclusifs + exercice mixte. Chaque personne est comptée une seule fois ; la part salariée est en %.")
    with tab_trends:
        active = proposed["dynamique"]["active"]
        start = proposed["dynamique"]["debut"]
        if active:
            st.info(f"Dynamique réellement appliquée : OUI, dès {anchor + 1}. Pente d'effectif estimée sur {start}–{anchor} puis raccordée à l'observation {anchor}.")
        else:
            st.warning("Dynamique réellement appliquée : NON. La tendance nationale des arrivées étrangères reste active, même avec le curseur à 0 %. Elle peut produire des écarts territoriaux. Des projections régionales détaillées peuvent aussi avoir leur propre évolution.")
        st.write(f"Les courbes montrent un écart en nombre de psychiatres. Chaque effectif est comparé au total national après correction de l’internat, avant tendance étrangère et hausses des places ou des arrivées, multiplié par la part du département observée en {anchor}. Cette référence est commune aux deux vues.")
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
                    st.caption(f"Ouvertures validées : premiers effets en {proposed['formation']['rentree'] + proposed['formation']['delai']}. Ce délai concerne les places de formation ; la hausse des arrivées étrangères intervient sans délai.")
                else:
                    st.caption("Aucune hausse de places validée : les deux vues coïncident.")
            else:
                st.caption("Cette vue exclut les ajouts de places et inclut la tendance des arrivées étrangères ainsi que la hausse validée par le curseur. La carte conserve le scénario complet validé.")
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
        st.caption(f"Projection linéaire : {100 * calibration['pente_annuelle']:+.3f} point de pourcentage par an, avec raccord au dernier taux observé. Les taux sont limités à l’intervalle 0–100 %.")
        st.plotly_chart(vacancy_figure(state["vacancy_curve"], calibration), width="stretch")
        if state["vacancy_curve"].borne_atteinte.any():
            st.warning("La droite extrapolée sort de l’intervalle 0–100 % à certains horizons. Les taux y sont maintenus à la borne correspondante.")
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
    with tab_mobility:
        st.write("Les comptes annuels de l’onglet Effectifs par lieu diplome sont extrapolés séparément pour chaque origine de diplôme et région d’exercice, avec une pente linéaire raccordée à 2026 et un plancher à zéro.")
        st.write("Installations estimées = max(0, effectif de l’année − (1 − taux de retraite) × effectif précédent). Leur répartition est normalisée pour chaque origine et année. L’autre méthode utilise directement les parts des effectifs projetés, sans delta pour la répartition.")
        st.caption("Le taux de retraite sert à cette reconstitution et au volume du flux étranger ; il n’est pas déduit une seconde fois du stock DREES. Les stocks peuvent refléter des mobilités de carrière et d'autres sorties : les installations ne sont pas directement observées.")
        available_origins = sorted(state["mobility_panel"].origine.unique())
        origin_labels = {**REGIONS, "01": "Antilles-Guyane", "98": "Région de diplôme inconnue", "99": "Diplôme obtenu à l’étranger"}
        origin = st.selectbox("Lieu de diplôme à examiner", available_origins,
                              index=available_origins.index("99"),
                              format_func=lambda code: f"{code} · {origin_labels.get(code, code)}", key="mobility_origin")
        inspection_year = st.slider("Année de la répartition projetée", anchor + 1, int(df.annee.max()),
                                     max(anchor + 1, year), key="mobility_year")
        panel = state["mobility_panel"]
        inspection = foreign_france_distribution(panel) if origin == "99" else panel[panel.origine == origin]
        detail = inspection[inspection.annee == inspection_year].copy()
        destination_labels = {**REGIONS, "01": "Guadeloupe", "02": "Martinique", "03": "Guyane", "04": "La Réunion", "06": "Mayotte", "99": "COM ou étranger"}
        detail["Région d’exercice"] = detail.destination.map(destination_labels).fillna(detail.destination)
        detail["Part (%)"] = 100 * detail.part
        st.dataframe(detail[["destination", "Région d’exercice", "precedent", "effectif", "installations", "Part (%)", "repli"]].rename(
            columns={"destination": "Code", "precedent": "Effectif précédent", "effectif": "Effectif projeté", "installations": "Installations estimées", "repli": "Règle de repli"}), hide_index=True, width="stretch")
        distribution = inspection.copy()
        distribution["Part (%)"] = 100 * distribution.part
        distribution["Région d’exercice"] = distribution.destination.map(destination_labels).fillna(distribution.destination)
        figure = px.line(distribution, x="annee", y="Part (%)", color="Région d’exercice", labels={"annee": "Année d’installation"})
        figure.update_layout(height=450, yaxis_range=[0, 100])
        st.plotly_chart(figure, width="stretch")
        if (detail.repli != "Aucun").any():
            st.info("Le flux estimé est nul ou le stock projeté est épuisé pour cette origine : la répartition utilise les effectifs projetés, puis si nécessaire les parts de 2026.")
        st.caption("Les proportions incluent les destinations hors métropole. Pour le code 99, elles sont conditionnées à une installation en France, DROM inclus, en excluant la destination COM ou étranger. Les parts départementales au sein de chaque région restent celles de 2026. Le code 98 n’est pas assimilé au code 99.")
        st.write("Code 99 : les effectifs de diplômés à l’étranger exerçant en France sont d’abord agrégés. Le flux annuel vaut max(0, total courant − (1 − retraite) × total précédent). Le max s’applique au total national, et non à chaque région séparément.")
        st.write(f"La tendance ajoute chaque année l’écart au flux estimé de {anchor}, puis cumule ces écarts à partir de {anchor + 1}. L’ajustement vaut zéro en {anchor}. Il reste actif à 0 % de hausse ; le curseur ajoute séparément un pourcentage du flux annuel, sans délai ni correction des postes vacants français.")
        st.dataframe(state["foreign_cohorts"].query("scenario == @scenario")[["annee", "effectif", "flux_reference", "flux_ancrage", "ecart_flux", "flux_hausse", "flux_ajoute"]].rename(
            columns={"annee": "Année d’arrivée", "effectif": "Diplômés étrangers exerçant en France", "flux_reference": "Flux estimé France", "flux_ancrage": f"Flux de référence {anchor}", "ecart_flux": "Écart annuel au flux de référence", "flux_hausse": "Hausse annuelle du curseur", "flux_ajoute": "Ajustement annuel total France"}), hide_index=True, width="stretch")
        st.caption("Un écart négatif signifie moins d'arrivées que dans le scénario à flux constant de référence, pas un nombre négatif d'installations. La référence 2026 est une convention de scénario supplémentaire ; le flux étranger exactement intégré dans la DREES n'est pas isolé ici.")
        if state["foreign_trend_limited_rows"]:
            st.warning("L’écart cumulé aurait rendu certains effectifs départementaux négatifs. Leur correction a été limitée à un effectif nul ; la décomposition montre le retrait effectivement appliqué.")
    with tab_demand:
        render_demand_tab()
    with tab_method:
        first = proposed["formation"]["rentree"] + proposed["formation"]["delai"]
        st.write(f"Places en deuxième année : hausse annuelle permanente à partir de {proposed['formation']['rentree']}, avec premiers effets en {first}. Délai validé : {proposed['formation']['delai']} ans. Le défaut de dix ans correspond à cinq ans avant le DES et cinq ans de DES.")
        st.write("Le stock national d'internes de psychiatrie de chaque année de sortie est divisé par 5 pour approcher un flux annuel. La géographie des places est approchée par les stocks régionaux d'internes toutes spécialités de la dernière année observée. Le flux est ensuite corrigé par le coefficient des postes non pourvus à l’entrée en internat, cinq ans plus tôt, si cette modulation est active.")
        st.write("Chaque cohorte supplémentaire est répartie selon la matrice diplôme → exercice prédite pour son année d’installation, puis selon les parts départementales observées dans chaque région d'exercice. Les destinations hors métropole restent au dénominateur. Les cohortes sont cumulées sans attrition ni retraite supplémentaire jusqu'en 2070 ; les cohortes déjà arrivées ne sont pas redistribuées à chaque nouvelle année.")
        st.caption("X dépend de la région, du scénario et de l'année. Le tableau tient compte de la correction de l’internat et donne l’effet de +1 % dans une seule région pour la première cohorte diplômée ; ensuite les cohortes s'accumulent. Les nombres fractionnaires sont des effectifs attendus.")
        st.dataframe(calibration_table(reference, context, scenario, proposed, state["national_flows"], state["routing"]), hide_index=True, width="stretch")
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
