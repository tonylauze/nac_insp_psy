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

Densités futures : population.csv facultatif (dep,annee,population). Sans lui,
la population départementale de la dernière année RPPS est maintenue constante.
Les densités standardisées DREES ne sont pas assimilées à des densités brutes.

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
POPULATION_PATH = ROOT / "population.csv"  # Facultatif.
GEOJSON_URL = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements-version-simplifiee.geojson"
COLUMNS = ["dep", "annee", "scenario", "indicateur", "valeur"]
COUNT = "Effectif de psychiatres"
DENSITY = "Psychiatres pour 100 000 hab."
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
    """Extraire les lignes agrégées, sans sommer hommes/femmes ou classes d'âge."""
    require_excel_reader()
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    frames, names = [], []
    try:
        for sheet_name, prefix, metric in [("Effectifs", "effectif_", "effectif"), ("Densités", "densite_", "densite")]:
            rows = book[sheet_name].iter_rows(values_only=True)
            pos = {v: i for i, v in enumerate(next(rows)) if v is not None}
            records = []
            for row in rows:
                if any(row[pos[k]] != v for k, v in [("exercice", "0-Ensemble"), ("tranche_age", "00-Ensemble"), ("sexe", "0-Ensemble")]):
                    continue
                if normalize(row[pos["territoire"]]) != "francemetropolitaine":
                    continue
                # Exemples RPPS : « 075 -Paris », « 001 -Ain », « 02A -Corse-du-Sud ».
                raw_dep = str(row[pos["departement"]]).split("-", 1)[0].strip()
                dep = raw_dep[-2:] if len(raw_dep) == 3 and raw_dep.startswith("0") else raw_dep
                if dep not in METRO:
                    continue
                spec = normalize(row[pos["specialites"]])
                kind = "psy" if spec == "psychiatrie" else "tous"
                if kind == "tous" and not (spec == "ensemble" and normalize(row[pos["specialites_agregees"]]) == "ensemble"):
                    continue
                if sheet_name == "Effectifs" and kind == "psy":
                    names.append({"dep": dep, "nom": str(row[pos["departement"]]).split("-", 1)[1].strip(),
                                  "region_code": str(row[pos["region"]]).split("-", 1)[0].strip().zfill(2),
                                  "region_rpps": normalize(row[pos["region"]])})
                for column, index in pos.items():
                    if re.fullmatch(prefix + r"\d{4}", column):
                        value = row[index]
                        if not isinstance(value, (int, float)) or not np.isfinite(value) or value < 0:
                            raise ValueError(f"Donnée RPPS indisponible : {dep}, {column}, {value!r}.")
                        records.append({"dep": dep, "annee": int(column[-4:]), "kind": kind, metric: float(value)})
            frame = pd.DataFrame(records)
            if frame.empty or frame.duplicated(["dep", "annee", "kind"]).any():
                raise ValueError(f"Données absentes ou dupliquées : {sheet_name}.")
            frames.append(frame)
    finally:
        book.close()
    merged = frames[0].merge(frames[1], on=["dep", "annee", "kind"], how="outer", validate="one_to_one")
    if merged.isna().any().any():
        raise ValueError("Les effectifs et les densités RPPS ne couvrent pas les mêmes cellules.")
    all_doctors = merged[merged.kind == "tous"].copy()
    if (all_doctors.densite <= 0).any():
        raise ValueError("Population non reconstituable depuis les densités RPPS.")
    all_doctors["population"] = all_doctors.effectif / all_doctors.densite * 100_000
    history = merged[merged.kind == "psy"].drop(columns="kind").merge(
        all_doctors[["dep", "annee", "population"]], on=["dep", "annee"], validate="one_to_one")
    years = sorted(history.annee.unique())
    if not years or years != list(range(min(years), max(years) + 1)) or len(history) != 96 * len(years):
        raise ValueError("Historique départemental incomplet.")
    return history, pd.DataFrame(names)


def assemble_reference(history, names, regional, population=None):
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
    future = future_reg.merge(baseline[["dep", "region_source", "part_departement", "population"]], on="region_source", validate="many_to_many")
    # Le tableau régional peut provenir des projections détaillées DREES ou du
    # repli national_to_regions. Chaque département reçoit seulement sa part.
    future["effectif"] = future.effectif_region * future.part_departement
    if population is not None:
        if list(population.columns) != ["dep", "annee", "population"] or population.duplicated(["dep", "annee"]).any():
            raise ValueError("population.csv doit contenir dep,annee,population, avec des clés uniques.")
        future = future.drop(columns="population").merge(population, on=["dep", "annee"], how="left", validate="many_to_one")
    if future.population.isna().any() or not np.isfinite(future.population).all() or (future.population <= 0).any():
        raise ValueError("Population future manquante ou invalide.")
    future["densite"] = future.effectif / future.population * 100_000
    future["projete"] = True
    observed = history.merge(names[["dep", "region_source"]], on="dep", validate="many_to_one")
    observed = observed.merge(future_reg[["scenario"]].drop_duplicates(), how="cross")
    observed["projete"] = False
    columns = ["dep", "region_source", "annee", "scenario", "effectif", "densite", "population", "projete"]
    return pd.concat([observed[columns], future[columns]], ignore_index=True), anchor


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


def build_context(history, names, stocks, weights, mobility):
    anchor = int(history.annee.max())
    baseline = history[history.annee == anchor].merge(names, on="dep", validate="one_to_one")
    baseline["part_nationale"] = baseline.effectif / baseline.effectif.sum()
    baseline["part_regionale"] = baseline.effectif / baseline.groupby("region_code").effectif.transform("sum")
    h = history.merge(baseline[["dep", "part_nationale"]], on="dep", validate="many_to_one")
    h["ecart"] = h.effectif - h.part_nationale * h.groupby("annee").effectif.transform("sum")
    x = h.annee - h.annee.mean()
    h["produit"], h["carre"] = x * h.ecart, x * x
    sums = h.groupby("dep")[["produit", "carre"]].sum()
    if (sums.carre <= 0).any():
        raise ValueError("Au moins deux années observées sont nécessaires pour estimer les dynamiques.")
    slopes = sums.produit / sums.carre
    slopes -= slopes.mean()  # Somme nulle, y compris après arrondis flottants.
    baseline["pente"] = baseline.dep.map(slopes)
    routing = mobility.merge(baseline[["dep", "region_code", "part_regionale"]],
                             left_on="destination", right_on="region_code", validate="many_to_many")
    routing["part"] *= routing.part_regionale
    available = set(mobility.origine) & set(weights.loc[weights.poids > 0, "origine"])
    return {"anchor": anchor, "baseline": baseline, "stocks": stocks,
            "weights": weights, "mobility": mobility, "routing": routing,
            "available": available, "observed_start": int(history.annee.min())}


def department_dynamics(projected, parameters, context):
    if not parameters["active"]:
        return projected
    result = projected.copy()
    base = context["baseline"].set_index("dep")
    total = result.groupby(["annee", "scenario"]).effectif.transform("sum")
    raw = total * result.dep.map(base.part_nationale) + result.dep.map(base.pente) * (result.annee - context["anchor"])
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
    stocks["flux_psy"] = np.where(stocks.annee >= first_graduation, stocks.psychiatrie / DES_YEARS, 0.0)
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

    La fonction apply(projected, parameters, context) reçoit des copies. Seuls
    effectif et population sont modifiables ; le moteur recalcule la densité.
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
                   "Extrapolation des écarts historiques à la tendance nationale, avec saturation près de zéro.",
                   (Control("active", "Extrapoler les dynamiques départementales", "checkbox", True),), department_dynamics),
        Modulation("formation", "Places de formation par région",
                   "Première rentrée au niveau du bac ; hausse reconduite pour chaque cohorte. Premiers psychiatres 11 ans plus tard.",
                   (Control("rentree", "Première rentrée augmentée (année du bac)", "number", context["anchor"] + 1,
                            context["anchor"] + 1, int(context["stocks"].annee.max()) - STUDY_YEARS, 1),) + regional_controls,
                   training_additions),
    )


def default_parameters(registry):
    return {m.key: {c.key: c.default for c in m.controls} for m in registry}


def apply_modulations(reference, parameters, context):
    result = reference.copy(deep=True)
    projected = result[result.projete].copy()
    immutable = [c for c in projected if c not in {"effectif", "population"}]
    for modulation in modulation_registry(context):
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
        changed = modulation.apply(projected.copy(deep=True), dict(values), deepcopy(context))
        if not isinstance(changed, pd.DataFrame) or list(changed.columns) != list(projected.columns) or not changed.index.equals(projected.index):
            raise ValueError(f"{modulation.key} a modifié la structure des données.")
        if not changed[immutable].equals(projected[immutable]):
            raise ValueError(f"{modulation.key} a modifié une clé ou une colonne réservée.")
        if not np.isfinite(changed[["effectif", "population"]]).all().all() or (changed.effectif < 0).any() or (changed.population <= 0).any():
            raise ValueError(f"Résultat invalide : {modulation.key}.")
        projected = changed
    projected["densite"] = projected.effectif / projected.population * 100_000
    result.loc[projected.index, ["effectif", "population", "densite"]] = projected[["effectif", "population", "densite"]]
    return result


def to_projections_frame(panel):
    result = panel.melt(id_vars=["dep", "annee", "scenario"], value_vars=["effectif", "densite"], var_name="indicateur", value_name="valeur")
    result.indicateur = result.indicateur.map({"effectif": COUNT, "densite": DENSITY})
    result = result[COLUMNS].sort_values(COLUMNS[:-1]).reset_index(drop=True)
    if result.empty or result.isna().any().any() or result.duplicated(COLUMNS[:-1]).any():
        raise ValueError("Table de projections vide, incomplète ou dupliquée.")
    if not np.isfinite(result.valeur).all() or (result.valeur < 0).any():
        raise ValueError("Valeur projetée invalide.")
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
    if POPULATION_PATH.exists():
        paths.append(POPULATION_PATH)
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
    population = pd.read_csv(POPULATION_PATH, dtype={"dep": str}) if POPULATION_PATH.exists() else None
    reference, anchor = assemble_reference(history, names, regional, population)
    stocks, weights = training_students(DREES_PATH, anchor, names)
    mobility = diploma_mobility(RPPS_PATH, anchor)
    context = build_context(history, names, stocks, weights, mobility)
    expected = pd.MultiIndex.from_frame(reference[["annee", "scenario"]].drop_duplicates())
    future = expected[expected.get_level_values("annee") >= anchor]
    actual = pd.MultiIndex.from_frame(stocks[["annee", "scenario"]])
    if not future.isin(actual).all() or set(reference.scenario) != set(weights.scenario):
        raise ValueError("Calibration des étudiants incomplète sur les années/scénarios projetés.")
    return reference, {"annee_reference": anchor, "onglet_drees": sheet,
                       "population_fournie": population is not None,
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
    panel = apply_modulations(reference, parameters, context)
    output = to_projections_frame(panel)
    # Remplacement de l'état seulement lorsque le calcul a entièrement réussi.
    st.session_state.psy_projection = {
        "key": key, "parameters": deepcopy(parameters),
        "metadata": deepcopy(metadata), "projections": output, "panel": panel,
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


def dynamics_figure(panel, context, scenario, departments):
    """Écart à la part 2026 du total métropolitain de l'année, après modulations."""
    selected = panel[panel.scenario == scenario].copy()
    base = context["baseline"].set_index("dep")
    selected["ecart"] = selected.effectif - selected.dep.map(base.part_nationale) * selected.groupby("annee").effectif.transform("sum")
    figure = go.Figure()
    colors = px.colors.qualitative.Dark24
    for index, dep in enumerate(departments):
        series = selected[selected.dep == dep].sort_values("annee")
        label = f"{dep} · {base.loc[dep, 'nom']}"
        for projected, label_period, dash in ((False, "observé", "solid"), (True, "projeté", "dash")):
            part = series[series.annee >= context["anchor"]] if projected else series[~series.projete]
            figure.add_trace(go.Scatter(x=part.annee, y=part.ecart, mode="lines",
                                       name=f"{label} — {label_period}", legendgroup=dep,
                                       line={"color": colors[index % len(colors)], "dash": dash},
                                       hovertemplate="%{x} : %{y:.2f} psychiatres<extra>%{fullData.name}</extra>"))
    figure.add_hline(y=0, line_color="#64748b", line_width=1,
                    annotation_text=f"Même trajectoire que le total national, à part {context['anchor']} constante")
    figure.add_vline(x=context["anchor"], line_color="#94a3b8", line_dash="dot")
    figure.update_layout(height=570, xaxis_title="Année", yaxis_title="Écart en nombre de psychiatres à la tendance nationale",
                         margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified")
    return figure


def calibration_table(reference, context, scenario, parameters):
    first = parameters["formation"]["rentree"] + STUDY_YEARS
    stocks = context["stocks"]
    stock = float(stocks.loc[(stocks.scenario == scenario) & (stocks.annee == first), "psychiatrie"].iloc[0])
    national = reference[(reference.scenario == scenario) & (reference.annee == first)].effectif.sum()
    weights = context["weights"]
    weights = weights[weights.scenario == scenario].set_index("origine").poids
    retention = context["routing"].groupby("origine").part.sum()
    records = []
    for code, label in REGIONS.items():
        available = code in context["available"]
        additions = 0.01 * weights.get(code, 0) * stock / DES_YEARS * retention.get(code, 0)
        records.append({"Région de formation": label, "Calibration": "Disponible" if available else "Indisponible",
                        "Hausse validée (%)": parameters["formation"][f"region_{code}"],
                        "Poids formation France (%)": 100 * weights.get(code, 0) if available else np.nan,
                        "Exercice métropolitain (%)": 100 * retention.get(code, 0) if available else np.nan,
                        f"Psychiatres métro. en plus en {first} pour +1 %": additions if available else np.nan,
                        f"X (%) du total métro. en {first} pour +1 %": 100 * additions / national if available and national > 0 else np.nan})
    return pd.DataFrame(records)


def main():
    st.set_page_config(page_title="Projections psychiatres", layout="wide")
    st.title("Projections des psychiatres par département")
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
        st.caption(f"Dynamiques actives : total de référence réparti selon les parts de {anchor}, puis écarts départementaux extrapolés. Le lissage conserve ce total ; les nouvelles places ajoutent ensuite leurs cohortes.")
    with st.sidebar:
        indicator = st.selectbox("Indicateur", sorted(df.indicateur.unique()))
        scenario = st.selectbox("Scénario", sorted(df.scenario.unique()))
        year = st.slider("Année", int(df.annee.min()), int(df.annee.max()), anchor)
    tab_map, tab_trends, tab_method = st.tabs(["Carte et effectifs", "Dynamiques départementales", "Calibration et hypothèses"])
    names = context["baseline"].set_index("dep").nom.to_dict()
    with tab_map:
        selected = df[(df.indicateur == indicator) & (df.scenario == scenario) & (df.annee == year)].copy()
        selected["nom"] = selected.dep.map(names)
        try:
            gj = load_geojson()
        except (OSError, ValueError) as exc:
            gj = None
            st.warning(f"Carte indisponible ; les tableaux et les courbes restent accessibles : {exc}")
        vmin, vmax = df.loc[df.indicateur == indicator, "valeur"].quantile([0.02, 0.98])
        if vmin == vmax:
            vmax = vmin + 1
        map_col, table_col = st.columns([2, 1])
        if gj is not None:
            fig = px.choropleth(selected, geojson=gj, locations="dep", featureidkey="properties.code", color="valeur", color_continuous_scale="Blues", range_color=(vmin, vmax), hover_name="nom", labels={"valeur": indicator})
            fig.update_geos(fitbounds="locations", visible=False)
            fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=650)
            map_col.plotly_chart(fig, width="stretch")
        table_col.dataframe(selected.sort_values("valeur")[["dep", "nom", "valeur"]], hide_index=True, width="stretch")
    with tab_trends:
        st.write(f"Chaque courbe montre l'écart entre l'effectif du département et celui qu'il aurait en conservant sa part du total national de {anchor}. Le total national désigne ici l'ensemble des départements métropolitains.")
        departments = st.multiselect("Départements à comparer", sorted(names), default=["75"] if "75" in names else [sorted(names)[0]],
                                     format_func=lambda dep: f"{dep} · {names[dep]}")
        if departments:
            st.plotly_chart(dynamics_figure(st.session_state.psy_projection["panel"], context, scenario, departments), width="stretch")
            st.caption(f"Trait plein : observations. Pointillés : simulation validée, y compris les effets des nouvelles places après le délai d'études. Toutes les courbes valent zéro à l'ancrage {anchor}. Le filtre Année de la carte ne coupe pas les courbes.")
            slopes = context["baseline"].loc[context["baseline"].dep.isin(departments), ["dep", "nom", "pente"]]
            st.dataframe(slopes.rename(columns={"pente": "Écart de croissance estimé (psychiatres/an)"}), hide_index=True, width="stretch")
        else:
            st.info("Sélectionner au moins un département pour afficher les courbes.")
    with tab_method:
        first = proposed["formation"]["rentree"] + STUDY_YEARS
        st.write(f"Places : hausse annuelle permanente à partir de {proposed['formation']['rentree']}, avec premiers effets en {first}. Le délai retenu est de {STUDY_YEARS} ans après le bac : 6 ans avant le DES et 5 ans de DES. C'est une durée nominale, pas une moyenne mesurée.")
        st.write("Le stock national d'internes de psychiatrie de chaque année de sortie est divisé par 5 pour approcher un flux annuel. La géographie des places est approchée par les stocks régionaux d'internes toutes spécialités de la dernière année observée. Le choix de spécialité, la réussite et la mobilité sont supposés inchangés.")
        st.write("Chaque cohorte supplémentaire est répartie selon la matrice diplôme → exercice des psychiatres RPPS, puis selon les parts départementales observées dans chaque région d'exercice. Les destinations hors métropole restent au dénominateur : elles ne sont pas redistribuées aux départements affichés. Les cohortes sont cumulées sans attrition ni retraite supplémentaire jusqu'en 2070.")
        st.caption("X dépend de la région, du scénario et de l'année. Le tableau donne l'effet de +1 % dans une seule région pour la première cohorte diplômée ; ensuite les cohortes s'accumulent. Les nombres fractionnaires sont des effectifs attendus.")
        st.dataframe(calibration_table(reference, context, scenario, proposed), hide_index=True, width="stretch")
        missing = [REGIONS[code] for code in REGIONS if code not in context["available"]]
        if missing:
            st.warning("Calibration de formation absente pour : " + ", ".join(missing) + ". Une ouverture ex nihilo nécessite des hypothèses supplémentaires ; le curseur reste désactivé. Ces régions peuvent recevoir des diplômés formés ailleurs.")
        st.write(f"Dynamiques : régression linéaire des écarts départementaux à la tendance nationale sur {context['observed_start']}–{anchor}. L'écart extrapolé vaut pente × années depuis {anchor}. Sous 1 psychiatre, la courbe rejoint progressivement zéro ; une renormalisation préserve le total national de référence.")
        if metadata["mode_projection"] == "regional":
            st.write(f"Lorsque les dynamiques sont activées, le total des projections régionales sert de trajectoire nationale : leur ventilation régionale est remplacée par les parts {anchor} et les écarts départementaux. Désactiver les dynamiques rétablit la ventilation régionale de référence.")
        with st.expander("Sources, champ et densités"):
            st.write(f"Onglet de projection des actifs : {metadata['onglet_drees']}.")
            if metadata["mode_projection"] == "national":
                st.write(f"Référence départementale = effectif RPPS en {anchor} × effectif national DREES futur / effectif national DREES en {anchor}. La croissance du champ DREES (70 ans ou moins) est appliquée au RPPS tous âges ; les plus de 70 ans ne font pas l'objet d'une projection distincte.")
            else:
                st.write("Les effectifs régionaux DREES de référence sont utilisés directement. Le RPPS comprend tous les âges, contre 70 ans ou moins pour les actifs DREES : une rupture de champ peut apparaître au raccord.")
            st.write("Population future fournie par population.csv." if metadata["population_fournie"] else f"Population future maintenue au niveau {anchor}, estimée à partir des effectifs et densités RPPS tous médecins. Les arrondis des densités introduisent une approximation.")
            st.write("Densités historiques reprises du RPPS ; densités futures recalculées pour 100 000 habitants après toutes les modulations. Les densités standardisées DREES ne sont pas utilisées.")
            st.write("Aucun fichier de projections n'est écrit. Le calcul de référence est partagé en mémoire ; les paramètres et résultats simulés appartiennent à chaque session.")
        st.caption("Les formules, limites et règles d'extension sont détaillées dans README.md.")


if __name__ == "__main__":
    main()
