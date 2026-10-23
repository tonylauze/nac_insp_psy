"""Carte des psychiatres : un seul script, construction automatique du CSV.

Installation : pip install streamlit plotly pandas numpy openpyxl
Lancement   : streamlit run app.py

Placer les deux classeurs à côté de ce script. projections.csv est créé ou
recalculé si les sources, ce script ou les paramètres VALIDÉS changent, ou si
le CSV disparaît/est modifié. Changer seulement l'année affichée ne le reconstruit
pas. Un petit fichier .projections_state.json conserve les paramètres appliqués.

IMPORTANT : la projection régionale exige des lignes « Psychiatrie » croisées
avec « Région ». L'onglet « Effectif rég. par spé. ag. » de la copie fournie ne
contient que Ensemble / Médecine générale / Autres spécialités. Avec cette copie,
le script affiche une erreur explicite et ne produit pas de projection inventée.

Densités futures : population.csv facultatif (dep,annee,population). Sans lui,
la population départementale de la dernière année RPPS est maintenue constante.
Les densités standardisées DREES ne sont pas assimilées à des densités brutes.

Extension : ajouter une fonction et une entrée à MODULATIONS, dans ce fichier.
"""
from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Callable
import json
import os
import re
import unicodedata
import urllib.request

import numpy as np
import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st


# Chemins à adapter seulement si les fichiers sont dans un autre dossier.
ROOT = Path(__file__).resolve().parent
RPPS_PATH = ROOT / "Médecins RPPS 2012-2026.xlsx"
DREES_PATH = ROOT / "DREES_projections_medecins_2026.xlsx"
POPULATION_PATH = ROOT / "population.csv"  # Facultatif.
CSV_PATH = ROOT / "projections.csv"
STATE_PATH = ROOT / ".projections_state.json"
GEOJSON_URL = "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements-version-simplifiee.geojson"
COLUMNS = ["dep", "annee", "scenario", "indicateur", "valeur"]
COUNT = "Effectif de psychiatres"
DENSITY = "Psychiatres pour 100 000 hab."
METRO = {f"{n:02d}" for n in range(1, 96) if n != 20} | {"2A", "2B"}


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
            raise ValueError("Aucune projection régionale de psychiatrie dans ce classeur. " + detail)
        raise ValueError("Plusieurs onglets régionaux de psychiatrie trouvés : " + ", ".join(n for n, _ in candidates))
    sheet_name, result = candidates[0]
    if result.duplicated(["region_source", "annee", "scenario"]).any():
        raise ValueError("L'onglet régional contient plusieurs valeurs pour une même région, année et scénario.")
    if not np.isfinite(result.effectif_region).all() or (result.effectif_region < 0).any():
        raise ValueError("Effectif régional invalide.")
    return result, sheet_name


def historical_psychiatry(path):
    """Extraire les lignes agrégées, sans sommer hommes/femmes ou classes d'âge."""
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
                raw_dep = str(row[pos["departement"]]).split("-", 1)[0]
                dep = raw_dep[-2:] if len(raw_dep) == 3 and raw_dep.startswith("0") else raw_dep
                if dep not in METRO:
                    continue
                spec = normalize(row[pos["specialites"]])
                kind = "psy" if spec == "psychiatrie" else "tous"
                if kind == "tous" and not (spec == "ensemble" and normalize(row[pos["specialites_agregees"]]) == "ensemble"):
                    continue
                if sheet_name == "Effectifs" and kind == "psy":
                    names.append({"dep": dep, "nom": str(row[pos["departement"]]).split("-", 1)[1],
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
    # Les effectifs DREES régionaux sont utilisés directement, sans leur substituer
    # une trajectoire nationale ni recopier le total régional dans chaque département.
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


@dataclass(frozen=True)
class Modulation:
    key: str
    controls: tuple
    apply: Callable


def hausse_nationale(projected, parameters):
    result = projected.copy()
    result["effectif"] *= 1 + parameters["hausse_pct"] / 100
    return result


MODULATIONS = (
    Modulation("hausse_nationale", (
        Control("hausse_pct", "Hausse nationale des effectifs (%)", "slider", 0, 0, 100, 1),
    ), hausse_nationale),
)
# Exemple d'extension : écrire def ma_modulation(projected, parameters), qui
# cible projected.dep / projected.region_source / projected.annee / projected.scenario,
# puis ajouter Modulation("ma_modulation", (...Control...), ma_modulation) au tuple.
# Seuls effectif et population sont modifiables ; densite est recalculée à la fin.
# L'ordre dans MODULATIONS est l'ordre d'application. Les contrôles du registre
# apparaissent automatiquement dans le même formulaire et utilisent le même Valider.


def default_parameters():
    return {m.key: {c.key: c.default for c in m.controls} for m in MODULATIONS}


def apply_modulations(reference, parameters):
    result = reference.copy(deep=True)
    projected = result[result.projete].copy()
    immutable = [c for c in projected if c not in {"effectif", "population"}]
    for modulation in MODULATIONS:
        values = parameters[modulation.key]
        for control in modulation.controls:
            value = values[control.key]
            if control.kind in {"slider", "number"}:
                if isinstance(value, bool) or not isinstance(value, (int, float)) or not np.isfinite(value):
                    raise ValueError(f"Valeur numérique invalide : {control.label}.")
                if (control.minimum is not None and value < control.minimum) or (control.maximum is not None and value > control.maximum):
                    raise ValueError(f"Valeur hors limites : {control.label}.")
            elif control.kind == "checkbox" and not isinstance(value, bool):
                raise ValueError(f"Valeur booléenne attendue : {control.label}.")
            elif control.kind == "select" and value not in control.options:
                raise ValueError(f"Option invalide : {control.label}.")
        changed = modulation.apply(projected.copy(deep=True), dict(values))
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


def to_csv_frame(panel):
    result = panel.melt(id_vars=["dep", "annee", "scenario"], value_vars=["effectif", "densite"], var_name="indicateur", value_name="valeur")
    result.indicateur = result.indicateur.map({"effectif": COUNT, "densite": DENSITY})
    result = result[COLUMNS].sort_values(COLUMNS[:-1]).reset_index(drop=True)
    if result.empty or result.isna().any().any() or result.duplicated(COLUMNS[:-1]).any():
        raise ValueError("CSV vide, incomplet ou dupliqué.")
    if not np.isfinite(result.valeur).all() or (result.valeur < 0).any():
        raise ValueError("Valeur CSV invalide.")
    return result


# ---------- Cache des sources et invalidation du CSV ----------
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


@st.cache_data(show_spinner=False)
def load_reference(signature):
    # signature fait partie de la clé du cache, même si elle n'est pas utilisée
    # dans la lecture. Lire les projections d'abord évite une lecture RPPS coûteuse
    # lorsque le croisement région × psychiatrie manque.
    regional, sheet = regional_psychiatry(DREES_PATH)
    history, names = historical_psychiatry(RPPS_PATH)
    population = pd.read_csv(POPULATION_PATH, dtype={"dep": str}) if POPULATION_PATH.exists() else None
    reference, anchor = assemble_reference(history, names, regional, population)
    return reference, {"annee_reference": anchor, "onglet_drees": sheet,
                       "population_fournie": population is not None}


def atomic_write(path, content):
    temporary = None
    try:
        with NamedTemporaryFile(dir=path.parent, prefix=path.name + ".", suffix=".tmp", delete=False) as handle:
            temporary = Path(handle.name)
            handle.write(content)
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def read_state():
    try:
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
        return state if isinstance(state, dict) else {}
    except (OSError, ValueError):
        return {}


def ensure_projections(signature, parameters):
    """Ne reconstruit que si le résultat existant ne correspond plus à la demande."""
    key = sha256(json.dumps({"sources": signature, "parameters": parameters}, sort_keys=True).encode()).hexdigest()
    state = read_state()
    if (state.get("key") == key and CSV_PATH.exists()
            and state.get("csv_hash") == file_hash(CSV_PATH)):
        return pd.read_csv(CSV_PATH, dtype={"dep": str}), state["metadata"], False
    reference, metadata = load_reference(signature)
    # Toujours repartir de la référence non modulée : aucun cumul entre clics.
    output = to_csv_frame(apply_modulations(reference, parameters))
    content = output.to_csv(index=False).encode("utf-8")
    new_state = {"key": key, "csv_hash": sha256(content).hexdigest(),
                 "parameters": parameters, "metadata": metadata}
    atomic_write(CSV_PATH, content)
    atomic_write(STATE_PATH, json.dumps(new_state, ensure_ascii=False, indent=2).encode("utf-8"))
    return output, metadata, True


def render_parameters(applied):
    draft = {}
    for modulation in MODULATIONS:
        values = {}
        for control in modulation.controls:
            value = applied.get(modulation.key, {}).get(control.key, control.default)
            common = {"label": control.label, "key": f"param_{modulation.key}_{control.key}"}
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


@st.cache_data
def load_geojson():
    with urllib.request.urlopen(GEOJSON_URL, timeout=20) as handle:
        gj = json.load(handle)
    gj["features"] = [f for f in gj["features"] if f["properties"]["code"] in METRO]
    return gj


def main():
    st.set_page_config(page_title="Projections psychiatres", layout="wide")
    st.title("Projections des psychiatres par département")
    saved_parameters = read_state().get("parameters", {})
    defaults = default_parameters()
    for modulation in MODULATIONS:
        defaults[modulation.key].update({k: v for k, v in saved_parameters.get(modulation.key, {}).items() if k in defaults[modulation.key]})
    if "applied_parameters" not in st.session_state:
        st.session_state.applied_parameters = defaults
    with st.sidebar:
        with st.form("modulations"):
            st.subheader("Modulations")
            draft = render_parameters(st.session_state.applied_parameters)
            submitted = st.form_submit_button("Valider", type="primary")
        st.caption("Hausse de niveau appliquée à toutes les années projetées et à tous les scénarios. Les observations restent inchangées.")
    proposed = draft if submitted else st.session_state.applied_parameters
    try:
        signature = source_signature()
        with st.spinner("Vérification des données et recalcul si nécessaire…"):
            df, metadata, recalculated = ensure_projections(signature, proposed)
        # Enregistrer les paramètres seulement après calcul et écriture réussis.
        st.session_state.applied_parameters = proposed
    except FileNotFoundError as exc:
        st.error(f"Fichier absent : {exc.filename}. Placer les deux classeurs à côté de app.py.")
        st.stop()
    except (ValueError, KeyError, OSError) as exc:
        st.error(str(exc))
        st.stop()
    if recalculated:
        st.success("projections.csv a été recalculé.")
    anchor = metadata["annee_reference"]
    st.caption(f"Observé : {df.annee.min()}–{anchor}. Projeté : {anchor + 1}–{df.annee.max()}.")
    st.caption("Après l'historique, les effectifs régionaux DREES sont répartis selon les parts départementales de la dernière année RPPS.")
    with st.expander("Champ des données et densités"):
        st.write(f"Onglet utilisé : {metadata['onglet_drees']}. Les niveaux régionaux projetés sont utilisés directement. Vérifier leur champ d'âge : dans le classeur joint, la DREES limite son champ à 70 ans, tandis que l'historique RPPS comprend tous les âges. Ce changement de champ peut créer une rupture au raccord.")
        st.write("Population future fournie par population.csv." if metadata["population_fournie"] else f"Population future constante au niveau de {anchor}, reconstituée à partir des effectifs et densités RPPS tous médecins. Il s'agit d'une approximation liée aux arrondis des densités.")
        st.write("Les densités historiques sont reprises du RPPS ; les densités futures sont calculées pour 100 000 habitants. Les densités standardisées DREES ne sont pas utilisées.")
    with st.sidebar:
        indicator = st.selectbox("Indicateur", sorted(df.indicateur.unique()))
        scenario = st.selectbox("Scénario", sorted(df.scenario.unique()))
        year = st.slider("Année", int(df.annee.min()), int(df.annee.max()), anchor)
    selected = df[(df.indicateur == indicator) & (df.scenario == scenario) & (df.annee == year)].copy()
    try:
        gj = load_geojson()
        names = {f["properties"]["code"]: f["properties"]["nom"] for f in gj["features"]}
    except (OSError, ValueError) as exc:
        gj, names = None, {}
        st.warning(f"Carte indisponible ; le tableau reste accessible : {exc}")
    selected["nom"] = selected.dep.map(names).fillna(selected.dep)
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
    st.download_button("Télécharger projections.csv", df.to_csv(index=False).encode("utf-8"), file_name="projections.csv", mime="text/csv")


if __name__ == "__main__":
    main()
