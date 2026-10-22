"""Carte des projections de psychiatres par département (métropole).

Lancer :  pip install streamlit plotly pandas numpy
          streamlit run app.py

Données attendues (optionnel) : projections.csv avec les colonnes
    dep (code département en texte : "01", "2A"...), annee, scenario, indicateur, valeur
Sans ce fichier, l'application génère des données fictives pour tester l'affichage.
"""
import json
import urllib.request

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

GEOJSON_URL = (
    "https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/"
    "departements-version-simplifiee.geojson"
)


@st.cache_data
def load_geojson():
    with urllib.request.urlopen(GEOJSON_URL) as f:
        gj = json.load(f)
    # Métropole uniquement (les DROM ont des codes en 97x)
    gj["features"] = [
        feat for feat in gj["features"] if not feat["properties"]["code"].startswith("97")
    ]
    return gj


@st.cache_data
def load_projections(codes):
    try:
        return pd.read_csv("projections.csv", dtype={"dep": str})
    except FileNotFoundError:
        # Données fictives : à remplacer par la sortie de votre modèle
        rng = np.random.default_rng(0)
        rows = []
        for dep in codes:
            base = rng.uniform(5, 40)
            for scenario, croissance in [("Tendanciel", 0.010), ("Hausse des internes", 0.025)]:
                for annee in range(2026, 2051):
                    rows.append({
                        "dep": dep,
                        "annee": annee,
                        "scenario": scenario,
                        "indicateur": "Psychiatres pour 100 000 hab.",
                        "valeur": round(base * (1 + croissance) ** (annee - 2026), 1),
                    })
        return pd.DataFrame(rows)


st.set_page_config(page_title="Projections psychiatres", layout="wide")
st.title("Projections des psychiatres par département")

gj = load_geojson()
noms = {feat["properties"]["code"]: feat["properties"]["nom"] for feat in gj["features"]}
df = load_projections(list(noms))

with st.sidebar:
    indicateur = st.selectbox("Indicateur", sorted(df["indicateur"].unique()))
    scenario = st.selectbox("Scénario", sorted(df["scenario"].unique()))
    annee = st.slider("Année", int(df["annee"].min()), int(df["annee"].max()), int(df["annee"].min()))

sel = df[(df["indicateur"] == indicateur) & (df["scenario"] == scenario) & (df["annee"] == annee)].copy()
sel["nom"] = sel["dep"].map(noms)

# Échelle de couleurs fixe sur toutes les années pour voir l'évolution
vmin, vmax = df.loc[df["indicateur"] == indicateur, "valeur"].quantile([0.02, 0.98])

fig = px.choropleth(
    sel,
    geojson=gj,
    locations="dep",
    featureidkey="properties.code",
    color="valeur",
    color_continuous_scale="Blues",
    range_color=(vmin, vmax),
    hover_name="nom",
    labels={"valeur": indicateur},
)
fig.update_geos(fitbounds="locations", visible=False)
fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=650)

col_carte, col_table = st.columns([2, 1])
col_carte.plotly_chart(fig, width="stretch")
col_table.dataframe(
    sel.sort_values("valeur")[["dep", "nom", "valeur"]], hide_index=True, width="stretch"
)
