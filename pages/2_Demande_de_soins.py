"""Page Streamlit autonome de visualisation des projections de demande."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from demand_projection.run import build_projection


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "dataDemand"
OUTPUT_DIR = DATA_DIR / "output"
SUMMARY_PATH = OUTPUT_DIR / "demand_projection_summary.csv"


st.set_page_config(page_title="Demande de soins", page_icon="🧠", layout="wide")


@st.cache_data(show_spinner=False)
def load_summary(path: str, modified_at_ns: int) -> pd.DataFrame:
    """Cache invalidated when the generated CSV changes."""
    del modified_at_ns
    return pd.read_csv(path)


def current_data() -> pd.DataFrame | None:
    if not SUMMARY_PATH.exists():
        return None
    return load_summary(str(SUMMARY_PATH), SUMMARY_PATH.stat().st_mtime_ns)


st.title("Demande de soins en psychiatrie")
st.caption(
    "Projection démographique-épidémiologique indépendante de la projection de l'offre. "
    "Elle ne mobilise aucune donnée SNDS, PMSI, ATIH ou d'activité de soins."
)

top_left, top_right = st.columns([1, 3])
with top_left:
    recalculate = st.button("Recalculer les projections", type="primary")
with top_right:
    st.caption(
        "Le recalcul lit les onze classeurs Omphale dans `dataDemand/` et applique les "
        "paramètres définis dans `prevalence_parameters.csv` et `need_scenarios.csv`."
    )

if recalculate:
    with st.spinner("Calcul des projections de demande…"):
        build_projection(DATA_DIR, OUTPUT_DIR, base_year=2026)
    st.cache_data.clear()
    st.success("Projections recalculées.")

data = current_data()
if data is None:
    st.warning("Aucune projection n'a encore été générée.")
    st.code("py -m demand_projection.run", language="powershell")
    st.stop()

with st.sidebar:
    st.header("Filtres")
    demographic_scenario = st.selectbox(
        "Scénario démographique Insee",
        sorted(data["demographic_scenario"].unique()),
        index=sorted(data["demographic_scenario"].unique()).index("Central"),
    )
    need_scenario = st.selectbox(
        "Scénario de besoin",
        sorted(data["need_scenario"].unique()),
        index=sorted(data["need_scenario"].unique()).index("stable"),
    )
    indicator = st.selectbox(
        "Indicateur",
        sorted(data["indicator"].unique()),
        format_func=lambda value: data.loc[data["indicator"] == value, "label"].iloc[0],
    )

filtered = data.loc[
    (data["demographic_scenario"] == demographic_scenario)
    & (data["need_scenario"] == need_scenario)
    & (data["indicator"] == indicator)
].copy()
filtered["persons_with_need"] = filtered["persons_with_need"].round().astype(int)

departments = sorted(filtered["department"].unique())
default_departments = ["Paris"] if "Paris" in departments else [departments[0]]
with st.sidebar:
    selected_departments = st.multiselect(
        "Départements à comparer", departments, default=default_departments
    )

selected = filtered.loc[filtered["department"].isin(selected_departments)].copy()
if selected.empty:
    st.info("Sélectionnez au moins un département.")
    st.stop()

national = filtered.groupby("year", as_index=False)["persons_with_need"].sum()
base_value = int(national.loc[national["year"] == 2026, "persons_with_need"].iloc[0])
end_value = int(national.loc[national["year"] == 2070, "persons_with_need"].iloc[0])
variation = (end_value / base_value - 1) if base_value else 0.0
metric_a, metric_b, metric_c = st.columns(3)
metric_a.metric("Personnes avec besoin projeté — France, 2026", f"{base_value:,}".replace(",", " "))
metric_b.metric("Personnes avec besoin projeté — France, 2070", f"{end_value:,}".replace(",", " "))
metric_c.metric("Évolution 2026–2070", f"{variation:+.1%}")

figure = px.line(
    selected,
    x="year",
    y="persons_with_need",
    color="department",
    labels={
        "year": "Année",
        "persons_with_need": "Personnes avec besoin projeté",
        "department": "Département",
    },
    title=data.loc[data["indicator"] == indicator, "label"].iloc[0],
)
figure.update_layout(legend_title_text="Département", hovermode="x unified")
st.plotly_chart(figure, width="stretch")

st.subheader("Détail par département")
display = selected.loc[
    selected["year"].isin([2026, 2030, 2040, 2050, 2060, 2070]),
    ["department", "year", "population_at_risk", "prevalence", "persons_with_need"],
].copy()
display["population_at_risk"] = display["population_at_risk"].round().astype(int)
display["prevalence"] = (display["prevalence"] * 100).round(2)
st.dataframe(
    display.rename(
        columns={
            "department": "Département",
            "year": "Année",
            "population_at_risk": "Population concernée",
            "prevalence": "Prévalence (%)",
            "persons_with_need": "Personnes avec besoin",
        }
    ),
    hide_index=True,
    width="stretch",
)

with st.expander("Périmètre et limites", expanded=False):
    st.markdown(
        """
        Les résultats correspondent à des personnes ayant un besoin potentiel,
        estimé par population projetée × prévalence. Ils ne sont ni des
        consommations de soins observées, ni une mesure directe du nombre de
        psychiatres requis.

        Les indicateurs dépression et anxiété sont affichés séparément, car ils
        peuvent concerner une même personne. Dans la version initiale, les
        paramètres disponibles couvrent les 18–79 ans. Les enfants-adolescents,
        les 80 ans et plus, les addictions et les troubles sévères nécessitent
        des paramètres sourcés complémentaires avant d'être ajoutés.
        """
    )
