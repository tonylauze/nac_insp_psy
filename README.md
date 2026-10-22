# Projections des psychiatres par département

Site de visualisation des projections d'effectifs de psychiatres en France métropolitaine, par département.

## Lancer le site sur son ordinateur

```
python -m venv .venv
source .venv/bin/activate        # Windows : .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Données

L'application lit `projections.csv` s'il est présent dans le dossier, avec les colonnes :

| Colonne | Contenu |
| --- | --- |
| `dep` | Code département en texte (`01`, `2A`, `2B`…) |
| `annee` | Année de projection |
| `scenario` | Nom du scénario |
| `indicateur` | Nom de l'indicateur (ex. psychiatres pour 100 000 hab.) |
| `valeur` | Valeur |

Sans ce fichier, des données fictives sont générées pour tester l'affichage.

Fond de carte : [france-geojson](https://github.com/gregoiredavid/france-geojson) (départements, version simplifiée).
