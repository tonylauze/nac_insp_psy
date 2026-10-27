# Données et paramètres de projection de demande

Ce dossier est autonome : il ne dépend ni de `app.py` ni des fichiers de
projection d'offre. Le script `demand_projection/run.py` exploite les onze
classeurs Insee Omphale présents ici.

## Données mobilisées

- **Population** : classeurs `donnees_det_*.xlsx`, Insee, Omphale 2022,
  populations départementales annuelles par sexe et tranche d'âge, 2018-2070.
- **Paramètres épidémiologiques** : `prevalence_parameters.csv`. Les valeurs
  initiales correspondent au Baromètre Santé publique France 2024 : épisode
  dépressif caractérisé sur douze mois (18,2 % des femmes et 12,8 % des hommes)
  et trouble anxieux généralisé (6,3 % des 18-79 ans).

Sources :

- https://www.insee.fr/fr/statistiques/7747107
- https://www.santepubliquefrance.fr/sante-mentale/donnees

Les indicateurs ne sont **pas additionnés** : dépression et anxiété peuvent
concerner une même personne. Les paramètres actuels ne couvrent pas les moins
de 18 ans ni les 80 ans et plus ; ces populations ne doivent pas être
interprétées comme sans besoin. Elles attendent des paramètres sourcés dédiés.

## Scénarios

`need_scenarios.csv` contient des hypothèses exploratoires de variation
annuelle relative des prévalences. Elles sont distinctes des scénarios
démographiques Insee et sont plafonnées explicitement. Elles ne sont pas des
prévisions épidémiologiques.

Un fichier facultatif `territorial_modifiers.csv` peut être ajouté avec les
colonnes suivantes :

```csv
department,indicator,territorial_multiplier
Ain,episode_depressif_12_mois,1.05
```

À défaut, tous les multiplicateurs territoriaux valent 1. Ce mécanisme est
prévu pour des hypothèses documentées à partir d'indicateurs publics de
contexte ; il ne doit pas être calibré à partir de l'activité de soins.

## Exécution

Depuis la racine du dépôt :

```powershell
py -m demand_projection.run
```

Deux sorties sont créées :

- `dataDemand/output/demand_projection.csv` : niveau détaillé, y compris le
  sexe auquel s'applique chaque paramètre ;
- `dataDemand/output/demand_projection_summary.csv` : niveau directement
  exploitable, agrégé par département, année, scénarios et indicateur.

Elles contiennent les personnes à risque, la prévalence appliquée et le nombre
projeté de personnes avec besoin.

Le modèle n'utilise aucune donnée SNDS, PMSI, ATIH ou activité de soins.
