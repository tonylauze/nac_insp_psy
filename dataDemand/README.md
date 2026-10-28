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
  et trouble anxieux généralisé (6,3 % des 18-79 ans). Il inclut aussi, pour
  les enfants de 6 à 11 ans, l'indicateur « au moins un trouble probable de
  santé mentale » (13,0 %, étude Enabee 2022). Cet indicateur est déjà une
  mesure tous troubles confondus et ne doit pas être additionné à ses
  composantes (trouble émotionnel, trouble oppositionnel, TDAH probable).

Sources :

- https://www.insee.fr/fr/statistiques/7747107
- https://www.santepubliquefrance.fr/sante-mentale/donnees
- https://www.santepubliquefrance.fr/sante-mentale/depression-et-anxiete/enquetesetudes/premiers-resultats-de-letude-enabee-sur-le-bien-etre-et-la-sante-mentale-des-enfants-de-6-a-11-ans

Les sorties du moteur gardent les indicateurs séparés. Dans `app.py`, un cumul
dépression + anxiété est proposé pour les 18–79 ans sous l'hypothèse explicite
d'absence de recouvrement. Il ne comprend pas les enfants. La population adulte
n'est comptée qu'une fois dans son dénominateur et les deux composantes doivent
être présentes pour chaque département, année et scénario.

Le modèle couvre désormais les 6-11 ans par un
indicateur global ; les moins de 6 ans, les 12-17 ans et les 80 ans et plus attendent des paramètres
sourcés dédiés. Leur absence ne doit pas être interprétée comme une absence de
besoin.

Enabee porte sur les enfants scolarisés en élémentaire en France hexagonale.
Appliquer son taux national à tous les 6–11 ans projetés, y compris aux
territoires ultramarins, est une hypothèse d'extrapolation. Les tranches d'âge
quinquennales Omphale sont proratisées en supposant une distribution uniforme
à l'intérieur de chaque tranche. Les scénarios de besoin sont également
appliqués aux enfants ; ce sont des hypothèses exploratoires ancrées en 2026.

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

L'onglet « Demande de soins » de `app.py` utilise les sorties locales lorsqu'elles
sont plus récentes que les sources. À défaut, il calcule les résultats en mémoire
à partir des fichiers versionnés, avec cache. Il fonctionne donc après un clone
sans avoir à versionner le dossier `output/`. La commande ci-dessous reste
disponible pour exporter les résultats en CSV.

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
