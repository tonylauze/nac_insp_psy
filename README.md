# Projections de psychiatres par département

Application Streamlit pilotée par **`app.py`**, avec les modules `demand_projection/` pour la demande de soins. Les deux classeurs sont lus par le code ; les projections et les scénarios sont calculés **en mémoire, sans écriture de résultats sur le serveur**. La projection combine les dynamiques départementales, une correction nationale des arrivées, une répartition annuelle diplôme/exercice, les ouvertures régionales de formation, la tendance nationale des arrivées de diplômés à l’étranger et leur hausse supplémentaire, puis une répartition entre salariés et libéraux. Les dynamiques départementales et la correction de l'internat sont activées par défaut, et désactivables pour comparaison. Les curseurs de hausse de places et d’arrivées étrangères sont initialement nuls ; **la tendance des arrivées étrangères reste active à 0 %**, avec un ajustement nul en 2026. Le taux de retraite utilisé pour inférer les installations est réglable.

La projection démographique conserve les extrapolations de stocks. Les deux politiques publiques calculent séparément des pourcentages de destination et des volumes par lieu de diplôme à partir des installations historiques, puis les extrapolent. Le sélecteur de répartition par stocks agit seulement sur les pourcentages des cohortes supplémentaires.

Le rapprochement offre/demande ajoute une capacité annuelle en actes, un besoin annuel en actes et un solde signé par habitant. Il utilise les effectifs du scénario validé et conserve séparément les hypothèses de demande et de productivité. Voir la section 12 pour les calculs, valeurs initiales et fichiers attendus.

## 1. Installation et fichiers

Placer dans le même dossier du dépôt GitHub :

- `app.py`
- `requirements.txt` (inclut notamment `openpyxl>=3.1,<4`)
- `Médecins RPPS 2012-2026.xlsx`
- `DREES_projections_medecins_2026.xlsx`
- `demand_projection/` et `dataDemand/` pour la demande et son rapprochement avec l’offre ; leur absence laisse les projections d’effectifs accessibles.

Les noms des classeurs doivent correspondre exactement aux constantes `RPPS_PATH` et `DREES_PATH`. Pour les ranger ailleurs, adapter ces constantes dans `app.py`. Un pointeur Git LFS non résolu n'est pas un classeur : les véritables octets Excel doivent être accessibles à l'application.

Installation locale :

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sur Streamlit Cloud, committer également `requirements.txt` permet l'installation d'`openpyxl`. Le script intercepte l'absence de cette dépendance et affiche un message utile. Retirer seulement `import openpyxl` ne résoudrait pas la lecture Excel.

Aucun fichier `projections.csv` n'est requis, lu ou créé. Aucun résultat fictif n'est substitué aux sources absentes ou invalides. Le fond de carte est chargé depuis [france-geojson](https://github.com/gregoiredavid/france-geojson/blob/master/departements-version-simplifiee.geojson) ; si son téléchargement échoue, tableaux et courbes restent disponibles.

## 2. Structure, périodes et champ

### Demande de soins

L'onglet « Demande de soins » utilise `demand_projection/` et `dataDemand/`,
qui doivent également être présents dans le déploiement. Il propose les
indicateurs dépression et anxiété des 18–79 ans, leur cumul sous hypothèse
d'absence de recouvrement, ainsi que « Enfant (6-11) : trouble probable (Enabee) ».
Ce cumul adulte n'inclut pas les enfants. Les filtres de demande sont indépendants
des paramètres de l'offre ; aucune conversion en nombre de psychiatres requis
n'est effectuée.

Les CSV locaux à jour sont lus lorsqu'ils existent. Sinon, le calcul est réalisé
en mémoire et mis en cache à partir des sources versionnées ; les exports ne
sont pas nécessaires au déploiement. Voir [les sources et limites](dataDemand/README.md).

### Offre

Le DataFrame public conserve strictement ces cinq colonnes, dans cet ordre :

| Colonne | Contenu |
| --- | --- |
| `dep` | Code texte, zéros initiaux conservés, y compris `2A` et `2B` |
| `annee` | Année entière |
| `scenario` | Scénario DREES |
| `indicateur` | Les quatre indicateurs d’effectifs/statut ; quatre indicateurs d’activité et de rapprochement s’y ajoutent lorsque les données requises sont disponibles (section 12) |
| `valeur` | Effectif ou actes non négatifs, part de salariés entre 0 et 100 %, ou **solde d’actes signé par habitant et par an** |

Chaque combinaison département × année × scénario × indicateur est unique. Les variables de calibration restent dans des tables internes ; elles n'ajoutent aucune colonne à cette structure publique.

Avec les fichiers fournis : **96 départements métropolitains**, observations **2012–2026**, projections **2027–2070**, quatre scénarios DREES et quatre indicateurs d’effectifs/statut, soit **90 624 lignes d’offre**. Les observations sont répétées à l'identique dans chaque scénario. **2026 est observé**, sans double ligne projetée. L'ancrage `b` est détecté comme la dernière année du classeur historique ; la borne finale dépend des projections disponibles.

Le rapprochement garde les mêmes cinq colonnes publiques. Avec une demande et une population complètes sur 2026–2070, ses quatre indicateurs ajoutent **69 120 lignes**, soit **159 744 lignes** au total. Le champ commun et la disponibilité des données déterminent ce nombre ; aucune ligne de rapprochement n’est inventée pour 2012–2025.

Le RPPS porte sur tous les âges, sexes et modes d'exercice. Le classeur DREES déclare pour les médecins actifs un champ de **70 ans ou moins** ; les étudiants de troisième cycle constituent un statut distinct. Ces champs ne sont pas identiques. Les effectifs sont des personnes, pas des équivalents temps plein.

## 3. Observations et référence de projection

### Observations départementales

Le seul onglet utilisé pour les comptes historiques et les statuts est **`Effectifs`** du classeur fourni `Médecins RPPS 2012-2026.xlsx`. Aucune extraction distincte du RPPS n'est recherchée. On sélectionne Psychiatrie, les départements métropolitains, le sexe et la tranche d'âge « Ensemble ». Les codes tels que `075 -Paris` et `02A -Corse-du-Sud` sont normalisés correctement.

Le total vient de la ligne d'exercice « Ensemble ». Les quatre lignes d'exercice détaillées servent uniquement à répartir ce total : libéral, mixte, salarié hospitalier, autre salarié. Les lignes agrégées ne sont jamais additionnées aux sous-groupes. La somme des quatre catégories doit retrouver exactement le total publié chaque année. Une catégorie sans ligne est traitée comme zéro seulement si cette égalité reste vérifiée ; une cellule manquante dans une ligne présente ou un doublon provoque une erreur. La convention binaire et l'extrapolation sont détaillées en section 9.

### Priorité aux projections régionales de psychiatrie

Le lecteur cherche les en-têtes et le contenu, sans dépendre uniquement du titre de l'onglet : année × scénario × région × spécialité × statut × effectif. Il sélectionne **Psychiatrie** et **Médecins en activité**, en excluant les étudiants, les observations et les lignes de total régional.

Un onglet région × spécialité peut exister sans comporter de ligne Psychiatrie. C'est le cas du classeur fourni : `Effectif rég. par spé. ag.` contient des spécialités agrégées. Ce constat déclenche le repli national ; l'application ne confond pas « autres spécialités » et psychiatrie.

Si des projections régionales détaillées de psychiatrie sont présentes, elles constituent la référence. Leur répartition entre départements suit les parts d'effectifs RPPS de la dernière année observée à l'intérieur de la région source. Les anciennes régions DREES sont rapprochées des départements par leurs libellés ; leurs numéros ne sont pas traités comme des codes régionaux Insee.

Les niveaux régionaux DREES sont utilisés directement dans cette référence : le changement de champ d'âge peut produire une rupture au raccord. Un tableau régional psychiatrique présent mais incomplet, ambigu ou invalide provoque une erreur explicite, pas un repli qui masquerait le problème.

### Repli national rétabli

Si les lignes régionales psychiatriques sont absentes, lire les actifs psychiatres de `Effectif par spécialité`. Pour le département `d`, l'année `t` et le scénario `s` :

```text
Référence[d,t,s] = RPPS[d,b] × DREES_France[t,s] / DREES_France[b,s]
```

Ce mécanisme revient à appliquer la croissance nationale à chaque région, puis à répartir à parts départementales fixes. Il raccorde exactement les niveaux RPPS observés. Il utilise l'évolution relative DREES, sans assimiler son niveau national au total métropolitain RPPS. Il suppose aussi que le champ RPPS plus large suit la croissance du champ DREES ; les plus de 70 ans ne sont pas modélisés séparément.

Avec les sources fournies, l'ancrage métropolitain est de **15 777 psychiatres en 2026**. Les différences entre scénarios DREES, y compris leurs petites différences dues à la simulation, sont conservées.

## 4. Dynamiques départementales

### Estimation de l'écart annuel

Par défaut, toutes les années observées disponibles sont utilisées : **2012–2026** actuellement, avec le même poids pour chacune. Le formulaire permet maintenant de choisir la **première année d'estimation** : par exemple 2017 pour dix années, ou 2022 pour cinq années. Il faut conserver au moins cinq années et cliquer sur « Valider ». Aucun résultat futur n'entre dans l'estimation.

Ce choix est une hypothèse de projection : une pente sur quinze ans n'est pas la pente des deux dernières années. Une fenêtre courte peut réagir davantage aux fluctuations ou ruptures de données. La période complète reste le réglage initial ; aucune fenêtre récente n'est imposée pour rendre les courbes artificiellement plus convaincantes.

### Estimation en nombres de psychiatres

Tous les calculs utilisent les effectifs bruts. Pour chaque département `d`, année `t`, et dernière année observée `b` :

```text
n[d,t] = nombre de psychiatres observé
N[t] = somme_d n[d,t]                     (total métropolitain observé)
p[d] = n[d,b] / N[b]                     (part départementale à l'ancrage)
e[d,t] = n[d,t] - p[d] × N[t]            (écart à la référence nationale)
```

Par construction, `e[d,b] = 0`. Une régression linéaire avec constante de `e[d,t]` sur l'année estime `beta[d]`, en **psychiatres par an**, sur la fenêtre validée. C'est un écart annuel à l'évolution nationale, pas la croissance totale du département. Par exemple, une pente de +5 signifie cinq psychiatres supplémentaires par an par rapport à la trajectoire nationale ventilée à parts fixes.

Tous les départements utilisent les mêmes années. Les écarts somment à zéro pour chaque année, donc les pentes somment aussi à zéro. Le code soustrait leur moyenne pour éliminer le seul résidu d'arrondi numérique. Il vérifie que la fenêtre est complète, sans doublon et commune à tous les départements. Le choix d'une fenêtre récente conserve l'ancrage `b` et les parts de cette dernière année.

### Extrapolation et conservation du total national

La dynamique agit **dès 2027**, indépendamment du délai de formation. La constante de la régression sert à estimer la pente et à tracer la droite ajustée, mais n'est pas extrapolée : le futur repart du niveau réellement observé.

`B[t,s]` désigne le total métropolitain de référence du scénario, issu des projections DREES et du raccord décrit en section 3. Avant saturation :

```text
brut[d,t,s] = p[d] × B[t,s] + beta[d] × (t - b)
```

La somme des pentes étant nulle, les dynamiques redistribuent les effectifs sans modifier le total `B[t,s]`. La projection est raccordée aux observations à l'ancrage. Ce modèle est une extrapolation descriptive, sans identification causale ni bande d'incertitude ; prolonger une tendance jusqu'en 2070 est une hypothèse de stabilité.

**Articulation avec les régions :** lorsque les dynamiques sont activées, cette formule remplace la ventilation territoriale de référence. Si des projections psychiatriques régionales détaillées deviennent disponibles, leur somme métropolitaine fournit `B`. Désactiver les dynamiques rétablit la ventilation régionale d'origine. Les ajouts issus des places continuent à suivre la matrice diplôme/exercice puis les parts départementales observées dans chaque région.

### Saturation lissée près de zéro

Le lissage s'applique aux effectifs projetés, avec un seuil `epsilon = 1` psychiatre, réglable par `SMOOTH_FLOOR` :

```text
f(x) = x                            si x >= epsilon
f(x) = epsilon × exp(x/epsilon - 1)  sinon
```

La valeur et la dérivée sont continues au seuil ; la fonction tend vers zéro pour les extrapolations très négatives. Les nombres très faibles peuvent devenir exactement zéro par sous-flux numérique. Aucun effectif négatif n'est produit.

Le lissage relève les valeurs brutes les plus faibles. Pour ne pas créer artificiellement de psychiatres à l'échelle nationale :

```text
dynamique[d,t,s] = f(brut[d,t,s]) × B[t,s] / somme_d f(brut[d,t,s])
```

Si le total de référence est nul, les effectifs sont nuls. Le total national est donc conservé **à l'étape des dynamiques**. La correction de l'internat peut ensuite le réduire, la tendance des arrivées étrangères le modifier dans les deux sens, puis les hausses d'arrivées et les places supplémentaires l'augmenter. Cette renormalisation implique qu'après activation du plancher, les trajectoires finales ne sont plus strictement linéaires, même pour les autres départements. Les observations, dont 2026, ne sont jamais lissées ni modifiées.

## 5. Projections démographiques et simulations de politiques publiques

Les deux calculs sont séparés. La trajectoire démographique conserve l'extrapolation des stocks ; les curseurs d'ouverture de places et d'augmentation des flux migratoires utilisent des installations reconstituées sur les observations. **À curseurs nuls, les résultats sont identiques à ceux de la version fusionnée précédente**, avec la méthode démographique initiale. Le choix de répartition des politiques publiques ne modifie jamais cette trajectoire.

### Paramètres et sources

| Paramètre | Valeur initiale | Fonction |
| --- | --- | --- |
| Répartition des nouveaux arrivants | Installations estimées (delta + retraites) | Extrapolation des proportions historiques d'installations ; option alternative fondée sur les stocks |
| Taux annuel uniforme de retraite | 3 % | Hypothèse réglable entre 0 et 20 %, commune aux origines et destinations |
| Première année d'estimation | 2012 | Stocks à partir de 2012 ; installations disponibles à partir de 2013 |
| Hausse de places par région de diplôme | 0 % | De 0 à 100 %, maintenue pour chaque cohorte suivante |
| Première rentrée augmentée en deuxième année | 2027 | Année d'entrée de la première cohorte supplémentaire |
| Délai deuxième année → installation | 10 ans | Réglable entre 5 et 20 ans ; premiers effets en 2037 par défaut |
| Hausse des arrivées de diplômés à l'étranger | 0 % | De 0 à 100 %, maintenue chaque année, sans délai de formation |
| Première année de hausse du flux étranger | 2027 | Premiers effets la même année |

Le lecteur utilise toutes les colonnes **2012–2026** de l'onglet **`Effectifs par lieu diplome`**, pour Psychiatrie. La source donne la région de diplôme et la région d'exercice, sans âge ni année individuelle de diplôme. Les marges « Ensemble » servent aux contrôles ; elles ne sont pas additionnées au détail. Les cellules absentes ne sont complétées par zéro qu'après vérification des marges. Les doublons, valeurs manquantes ou négatives sont rejetés.

Il existe 15 origines : 12 régions métropolitaines de formation, Antilles-Guyane (`01`), région inconnue (`98`) et étranger (`99`). Les 12 origines métropolitaines disponibles alimentent les curseurs de formation. La Corse ne dispose pas d'une origine de diplôme calibrée dans cette source ; son curseur reste désactivé, mais elle peut recevoir des diplômés d'autres origines. Les codes `01` et `98` sont inspectables sans alimenter de curseur. **Le code 99 d'origine désigne le lieu du diplôme, pas la nationalité.** Il ne doit pas être confondu avec la destination `99-COM ou étranger`.

Pour les politiques publiques, tous les volumes et proportions portent sur les régions d'exercice françaises, **DROM inclus**, en excluant la destination `99-COM ou étranger`, dont le champ ne peut pas être isolé. Les parts ultramarines restent au dénominateur et ne sont pas réaffectées à la métropole.

### A. Trajectoire démographique : extrapolation des stocks conservée

Pour chaque origine de diplôme `o` et région d'exercice `j`, une régression linéaire sur les effectifs historiques estime une pente `beta_N[o,j]`. La projection reste raccordée au compte réel de `b=2026` :

```text
N_proj[o,j,t] = max(0, N[o,j,b] + beta_N[o,j] × (t-b))
D[o,j,t] = max(0, N_proj[o,j,t] - (1-r) × N_proj[o,j,t-1])
```

L'effectif précédent de 2027 est le compte réel de 2026. Le diagnostic de 2026 utilise 2025. Ces stocks et les flux qui en sont déduits après extrapolation conservent leur rôle dans la correction nationale de l'internat et dans la tendance démographique étrangère. Leur calcul est indépendant du sélecteur de répartition des politiques publiques.

Le total général reste issu de la projection DREES raccordée aux observations, avec les dynamiques départementales et la correction nationale décrites ailleurs. Il n'est pas reconstruit à partir des nouveaux flux de simulation.

### B. Politiques publiques : extrapolation des pourcentages historiques

Pour chaque couple origine de diplôme / région d'exercice, **on calcule d'abord les installations sur les stocks réellement observés**, et non sur les stocks projetés :

```text
I_hist[o,j,t] = max(0, N_obs[o,j,t] - (1-r) × N_obs[o,j,t-1])
p_hist[j|o,t] = I_hist[o,j,t] / somme_regions_francaises_k I_hist[o,k,t]
```

L'historique des installations commence en **2013**, faute de stock 2011 pour calculer le delta 2012. Par défaut, les proportions sont ajustées sur 2013–2026. Si la fenêtre est déplacée à 2019, les installations de 2019 utilisent aussi le compte de 2018, sans inclure 2018 dans la régression. Au moins cinq années sont requises.

Une régression linéaire avec constante et poids égaux par année est ajustée **sur les proportions historiques** de chaque couple. Sa pente `gamma[o,j]` est prolongée depuis la dernière proportion reconstituée en 2026 :

```text
p_brut[j|o,t] = p_hist[j|o,b] + gamma[o,j] × (t-b)
p_borne[j|o,t] = min(1, max(0, p_brut[j|o,t]))
P[j|o,t] = p_borne[j|o,t] / somme_regions_francaises_k p_borne[k|o,t]
```

Les proportions sont donc bornées à 0–100 %, puis renormalisées ensemble pour totaliser 100 % pour chaque origine et chaque année. Le bornage peut modifier la pente finale et la renormalisation peut ramener sous 100 % une proportion individuellement plafonnée à 100 %. Le point 2026 reste raccordé exactement. **La méthode n'est ni une moyenne des répartitions historiques, ni une extrapolation des stocks suivie d'un delta, ni le ratio de flux régionaux extrapolés séparément.**

Lorsque toutes les installations historiques d'une origine sont nulles une année, ses proportions utilisent les stocks observés de cette année. Pour une projection sans proportion disponible, on utilise les stocks projetés ; si tous ces stocks sont nuls, les parts de 2026 servent de repli. Les règles sont affichées.

### C. Politiques publiques : extrapolation séparée des volumes par origine

Le volume est calculé séparément, avec la même idée de delta corrigé des retraites, **appliquée après agrégation des stocks observés sur les régions françaises** :

```text
T_obs[o,t] = somme_regions_francaises_j N_obs[o,j,t]
F_hist[o,t] = max(0, T_obs[o,t] - (1-r) × T_obs[o,t-1])
```

On ajuste une régression linéaire de `F_hist` sur la même fenêtre historique que les proportions. Sa pente `beta_F[o]` est prolongée depuis le flux reconstitué en 2026 :

```text
F_politique[o,t] = max(0, F_hist[o,b] + beta_F[o] × (t-b))
Arrivees_reparties[o,j,t] = F_politique[o,t] × P[j|o,t]
```

**Cette règle vaut pour les deux politiques : ouverture de places et hausse des flux migratoires.** Le volume des cohortes supplémentaires françaises ne provient plus du stock DREES d'internes divisé par cinq ni des poids régionaux d'internes toutes spécialités. Ces anciennes calibrations restent utilisées uniquement pour la correction nationale de la trajectoire démographique, afin de la conserver.

Le maximum est pris après agrégation pour le volume : la somme des deltas régionaux positifs pourrait compter des mobilités internes comme de nouvelles arrivées nationales. Le volume national n'est donc pas nécessairement égal à la somme des `I_hist` régionaux utilisés pour les proportions. Les deux estimations sont combinées explicitement par multiplication. Les flux par origine sont identiques dans les quatre scénarios DREES pour des paramètres identiques.

### D. Option de répartition fondée sur les stocks

Le sélecteur permet d'utiliser les proportions des stocks extrapolés à la place des proportions d'installations :

```text
P_stock[j|o,t] = N_proj[o,j,t] / somme_regions_francaises_k N_proj[o,k,t]
```

Cette option change **uniquement les proportions de destination des cohortes supplémentaires**. Les volumes `F_politique[o,t]` continuent d'être calculés à partir des installations historiques. Elle ne modifie ni la trajectoire démographique, ni la correction nationale de l'internat, ni la composante étrangère active à curseur nul. Les parts de 2026 servent de repli si le stock projeté d'une origine devient nul.

### E. Répartition départementale, délais et cumul

La source diplôme/exercice étant régionale, les proportions départementales utilisent toujours les poids observés en 2026 à l'intérieur de la région d'exercice :

```text
q[d|j] = Effectif[d,b] / somme_departements_de_j Effectif[d,b]
R[d|o,t] = P[region(d)|o,t] × q[d|region(d)]
```

Il s'agit donc d'une évolution estimée au niveau régional, puis ventilée entre départements ; on n'estime pas de mobilités départementales absentes de la source. Les proportions des 96 départements métropolitains peuvent sommer à moins de 100 %, puisque les destinations ultramarines sont conservées.

Pour une hausse de places `u[o]` et une première arrivée `g = rentrée + délai` :

```text
Flux_places[o,t] = (u[o]/100) × F_politique[o,t] × c[t-5]       si t >= g
Ajout_places[d,T] = somme_t=g..T somme_origines_o Flux_places[o,t] × R[d|o,t]
```

Le coefficient des postes vacants `c` est appliqué une seule fois aux flux français, avant leur répartition et leur cumul. Il vaut 1 si la correction est désactivée. La hausse de places est maintenue pour chaque cohorte suivante, avec une élasticité supposée de 1 entre hausse relative des places et hausse relative des installations. Ce n'est pas un doublement exponentiel année après année.

La rentrée désigne la **deuxième année**. Le délai initial de dix ans correspond à cinq ans avant le DES, puis cinq ans de DES ; il est réglable et ne constitue pas une durée moyenne empirique. Une rentrée augmentée en **2027** produit ses premiers effets en **2037**. Sources de parcours consultées le 23 octobre 2026 : [Onisep](https://www.onisep.fr/formation/les-principaux-domaines-de-formation/les-etudes-de-sante/les-etudes-de-medecine) ; [DES de psychiatrie, dix semestres](https://www.legifrance.gouv.fr/loda/article_lc/LEGIARTI000047115341/2026-07-25). Les interruptions et redoublements ne sont pas simulés séparément ; la région d'ouverture de places est assimilée à la région de diplôme.

Pour une hausse `v` du flux migratoire, active à partir de `h` :

```text
Flux_migration[t] = (v/100) × F_politique[99,t]                 si t >= h
Ajout_migration[d,T] = somme_t=h..T Flux_migration[t] × R[d|99,t]
```

Aucun délai de formation ni coefficient de postes vacants français n'est appliqué aux arrivées étrangères. Chaque cohorte utilise la répartition de son année d'arrivée ; les cohortes antérieures ne sont jamais redistribuées. Les ajouts des deux politiques sont cumulés sans attrition ni retraite supplémentaire, comme précédemment. Le taux macro sert à reconstituer les installations historiques ; il ne constitue pas un modèle de survie des cohortes ajoutées.

Le tableau de calibration donne l'effet de +1 % de places à la première année d'installation, à partir du flux d'origine corrigé et de la part de destination métropolitaine. Son ratio `X` utilise comme dénominateur le total DREES après correction de l'internat, avant tendance étrangère et hausses. Le flux supplémentaire absolu est commun aux scénarios ; ce ratio dépend du total de chaque scénario.

### F. Code 99 : composante démographique conservée, distincte du curseur

La tendance active même à curseur nul conserve la projection des stocks du code 99, agrégée sur les régions d'exercice françaises :

```text
T_proj[t] = somme_regions_francaises_j N_proj[99,j,t]
F_demo[t] = max(0, T_proj[t] - (1-r) × T_proj[t-1])
F_demo[b] = max(0, T_obs[99,b] - (1-r) × T_obs[99,b-1])
Delta_F_demo[t] = F_demo[t] - F_demo[b]                         si t > b
Delta_F_demo[b] = 0
```

Son routage démographique reste celui des flux `D` déduits des stocks projetés, conditionné à la France, avec les mêmes replis qu'auparavant. Il n'utilise pas les nouveaux pourcentages des politiques publiques. L'ajustement cumule les écarts annuels au flux de 2026, et peut être négatif. Un retrait ne peut pas rendre un effectif départemental négatif ; une limitation éventuelle est signalée. Le curseur migratoire ajoute séparément son flux, selon les équations de la section E.

Avec 3 % de retraite, le flux de référence 2026 vaut `2 068 − 0,97 × 1 999 = 128,97`. Le flux démographique issu des stocks atteint environ 148,210 en 2040. L'effet cumulé de tendance, à curseur nul, reste environ **+95,878 en métropole en 2040 et +1 479,228 en 2070**.

Le flux utilisé par le curseur est désormais différent : la régression des installations historiques étrangères donne une pente d'environ **−3,569 installations par an** avec la fenêtre initiale. Raccordé à 128,97 en 2026, il vaut environ **125,401 en 2027**, **79,010 en 2040**, puis atteint le plancher zéro en **2063**. Il s'agit du résultat mécanique de cette hypothèse linéaire ; cela ne réduit pas la composante démographique conservée. L'interface présente les deux flux séparément.

### Limites d'interprétation

Le taux de retraite de 3 % est une hypothèse illustrative, non estimée dans ces sources. Les sorties hors retraite, décès, changements de champ et mobilités entre régions peuvent contribuer aux deltas. Les installations reconstituées ne sont donc pas des premières installations directement observées. Les différences annuelles et les petites cellules peuvent produire des proportions instables ; la fenêtre et le taux de retraite influencent les pentes. Les planchers et renormalisations garantissent des valeurs admissibles, pas la validité empirique des hypothèses.

Le maintien du flux de 2026 dans la comparaison démographique étrangère reste une convention de scénario additionnel ; le flux étranger exactement incorporé aux projections DREES n'est pas identifié ici. Les ajouts ne sont ni effacés par une renormalisation nationale ultérieure, ni multipliés par le facteur de raccord RPPS/DREES.


## 6. Périmètre de calcul

La version actuelle utilise les effectifs de psychiatres, leur répartition par statut, les stocks d'étudiants, la mobilité diplôme/exercice et les taux nationaux de postes non pourvus. Ces derniers sont intégrés au code avec leurs sources, sans fichier supplémentaire ni téléchargement de statistiques à l'exécution. Aucun CSV n’est nécessaire au moteur d’effectifs. La demande utilise les sources de `dataDemand/`. Le nouveau solde par habitant requiert la **population totale tous âges**, selon le scénario Insee sélectionné, en plus des populations à risque utilisées pour calculer le besoin. Il ne modifie pas les effectifs projetés.

## 7. Interface et graphique

- « Valider » applique ensemble les paramètres du formulaire. Bouger un curseur ou cocher une case ne modifie pas à lui seul les résultats affichés. Les bandeaux indiquent les **paramètres effectivement validés**, et non une éventuelle saisie en attente.
- L'état des dynamiques est désormais affiché explicitement : activées/désactivées, première année d'effet et fenêtre d'estimation. Si elles sont désactivées, le graphique avertit que la pente estimée n'est pas appliquée.
- Les filtres indicateur/scénario/année pilotent la carte et le tableau ; leur utilisation ne relance pas la simulation. Les quatre indicateurs de rapprochement sont sélectionnables au même endroit que les quatre indicateurs d’effectifs/statut. Ils ne sont calculés que sur la période commune disponible.
- « Valider l’offre / demande » applique séparément les paramètres du rapprochement. Les filtres de l’onglet « Demande de soins » ne les remplacent pas.
- L'échelle de couleurs est fixe entre années pour chaque indicateur d'effectifs : quantiles 2 % et 98 % sur toutes les années et tous les scénarios de la simulation validée. Pour la part salariée, elle reste entre 0 et 100 %. Le solde d’actes utilise une palette divergente rouge/bleu, centrée sur zéro, avec des bornes symétriques au quantile 98 % de sa valeur absolue sur toute la période et tous les scénarios. Les valeurs hors bornes restent dans le tableau et au survol.
- La carte conserve la projection conique conforme de Lambert, avec les parallèles 44° et 49°, un cadrage fixe incluant la Corse et le déplacement désactivé. Les anneaux des polygones GeoJSON sont orientés pour Plotly afin d'éviter un remplissage du globe. La palette bleue commence à une teinte visible sur le fond, les départements sont séparés par des contours blancs et les infobulles présentent le nom, le code et la valeur avec son unité.
- L'onglet « Dynamiques départementales » compare un ou plusieurs départements sur toute la période. Le filtre Année de la carte ne tronque pas ces courbes.
- L'onglet « Statuts d’exercice » montre les parts observées puis extrapolées et les deux comptes pour l'année sélectionnée.
- L'onglet « Diplômes et installations » expose les proportions utilisées par les politiques publiques, leurs pentes, les arrivées réparties et les règles de repli. Le graphique montre l'historique reconstitué en trait plein, la projection en tirets et une ligne de raccord en 2026, avec une couleur distincte par région. Les tableaux historiques permettent de vérifier les deltas ; le tableau du code 99 distingue le flux démographique issu des stocks et le flux extrapolé utilisé par le curseur.
- L'onglet « Internat national » montre la série observée, sa prolongation linéaire et les flux par scénario avant/après correction, avec les sources et les limites de calibration.

- L’onglet « Demande de soins » conserve les projections de besoin et leurs filtres indépendants. L’onglet « Offre / demande » expose les calculs, les courbes du solde, les totaux métropolitains et les hypothèses. L’application comporte huit onglets.

### Deux vues, une référence commune

La vue initiale est **« Sans ouverture de places »** : elle montre la référence, la dynamique départementale, la correction de l'internat, la tendance des arrivées étrangères et leur hausse supplémentaire validée, avant les ajouts de formation. La vue **« Avec les ouvertures de places »** montre le scénario complet validé. La carte présente toujours ce scénario complet, quelle que soit la vue choisie pour le graphique.

Dans les deux vues, la comparaison utilise le même total de référence, **après correction de l'internat et avant tendance des arrivées étrangères, hausse supplémentaire de ces arrivées et ouvertures de places** :

```text
Reference[d,t,s] = p[d] × B_corrige[t,s]
Ecart_affiche[d,t,s] = effectif_de_la_vue[d,t,s] - Reference[d,t,s]
```

`B_corrige` correspond au total observé sur l'historique. La correction nationale uniforme ne crée donc pas à elle seule un écart territorial. La tendance des arrivées étrangères, leur hausse supplémentaire et les ouvertures régionales ne déplacent pas cette référence. L'axe vertical, les survols et les droites ajustées sont en **nombre de psychiatres**. Une valeur négative indique moins de psychiatres que la référence, pas un effectif négatif. Toutes les courbes sont nulles à l'ancrage. Un plateau à zéro indique que le département suit la trajectoire nationale à part fixe ; son nombre de psychiatres peut évoluer.

Les observations sont en **trait plein**, le futur en **tirets**, avec raccord au dernier point observé. Une option affiche les régressions avec constante sur l'historique, en fins pointillés. Ces droites ajustées ne passent pas nécessairement par le dernier point ; la projection repart, elle, du point observé. La saturation et la renormalisation peuvent ensuite infléchir les trajectoires projetées.

Sous la courbe, le tableau distingue la **pente estimée** et la **pente appliquée avant saturation**, en psychiatres par an ; la seconde vaut zéro si la dynamique est désactivée. Un volet détaille l'année sélectionnée : **référence sans modulation + effet départemental après lissage + correction internat + tendance des arrivées étrangères + hausse supplémentaire de ces arrivées + ajout des places = effectif final**. Les contributions sont enregistrées pendant le calcul. La correction nationale peut réduire la pente et l'écart finalement visibles, même si la pente historique estimée reste identique.

### Lecture des trajectoires

Avec le repli national, désactiver les dynamiques départementales supprime leur contribution, mais la tendance des arrivées étrangères reste active même lorsque tous les curseurs de hausse valent zéro. Elle peut donc créer des écarts dès 2027. Les ouvertures de deuxième année agissent à partir de 2037 avec le calendrier initial ; une hausse supplémentaire du flux étranger agit dès 2027 par défaut. Les bandeaux affichent l'état effectivement validé pour distinguer ces situations.

Avec les fichiers fournis, les pentes de 2012–2026 sont d'environ −8,172 pour Paris, +5,505 pour le Rhône, +0,740 pour la Meurthe-et-Moselle et −0,273 pour les Vosges, en psychiatres par an relativement à la tendance nationale. Sur 2022–2026, elles deviennent respectivement −3,192, +2,874, +4,877 et −0,900. Le choix de fenêtre peut donc modifier sensiblement les projections.

## 8. Sessions, recalcul et architecture modulaire

`st.cache_data(persist=None)` conserve seulement la référence non modulée et les calibrations communes **en mémoire**, sans persistance disque. Streamlit restitue des copies indépendantes. Les paramètres et les résultats simulés sont conservés uniquement dans `st.session_state`, propre à chaque session. Aucun cache de ressources partagé ne contient un scénario utilisateur.

Une empreinte du contenu des deux classeurs et du code invalide la référence lorsqu'une source change. Un changement de code/source réinitialise les paramètres aux valeurs par défaut, notamment pour éviter de réutiliser des contrôles devenus incompatibles. Un nouveau navigateur/session démarre aussi avec les paramètres initiaux.

La simulation est recalculée au premier accès, après changement de sources ou lors d'une validation modifiant réellement les paramètres. Elle repart toujours de la référence non modulée : deux validations identiques ne cumulent pas deux fois les effets. L'état des résultats est remplacé seulement après réussite complète du calcul ; des paramètres invalides ne remplacent pas la dernière simulation.

Le moteur d’offre et le rapprochement sont dans `app.py` ; la demande appelle les modules existants `demand_projection.run` et `demand_projection.outcomes` :

| Partie | Rôle |
| --- | --- |
| Lecteurs RPPS/DREES | Extraire et valider les sources |
| `assemble_reference` | Assembler observations et projection de base |
| `build_context` | Parts d'ancrage, pentes, modèles de statut, étudiants, mobilité et routage |
| `logistic_fit`, `fit_status_models` | Ajuster les parts salariées et les raccorder aux observations |
| `vacancy_projection` | Ajuster la tendance linéaire des postes non pourvus et la raccorder au dernier taux observé |
| `diploma_history`, `project_diploma_distribution` | Lire les comptes annuels, extrapoler les cellules et inférer les installations |
| `installation_history`, `project_installation_distribution` | Calculer les installations sur les observations, puis extrapoler leurs proportions historiques bornées pour les politiques publiques |
| `origin_installation_flows`, `extrapolate_installations` | Agréger les stocks observés par origine, reconstituer les flux nationaux et extrapoler séparément leurs volumes |
| `annual_routing` | Ventiler les parts annuelles entre départements |
| `foreign_reference_flows`, `foreign_france_distribution` | Calculer le flux étranger sur le total français, son écart à 2026 et les proportions conditionnelles à la France |
| `prepare_national_flows` | Calibrer les coefficients de cohortes, les flux et les retraits cumulés |
| `prepare_arrivals`, `accumulate_arrivals` | Calibrer puis cumuler les cohortes françaises et étrangères après leur routage annuel |
| `Control` et `Modulation` | Décrire contrôles et transformation associée |
| `modulation_registry(context)` | Ordonner les modulations et générer leurs contrôles |
| `apply_modulations` | Appliquer les fonctions sur des copies et vérifier les invariants |
| `to_projections_frame` | Restituer exactement les cinq colonnes publiques |
| `ensure_projections` | Calcul et mémorisation par session |
| `demand_outcomes` | Charger un export de demande à jour ou appeler les modules de calcul de demande |
| `compute_supply_demand` | Rapprocher les effectifs, les besoins, la population totale et l’activité par psychiatre |
| `match_projections_frame` | Ajouter les quatre indicateurs de rapprochement, en conservant le solde signé |
| `ensure_matching` | Mémoriser le rapprochement par session et l’invalider après une modification d’offre ou de ses entrées |
| `main` | Formulaires, carte, courbes et méthode |

Les répartitions démographiques et celles des politiques publiques sont calculées séparément. `routing` et `mobility_panel` conservent la voie démographique ; `intervention_routing` et `installation_panel` alimentent uniquement les curseurs. Les volumes des politiques sont dans `origin_installations`. Ordre des effets sur le total : **dynamiques départementales → correction nationale de l’internat → tendance des arrivées étrangères → hausse supplémentaire des arrivées étrangères → places de formation → répartition par statut**. Le module `mobilite` décrit les contrôles de calibration ; sa contribution directe au stock est nulle. La tendance étrangère est toujours appliquée. La dernière étape produit les deux comptes à partir de la part salariée extrapolée.

Pour ajouter une modulation, écrire une fonction `ma_modulation(projected, parameters, context)` et ajouter une entrée `Modulation` dans `modulation_registry`. Les contrôles peuvent être `slider`, `checkbox`, `number` ou `select`. Le formulaire et sa validation sont communs. La fonction peut cibler département × année × scénario ; elle doit restituer les mêmes lignes, index et colonnes. Seule la colonne interne `effectif` peut être modifiée. Les observations ne sont pas transmises comme lignes modifiables aux fonctions.

Chaque modulation enregistre aussi sa contribution aux effectifs dans une table interne propre à la session. Ces colonnes de diagnostic ne sont jamais ajoutées aux cinq colonnes publiques. Les entrées et le contexte sont copiés. Le moteur refuse les clés altérées, les effectifs négatifs/non finis, les paramètres hors bornes. Un contrôle désactivé ne peut pas être activé par l'injection d'un paramètre. L'ancien curseur national n'est plus enregistré.

## 9. Statuts : comptes source et extrapolation logistique

### Deux groupes exclusifs

L'onglet `Effectifs` distingue quatre modes d'exercice. Pour construire une partition binaire exhaustive :

```text
Salariés = Salarié hospitalier + Autre Salarié
Libéraux = Libéral + Mixte
Part de salariés = Salariés / Ensemble
```

« Nombre brut de libéraux » inclut donc explicitement l'exercice mixte. « Part de salariés » signifie ici la part des salariés **exclusifs**, hors mixtes. Classer les mixtes dans les deux groupes ferait compter deux fois les mêmes personnes ; les retirer des deux ferait perdre des effectifs. Cette convention est rappelée dans l'interface.

En 2026, la source donne **15 777 psychiatres métropolitains : 9 103 salariés et 6 674 libéraux ou mixtes**. Pour Paris : 2 184 au total, dont 863 salariés et 1 321 libéraux ou mixtes. Ces comptes observés restent exacts dans tous les scénarios et pour tous les réglages.

### Ajustement par département

La fenêtre des statuts est toujours l'historique complet **2012–2026** ; le réglage de fenêtre des dynamiques d'effectifs ne la modifie pas. On ajuste, séparément dans chaque département, une régression logistique agrégée :

```text
logit(p[d,t]) = alpha[d] + beta_statut[d] × année_centrée
logit(p) = log(p / (1-p))
```

Les succès sont les comptes de salariés ; les essais sont les effectifs totaux. Les années sont ainsi pondérées par leurs effectifs, et non traitées comme quinze proportions de précision identique. Pour stabiliser les petites catégories ou les années à 0/100 %, chaque année reçoit une demi-observation dans chacun des deux groupes : succès `salariés + 0,5`, total `ensemble + 1`. Ce lissage numérique ne modifie pas les comptes affichés. Le calcul utilise Newton avec recherche de pas, des années centrées/réduites et une très faible pénalisation de la pente (`1e-8` sur l'échelle réduite). La pente restituée est annuelle. Aucune nouvelle dépendance n'est nécessaire.

Pour éviter une rupture entre observation et projection, on conserve la pente estimée mais on réancre la courbe sur la dernière part réelle :

```text
p[d,t] = expit(logit(p[d,2026]) + beta_statut[d] × (t-2026))
expit(z) = 1 / (1 + exp(-z))
```

Les parts restent entre 0 et 1 et sont stockées en fractions dans les tables internes ; la table publique et les graphiques les convertissent en pourcentages. Une part exactement nulle ou égale à un à l'ancrage est conservée à cette limite. Un département constamment dans un seul groupe, ou avec une seule année positive, a une pente nulle. Les années dont le total est nul ne servent pas à l'ajustement. Si 2026 avait un total nul, la part de la dernière année dont l'effectif total était positif servirait d'ancrage conventionnel en 2026 ; aucune extrapolation n'est possible sans une seule année positive. Un total simulé nul est affiché avec une part de 0 % par convention.

### Comptes futurs

Après toutes les modulations de l'effectif total :

```text
Salariés[d,t,s] = Effectif_final[d,t,s] × p[d,t]
Libéraux[d,t,s] = Effectif_final[d,t,s] - Salariés[d,t,s]
```

Les deux groupes reconstituent toujours le total. Les scénarios DREES, la correction nationale, les places et les arrivées étrangères modifient les nombres, mais utilisent la même part départementale extrapolée pour une année donnée. Les nouvelles cohortes reçoivent cette même répartition. Aucun changement individuel de statut ni répartition spécifique des jeunes diplômés n'est identifié dans les données exploitées. Il s'agit d'une prolongation descriptive, sans intervalle d'incertitude ou validation causale ; l'extrapolation jusqu'en 2070 peut amplifier les tendances historiques.

## 10. Postes non pourvus : coefficient national et arrivées

### Série nationale utilisée

Les taux de postes non pourvus ne figurent pas dans les deux classeurs. La série suivante est intégrée dans `VACANCY_OBSERVATIONS`, avec les URL dans `VACANCY_SOURCES`. Sources vérifiées le **24 octobre 2026**. Les dernières années sont calculées à partir des comptes, sans arrondir les taux utilisés par le modèle.

| Année | Postes non pourvus | Taux | Source |
| --- | --- | --- | --- |
| 2013 | Pourcentage publié | 2,3 % | CNUP, janvier 2024 |
| 2014 | Pourcentage publié | 4,3 % | CNUP, janvier 2024 |
| 2015 | Pourcentage publié | 2,9 % | CNUP, janvier 2024 |
| 2016 | Pourcentage publié | 1,6 % | CNUP, janvier 2024 |
| 2017 | Pourcentage publié | 1,4 % | CNUP, janvier 2024 |
| 2018 | Pourcentage publié | 3,6 % | CNUP, janvier 2024 |
| 2019 | Pourcentage publié | 17,5 % | CNUP, janvier 2024 |
| 2020 | Pourcentage publié | 10,5 % | CNUP, janvier 2024 |
| 2021 | Pourcentage publié | 13,2 % | CNUP, janvier 2024 |
| 2022 | Pourcentage publié | 6,0 % | CNUP, janvier 2024 |
| 2023 | 67 / 547 | 12,25 % | CNUP, janvier 2024 |
| 2024 | 65 / 489 | 13,29 % | Assemblée nationale |
| 2025 | 86 / 554 | 15,52 % | Communiqué CNUP |
| 2026 | 37 / 653 | 5,67 % | Olivier Bonnot, président du CNUP |

Sources détaillées :

- [Dossier de presse CNUP, janvier 2024](https://csa.eu/wp-content/uploads/2024/01/DossierPresse_ChoisirPsychiatrie_CNUP_Final_Janv_24.pdf), graphique historique des postes non pourvus.
- [Rapport d'information n° 714 de l'Assemblée nationale](https://www.assemblee-nationale.fr/dyn/17/rapports/cion-soc/l17b0714_rapport-information.pdf), p. 186, comptes de 2024.
- [Communiqué CNUP du 9 septembre 2025, reproduit par Santé mentale](https://www.santementale.fr/2025/09/internes-en-psychiatrie-15-des-postes-restent-vacants-en-2025/), 468 postes pourvus sur 554, soit 86 vacants, hors CESP.
- [Publication du président du CNUP sur les résultats de 2026](https://fr.linkedin.com/posts/olivier-bonnot-84804584_choisirpsychiatrie-nuitsdelapsychiatrie-activity-7510984553386209281-GRtp), 616 postes pourvus sur 653, soit 37 vacants, hors CESP ; les 11 postes CESP supplémentaires ne sont pas inclus dans ce dénominateur. [Reprise par Santé mentale](https://www.santementale.fr/2026/10/un-regain-dinteret-pour-la-psychiatrie/).

**La hausse n'est pas continue : 2026 marque une nette amélioration.** Cette observation est incluse dans l'ajustement et sert d'ancrage. On ne fabrique pas de point 2012 pour cette série. Les taux 2013–2022 sont publiés arrondis ; le traitement des postes CESP n'est pas explicité de façon uniforme dans les anciennes publications. Leur comparabilité est donc imparfaite. Les chiffres et leurs sources doivent être mis à jour ensemble dans le code lors de nouvelles publications.

### Extrapolation et référence de comparaison

Le modèle national ajuste une **régression linéaire des taux sur l'année, avec constante**, sur **2013–2026 par défaut**. L'utilisateur peut modifier l'année de début, avec au moins cinq observations, puis Valider. L'ajustement utilise les moindres carrés ordinaires, avec un poids égal par année, car les dénominateurs ne sont pas disponibles de façon homogène sur toute la période. La pente est exprimée en fractions par an dans le code et en points de pourcentage par an dans l'interface. La répartition salariés/libéraux conserve son propre modèle logistique, décrit en section 9.

La pente estimée est réancrée sur `v[b] = 37/653`, où `v` est la proportion non pourvue et `b = 2026` :

```text
v_brut[a] = v[b] + beta_v × (a-b)
v[a] = min(1, max(0, v_brut[a]))
c[a] = 1                                      si a <= b
c[a] = min(1, (1-v[a]) / (1-v[b]))             si a > b
```

`a` désigne l'année d'entrée en internat. Le coefficient compare le remplissage futur à un remplissage constant au niveau observé de 2026. Il reste entre 0 et 1. Si la tendance choisie s'améliore, il n'ajoute pas de psychiatres au scénario de référence : il vaut 1. Désactiver la correction impose également 1 pour toutes les cohortes, tout en laissant visible la courbe estimée.

La constante ajustée sert à estimer la pente, mais la projection repart du dernier taux réel : le niveau ajusté de 2026 ne remplace pas l'observation. Les taux historiques restent inchangés. Les bornes 0 et 100 % empêchent les proportions impossibles ; si la droite les franchit, la courbe reste à la borne correspondante et l'interface le signale. Le scénario par défaut n'atteint aucune borne avant 2070. Un taux de référence de 100 % provoquerait une erreur explicite, car le coefficient relatif de remplissage serait indéfini.

Avec la fenêtre complète, la pente vaut environ **+0,83536 point de pourcentage par an**. Pour une entrée en internat en 2027, le taux extrapolé vaut 6,50 % et `c ≈ 0,9911` ; en 2030, 9,01 % ; en 2040, 17,36 % ; en 2050, 25,71 % ; en 2070, 42,42 % et `c ≈ 0,6104`. Cette extrapolation suppose une variation annuelle constante ; ce n'est pas une prévision établie des comportements. Une estimation débutant en 2019 donne au contraire environ −0,60945 point par an, atteint le plancher de 0 % à long terme et n'entraîne aucune réduction supplémentaire. L'interface expose la fenêtre, la pente et toute la courbe pour permettre cette comparaison.

La [présentation DREES du modèle ProMédI, 6 octobre 2026](https://drees.solidarites-sante.gouv.fr/communique-de-presse-jeux-de-donnees/261006-projections-effectifs-m%C3%A9decins-moins-70-ans-activite-horizon-2070) indique que le modèle tient déjà compte des durées d'études, interruptions, retards et abandons. Sa documentation publique ne permet pas d'identifier ici le taux précis de postes non pourvus incorporé aux projections de stocks. **Le remplissage constant de 2026 est donc une convention de scénario additionnel, pas un paramètre DREES retrouvé.** Appliquer directement `1-v[a]` aux étudiants déjà inscrits soustrairait un niveau de non-remplissage potentiellement déjà intégré. Le ratio utilisé n'applique que la dégradation future relativement à cette convention ; il ne remplace pas une reconstruction du modèle DREES et ne garantit pas une correction structurelle exacte de ses hypothèses.

### De la cohorte au flux, puis au stock

Pour l'année d'arrivée `t`, le coefficient est celui de l'entrée en internat `a=t-5`. On utilise le stock DREES de troisième cycle dans **l'année d'arrivée** pour approcher le flux annuel de référence, conformément à la calibration de formation :

```text
F_reference[t,s] = S_psy[t,s] / 5
F_corrige[t,s] = F_reference[t,s] × c[t-5]
Manque_France[t,s] = F_reference[t,s] - F_corrige[t,s]
```

Les cohortes déjà entrées au plus tard en 2026 conservent leur flux de référence. Une entrée en 2027 a son premier effet sur l'effectif en **2032**, sans modifier les observations ni les arrivées de 2027–2031. Le délai de cinq ans concerne cette correction ; il reste distinct du délai réglable entre entrée en deuxième année et installation, dix ans par défaut.

Cette section concerne la correction de la trajectoire démographique, conservée à l'identique. Le stock d'étudiants couvre la France ; le résultat affiché couvre la métropole. Pour ne pas retrancher un flux France entière à un total métropolitain, on conserve les poids régionaux de formation et la mobilité démographique issue des stocks. Le sélecteur des politiques publiques ne les modifie pas :

```text
m[t,s] = somme_origines_metropolitaines_o w[o,s] × somme_destinations_metropolitaines_j P[j|o,t]
Manque_metro[t,s] = Manque_France[t,s] × m[t,s]
Manque_cumule[t,s] = somme_v=b+1..t Manque_metro[v,s]
Retrait[t,s] = min(Manque_cumule[t,s], B[t,s])
K[t,s] = (B[t,s] - Retrait[t,s]) / B[t,s]
Effectif_apres_internat[d,t,s] = Effectif_apres_dynamique[d,t,s] × K[t,s]
```

Le facteur `m[t,s]` varie désormais avec la répartition prédite de chaque année. Avec la méthode delta + retraites à 3 %, il vaut environ 0,9259 en 2027, 0,9186 en 2040 et 0,8964 en 2070 dans le scénario tendanciel. Les origines ultramarines ne sont pas calibrées par les curseurs/routages métropolitains actuels : leur éventuel flux de diplômés vers la métropole n'est pas ajouté à `m`. Cette approximation de champ doit être revue si l'on souhaite une modélisation complète des mobilités ultramarines. Le correcteur estime une perte de flux issu des formations françaises ; il ne modifie pas séparément les arrivées de médecins diplômés à l'étranger.

`K` est le **même pour tous les départements** d'une année et d'un scénario : il conserve leur répartition après dynamique. La conversion nationale ne varie pas selon le statut. Si `B=0`, aucun retrait n'est possible et le coefficient est conventionnellement 1. Si le manque cumulé dépasse le stock, le retrait est plafonné au stock et l'interface avertit que les hypothèses doivent être revues.

Les cohortes manquantes restent soustraites jusqu'en 2070 **sans sorties ni retraites**, comme les cohortes supplémentaires de places restent ajoutées sans attrition. Cette simplification peut amplifier les effets à long terme. Le flux `stock/5` n'identifie ni les entrées effectives, ni les sorties exactes, ni les changements de spécialité ; c'est une approximation disponible à partir des tableaux fournis.

Exemple de contrôle, scénario Tendanciel, fenêtre 2013–2026, sans nouvelles places : en 2032, `S=2 624`, le flux France de référence vaut 524,8 et le flux corrigé environ 520,153 ; avec la méthode delta + retraites à 3 %, le retrait métropolitain cumulé vaut environ 4,290. En 2070, le retrait cumulé atteint environ 3 557,606 pour une référence métropolitaine de 26 736,926, soit `K ≈ 0,8669`. L'arrivée de 2070 utilise **le coefficient de la cohorte 2065**, pas celui de l'entrée 2070.

Enfin, une hausse de places pour une rentrée de deuxième année en 2027 produit sa première arrivée en 2037 avec le délai initial de dix ans. Cette cohorte utilise `c[2032]`. Le correcteur est appliqué à son flux une seule fois avant routage avec les parts prédites de 2037, puis cumul. Les nombres finaux sont ensuite répartis par statut selon la section 9.

## 11. Vérifications de cette version

Les vérifications portent sur les deux classeurs fournis, des cas synthétiques et des sessions Streamlit simulées :

- lecture des observations depuis le seul onglet RPPS `Effectifs`, avec conservation exacte des valeurs ;
- retour à des pentes exprimées en psychiatres/an et comparaison avec les régressions d'effectifs ;
- ancrage observé conservé, dynamiques actives dès 2027 et conservation du total national à cette étape ;
- saturation en cas d'extrapolation négative, résultat fini et non négatif ;
- premiers ajouts de formation en 2037 pour la deuxième année 2027 et le délai de dix ans ; premières hausses d’arrivées étrangères en 2027 ;
- lecture des 15 années et 15 origines de diplôme, contrôle des marges, installations négatives ramenées à zéro et proportions sommant à un ;
- installations historiques 2013–2026 recalculées indépendamment pour chaque couple, puis régressions indépendantes sur leurs pourcentages et sur les volumes nationaux par origine ;
- pourcentages bornés à 0–100 %, renormalisation à 100 %, flux bornés à zéro, raccord exact en 2026 et fenêtres d'au moins cinq années ;
- option de répartition par stocks modifiant uniquement les proportions des politiques, sans changer les volumes ni la trajectoire démographique ;
- cumul indépendant des cohortes avec les proportions de leur année d’arrivée, pour les 96 départements et les quatre scénarios ;
- indépendance du flux étranger vis-à-vis de la correction des places françaises et du délai de formation ;
- calcul du maximum après agrégation nationale, sans transformer une mobilité entre deux régions en arrivée étrangère ;
- ajustement nul en 2026, actif à curseur nul après 2026, cumul des écarts annuels signés et ajustement nul pour un flux futur constant au niveau de 2026 ;
- séparation de la tendance étrangère et de la hausse supplémentaire, y compris quand le début de cette hausse est reporté ;
- graphique en nombres de psychiatres, trait plein/tirets, référence commune aux deux vues et décomposition des effets ;
- comptes de salariés et libéraux/mixte retrouvant les totaux de chaque année et département, dont les comptes 2026 et ceux de Paris ;
- pentes logistiques sur données synthétiques, raccord exact, parts bornées et cas à 0/100 % ;
- pente nationale linéaire comparée à un ajustement indépendant, raccord au dernier taux observé, incréments annuels constants avant bornage et cas synthétiques atteignant 0 ou 100 % ;
- structure publique de cinq colonnes conservée : 90 624 lignes d’effectifs/statut et, avec la demande de test 2026–2070, 69 120 lignes supplémentaires de rapprochement ;
- correction uniforme par département, premiers retraits en 2032, cohorte décalée de cinq ans, baisse récente de 2026 incluse et absence de réduction si la pente est négative ;
- avec les curseurs nuls, identité exacte avec la trajectoire démographique de la version fusionnée précédente ;
- volumes des deux curseurs issus des installations historiques agrégées par origine, sans recours au stock d'internes pour calibrer les hausses de places ;
- correction appliquée une seule fois aux nouvelles cohortes de places et conservation de la décomposition complète ;
- huit onglets, sélection des huit indicateurs, choix du lieu de diplôme et affichage des unités ;
- conservation de la projection Lambert, du cadrage, des couleurs et des infobulles ; orientation correcte des 118 anneaux des 96 départements, y compris après inversion de leur sens, sans modification des limites ;
- validation explicite, absence de cumul lors des clics répétés et indépendance de deux sessions ;
- absence d'écriture des projections ou paramètres sur disque.

- solde négatif conservé, carte centrée sur zéro, dénominateur tous âges requis, absence de double compte du total et de ses composantes ;
- moyenne nationale calculée par ratio des sommes, variation de l’activité et de la couverture, absence de cumul des actes entre années ;
- remplacement des hypothèses de productivité par une projection d’activité externe, contrôles des doublons, scénarios, années et données manquantes ;
- validation séparée du rapprochement, mise à jour de l’offre après une politique sans modifier les besoins, maintien de l’offre si la demande ou la population est absente.

**Limite de vérification :** les deux fichiers reçus pour cette fusion ne contiennent pas les dossiers `demand_projection/`, `dataDemand/`, ni une projection d’activité. Le moteur d’offre a été contrôlé sur les données réelles des classeurs déjà fournis ; le rapprochement et son interface ont été vérifiés avec une demande synthétique contrôlée. Ces données de test ne sont pas livrées comme des observations. L’exécution du module de demande et la jointure avec vos sources réelles restent à vérifier dans le dépôt complet.

Ces contrôles vérifient le fonctionnement du code ; ils ne constituent pas une validation empirique des hypothèses de comportement.


## 12. Rapprochement offre / demande en actes annuels

### Indicateurs et calcul

Le menu principal ajoute :

1. **Actes annuels par psychiatre** : moyenne pondérée par les effectifs libéraux/mixte et salariés du département.
2. **Actes disponibles pour le besoin sélectionné** : capacité annuelle affectée au champ du référentiel choisi.
3. **Actes nécessaires pour le besoin sélectionné** : nombre annuel d’actes associé à la couverture cible du besoin.
4. **Solde d’actes par habitant et par an** : différence entre actes disponibles et nécessaires, divisée par la population totale du département.

Pour un département `d`, une année `t` et un scénario d’offre `s` :

```text
Actes_totaux[d,t,s] = Libéraux[d,t,s] × Activité_libérale[d,t,s]
                    + Salariés[d,t,s] × Activité_salariée[d,t,s]
Actes_par_psychiatre = Actes_totaux / Effectif_total
Actes_disponibles = Actes_totaux × Part_offre_mobilisable
Personnes_ciblées = Personnes_avec_besoin × Couverture_cible
Actes_nécessaires = Personnes_ciblées × Actes_par_personne_ciblée
Solde_par_habitant = (Actes_disponibles − Actes_nécessaires) / Population_totale
```

Les pourcentages du formulaire sont convertis en fractions pour ces équations. Si l’effectif est nul, l’offre et la moyenne d’actes par psychiatre sont définies à zéro. Le solde reste calculable à condition que la population totale soit strictement positive.

**Négatif : manque théorique d’actes ; positif : marge théorique dans le périmètre choisi.** Le solde n’est pas tronqué à zéro. Il ne mesure ni les rendez-vous réellement disponibles, ni l’équilibre de l’ensemble de la psychiatrie. Les actes sont annuels : on ne cumule pas les déficits ni les actes entre années. Les effectifs utilisés incluent déjà les effets cumulés des cohortes simulées. Le solde métropolitain est `somme des soldes en actes / somme des populations`, pas une moyenne simple des ratios départementaux.

Un seul indicateur de besoin est sélectionné. Le cumul adulte dépression/anxiété n’est jamais additionné à ses composantes ou à l’indicateur enfant. Il conserve l’hypothèse explicite d’absence de recouvrement entre dépression et anxiété de la demande reçue. Les projections de demande et d’offre gardent des scénarios distincts : le scénario Insee et le scénario de besoin sont validés dans le formulaire, le scénario DREES reste choisi dans le menu de la carte.

### Paramètres initiaux, tous modifiables

| Paramètre | Défaut | Interprétation et origine |
| --- | ---: | --- |
| Consultations annuelles par psychiatre libéral ou mixte | 2 515 | Calibration approximative CNAM 2024, portée à 2026 par convention |
| Activité salariée / libérale | 0,65 | Hypothèse de conversion en équivalents de consultation ; pas une mesure récente de l’activité hospitalière |
| Évolution annuelle de l’activité par psychiatre | 0 % | Projection constante par défaut ; paramètre réglable entre −5 % et +5 % |
| Actes annuels par personne ciblée | 8 | Hypothèse descriptive de suivi, pas une recommandation clinique par trouble |
| Couverture cible du besoin par un psychiatre | 25 % | Hypothèse sur la part du besoin potentiel à couvrir par un psychiatre, pas un taux de recours observé |
| Part des actes mobilisable pour ce besoin | 50 % | Hypothèse réservant une part de l’offre aux troubles et âges hors du référentiel sélectionné |

La [fiche CNAM 2024 sur les psychiatres et neuropsychiatres libéraux](https://www.assurance-maladie.ameli.fr/sites/default/files/2024_fiche_psychiatres-et-neuro-psychiatres-liberaux.pdf) indique 15,865 millions de consultations et téléconsultations et 6 308 praticiens en fin d’année, soit environ 2 515 par praticien. Le volume annuel comprend aussi l’activité de praticiens ayant cessé en cours d’année : ce ratio est une approximation, pas une mesure individuelle de productivité. La fiche indique aussi environ 318 patients différents par praticien, donnant un ordre de grandeur descriptif proche de huit consultations par patient. Ce nombre ne définit pas les soins nécessaires pour chaque trouble ou gravité. Les [recommandations HAS sur la dépression de l’adulte](https://www.has-sante.fr/jcms/pprd_2974237/fr/depression-de-l-adulte-reperage-et-prise-en-charge-initiale) décrivent un suivi adapté à la situation clinique ; elles ne justifient pas un quota universel de huit actes.

Les valeurs de 0,65, 25 % et 50 % sont des **hypothèses de scénario à calibrer**, non des paramètres estimés dans les fichiers fournis. L’activité est exprimée en équivalents de consultation afin d’aligner les unités ; hospitalisations, travail en équipe et consultations facturées ne sont pas directement interchangeables. Il convient d’appliquer la même définition d’acte à l’offre et au besoin lors d’une calibration métier.

Sans projection d’activité externe, l’activité suit :

```text
Activité_libérale[t] = Activité_libérale[2026] × (1 + Croissance_annuelle)^(t − 2026)
Activité_salariée[t] = Activité_libérale[t] × Ratio_salarié
```

Il s’agit d’une projection paramétrique, pas d’une tendance d’activité estimée à partir d’une série historique. Les paramètres prennent effet avec « Valider l’offre / demande ». Le changement des curseurs d’effectifs après « Valider » actualise aussi les actes disponibles, en conservant la demande validée.

### Population totale : entrée obligatoire pour un solde par habitant

Deux entrées sont acceptées :

- une colonne `population_total` (ou `total_population`, `population_totale`) dans le résultat de demande ;
- un CSV `dataDemand/output/population_totale.csv`, ou un CSV chargé dans le volet « Paramètres offre / demande ».

Colonnes du fichier séparé :

| Colonne | Contenu |
| --- | --- |
| `dep` ou `department` | Code départemental ou nom reconnu |
| `annee` ou `year` | Année entière |
| `demographic_scenario` | Nom exact du scénario Insee, par exemple `Central` |
| `population_total` | Population résidente totale, tous âges, strictement positive |

Le nom `population` est aussi accepté dans ce **fichier explicitement déclaré de population totale**. Le scénario Insee est obligatoire pour éviter une association involontaire de populations et de besoins issus de scénarios différents. Les codes `075`, `02A`, etc. sont normalisés. Si la population est répétée dans plusieurs lignes d’indicateurs de demande, ces répétitions ne sont dédupliquées que lorsque les valeurs concordent.

**`population_at_risk` n’est jamais utilisée comme dénominateur par habitant.** Les 18–79 ans ou les 6–11 ans ne représentent pas la population totale, et des populations à risque de plusieurs troubles peuvent se recouvrir. Aucune somme de ces populations n’est effectuée pour reconstituer artificiellement un total.

Le champ du rapprochement est celui de l’offre, soit 96 départements métropolitains. Les éventuelles lignes ultramarines sont exclues du rapprochement sans être retirées de l’onglet de demande indépendant. Sur la période commune à l’offre et à la demande, à partir de 2026, chaque couple département/année doit disposer d’un besoin et d’une population. Les trous, doublons incohérents, années manquantes et valeurs non finies provoquent un message explicite : aucun zéro, report ou interpolation n’est substitué.

### Projection d’activité externe facultative

Un CSV `dataSupply/actes_par_psychiatre.csv`, ou un CSV chargé dans le même volet, peut remplacer les hypothèses de productivité. Il doit contenir :

| Colonne | Rôle |
| --- | --- |
| `annee` ou `year` | Obligatoire : année de l’activité projetée |
| `actes_liberal` et `actes_salarie` | Deux volumes annuels par psychiatre, un par statut |
| `actes_par_psychiatre` | Alternative aux deux colonnes précédentes : même moyenne appliquée aux deux statuts |
| `dep` ou `department` | Facultative : si absente, les valeurs s’appliquent à tous les départements |
| `scenario` | Facultative : si absente, les valeurs s’appliquent à tous les scénarios DREES |

Les volumes doivent être non négatifs et couvrir toute la période du rapprochement. La projection fournie **remplace** la valeur initiale, le ratio salarié et la croissance annuelle du formulaire : ces trois paramètres ne sont pas appliqués une seconde fois. La couverture, les actes par personne ciblée et la part d’offre mobilisable continuent de s’appliquer.

### Contrat d’intégration avec la demande et fusion

Les ajouts reçus sont conservés : `demand_projection.run` fournit `SCENARIO_FILES`, `build_projection` et `summarize_projection` ; `demand_projection.outcomes` fournit `prepare_outcomes`. Le résultat attendu contient `department`, `year`, `demographic_scenario`, `need_scenario`, `indicator`, `label`, `population_at_risk`, `prevalence` et `persons_with_need`. La population totale peut y être ajoutée ou être fournie séparément comme ci-dessus. Les codes des indicateurs sont lus dans les résultats, pas supposés dans le rapprochement.

`dataDemand/` conserve les fichiers Insee indiqués par `SCENARIO_FILES`, `prevalence_parameters.csv`, `need_scenarios.csv` et, s’il existe, `territorial_modifiers.csv`. L’export `dataDemand/output/demand_projection_summary.csv` n’est utilisé que s’il est à jour ; sinon la demande est recalculée en mémoire. Les imports de demande sont différés : une dépendance absente affiche une erreur dans l’onglet concerné sans empêcher l’offre de fonctionner.

Pour intégrer la fusion, remplacer `app.py` et `README.md` dans le dépôt et conserver les modules, données et dépendances existants. Les deux fichiers fusionnés ne recréent pas les dossiers de demande absents des pièces jointes. Aucun commit n’est effectué. Les calculs de stocks démographiques, des installations pour les politiques publiques, des statuts et de l’internat sont conservés ; le rapprochement est une étape en aval qui ne modifie pas les effectifs.
