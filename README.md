# Projections de psychiatres par département

Application Streamlit contenue dans **un seul `app.py`**. Les deux classeurs sont lus par le code ; les projections et les scénarios sont calculés **en mémoire, sans écriture de résultats sur le serveur**. La projection combine les dynamiques départementales, une correction nationale des arrivées, une répartition annuelle diplôme/exercice, les ouvertures régionales de formation, la tendance nationale des arrivées de diplômés à l’étranger et leur hausse supplémentaire, puis une répartition entre salariés et libéraux. Les dynamiques départementales et la correction de l'internat sont activées par défaut, et désactivables pour comparaison. Les curseurs de hausse de places et d’arrivées étrangères sont initialement nuls ; **la tendance des arrivées étrangères reste active à 0 %**, avec un ajustement nul en 2026. Le taux de retraite utilisé pour inférer les installations est réglable.

## 1. Installation et fichiers

Placer dans le même dossier du dépôt GitHub :

- `app.py`
- `requirements.txt` (inclut notamment `openpyxl>=3.1,<4`)
- `Médecins RPPS 2012-2026.xlsx`
- `DREES_projections_medecins_2026.xlsx`

Les noms des classeurs doivent correspondre exactement aux constantes `RPPS_PATH` et `DREES_PATH`. Pour les ranger ailleurs, adapter ces constantes dans `app.py`. Un pointeur Git LFS non résolu n'est pas un classeur : les véritables octets Excel doivent être accessibles à l'application.

Installation locale :

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sur Streamlit Cloud, committer également `requirements.txt` permet l'installation d'`openpyxl`. Le script intercepte l'absence de cette dépendance et affiche un message utile. Retirer seulement `import openpyxl` ne résoudrait pas la lecture Excel.

Aucun fichier `projections.csv` n'est requis, lu ou créé. Aucun résultat fictif n'est substitué aux sources absentes ou invalides. Le fond de carte est chargé depuis [france-geojson](https://github.com/gregoiredavid/france-geojson/blob/master/departements-version-simplifiee.geojson) ; si son téléchargement échoue, tableaux et courbes restent disponibles.

## 2. Structure, périodes et champ

Le DataFrame public conserve strictement ces cinq colonnes, dans cet ordre :

| Colonne | Contenu |
| --- | --- |
| `dep` | Code texte, zéros initiaux conservés, y compris `2A` et `2B` |
| `annee` | Année entière |
| `scenario` | Scénario DREES |
| `indicateur` | `Effectif de psychiatres`, `Part de salariés`, `Nombre brut de libéraux` ou `Nombre brut de salariés` |
| `valeur` | Effectif réel non négatif, ou part de salariés exprimée en pourcentage entre 0 et 100 |

Chaque combinaison département × année × scénario × indicateur est unique. Les variables de calibration restent dans des tables internes ; elles n'ajoutent aucune colonne à cette structure publique.

Avec les fichiers fournis : **96 départements métropolitains**, observations **2012–2026**, projections **2027–2070**, quatre scénarios DREES et quatre indicateurs, soit **90 624 lignes**. Les observations sont répétées à l'identique dans chaque scénario. **2026 est observé**, sans double ligne projetée. L'ancrage `b` est détecté comme la dernière année du classeur historique ; la borne finale dépend des projections disponibles.

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

## 5. Installations, places en deuxième année et diplômés à l'étranger

### Paramètres validés

| Paramètre | Défaut | Effet |
| --- | --- | --- |
| Méthode de répartition | Installations estimées (delta + retraites) | Utiliser les installations reconstituées, ou directement les parts des effectifs projetés |
| Taux annuel uniforme de retraite | 3 % | Hypothèse modifiable entre 0 et 20 %, par pas de 0,1 point ; identique pour toutes les origines et destinations |
| Première année d'estimation diplôme/exercice | 2012 | Régressions jusqu'en 2026, avec au moins cinq années |
| Hausses régionales de places | 0 % | De 0 à 100 %, reconduites pour chaque cohorte de deuxième année |
| Première rentrée augmentée en deuxième année | 2027 | Année de départ des cohortes concernées |
| Délai deuxième année → installation | 10 ans | Modifiable de 5 à 20 ans ; la première installation doit rester dans l'horizon disponible |
| Hausse du flux de diplômés à l'étranger, code 99 | 0 % | De 0 à 100 %, appliquée chaque année, sans délai de formation |
| Première année de hausse du flux étranger | 2027 | Premiers effets cette même année |

Le taux de retraite de 3 % est un **paramètre illustratif, non estimé dans les sources**. Il sert à reconstituer un flux à partir des stocks. Il n'est pas soustrait une seconde fois à la référence DREES et ne modifie pas directement ses hypothèses de retraite. Les variations de ce taux modifient les installations inférées, donc leur répartition, ainsi que le volume étranger estimé. En mode de répartition par les effectifs projetés, il n'affecte pas les proportions mais continue d'affecter le volume du flux étranger.

### Historique complet par origine de diplôme et destination d'exercice

Le lecteur exploite toutes les colonnes annuelles **2012–2026** de l'onglet **`Effectifs par lieu diplome`**, pour Psychiatrie. Le tableau fourni contient les régions de diplôme et les régions d'exercice, sans ventilation par âge ou année individuelle de diplôme. Les totaux « Ensemble » servent uniquement au contrôle des marges. Pour chaque origine et année, la somme des destinations doit retrouver le total publié ; les cellules absentes peuvent être complétées par zéro uniquement après ce contrôle. Les doublons, valeurs manquantes et valeurs négatives sont rejetés.

Les sources actuelles donnent 15 origines, dont 12 régions métropolitaines de formation, Antilles-Guyane (`01`), région inconnue (`98`) et étranger (`99`). Toutes sont extrapolées pour permettre l'inspection. Le code `98` n'est pas fusionné avec `99` et n'alimente aucun curseur. Les destinations hors métropole restent dans la normalisation des proportions ; pour l'origine étrangère, le calcul national conditionne ensuite ces proportions aux destinations françaises, comme décrit plus bas. **Le code 99 d'origine désigne le lieu du diplôme, pas la nationalité ; le code 99 de destination est une autre variable.**

Pour l'origine `o`, la région d'exercice `j` et l'année `t`, on note `N[o,j,t]` l'effectif. Une régression linéaire avec constante sur la fenêtre validée estime une pente `beta[o,j]`, en personnes par an. La projection est raccordée au dernier compte réel `b=2026` :

```text
N_proj[o,j,t] = max(0, N[o,j,b] + beta[o,j] × (t-b))
```

Chaque cellule est extrapolée séparément, sans modifier les comptes observés. Le plancher empêche les effectifs négatifs. Cette projection des cellules n'est pas renormalisée au total national DREES : elle sert à estimer les répartitions d'installation et les flux du code 99. Le total général de l'application conserve la référence DREES et les dynamiques départementales décrites plus haut.

### Deux méthodes de répartition, année par année

**Méthode 1 — installations estimées.** À partir des comptes extrapolés et du taux de retraite `r` :

```text
I_brut[o,j,t] = N_proj[o,j,t] - (1-r) × N_proj[o,j,t-1]
I[o,j,t] = max(0, I_brut[o,j,t])
P[j|o,t] = I[o,j,t] / somme_destinations I[o,destination,t]
```

Pour la première projection (2027), l'effectif précédent est le compte réel de 2026. Le diagnostic à l'ancrage utilise les comptes réels de 2026 et 2025. Aucune installation négative n'est conservée. Le dénominateur inclut toutes les destinations, pas seulement celles affichées sur la carte métropolitaine.

**Méthode 2 — répartition des effectifs projetés.** Les proportions n'utilisent pas le delta :

```text
P[j|o,t] = N_proj[o,j,t] / somme_destinations N_proj[o,destination,t]
```

En méthode 1, si tous les flux d'une origine sont nuls, on utilise les parts des effectifs projetés. Si tous les effectifs projetés sont eux-mêmes nuls, on utilise les parts observées en 2026. En méthode 2, ce dernier repli s'applique également. L'interface signale le recours à ces règles. Une origine sans effectif à l'ancrage et sans calibration n'est pas dotée arbitrairement d'une répartition.

Cette reconstitution suppose que les sorties sont approchées par `r × N_précédent`. **Elle n'identifie pas directement les premières installations** : les changements de région en cours de carrière, les cessations hors retraite, les décès et les changements de champ peuvent également expliquer les variations de stocks. Pour les proportions territoriales, le maximum est pris cellule par cellule ; sa somme peut dépasser le maximum du delta agrégé, notamment lors de mobilités entre régions. **Le volume national étranger est désormais calculé séparément, en appliquant le maximum après agrégation**, pour éviter de compter ces mouvements internes comme des arrivées nationales.

### Volume des diplômés formés en France

Le nouveau calcul détermine la répartition géographique. Le volume national reste calibré sur le statut **Étudiants en 3ème cycle** en Psychiatrie dans `Effectif par spécialité` :

```text
F_reference[t,s] = S_psy[t,s] / 5
F_corrige[t,s] = F_reference[t,s] × c[t-5]
```

`S_psy` est le stock d'internes de l'année d'arrivée ; sa division par cinq approche un flux annuel de sortie, sous hypothèse de cohortes équilibrées. Elle ne reconstitue pas exactement les sorties lorsque les tailles de cohortes ou durées changent. Le coefficient de postes non pourvus est celui de l'entrée en internat, cinq ans plus tôt ; il vaut 1 lorsque la correction est désactivée. Sa définition figure en section 10. Il ne s'applique pas aux diplômés à l'étranger.

Faute d'effectifs régionaux d'internes spécialisés en psychiatrie dans les tables exploitées, les poids régionaux de formation restent ceux des stocks DREES d'internes **toutes spécialités** à l'ancrage :

```text
w[o,s] = S_tous_region[o,b,s] / S_tous_France[b,s]
F_origine[o,t,s] = w[o,s] × F_corrige[t,s]
```

Les poids sont propres au scénario et fixes dans le temps. Les anciennes régions DREES sont regroupées en régions actuelles. Le dénominateur est France entière ; les seules régions métropolitaines ne sont pas renormalisées à 100 %. Le ratio des stocks Psychiatrie/Ensemble n'est pas traité comme une probabilité individuelle de spécialisation, car les durées de DES diffèrent.

Cette calibration maintient l'hypothèse qu'une hausse relative de places de deuxième année se transmet proportionnellement au flux futur de psychiatres : élasticité de 1, composition en spécialités et réussite inchangées. Les effectifs historiques par lieu de diplôme servent à prévoir **où** s'installera chaque cohorte, sans remplacer le volume DREES français par la somme des deltas.

### Délai, routage annuel et cumul

La rentrée désigne désormais **la deuxième année**, et non l'année du bac. Le délai initial de dix ans correspond à cinq ans restant avant le DES, puis cinq ans de DES. Il est réglable et ne constitue pas une durée moyenne empirique. Une rentrée augmentée en **2027** produit donc ses premiers effets en **2037** avec le réglage initial. Le délai de cinq ans du DES employé pour le coefficient de non-remplissage reste distinct et fixe.

Sources sur le parcours consultées le 23 octobre 2026 : [Onisep](https://www.onisep.fr/formation/les-principaux-domaines-de-formation/les-etudes-de-sante/les-etudes-de-medecine) ; [maquette du DES de psychiatrie, dix semestres](https://www.legifrance.gouv.fr/loda/article_lc/LEGIARTI000047115341/2026-07-25). Les interruptions, redoublements et changements de région en cours d'études ne sont pas simulés séparément. On assimile région d'ouverture des places et région de diplôme.

À l'intérieur de la région d'exercice, la part du département est maintenue à son niveau de 2026 :

```text
q[d|j] = Effectif[d,b] / somme_departements_de_j Effectif[d,b]
R[d|o,t] = P[region(d)|o,t] × q[d|region(d)]
```

Pour une hausse régionale `u[o]` en pourcentage et une première arrivée `g = rentrée + délai` :

```text
Ajout_annuel[d,t,s] = somme_origines_o (u[o]/100) × F_origine[o,t,s] × R[d|o,t]   si t >= g
Ajout_places[d,T,s] = somme_t=g..T Ajout_annuel[d,t,s]
```

**Chaque cohorte utilise la répartition de son année d'arrivée, puis est cumulée.** Les cohortes déjà arrivées ne sont pas redistribuées selon la matrice des années suivantes. La hausse des places est permanente pour les cohortes suivantes ; +100 % signifie doubler chaque cohorte de référence, sans doublement exponentiel d'une année sur l'autre. Le coefficient de postes vacants n'est appliqué qu'une fois, avant le cumul.

Les cohortes supplémentaires restent cumulées sans attrition ni retraite supplémentaire jusqu'en 2070, conformément à l'hypothèse de cumul du modèle. Le taux de retraite macro sert à inférer les installations dans la matrice ; il ne constitue pas ici un modèle de survie des nouvelles cohortes. Cette simplification peut majorer les effets aux horizons lointains. Les ajouts ne sont ni multipliés par le facteur de raccord RPPS/DREES, ni effacés par une renormalisation ultérieure.

Le tableau de calibration donne pour +1 % dans une origine, à la première année d'installation `g` :

```text
Delta_metro[o,g,s] = 0,01 × F_origine[o,g,s] × somme_departements_metro R[d|o,g]
X[o,g,s] (%) = 100 × Delta_metro[o,g,s] / B_corrige[g,s]
```

`B_corrige` est le total après correction de l'internat, avant la tendance des arrivées étrangères et les hausses de places ou d'arrivées. X dépend de l'origine, de la méthode, du taux de retraite, de l'année et du scénario. La Corse n'a pas de calibration d'origine de formation exploitable dans ces sources : son curseur reste désactivé. Elle peut recevoir des diplômés des autres origines. Une création ex nihilo demanderait un nombre absolu de places et des hypothèses supplémentaires.

### Code 99 : tendance nationale ancrée en 2026 et hausse supplémentaire

Les comptes par région d'exercice des diplômés à l'étranger sont extrapolés comme les autres origines. On les agrège ensuite sur les régions d'exercice françaises, **DROM inclus**, pour obtenir le total national `T[t]`. La destination `99-COM ou étranger`, qui ne permet pas d'isoler une région française, est exclue de ce calcul. Son effectif est nul dans la série étrangère fournie. Le flux annuel est obtenu en appliquant le maximum **une seule fois au total national** :

```text
T[t] = somme_regions_francaises_j N_proj[99,j,t]
F_99[t] = max(0, T[t] - (1-r) × T[t-1])
F_99[b] = max(0, T_observe[b] - (1-r) × T_observe[b-1])     avec b = 2026
```

Le flux de 2026 utilise donc les comptes réellement observés de 2026 et 2025, et non une différence entre deux niveaux ajustés. Les flux `F_99` sont toujours non négatifs. Le volume est identique dans les quatre scénarios DREES pour les mêmes paramètres, puisqu'il vient du même historique et du même taux macro. Le choix de méthode de répartition ne modifie pas ce volume.

**Même avec le curseur à 0 %, on ajoute désormais l'évolution du flux par rapport au niveau de 2026.** Le flux annuel supplémentaire de tendance est :

```text
Delta_F[t] = F_99[t] - F_99[b]                 si t > b
Delta_F[b] = 0
```

On cumule les écarts annuels au flux constant de référence : on ne cumule pas tous les `F_99[t]`, et on ne soustrait pas le stock `T[b]` à un flux. Cette distinction stock/flux conserve une référence cohérente. Un flux futur identique à celui de 2026 produit un ajustement nul, quelle que soit la durée de projection. **Un écart peut être négatif** lorsque le flux prévu est inférieur au flux de 2026 ; cela représente moins d'arrivées que dans la référence, pas un nombre négatif d'installations.

Les proportions `P[j|99,t]` restent déterminées par la méthode sélectionnée (delta + retraites ou parts d'effectifs projetés). Comme le volume porte désormais explicitement sur la France, les proportions sont conditionnées à ce même champ :

```text
P_France[j|99,t] = P[j|99,t] / somme_regions_francaises_k P[k|99,t]
R_France[d|99,t] = P_France[region(d)|99,t] × q[d|region(d)]
```

Si les parts françaises initiales somment à zéro, on utilise les parts des effectifs français projetés, puis celles de 2026 si nécessaire. Les destinations ultramarines restent au dénominateur : leur part n'est pas réaffectée aux départements métropolitains affichés. Le tableau et le graphique du code 99 affichent cette répartition conditionnelle à une installation en France.

L'ajustement de tendance départemental est calculé en répartissant chaque écart annuel selon les proportions de son année, puis en cumulant :

```text
Ajustement_tendance[d,T] = somme_t=b+1..T Delta_F[t] × R_France[d|99,t]
Ajustement_tendance[d,b] = 0
```

Le curseur de hausse est conservé comme simulation supplémentaire. Pour un pourcentage `u` et une première année choisie `g` :

```text
Hausse_annuelle[d,t] = (u/100) × F_99[t] × R_France[d|99,t]     si t >= g
Ajout_curseur[d,T] = somme_t=g..T Hausse_annuelle[d,t]
Effectif_apres_etranger[d,T] = max(0, Effectif_avant_etranger[d,T] + Ajustement_tendance[d,T])
                             + Ajout_curseur[d,T]
```

Le changement de l'année de début du curseur n'affecte pas la tendance de base, qui commence toujours après 2026. Le curseur porte sur le flux annuel complet `F_99[t]`, pas seulement sur son écart à 2026. Aucun délai de formation ni coefficient de postes vacants français n'est appliqué. Les deux composantes sont présentées séparément dans la décomposition des effectifs. Si l'écart cumulé rendrait un effectif départemental négatif, son retrait est limité à ce stock, l'interface le signale et la décomposition affiche la correction effectivement appliquée.

Ces écarts et hausses restent cumulés sans attrition des cohortes supplémentaires, comme les ouvertures de places. **Le scénario de comparaison à flux étranger constant au niveau de 2026 est une convention**, et non une identification du flux exact déjà inclus dans la DREES. La projection DREES reste la base ; on lui applique cette correction additionnelle. Le code 99 caractérise le lieu du diplôme, pas la nationalité, les visas ou des passages de frontière directement observés.

Avec les sources actuelles et 3 % de retraite : `T[2025] = 1 999`, `T[2026] = 2 068`, donc `F_99[2026] = 128,97`. Le flux estimé vaut environ 124,033 en 2027, 148,210 en 2040 et 204,214 en 2070. Les écarts annuels à 2026 valent respectivement −4,937, +19,240 et +75,244. Avec la répartition par installations estimées et un curseur nul, l'effet cumulé en métropole vaut environ **−4,733 en 2027, −8,234 en 2030, +95,878 en 2040 et +1 479,228 en 2070**. Le plancher à zéro ne limite aucun de ces ajustements dans ce réglage initial.

## 6. Périmètre de calcul

La version actuelle utilise les effectifs de psychiatres, leur répartition par statut, les stocks d'étudiants, la mobilité diplôme/exercice et les taux nationaux de postes non pourvus. Ces derniers sont intégrés au code avec leurs sources, sans fichier supplémentaire ni téléchargement de statistiques à l'exécution. Aucun fichier CSV n'est nécessaire. L'intégration de projections démographiques territoriales est reportée à une étape ultérieure et n'entre dans aucun calcul actuel.

## 7. Interface et graphique

- « Valider » applique ensemble les paramètres du formulaire. Bouger un curseur ou cocher une case ne modifie pas à lui seul les résultats affichés. Les bandeaux indiquent les **paramètres effectivement validés**, et non une éventuelle saisie en attente.
- L'état des dynamiques est désormais affiché explicitement : activées/désactivées, première année d'effet et fenêtre d'estimation. Si elles sont désactivées, le graphique avertit que la pente estimée n'est pas appliquée.
- Les filtres indicateur/scénario/année pilotent la carte et le tableau ; leur utilisation ne relance pas la simulation. Les trois nouveaux indicateurs sont sélectionnables au même endroit que le total.
- L'échelle de couleurs est fixe entre années pour chaque indicateur d'effectifs : quantiles 2 % et 98 % sur toutes les années et tous les scénarios de la simulation validée. Pour la part salariée, elle reste entre 0 et 100 %. Les valeurs hors bornes restent dans le tableau et au survol.
- La carte conserve la projection conique conforme de Lambert, avec les parallèles 44° et 49°, un cadrage fixe incluant la Corse et le déplacement désactivé. Les anneaux des polygones GeoJSON sont orientés pour Plotly afin d'éviter un remplissage du globe. La palette bleue commence à une teinte visible sur le fond, les départements sont séparés par des contours blancs et les infobulles présentent le nom, le code et la valeur avec son unité.
- L'onglet « Dynamiques départementales » compare un ou plusieurs départements sur toute la période. Le filtre Année de la carte ne tronque pas ces courbes.
- L'onglet « Statuts d’exercice » montre les parts observées puis extrapolées et les deux comptes pour l'année sélectionnée.
- L'onglet « Diplômes et installations » expose, pour chaque origine et année, les effectifs précédents et projetés, les installations non négatives, les proportions et les éventuelles règles de repli. Il permet aussi de suivre le flux estimé du code 99.
- L'onglet « Internat national » montre la série observée, sa prolongation linéaire et les flux par scénario avant/après correction, avec les sources et les limites de calibration.

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

Tout reste dans `app.py` :

| Partie | Rôle |
| --- | --- |
| Lecteurs RPPS/DREES | Extraire et valider les sources |
| `assemble_reference` | Assembler observations et projection de base |
| `build_context` | Parts d'ancrage, pentes, modèles de statut, étudiants, mobilité et routage |
| `logistic_fit`, `fit_status_models` | Ajuster les parts salariées et les raccorder aux observations |
| `vacancy_projection` | Ajuster la tendance linéaire des postes non pourvus et la raccorder au dernier taux observé |
| `diploma_history`, `project_diploma_distribution` | Lire les comptes annuels, extrapoler les cellules et inférer les installations |
| `annual_routing` | Ventiler les parts annuelles entre départements |
| `foreign_reference_flows`, `foreign_france_distribution` | Calculer le flux étranger sur le total français, son écart à 2026 et les proportions conditionnelles à la France |
| `prepare_national_flows` | Calibrer les coefficients de cohortes, les flux et les retraits cumulés |
| `prepare_arrivals`, `accumulate_arrivals` | Calibrer puis cumuler les cohortes françaises et étrangères après leur routage annuel |
| `Control` et `Modulation` | Décrire contrôles et transformation associée |
| `modulation_registry(context)` | Ordonner les modulations et générer leurs contrôles |
| `apply_modulations` | Appliquer les fonctions sur des copies et vérifier les invariants |
| `to_projections_frame` | Restituer exactement les cinq colonnes publiques |
| `ensure_projections` | Calcul et mémorisation par session |
| `main` | Formulaire, carte, courbes et méthode |

Les répartitions annuelles sont d’abord calibrées avec la méthode et le taux de retraite validés. Ordre des effets sur le total : **dynamiques départementales → correction nationale de l’internat → tendance des arrivées étrangères → hausse supplémentaire des arrivées étrangères → places de formation → répartition par statut**. Le module `mobilite` décrit les contrôles de calibration ; sa contribution directe au stock est nulle. La tendance étrangère est un module sans contrôle d'activation, toujours appliqué. La dernière étape produit les deux comptes à partir de la part salariée extrapolée.

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

Le stock d'étudiants couvre la France ; le résultat affiché couvre la métropole. Pour ne pas retrancher un flux France entière à un total métropolitain, on reprend les poids régionaux de formation et la mobilité diplôme/exercice :

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
- deux méthodes de répartition, sensibilité au taux de retraite, cas de flux nuls et règles de repli ;
- cumul indépendant des cohortes avec les proportions de leur année d’arrivée, pour les 96 départements et les quatre scénarios ;
- indépendance du flux étranger vis-à-vis de la correction des places françaises et du délai de formation ;
- calcul du maximum après agrégation nationale, sans transformer une mobilité entre deux régions en arrivée étrangère ;
- ajustement nul en 2026, actif à curseur nul après 2026, cumul des écarts annuels signés et ajustement nul pour un flux futur constant au niveau de 2026 ;
- séparation de la tendance étrangère et de la hausse supplémentaire, y compris quand le début de cette hausse est reporté ;
- graphique en nombres de psychiatres, trait plein/tirets, référence commune aux deux vues et décomposition des effets ;
- comptes de salariés et libéraux/mixte retrouvant les totaux de chaque année et département, dont les comptes 2026 et ceux de Paris ;
- pentes logistiques sur données synthétiques, raccord exact, parts bornées et cas à 0/100 % ;
- pente nationale linéaire comparée à un ajustement indépendant, raccord au dernier taux observé, incréments annuels constants avant bornage et cas synthétiques atteignant 0 ou 100 % ;
- structure publique de cinq colonnes avec quatre indicateurs et 90 624 lignes, unité en pourcentage pour la part ;
- correction uniforme par département, premiers retraits en 2032, cohorte décalée de cinq ans, baisse récente de 2026 incluse et absence de réduction si la pente est négative ;
- avec des curseurs nuls, écart à la version précédente exactement égal à la nouvelle contribution de tendance étrangère ;
- correction appliquée une seule fois aux nouvelles cohortes de places et conservation de la décomposition complète ;
- six onglets, sélection des quatre indicateurs, choix du lieu de diplôme et affichage des unités ;
- conservation de la projection Lambert, du cadrage, des couleurs et des infobulles ; orientation correcte des 118 anneaux des 96 départements, y compris après inversion de leur sens, sans modification des limites ;
- validation explicite, absence de cumul lors des clics répétés et indépendance de deux sessions ;
- absence d'écriture des projections ou paramètres sur disque.

Ces contrôles vérifient le fonctionnement du code ; ils ne constituent pas une validation empirique des hypothèses de comportement.
