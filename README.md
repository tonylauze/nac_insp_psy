# Projections de psychiatres par département

Application Streamlit contenue dans **un seul `app.py`**. Les deux classeurs sont lus par le code ; les projections et les scénarios sont calculés **en mémoire, sans écriture de résultats sur le serveur**. La projection combine les dynamiques départementales, une correction nationale des arrivées, les ouvertures régionales de formation et une répartition entre salariés et libéraux. Les dynamiques départementales et la correction nationale sont activées par défaut, et désactivables pour comparaison. Les hausses de places sont initialement nulles.

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

Si le total de référence est nul, les effectifs sont nuls. Le total national est donc conservé **à l'étape des dynamiques**. La correction de l'internat peut ensuite le réduire, puis les places supplémentaires l'augmenter. Cette renormalisation implique qu'après activation du plancher, les trajectoires finales ne sont plus strictement linéaires, même pour les autres départements. Les observations, dont 2026, ne sont jamais lissées ni modifiées.

## 5. Places de formation par région

### Sens des curseurs et calendrier

Chaque curseur représente une **hausse relative de 0 à 100 % des places d'une région**, appliquée à chaque nouvelle cohorte à partir de la rentrée choisie. Ce n'est ni un nombre absolu de places ni une seule cohorte ponctuelle. Les curseurs sont initialisés à zéro. La rentrée initiale est **2027 par défaut**, exprimée comme **année du bac/début du parcours**, et peut être décalée dans le formulaire.

On suppose que cette hausse se transmet proportionnellement aux futurs flux de psychiatres : choix de spécialité, réussite et conditions de formation restent constants. Il s'agit d'une hypothèse d'élasticité égale à 1, pas d'un effet causal mesuré dans les classeurs.

Délai retenu : **N = 11 ans**, soit 6 ans avant le DES et un DES de psychiatrie de 5 ans. Il s'agit de la **durée nominale actuelle**, utilisée faute de moyenne empirique de durée d'études dans les sources. Les admissions tardives, redoublements, interruptions et options prolongeant les études ne sont pas simulés.

Par convention annuelle, une rentrée au niveau du bac en 2027 produit sa première cohorte active en **2038**. Ce calendrier ne correspond pas à une ouverture en troisième cycle, ni directement à une rentrée de deuxième année. Les libellés de dates du modèle DREES sont utilisés tels quels pour les stocks : cette couche de simulation n'est pas une reconstruction de son calendrier individuel.

Sources de durée consultées le 23 octobre 2026 : [parcours des études de médecine, Onisep](https://www.onisep.fr/formation/les-principaux-domaines-de-formation/les-etudes-de-sante/les-etudes-de-medecine) ; [maquette du DES de psychiatrie, 10 semestres, Légifrance](https://www.legifrance.gouv.fr/loda/article_lc/LEGIARTI000047115341/2026-07-25). Les constantes `DES_YEARS` et `STUDY_YEARS` centralisent ce choix.

### Conversion des étudiants en flux : comment X est calculé

Dans `Effectif par spécialité`, on sélectionne le statut **Étudiants en 3ème cycle**, avec les spécialités Psychiatrie et Ensemble. Le nombre d'internes est un **stock sur plusieurs cohortes**, pas un nombre d'entrants annuels.

L'application approche le flux annuel de nouveaux psychiatres par :

```text
F_reference[t,s] = S_psy[t,s] / 5
F_corrige[t,s] = F_reference[t,s] × c[t - 5]
```

`S_psy[t,s]` est le stock projeté national d'internes de psychiatrie de l'année de sortie. Cette conversion suppose des cohortes approximativement équilibrées et une durée de DES de cinq ans. Elle ne permet pas de reconstituer exactement les sorties si les cohortes changent fortement de taille. Le coefficient `c` porte sur l'année d'entrée en internat, cinq ans plus tôt ; il vaut 1 si la correction est désactivée. Sa formule figure en section 10. Le ratio brut `S_psy / S_tous` n'est pas interprété comme une probabilité individuelle de devenir psychiatre : les durées de DES diffèrent entre spécialités.

Faute d'entrées en médecine par région dans les tables exploitées, la géographie initiale des places est approchée par celle des **stocks d'internes toutes spécialités** dans `Effectif régional`, à l'ancrage :

```text
w[r,s] = S_tous_region[r,b,s] / S_tous_France[b,s]
```

Les anciennes régions sont regroupées dans les régions actuelles. Les poids restent fixes dans le futur et sont propres à chaque scénario. Ils utilisent un dénominateur France entière : les poids des seules régions métropolitaines ne sont pas renormalisés à 100 %. Les écarts d'arrondi entre somme des lignes et total régional sont tolérés à hauteur d'une unité par ligne, sans corriger les valeurs sources.

Pour une hausse `u[r]` exprimée en pourcentage :

```text
nouveaux_psy_formés[r,t,s] = (u[r]/100) × w[r,s] × F_corrige[t,s]
```

Ce flux vaut zéro avant `rentrée + 11`. Cette calibration combine donc un flux national spécialisé et un poids régional toutes spécialités ; elle suppose la même composition en spécialités dans les régions et ne prétend pas observer les internes psychiatres de chacune.

### Mobilité diplôme → exercice

Onglet RPPS **`Effectifs par lieu diplome`**, spécialité Psychiatrie, effectifs de la dernière année observée. Les codes numériques RPPS identifient les régions, pour éviter les différences d'orthographe des libellés. Pour l'origine de diplôme `r` et la région d'exercice `j` :

```text
M[j|r] = effectif_psy[diplôme=r, exercice=j, b]
         / somme_toutes_destinations effectif_psy[diplôme=r, exercice=destination, b]
```

On normalise bien **par origine de diplôme**. Les marges « Ensemble » ne sont pas additionnées aux cellules ; le code vérifie que ces cellules reconstituent chaque marge. Les destinations ultramarines et COM/étranger restent au dénominateur. Leur part n'est pas réaffectée à la métropole. Les origines étrangères ou inconnues ne sont pas des régions où un curseur ouvre des places.

Cette matrice décrit les psychiatres actuellement actifs, toutes générations confondues. Elle n'est pas une mesure longitudinale de mobilité des nouveaux diplômés. On la maintient constante et l'on assimile **région où les places sont ouvertes et région de diplôme** ; les mobilités en cours d'études ne sont pas modélisées séparément.

À l'intérieur d'une région d'exercice :

```text
q[d|j] = n[d,b] / somme_des_départements_de_j n[d,b]
```

Ces parts restent celles de la dernière année observée, y compris lorsque les dynamiques départementales sont activées.

### Cumul et impact final

Les cohortes supplémentaires sont cumulées à partir de la première année de sortie :

```text
Ajout[d,t,s] = somme_r (u[r]/100) × w[r,s] × M[region(d)|r] × q[d|region(d)]
              × somme_v=rentrée+11..t (S_psy[v,s]/5 × c[v-5])

Effectif_final = Effectif_après_dynamique_et_correction_internat + Ajout
```

Les ajouts sont des flux de personnes formées ; ils ne sont pas multipliés par le facteur de raccord RPPS/DREES. Ils viennent **après** le lissage des dynamiques et la correction du stock de référence. Le coefficient de remplissage est appliqué **une seule fois à chaque nouvelle cohorte avant cumul**. On ne multiplie pas à nouveau ce cumul par la correction du stock. Aucune renormalisation ultérieure ne les efface. L'application conserve les nombres fractionnaires comme effectifs attendus, sans arrondir les calculs.

Aucune attrition, sortie d'activité, retraite ni nouvelle migration n'est appliquée à ces cohortes supplémentaires jusqu'en 2070. Avec le calendrier par défaut, la première cohorte arrive seulement en 2038. Cette simplification concerne les ajouts ; le scénario DREES de référence conserve ses propres hypothèses de démographie professionnelle.

**X n'est donc pas une constante universelle.** Pour +1 % dans une seule région, au premier millésime de sortie `g` :

```text
Delta_metro[r,g,s] = 0,01 × w[r,s] × F_corrige[g,s] × somme_j_metropole M[j|r]
X[r,g,s] (%) = 100 × Delta_metro[r,g,s] / B_corrige[g,s]
```

`B_corrige` est le total métropolitain après correction de l'internat, avant nouvelles places. L'onglet « Calibration et hypothèses » affiche ces deux valeurs par région, pour le scénario et la rentrée validés. Aux années suivantes, le cumul des cohortes augmente l'effet sur le stock.

### Région non calibrable

Dans les fichiers actuels, la **Corse n'a ni stock régional d'internes exploitable ni origine de diplôme permettant cette calibration**. Son curseur est affiché mais désactivé, avec une explication. Cela ne signifie pas qu'il est impossible d'y ouvrir une formation : une création ex nihilo exige un nombre absolu de places et des hypothèses de diplôme/mobilité supplémentaires. Les départements corses peuvent recevoir les diplômés formés ailleurs et ont leurs propres dynamiques départementales.

## 6. Périmètre de calcul

La version actuelle utilise les effectifs de psychiatres, leur répartition par statut, les stocks d'étudiants, la mobilité diplôme/exercice et les taux nationaux de postes non pourvus. Ces derniers sont intégrés au code avec leurs sources, sans fichier supplémentaire ni téléchargement de statistiques à l'exécution. Aucun fichier CSV n'est nécessaire. L'intégration de projections démographiques territoriales est reportée à une étape ultérieure et n'entre dans aucun calcul actuel.

## 7. Interface et graphique

- « Valider » applique ensemble les paramètres du formulaire. Bouger un curseur ou cocher une case ne modifie pas à lui seul les résultats affichés. Les bandeaux indiquent les **paramètres effectivement validés**, et non une éventuelle saisie en attente.
- L'état des dynamiques est désormais affiché explicitement : activées/désactivées, première année d'effet et fenêtre d'estimation. Si elles sont désactivées, le graphique avertit que la pente estimée n'est pas appliquée.
- Les filtres indicateur/scénario/année pilotent la carte et le tableau ; leur utilisation ne relance pas la simulation. Les trois nouveaux indicateurs sont sélectionnables au même endroit que le total.
- L'échelle de couleurs est fixe entre années pour chaque indicateur d'effectifs : quantiles 2 % et 98 % sur toutes les années et tous les scénarios de la simulation validée. Pour la part salariée, elle reste entre 0 et 100 %. Les valeurs hors bornes restent dans le tableau et au survol.
- L'onglet « Dynamiques départementales » compare un ou plusieurs départements sur toute la période. Le filtre Année de la carte ne tronque pas ces courbes.
- L'onglet « Statuts d’exercice » montre les parts observées puis extrapolées et les deux comptes pour l'année sélectionnée.
- L'onglet « Internat national » montre la série observée, sa prolongation logistique et les flux par scénario avant/après correction, avec les sources et les limites de calibration.

### Deux vues, une référence commune

La vue initiale est **« Sans ouverture de places »** : elle montre la référence et la dynamique départementale effectivement appliquée, avant les ajouts de formation. La vue **« Avec les ouvertures de places »** montre le scénario complet validé. La carte présente toujours ce scénario complet, quelle que soit la vue choisie pour le graphique.

Dans les deux vues, la comparaison utilise le même total de référence, **après correction de l'internat et avant nouvelles places** :

```text
Reference[d,t,s] = p[d] × B_corrige[t,s]
Ecart_affiche[d,t,s] = effectif_de_la_vue[d,t,s] - Reference[d,t,s]
```

`B_corrige` correspond au total observé sur l'historique. La correction nationale uniforme ne crée donc pas à elle seule un écart territorial. Les ouvertures régionales ne déplacent pas cette référence. L'axe vertical, les survols et les droites ajustées sont en **nombre de psychiatres**. Une valeur négative indique moins de psychiatres que la référence, pas un effectif négatif. Toutes les courbes sont nulles à l'ancrage. Un plateau à zéro indique que le département suit la trajectoire nationale à part fixe ; son nombre de psychiatres peut évoluer.

Les observations sont en **trait plein**, le futur en **tirets**, avec raccord au dernier point observé. Une option affiche les régressions avec constante sur l'historique, en fins pointillés. Ces droites ajustées ne passent pas nécessairement par le dernier point ; la projection repart, elle, du point observé. La saturation et la renormalisation peuvent ensuite infléchir les trajectoires projetées.

Sous la courbe, le tableau distingue la **pente estimée** et la **pente appliquée avant saturation**, en psychiatres par an ; la seconde vaut zéro si la dynamique est désactivée. Un volet détaille l'année sélectionnée : **référence sans modulation + effet départemental après lissage + correction internat + ajout des places = effectif final**. Les contributions sont enregistrées pendant le calcul. La correction nationale peut réduire la pente et l'écart finalement visibles, même si la pente historique estimée reste identique.

### Lecture des trajectoires

Avec le repli national, désactiver les dynamiques maintient les écarts territoriaux à zéro en l'absence de nouvelles places, même si la correction nationale réduit les effectifs. Les ouvertures peuvent ensuite faire diverger les courbes à partir de 2038 avec le calendrier initial. Les bandeaux affichent l'état effectivement validé pour distinguer ces situations.

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
| `logistic_fit`, `fit_status_models` | Ajuster les tendances de proportions et les raccorder aux observations |
| `prepare_national_flows` | Calibrer les coefficients de cohortes, les flux et les retraits cumulés |
| `Control` et `Modulation` | Décrire contrôles et transformation associée |
| `modulation_registry(context)` | Ordonner les modulations et générer leurs contrôles |
| `apply_modulations` | Appliquer les fonctions sur des copies et vérifier les invariants |
| `to_projections_frame` | Restituer exactement les cinq colonnes publiques |
| `ensure_projections` | Calcul et mémorisation par session |
| `main` | Formulaire, carte, courbes et méthode |

Ordre actuel : **dynamiques départementales → correction nationale de l'internat → places de formation → répartition par statut**. Les trois premières transformations portent sur le total ; la dernière produit les deux comptes à partir de la part salariée extrapolée.

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

Les deux groupes reconstituent toujours le total. Les scénarios DREES, la correction nationale et les places modifient les nombres, mais utilisent la même part départementale extrapolée pour une année donnée. Les nouvelles cohortes reçoivent cette même répartition. Aucun changement individuel de statut ni répartition spécifique des jeunes diplômés n'est identifié dans les données exploitées. Il s'agit d'une prolongation descriptive, sans intervalle d'incertitude ou validation causale ; l'extrapolation jusqu'en 2070 peut amplifier les tendances historiques.

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

Le modèle national ajuste une logistique sur **2013–2026 par défaut**. L'utilisateur peut modifier l'année de début, avec au moins cinq observations, puis Valider. L'ajustement porte sur les taux, avec un poids égal par année, car les dénominateurs ne sont pas disponibles de façon homogène sur toute la période. Il utilise le même solveur que les statuts, sans demi-observation. Ce choix estime une tendance des taux publiés ; il ne suppose pas que chaque taux ancien correspond à un essai binomial unique.

La pente estimée est réancrée sur `v[b] = 37/653`, où `v` est la proportion non pourvue et `b = 2026` :

```text
v[a] = expit(logit(v[b]) + beta_v × (a-b))
c[a] = 1                                      si a <= b
c[a] = min(1, (1-v[a]) / (1-v[b]))             si a > b
```

`a` désigne l'année d'entrée en internat. Le coefficient compare le remplissage futur à un remplissage constant au niveau observé de 2026. Il reste entre 0 et 1. Si la tendance choisie s'améliore, il n'ajoute pas de psychiatres au scénario de référence : il vaut 1. Désactiver la correction impose également 1 pour toutes les cohortes, tout en laissant visible la courbe estimée.

Avec la fenêtre complète, la pente logit vaut environ **0,12033 par an**. Pour une entrée en internat en 2027, le taux extrapolé vaut 6,34 % et `c ≈ 0,9928` ; en 2050, environ 51,89 % et `c ≈ 0,5100` ; en 2070, environ 92,29 % et `c ≈ 0,0817`. **Ces valeurs lointaines illustrent la force de l'hypothèse, pas une prévision établie des comportements.** Une estimation débutant en 2019 a au contraire une pente négative avec cette série et n'entraîne aucune réduction supplémentaire. L'interface expose la fenêtre et toute la courbe pour permettre cette comparaison.

La [présentation DREES du modèle ProMédI, 6 octobre 2026](https://drees.solidarites-sante.gouv.fr/communique-de-presse-jeux-de-donnees/261006-projections-effectifs-m%C3%A9decins-moins-70-ans-activite-horizon-2070) indique que le modèle tient déjà compte des durées d'études, interruptions, retards et abandons. Sa documentation publique ne permet pas d'identifier ici le taux précis de postes non pourvus incorporé aux projections de stocks. **Le remplissage constant de 2026 est donc une convention de scénario additionnel, pas un paramètre DREES retrouvé.** Appliquer directement `1-v[a]` aux étudiants déjà inscrits soustrairait un niveau de non-remplissage potentiellement déjà intégré. Le ratio utilisé n'applique que la dégradation future relativement à cette convention ; il ne remplace pas une reconstruction du modèle DREES et ne garantit pas une correction structurelle exacte de ses hypothèses.

### De la cohorte au flux, puis au stock

Pour l'année d'arrivée `t`, le coefficient est celui de l'entrée en internat `a=t-5`. On utilise le stock DREES de troisième cycle dans **l'année d'arrivée** pour approcher le flux annuel de référence, conformément à la calibration de formation :

```text
F_reference[t,s] = S_psy[t,s] / 5
F_corrige[t,s] = F_reference[t,s] × c[t-5]
Manque_France[t,s] = F_reference[t,s] - F_corrige[t,s]
```

Les cohortes déjà entrées au plus tard en 2026 conservent leur flux de référence. Une entrée en 2027 a son premier effet sur l'effectif en **2032**, sans modifier les observations ni les arrivées de 2027–2031. Le délai de cinq ans concerne cette correction ; il reste distinct des onze ans entre ouverture de places au niveau du bac et première arrivée.

Le stock d'étudiants couvre la France ; le résultat affiché couvre la métropole. Pour ne pas retrancher un flux France entière à un total métropolitain, on reprend les poids régionaux de formation et la mobilité diplôme/exercice :

```text
m[s] = somme_origines_metropolitaines_r w[r,s] × somme_destinations_metropolitaines_j M[j|r]
Manque_metro[t,s] = Manque_France[t,s] × m[s]
Manque_cumule[t,s] = somme_v=b+1..t Manque_metro[v,s]
Retrait[t,s] = min(Manque_cumule[t,s], B[t,s])
K[t,s] = (B[t,s] - Retrait[t,s]) / B[t,s]
Effectif_apres_internat[d,t,s] = Effectif_apres_dynamique[d,t,s] × K[t,s]
```

Le facteur `m` vaut environ 0,9434 avec les fichiers fournis. Les origines ultramarines ne sont pas calibrées par les curseurs/routages métropolitains actuels : leur éventuel flux de diplômés vers la métropole n'est pas ajouté à `m`. Cette approximation de champ doit être revue si l'on souhaite une modélisation complète des mobilités ultramarines. Le correcteur estime une perte de flux issu des formations françaises ; il ne modifie pas séparément les arrivées de médecins diplômés à l'étranger.

`K` est le **même pour tous les départements** d'une année et d'un scénario : il conserve leur répartition après dynamique. La conversion nationale ne varie pas selon le statut. Si `B=0`, aucun retrait n'est possible et le coefficient est conventionnellement 1. Si le manque cumulé dépasse le stock, le retrait est plafonné au stock et l'interface avertit que les hypothèses doivent être revues.

Les cohortes manquantes restent soustraites jusqu'en 2070 **sans sorties ni retraites**, comme les cohortes supplémentaires de places restent ajoutées sans attrition. Cette simplification peut amplifier les effets à long terme. Le flux `stock/5` n'identifie ni les entrées effectives, ni les sorties exactes, ni les changements de spécialité ; c'est une approximation disponible à partir des tableaux fournis.

Exemple de contrôle, scénario Tendanciel, fenêtre 2013–2026, sans nouvelles places : en 2032, `S=2 624`, le flux France de référence vaut 524,8 et le flux corrigé environ 521,025 ; le retrait métropolitain cumulé vaut environ 3,561. En 2070, le retrait cumulé atteint environ 8 252,975 pour une référence métropolitaine de 26 736,926, soit `K ≈ 0,6913`. L'arrivée de 2070 utilise **le coefficient de la cohorte 2065**, pas celui de l'entrée 2070.

Enfin, une hausse de places pour une rentrée au bac en 2027 produit sa première arrivée en 2038. Cette cohorte utilise `c[2033]`. Le correcteur est appliqué à son flux une seule fois avant cumul, sans modifier les parts territoriales de routage. Les nombres finaux sont ensuite répartis par statut selon la section 9.

## 11. Vérifications de cette version

Les vérifications portent sur les deux classeurs fournis, des cas synthétiques et des sessions Streamlit simulées :

- lecture des observations depuis le seul onglet RPPS `Effectifs`, avec conservation exacte des valeurs ;
- retour à des pentes exprimées en psychiatres/an et comparaison avec les régressions d'effectifs ;
- ancrage observé conservé, dynamiques actives dès 2027 et conservation du total national à cette étape ;
- saturation en cas d'extrapolation négative, résultat fini et non négatif ;
- premiers ajouts de formation en 2038 avec la rentrée par défaut, avec la même calibration et le même routage ;
- graphique en nombres de psychiatres, trait plein/tirets, référence commune aux deux vues et décomposition des effets ;
- comptes de salariés et libéraux/mixte retrouvant les totaux de chaque année et département, dont les comptes 2026 et ceux de Paris ;
- pentes logistiques sur données synthétiques, raccord exact, parts bornées et cas à 0/100 % ;
- structure publique de cinq colonnes avec quatre indicateurs et 90 624 lignes, unité en pourcentage pour la part ;
- correction uniforme par département, premiers retraits en 2032, cohorte décalée de cinq ans, baisse récente de 2026 incluse et absence de réduction si la pente est négative ;
- retour aux totaux de la version précédente lorsque la correction de l'internat est désactivée ;
- correction appliquée une seule fois aux nouvelles cohortes de places et conservation de la décomposition complète ;
- cinq onglets, sélection des quatre indicateurs et affichage des unités ;
- validation explicite, absence de cumul lors des clics répétés et indépendance de deux sessions ;
- absence d'écriture des projections ou paramètres sur disque.

Ces contrôles vérifient le fonctionnement du code ; ils ne constituent pas une validation empirique des hypothèses de comportement.
