# Projections de psychiatres par département

Application Streamlit contenue dans **un seul `app.py`**. Les deux classeurs sont lus par le code ; les projections et les scénarios sont calculés **en mémoire, sans écriture de résultats sur le serveur**. Le curseur national uniforme est remplacé par des curseurs régionaux de formation. Les dynamiques départementales sont activées par défaut et peuvent être désactivées pour comparaison.

## 1. Installation et fichiers

Placer dans le même dossier du dépôt GitHub :

- `app.py`
- `requirements.txt` (inclut notamment `openpyxl>=3.1,<4`)
- `Médecins RPPS 2012-2026.xlsx`
- `DREES_projections_medecins_2026.xlsx`
- facultativement `population.csv`, décrit plus bas.

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
| `indicateur` | `Effectif de psychiatres` ou `Psychiatres pour 100 000 hab.` |
| `valeur` | Nombre réel, non négatif |

Chaque combinaison département × année × scénario × indicateur est unique. Les variables de calibration restent dans des tables internes ; elles n'ajoutent aucune colonne à cette structure publique.

Avec les fichiers fournis : **96 départements métropolitains**, observations **2012–2026**, projections **2027–2070**, quatre scénarios DREES, deux indicateurs, soit **45 312 lignes**. Les observations sont répétées à l'identique dans chaque scénario. **2026 est observé**, sans double ligne projetée. L'ancrage `b` est détecté comme la dernière année RPPS ; la borne finale dépend des projections disponibles.

Le RPPS porte sur tous les âges, sexes et modes d'exercice. Le classeur DREES déclare pour les médecins actifs un champ de **70 ans ou moins** ; les étudiants de troisième cycle constituent un statut distinct. Ces champs ne sont pas identiques. Les effectifs sont des personnes, pas des équivalents temps plein.

## 3. Observations et référence de projection

### Observations départementales

Onglets RPPS `Effectifs` et `Densités`, spécialité Psychiatrie, France métropolitaine, exercice/sexe/tranche d'âge « Ensemble ». On utilise les lignes agrégées existantes, sans additionner les sous-groupes. Les codes tels que `075 -Paris` et `02A -Corse-du-Sud` sont normalisés correctement.

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

Toutes les années observées disponibles sont utilisées : **2012–2026** actuellement, avec le même poids pour chacune. Aucun résultat futur n'entre dans l'estimation.

Notations :

```text
n[d,t] : effectif observé du département
N[t]   : somme des effectifs observés des 96 départements
p[d]   : n[d,b] / N[b]  (part départementale à l'ancrage)
e[d,t] = n[d,t] - p[d] × N[t]
```

Une régression linéaire avec constante de `e[d,t]` sur l'année estime une pente `beta[d]`, en **psychiatres par an**. C'est aussi la pente départementale moins la part 2026 de la pente nationale. Les pentes ont une somme nulle ; le code retire leur moyenne pour éliminer un éventuel résidu d'arrondi.

La constante estimée sert à estimer la pente, mais n'est pas extrapolée : le futur est **réancré exactement en 2026**, où l'écart est nul. La projection brute est :

```text
brut[d,t,s] = p[d] × B[t,s] + beta[d] × (t - b)
```

`B[t,s]` est le total métropolitain de la référence, avant ajout de places. Ce modèle extrapole une différence de croissance absolue, pas un taux composé. Il ne s'agit pas d'une estimation causale ; aucune bande d'incertitude statistique n'est produite.

**Articulation avec les régions :** lorsque les dynamiques sont activées, cette formule nationale remplace la ventilation territoriale de référence, conformément à la demande de tendance nationale + effet départemental. Si des projections psychiatriques régionales deviennent disponibles, leur somme métropolitaine fournit `B`, mais leurs profils régionaux ne sont pas ajoutés une deuxième fois. Désactiver les dynamiques rétablit la ventilation régionale d'origine. Avec le repli national actuel, la référence utilise déjà `p[d] × B[t,s]`.

### Saturation lissée près de zéro

Seuil `epsilon = 1` psychiatre, réglable par `SMOOTH_FLOOR` :

```text
f(x) = x                            si x >= epsilon
f(x) = epsilon × exp(x/epsilon - 1)  sinon
```

La valeur et la dérivée sont continues au seuil ; la fonction tend vers zéro pour les extrapolations très négatives. Les nombres très faibles peuvent devenir exactement zéro par sous-flux numérique. Aucun effectif négatif n'est produit.

Le lissage relève les valeurs brutes les plus faibles. Pour ne pas créer artificiellement de psychiatres à l'échelle nationale :

```text
dynamique[d,t,s] = f(brut[d,t,s]) × B[t,s] / somme_d f(brut[d,t,s])
```

Si le total de référence est nul, les effectifs sont nuls. Le total national est donc conservé avant l'ajout des places. Cette renormalisation implique qu'après activation du plancher, les trajectoires finales ne sont plus strictement linéaires, même pour les autres départements. Les observations, dont 2026, ne sont jamais lissées ni modifiées.

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
F[t,s] = S_psy[t,s] / 5
```

`S_psy[t,s]` est le stock projeté national d'internes de psychiatrie de l'année de sortie. Cette conversion suppose des cohortes approximativement équilibrées et une durée de DES de cinq ans. Elle ne permet pas de reconstituer exactement les sorties si les cohortes changent fortement de taille. Le ratio brut `S_psy / S_tous` n'est pas interprété comme une probabilité individuelle de devenir psychiatre : les durées de DES diffèrent entre spécialités.

Faute d'entrées en médecine par région dans les tables exploitées, la géographie initiale des places est approchée par celle des **stocks d'internes toutes spécialités** dans `Effectif régional`, à l'ancrage :

```text
w[r,s] = S_tous_region[r,b,s] / S_tous_France[b,s]
```

Les anciennes régions sont regroupées dans les régions actuelles. Les poids restent fixes dans le futur et sont propres à chaque scénario. Ils utilisent un dénominateur France entière : les poids des seules régions métropolitaines ne sont pas renormalisés à 100 %. Les écarts d'arrondi entre somme des lignes et total régional sont tolérés à hauteur d'une unité par ligne, sans corriger les valeurs sources.

Pour une hausse `u[r]` exprimée en pourcentage :

```text
nouveaux_psy_formés[r,t,s] = (u[r]/100) × w[r,s] × F[t,s]
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
              × somme_v=rentrée+11..t (S_psy[v,s]/5)

Effectif_final = Effectif_après_dynamique + Ajout
```

Les ajouts sont des flux de personnes formées ; ils ne sont pas multipliés par le facteur de raccord RPPS/DREES. Ils viennent **après** le lissage des dynamiques ; aucune renormalisation ultérieure ne les efface. L'application conserve les nombres fractionnaires comme effectifs attendus, sans arrondir les calculs.

Aucune attrition, sortie d'activité, retraite ni nouvelle migration n'est appliquée à ces cohortes supplémentaires jusqu'en 2070. Avec le calendrier par défaut, la première cohorte arrive seulement en 2038. Cette simplification concerne les ajouts ; le scénario DREES de référence conserve ses propres hypothèses de démographie professionnelle.

**X n'est donc pas une constante universelle.** Pour +1 % dans une seule région, au premier millésime de sortie `g` :

```text
Delta_metro[r,g,s] = 0,01 × w[r,s] × S_psy[g,s]/5 × somme_j_metropole M[j|r]
X[r,g,s] (%) = 100 × Delta_metro[r,g,s] / B[g,s]
```

L'onglet « Calibration et hypothèses » affiche ces deux valeurs par région, pour le scénario et la rentrée validés. Aux années suivantes, le cumul des cohortes augmente l'effet sur le stock.

### Région non calibrable

Dans les fichiers actuels, la **Corse n'a ni stock régional d'internes exploitable ni origine de diplôme permettant cette calibration**. Son curseur est affiché mais désactivé, avec une explication. Cela ne signifie pas qu'il est impossible d'y ouvrir une formation : une création ex nihilo exige un nombre absolu de places et des hypothèses de diplôme/mobilité supplémentaires. Les départements corses peuvent recevoir les diplômés formés ailleurs et ont leurs propres dynamiques départementales.

## 6. Population et densités

Les densités historiques de psychiatrie sont conservées telles que publiées. Pour estimer une population départementale de référence, le code utilise les lignes RPPS **tous médecins** :

```text
Population[d,t] ≈ effectif_tous_médecins[d,t] / densité_tous_médecins[d,t] × 100 000
```

Ce choix limite l'effet des arrondis par rapport à une reconstruction fondée sur les seuls psychiatres. Il reste une approximation. Sans autre source, la population de chaque département est maintenue au niveau de l'ancrage jusqu'à la fin de la projection.

Un fichier d'entrée facultatif `population.csv` peut fournir exactement les colonnes `dep,annee,population`, dans cet ordre, avec une population positive et une clé unique par département et année. Il doit couvrir toutes les cellules projetées ; un manque ne déclenche pas de remplissage silencieux. Cette population est commune aux scénarios. Ce fichier d'entrée n'est pas un export des résultats.

Après **toutes** les modulations : `densité = effectif / population × 100 000`. Les densités standardisées DREES ne sont pas utilisées comme densités brutes.

## 7. Interface et graphique

- « Valider » applique ensemble les paramètres du formulaire. Bouger un curseur seul ne modifie pas les résultats affichés.
- Les filtres indicateur/scénario/année pilotent la carte et le tableau ; leur utilisation ne relance pas la simulation.
- L'échelle de couleurs est fixe entre années pour l'indicateur choisi : quantiles 2 % et 98 % sur toutes les années et tous les scénarios de la simulation validée. Les valeurs hors bornes restent dans le tableau et au survol.
- L'onglet « Dynamiques départementales » compare un ou plusieurs départements sur toute la période, dans le scénario sélectionné. Le filtre Année de la carte ne tronque pas ces courbes.
- L'ordonnée du graphique est `effectif[d,t] - p[d] × total_métropolitain[t]`. Au-dessus de zéro, le département dépasse la trajectoire à part 2026 constante ; en dessous, il reste en deçà. Toutes les courbes sont nulles à l'ancrage.
- Les observations sont en **trait plein** ; le futur en **pointillés**, avec le point d'ancrage commun pour assurer le raccord. Le futur inclut les paramètres validés, donc les ajouts de places lorsqu'ils arrivent. La référence nationale du graphique utilise également le total final avec ces ajouts.
- Le tableau sous le graphique donne la pente historique estimée, pas la dérivée de la courbe finale après lissage et ajout de cohortes.

## 8. Sessions, recalcul et architecture modulaire

`st.cache_data(persist=None)` conserve seulement la référence non modulée et les calibrations communes **en mémoire**, sans persistance disque. Streamlit restitue des copies indépendantes. Les paramètres et les résultats simulés sont conservés uniquement dans `st.session_state`, propre à chaque session. Aucun cache de ressources partagé ne contient un scénario utilisateur.

Une empreinte du contenu des deux classeurs, du code et de l'éventuel fichier population invalide la référence lorsqu'une source change. Un changement de code/source réinitialise les paramètres aux valeurs par défaut, notamment pour éviter de réutiliser des contrôles devenus incompatibles. Un nouveau navigateur/session démarre aussi avec les paramètres initiaux.

La simulation est recalculée au premier accès, après changement de sources ou lors d'une validation modifiant réellement les paramètres. Elle repart toujours de la référence non modulée : deux validations identiques ne cumulent pas deux fois les effets. L'état des résultats est remplacé seulement après réussite complète du calcul ; des paramètres invalides ne remplacent pas la dernière simulation.

Tout reste dans `app.py` :

| Partie | Rôle |
| --- | --- |
| Lecteurs RPPS/DREES | Extraire et valider les sources |
| `assemble_reference` | Assembler observations et projection de base |
| `build_context` | Parts d'ancrage, pentes, étudiants, mobilité et routage |
| `Control` et `Modulation` | Décrire contrôles et transformation associée |
| `modulation_registry(context)` | Ordonner les modulations et générer leurs contrôles |
| `apply_modulations` | Appliquer les fonctions sur des copies et vérifier les invariants |
| `to_projections_frame` | Restituer exactement les cinq colonnes publiques |
| `ensure_projections` | Calcul et mémorisation par session |
| `main` | Formulaire, carte, courbes et méthode |

Ordre actuel : **dynamiques départementales → places de formation → recalcul des densités**.

Pour ajouter une modulation, écrire une fonction `ma_modulation(projected, parameters, context)` et ajouter une entrée `Modulation` dans `modulation_registry`. Les contrôles peuvent être `slider`, `checkbox`, `number` ou `select`. Le formulaire et sa validation sont communs. La fonction peut cibler département × année × scénario ; elle doit restituer les mêmes lignes, index et colonnes. Seules les colonnes internes `effectif` et `population` peuvent être modifiées ; `densite` est recalculée par le moteur. Les observations ne sont pas transmises comme lignes modifiables aux fonctions.

Les entrées et le contexte sont copiés. Le moteur refuse les clés altérées, les effectifs négatifs/non finis, les populations non positives et les paramètres hors bornes. Un contrôle désactivé ne peut pas être activé par l'injection d'un paramètre. L'ancien curseur national n'est plus enregistré.

## 9. Vérifications de cette version

Les tests de développement utilisent les deux classeurs fournis ainsi que des cas extrêmes et des sessions Streamlit simulées. Ils vérifient le raccord observé, le délai, la conservation du total par les dynamiques, le cumul des cohortes, le routage à parts départementales fixes, la non-négativité, les années de calibration manquantes, la structure publique, l'absence de cumul entre validations et l'indépendance des sessions. Ils ne constituent pas une validation empirique des hypothèses de comportement.
