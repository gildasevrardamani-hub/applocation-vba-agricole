# Règles de calcul — AGRIECO PRO

## Statut du document

Ce document est le manuel officiel des règles de calcul d'AGRIECO PRO. Il prépare le futur moteur économique, mais ne contient aucun code VBA, aucun UserForm, aucun classeur Excel et aucune implémentation technique.

Le document n'est pas limité à la V1. Les règles indispensables à la première version sont indiquées comme **V1**. Les règles utiles pour l'architecture évolutive, mais non obligatoires dans la première version, sont indiquées comme **préparées pour versions futures**.

## Convention de description d'une règle

Chaque règle métier ou économique doit être interprétée selon la grille suivante :

- **Définition** : signification exacte de la règle ou de l'indicateur.
- **Objectif** : raison de son calcul.
- **Variables utilisées** : données nécessaires.
- **Formule mathématique** : formule de référence indépendante de toute technologie.
- **Explication détaillée** : interprétation fonctionnelle.
- **Exemple numérique complet** : cas chiffré reproductible.
- **Cas particuliers** : adaptations ou exceptions.
- **Règles de validation** : contrôles avant calcul.
- **Cas d'erreur** : situations où le calcul est bloqué, averti ou non applicable.
- **Impact sur les autres indicateurs** : indicateurs dépendants.

---

# 1. Principes généraux

## 1.1 Cohérence des données

### Définition

La cohérence des données signifie que les informations saisies et calculées ne doivent pas se contredire entre elles : une quantité vendue ne doit pas dépasser la quantité disponible, une superficie validée ne doit pas être nulle pour une culture, et une charge ne doit pas avoir de montant négatif sauf règle explicitement documentée.

### Objectif

Garantir que les résultats économiques reposent sur des données crédibles et contrôlables.

### Variables utilisées

- Quantités produites.
- Quantités vendues.
- Stocks entrants et sortants.
- Pertes.
- Superficie.
- Effectif.
- Prix unitaires.
- Charges.
- Dates ou périodes.

### Formule mathématique

```text
Quantité disponible = Stock entrant + Production récoltée - Pertes
Quantité utilisée = Quantité vendue + Autoconsommation + Stock sortant
Écart physique = Quantité disponible - Quantité utilisée
```

Un compte est physiquement cohérent si :

```text
Écart physique >= 0
```

ou si l'écart est justifié par une observation validée.

### Explication détaillée

La cohérence ne signifie pas que toutes les données doivent être parfaites. Elle impose que les contradictions majeures soient signalées et que les calculs critiques ne soient pas validés sans justification.

### Exemple numérique complet

Un producteur déclare :

- stock entrant : 100 kg ;
- production récoltée : 1 000 kg ;
- pertes : 80 kg ;
- ventes : 900 kg ;
- autoconsommation : 50 kg ;
- stock sortant : 70 kg.

Calcul :

```text
Quantité disponible = 100 + 1 000 - 80 = 1 020 kg
Quantité utilisée = 900 + 50 + 70 = 1 020 kg
Écart physique = 1 020 - 1 020 = 0 kg
```

Le compte est cohérent.

### Cas particuliers

- Une production nulle est autorisée si elle correspond à un échec de production documenté.
- Une récolte partielle peut coexister avec un stock sur pied si la spéculation le justifie.
- Les cultures pérennes peuvent avoir des années sans production.

### Règles de validation

- Les quantités physiques ne peuvent pas être négatives.
- Les prix unitaires ne peuvent pas être négatifs.
- Les superficies de comptes de cultures validés doivent être strictement positives.
- Les effectifs ou biomasses de comptes d'élevage validés doivent être strictement positifs, sauf règle spécifique.

### Cas d'erreur

- Stock sortant négatif : erreur bloquante.
- Pertes supérieures à la production disponible sans justification : erreur bloquante pour validation.
- Quantités vendues supérieures à la quantité disponible : erreur bloquante ou correction requise.

### Impact sur les autres indicateurs

La cohérence physique conditionne le calcul fiable des rendements, coûts unitaires, marges, produit brut, stocks, pertes et alertes.

## 1.2 Traçabilité

### Définition

La traçabilité signifie que chaque résultat calculé doit pouvoir être relié aux données saisies, aux hypothèses utilisées, à la période, à la spéculation et à la version des paramètres de calcul.

### Objectif

Permettre la vérification, l'audit, la correction et la reproduction des résultats.

### Variables utilisées

- Identifiant du producteur.
- Identifiant de l'exploitation.
- Identifiant de la parcelle ou atelier.
- Identifiant de la spéculation.
- Campagne, exercice et cycle.
- Version des paramètres.
- Date de calcul.
- Statut du compte.

### Formule mathématique

La traçabilité n'est pas un indicateur financier, mais une condition logique :

```text
Résultat traçable = Données sources + Paramètres + Version + Date de calcul + Statut
```

### Explication détaillée

Un rapport ne doit jamais afficher un résultat isolé sans indiquer les hypothèses majeures : superficie, quantité, prix, charges, cycle et mode financier ou économique.

### Exemple numérique complet

Un bénéfice de 250 000 FCFA doit pouvoir être expliqué par :

```text
Produit brut = 900 000 FCFA
Charges totales = 650 000 FCFA
Bénéfice = 900 000 - 650 000 = 250 000 FCFA
```

Le rapport doit indiquer la campagne, la spéculation, le producteur et la date de calcul.

### Cas particuliers

- Si les pondérations IPE changent, le score historique doit conserver les anciennes pondérations.
- Si un compte validé est corrigé, l'ancien résultat doit rester consultable ou archivé.

### Règles de validation

- Aucun compte validé ne doit être sans producteur, exploitation, campagne, spéculation et cycle.
- Les paramètres critiques doivent être historisés ou recopiés dans le résultat calculé.

### Cas d'erreur

- Résultat sans compte source : erreur bloquante.
- Compte sans période : validation interdite.

### Impact sur les autres indicateurs

La traçabilité influence l'audit, les comparaisons, les rapports, l'historique, l'IPE et les corrections.

## 1.3 Reproductibilité

### Définition

La reproductibilité signifie qu'un même jeu de données et les mêmes paramètres doivent toujours produire les mêmes résultats.

### Objectif

Garantir que le moteur économique peut être implémenté en VBA, SQL, Web ou Android avec des résultats identiques.

### Variables utilisées

- Données sources.
- Paramètres de calcul.
- Règles d'arrondi.
- Version des formules.

### Formule mathématique

```text
Résultat = Fonction(Données sources, Paramètres, Version règles)
```

### Explication détaillée

Les formules ne doivent pas dépendre de cellules cachées, de manipulations manuelles ou d'interprétations implicites.

### Exemple numérique complet

Si le produit brut est 1 000 000 FCFA et les charges totales 750 000 FCFA, le bénéfice doit toujours être :

```text
Bénéfice = 1 000 000 - 750 000 = 250 000 FCFA
```

### Cas particuliers

- Les arrondis peuvent différer selon les plateformes ; la règle d'arrondi doit donc être définie.
- Les conversions d'unités doivent conserver le coefficient utilisé.

### Règles de validation

- Les paramètres de calcul doivent être enregistrés avant calcul.
- Les montants doivent être arrondis selon une règle unique.

### Cas d'erreur

- Paramètre de conversion manquant : calcul bloqué pour l'unité concernée.
- Version de pondération IPE inconnue : score IPE non calculable.

### Impact sur les autres indicateurs

Tous les indicateurs économiques dépendent de la reproductibilité.

## 1.4 Transparence

### Définition

La transparence impose que chaque indicateur affiché soit compréhensible par un conseiller agricole et explicable à un producteur.

### Objectif

Faciliter l'appropriation de l'outil et éviter les décisions fondées sur des calculs opaques.

### Variables utilisées

- Libellés des indicateurs.
- Hypothèses de calcul.
- Sources des données.
- Paramètres utilisés.

### Formule mathématique

La transparence se traduit par une règle documentaire :

```text
Indicateur affiché = Valeur + Formule + Hypothèses principales + Période + Unité
```

### Explication détaillée

Un rapport doit pouvoir expliquer pourquoi une activité est rentable, déficitaire ou risquée.

### Exemple numérique complet

Pour une marge brute de 450 000 FCFA :

```text
Marge brute = Produit brut - Charges variables
Marge brute = 1 200 000 - 750 000 = 450 000 FCFA
```

### Cas particuliers

- Certains indicateurs avancés comme VAN et TRI peuvent être préparés pour versions futures, mais doivent rester documentés.

### Règles de validation

- Chaque indicateur majeur doit avoir un libellé, une unité et une formule officielle.

### Cas d'erreur

- Indicateur sans unité : alerte de documentation.
- Indicateur calculé avec donnée estimée non signalée : alerte de qualité.

### Impact sur les autres indicateurs

La transparence influence l'acceptation des rapports et les recommandations.

## 1.5 Séparation entre données saisies et données calculées

### Définition

Les données saisies sont fournies par l'utilisateur ou un référentiel. Les données calculées sont produites par les règles de calcul.

### Objectif

Éviter que l'utilisateur modifie directement un résultat calculé et rompe la cohérence du compte.

### Variables utilisées

- Données saisies : quantité, prix, charge, date, unité.
- Données calculées : montant, produit brut, marge, ROI, IPE.

### Formule mathématique

```text
Montant ligne = Quantité saisie × Prix unitaire saisi
```

### Explication détaillée

L'utilisateur saisit la quantité et le prix. Le montant doit être calculé automatiquement et non saisi manuellement, sauf cas d'ajustement documenté.

### Exemple numérique complet

```text
Quantité vendue = 2 000 kg
Prix unitaire = 250 FCFA/kg
Montant vente = 2 000 × 250 = 500 000 FCFA
```

### Cas particuliers

- Si seul le montant total est connu, le compte doit signaler que la quantité ou le prix est estimé ou manquant.

### Règles de validation

- Une valeur calculée ne doit pas être modifiée manuellement sans statut de correction.
- Les données estimées doivent être marquées comme telles.

### Cas d'erreur

- Montant incohérent avec quantité et prix : recalcul ou alerte.

### Impact sur les autres indicateurs

Cette séparation affecte tous les montants agrégés, la qualité des données et l'audit.

## 1.6 Distinction entre flux physiques, financiers et économiques

### Définition

- **Flux physique** : mouvement de quantité réelle, par exemple kg récoltés, têtes vendues, sacs stockés.
- **Flux financier** : encaissement ou décaissement monétaire réel.
- **Flux économique** : valeur économique incluant aussi les éléments valorisés sans paiement, comme autoconsommation ou main-d'œuvre familiale.

### Objectif

Produire séparément le résultat financier et le résultat économique.

### Variables utilisées

- Quantités physiques.
- Paiements réels.
- Valeurs de marché.
- Charges valorisées.
- Autoconsommation.
- Main-d'œuvre familiale.

### Formule mathématique

```text
Résultat financier = Encaissements réels d'exploitation - Décaissements réels d'exploitation - Frais financiers retenus
Résultat économique = Produits économiques - Charges économiques
```

### Explication détaillée

Un producteur peut ne pas payer la main-d'œuvre familiale, mais cette main-d'œuvre a une valeur économique. Le résultat financier mesure la trésorerie d'exploitation ; le résultat économique mesure la rentabilité complète.

### Exemple numérique complet

```text
Ventes encaissées = 800 000 FCFA
Autoconsommation valorisée = 100 000 FCFA
Charges payées = 500 000 FCFA
Main-d'œuvre familiale valorisée = 150 000 FCFA
Résultat financier = 800 000 - 500 000 = 300 000 FCFA
Résultat économique = (800 000 + 100 000) - (500 000 + 150 000) = 250 000 FCFA
```

### Cas particuliers

- Les crédits reçus ne sont pas des produits d'exploitation.
- Les remboursements de capital ne sont pas des charges de production.
- Les intérêts sont présentés séparément comme frais financiers.

### Règles de validation

- Chaque flux doit être catégorisé comme physique, financier, économique ou mixte.
- Les flux valorisés doivent avoir une méthode de valorisation.

### Cas d'erreur

- Crédit enregistré comme vente : erreur de classification.
- Main-d'œuvre familiale non typée : alerte de qualité.

### Impact sur les autres indicateurs

Cette distinction affecte produit brut, charges, marges, résultats, trésorerie, IPE et rapports.

## 1.7 Moteur économique commun aux modes individuel et organisation

### Définition

AGRIECO PRO doit utiliser un moteur économique unique pour le mode individuel et le mode conseiller ou organisation. Les règles de calcul ne changent pas selon le profil utilisateur ; seuls changent le parcours, le niveau de détail, les agrégations et la présentation.

### Objectif

Garantir que le même compte d'exploitation produit les mêmes résultats, qu'il soit saisi par un entrepreneur agricole pour son exploitation ou par un conseiller pour un producteur accompagné.

### Variables utilisées

- Mode d'utilisation : individuel, conseiller ou organisation.
- Niveau de saisie : simplifié ou détaillé.
- Périmètre d'analyse : activité personnelle, producteur, exploitation, portefeuille ou organisation.
- Données économiques communes : charges, produits, stocks, trésorerie, investissements, dettes, créances, cycles et objectifs.

### Formule mathématique

```text
Résultat économique = Fonction commune(Données économiques, Paramètres de calcul)
Présentation = Fonction d'affichage(Résultat économique, Mode utilisateur, Niveau de détail)
```

### Explication détaillée

Le mode individuel peut afficher « Mes dépenses », « Mes ventes » ou « Ma trésorerie ». Le mode organisation peut afficher « Dépenses des bénéficiaires », « Ventes consolidées » ou « Statistiques par zone ». Ces libellés et agrégations ne doivent pas modifier les formules de produit brut, charges, marge, trésorerie, rentabilité, IPE ou IQD.

### Exemple numérique complet

Un entrepreneur agricole saisit pour la tomate :

```text
Produit brut = 1 200 000 FCFA
Charges économiques = 850 000 FCFA
Résultat économique = 350 000 FCFA
```

Un conseiller saisit exactement les mêmes données pour un producteur accompagné. Le résultat doit rester :

```text
Résultat économique = 1 200 000 - 850 000 = 350 000 FCFA
```

La différence porte uniquement sur le rapport : personnel pour l'entrepreneur, consolidable pour l'organisation.

### Cas particuliers

- En saisie simplifiée, certaines données peuvent être estimées par référentiel ; l'IQD doit alors signaler une précision moindre.
- En mode organisation, les résultats peuvent être agrégés par producteur, zone, conseiller, campagne ou spéculation.

### Règles de validation

- Les formules ne doivent jamais dépendre du mode utilisateur.
- Les agrégations organisationnelles doivent provenir de résultats individuels déjà calculés et traçables.
- Les comptes simplifiés et détaillés doivent indiquer leur niveau de saisie.

### Cas d'erreur

- Deux résultats différents pour les mêmes données sources et paramètres : erreur critique de reproductibilité.
- Comparaison de comptes simplifiés et détaillés sans indication IQD : alerte de qualité.

### Impact sur les autres indicateurs

Cette règle impacte tous les indicateurs, notamment les tableaux de bord personnels, les consolidations organisationnelles, les simulations, l'IPE, l'IQD et les rapports.

---

# 2. Gestion des quantités

## 2.1 Superficie

### Définition

La superficie est la surface utilisée par une culture ou une spéculation végétale, exprimée en hectare comme unité de référence.

### Objectif

Calculer les rendements, coûts par hectare, revenus par hectare et bénéfices par hectare.

### Variables utilisées

- Superficie saisie.
- Unité de superficie saisie.
- Coefficient de conversion vers hectare.

### Formule mathématique

```text
Superficie_ha = Superficie_saisie × Coefficient_conversion_vers_ha
```

### Explication détaillée

Toutes les superficies doivent être converties en hectares pour permettre les comparaisons. Les unités locales sont acceptables si leur conversion est validée.

### Exemple numérique complet

```text
Superficie saisie = 2,5 ha
Coefficient = 1
Superficie_ha = 2,5 × 1 = 2,5 ha
```

Si l'utilisateur saisit 25 000 m² :

```text
Coefficient m² vers ha = 0,0001
Superficie_ha = 25 000 × 0,0001 = 2,5 ha
```

### Cas particuliers

- Une parcelle partagée entre plusieurs cultures doit utiliser une superficie affectée par spéculation.
- Les cultures associées nécessitent une règle d'affectation de superficie ou une analyse spécifique.

### Règles de validation

- Superficie strictement positive pour un compte de culture validé.
- Coefficient de conversion obligatoire si unité différente de l'hectare.

### Cas d'erreur

- Superficie nulle dans un compte validé : erreur bloquante.
- Unité inconnue : calcul bloqué.

### Impact sur les autres indicateurs

Impact direct sur rendement, coût par hectare, revenu par hectare, bénéfice par hectare et rentabilité par hectare.

## 2.2 Effectifs et biomasse

### Définition

L'effectif est le nombre d'animaux ou d'unités biologiques engagés dans une activité d'élevage. La biomasse est la masse totale vivante, notamment utile en pisciculture.

### Objectif

Calculer les coûts par tête, produits par tête, mortalité, productivité et résultats d'élevage.

### Variables utilisées

- Effectif initial.
- Effectif acheté ou né.
- Mortalité.
- Effectif vendu.
- Effectif final.
- Poids moyen ou biomasse.

### Formule mathématique

```text
Effectif disponible = Effectif initial + Entrées - Mortalité
Effectif final théorique = Effectif disponible - Effectif vendu - Autoconsommation - Pertes
Taux mortalité = Mortalité / (Effectif initial + Entrées) × 100
```

### Explication détaillée

Les comptes d'élevage doivent vérifier que les sorties d'animaux ne dépassent pas les effectifs disponibles.

### Exemple numérique complet

```text
Effectif initial = 500 poussins
Entrées = 0
Mortalité = 25
Vendus = 460
Autoconsommation = 5
Effectif final théorique = 500 - 25 - 460 - 5 = 10 poulets
Taux mortalité = 25 / 500 × 100 = 5 %
```

### Cas particuliers

- En pisciculture, l'effectif peut être moins fiable que la biomasse.
- Pour les naissances, l'effectif peut augmenter pendant le cycle.

### Règles de validation

- Effectif initial strictement positif pour un élevage validé, sauf suivi uniquement en biomasse validée.
- Mortalité non négative.
- Ventes et pertes non négatives.

### Cas d'erreur

- Effectif vendu supérieur à effectif disponible : erreur bloquante.
- Taux de mortalité supérieur à 100 % : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur coût par tête, produit par tête, productivité, rendement zootechnique, risque et alertes.

## 2.3 Unités et conversions

### Définition

Une unité est une mesure physique ou monétaire. Une conversion transforme une valeur d'une unité source vers une unité de référence.

### Objectif

Comparer et agréger des données saisies dans des unités différentes.

### Variables utilisées

- Unité source.
- Unité cible.
- Coefficient de conversion.
- Quantité source.

### Formule mathématique

```text
Quantité_cible = Quantité_source × Coefficient_conversion
```

### Explication détaillée

Les conversions doivent être paramétrables et traçables. Certaines unités locales varient selon le produit ; le coefficient doit alors dépendre de la spéculation ou du produit.

### Exemple numérique complet

```text
Quantité source = 20 sacs de maïs
Coefficient sac maïs vers kg = 100 kg/sac
Quantité cible = 20 × 100 = 2 000 kg
```

### Cas particuliers

- Un sac de riz peut avoir un poids différent d'un sac de maïs.
- Une caisse de tomate peut varier selon le marché.

### Règles de validation

- Toute conversion doit avoir un coefficient validé.
- La conversion doit indiquer sa source ou son référentiel.

### Cas d'erreur

- Coefficient manquant : calcul bloqué.
- Coefficient nul ou négatif : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur rendements, coûts unitaires, ventes, stocks et comparaisons.

## 2.4 Rendement

### Définition

Le rendement mesure la quantité produite par unité de surface, d'effectif, de biomasse ou autre base physique.

### Objectif

Évaluer la performance technique d'une spéculation.

### Variables utilisées

- Quantité produite vendable.
- Superficie en hectare.
- Effectif initial ou moyen.
- Biomasse initiale ou finale.

### Formule mathématique

Pour une culture :

```text
Rendement_ha = Quantité récoltée / Superficie_ha
```

Pour un élevage par tête :

```text
Production_par_tête = Quantité produite / Effectif initial
```

### Explication détaillée

Le rendement peut être calculé brut ou net selon que les pertes sont incluses ou exclues.

### Exemple numérique complet

```text
Quantité récoltée = 6 000 kg
Superficie = 2 ha
Rendement = 6 000 / 2 = 3 000 kg/ha
```

### Cas particuliers

- Production nulle : rendement nul, mais coût unitaire non calculable si divisé par production.
- Culture pérenne en année improductive : rendement non applicable ou égal à zéro selon phase.

### Règles de validation

- Superficie positive pour rendement par hectare.
- Effectif positif pour rendement par tête.

### Cas d'erreur

- Division par zéro : indicateur non applicable.
- Rendement extrêmement élevé : alerte de cohérence.

### Impact sur les autres indicateurs

Impact sur coût de production, alertes, IPE, comparaison et diagnostic.

## 2.5 Plusieurs récoltes

### Définition

Plusieurs récoltes correspondent à plusieurs sorties physiques de production pendant un même cycle.

### Objectif

Mesurer correctement la production totale d'un cycle lorsque la récolte est fractionnée.

### Variables utilisées

- Quantité récoltée par récolte.
- Date de récolte.
- Qualité ou catégorie.
- Pertes associées.

### Formule mathématique

```text
Production totale récoltée = Somme(Quantité récolte i)
```

### Explication détaillée

Chaque récolte est une ligne physique. Les ventes peuvent être liées ou non à une récolte spécifique.

### Exemple numérique complet

```text
Récolte 1 = 800 kg
Récolte 2 = 1 200 kg
Récolte 3 = 1 000 kg
Production totale = 800 + 1 200 + 1 000 = 3 000 kg
```

### Cas particuliers

- Récoltes de qualités différentes avec prix différents.
- Récolte partielle laissant un stock sur pied.

### Règles de validation

- Dates de récolte dans la période du cycle ou justification.
- Quantités non négatives.

### Cas d'erreur

- Récolte saisie sans cycle : erreur bloquante.
- Quantité négative : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur production totale, rendement, ventes, stocks, pertes et coût unitaire.

## 2.6 Plusieurs ventes

### Définition

Plusieurs ventes correspondent à plusieurs transactions avec des quantités, dates, clients ou prix différents.

### Objectif

Calculer précisément le chiffre d'affaires et le prix moyen pondéré.

### Variables utilisées

- Quantité vendue par vente.
- Prix unitaire par vente.
- Date de vente.
- Montant par vente.

### Formule mathématique

```text
Montant vente i = Quantité vendue i × Prix unitaire i
Chiffre d'affaires = Somme(Montant vente i)
Prix moyen pondéré = Chiffre d'affaires / Somme(Quantités vendues)
```

### Explication détaillée

Le prix moyen simple ne doit pas être utilisé si les quantités vendues diffèrent. Il faut utiliser un prix moyen pondéré.

### Exemple numérique complet

```text
Vente 1 = 500 kg × 250 FCFA = 125 000 FCFA
Vente 2 = 800 kg × 220 FCFA = 176 000 FCFA
Vente 3 = 700 kg × 300 FCFA = 210 000 FCFA
Chiffre d'affaires = 511 000 FCFA
Quantité vendue = 2 000 kg
Prix moyen pondéré = 511 000 / 2 000 = 255,5 FCFA/kg
```

### Cas particuliers

- Vente à crédit : produit économique reconnu, encaissement différé en trésorerie.
- Avance reçue : flux de trésorerie distinct de la vente finale.

### Règles de validation

- Quantité vendue positive.
- Prix unitaire non négatif.
- Total vendu inférieur ou égal à quantité disponible, sauf justification.

### Cas d'erreur

- Vente sans produit ou spéculation : erreur bloquante.
- Prix nul : alerte sauf don, perte ou transfert justifié.

### Impact sur les autres indicateurs

Impact sur chiffre d'affaires, produit brut, prix moyen, marge, rentabilité, trésorerie et alertes prix.

## 2.7 Stocks

### Définition

Un stock représente une quantité de produit disponible au début ou à la fin d'une période ou d'un cycle.

### Objectif

Éviter de confondre production, vente immédiate, conservation et variation de patrimoine.

### Variables utilisées

- Stock entrant.
- Production.
- Ventes.
- Autoconsommation.
- Pertes.
- Stock sortant.
- Prix de valorisation du stock.

### Formule mathématique

```text
Stock sortant théorique = Stock entrant + Production - Ventes - Autoconsommation - Pertes
Variation stock = Stock sortant - Stock entrant
Valeur stock sortant = Stock sortant × Prix valorisation stock
```

### Explication détaillée

Le stock sortant peut être valorisé dans le produit économique selon les règles retenues. Il doit rester séparé des ventes encaissées.

### Exemple numérique complet

```text
Stock entrant = 100 kg
Production = 1 000 kg
Ventes = 800 kg
Autoconsommation = 50 kg
Pertes = 30 kg
Stock sortant = 100 + 1 000 - 800 - 50 - 30 = 220 kg
Prix valorisation = 250 FCFA/kg
Valeur stock sortant = 220 × 250 = 55 000 FCFA
```

### Cas particuliers

- Stock périssable avec décote.
- Stock de qualité différente.
- Stock négatif interdit.

### Règles de validation

- Stock sortant non négatif.
- Prix de valorisation documenté si le stock entre dans le produit économique.

### Cas d'erreur

- Stock négatif : erreur bloquante.
- Variation de stock sans unité : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur produit brut économique, trésorerie future, rendement disponible et alertes.

## 2.8 Pertes

### Définition

Les pertes sont des quantités produites ou disponibles qui ne sont ni vendues, ni autoconsommées, ni stockées.

### Objectif

Mesurer l'impact technique et économique des pertes de production, pertes sanitaires ou pertes post-récolte.

### Variables utilisées

- Pertes avant récolte.
- Pertes post-récolte.
- Production brute.
- Production nette.
- Prix de valorisation éventuel.

### Formule mathématique

```text
Production nette = Production brute - Pertes
Taux de perte = Pertes / Production brute × 100
Valeur pertes = Pertes × Prix de valorisation
```

### Explication détaillée

Les pertes peuvent être physiques et économiques. Elles ne sont pas des ventes, mais leur valeur peut être utilisée pour évaluer le manque à gagner.

### Exemple numérique complet

```text
Production brute = 2 000 kg
Pertes = 200 kg
Production nette = 1 800 kg
Taux de perte = 200 / 2 000 × 100 = 10 %
Prix = 250 FCFA/kg
Valeur pertes = 200 × 250 = 50 000 FCFA
```

### Cas particuliers

- Production nulle : taux de perte non applicable si production brute égale zéro.
- Pertes supérieures à production : erreur sauf stock entrant inclus.

### Règles de validation

- Pertes non négatives.
- Pertes inférieures ou égales à quantité disponible.

### Cas d'erreur

- Pertes supérieures à production disponible : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur rendement net, produit brut potentiel, alertes, IPE et diagnostic.

---

# 3. Produits

## 3.1 Ventes

### Définition

Les ventes sont les quantités cédées contre un prix à un acheteur, qu'elles soient encaissées immédiatement ou ultérieurement.

### Objectif

Calculer le chiffre d'affaires et alimenter le résultat financier ou économique selon l'encaissement.

### Variables utilisées

- Quantité vendue.
- Prix unitaire.
- Date de vente.
- Date d'encaissement.
- Statut de paiement.

### Formule mathématique

```text
Chiffre d'affaires vente = Quantité vendue × Prix unitaire
Chiffre d'affaires total = Somme(Chiffre d'affaires vente i)
```

### Explication détaillée

Une vente à crédit est économiquement une vente, mais son encaissement appartient au calendrier de trésorerie.

### Exemple numérique complet

```text
Vente 1 = 1 000 kg × 300 FCFA = 300 000 FCFA
Vente 2 = 500 kg × 280 FCFA = 140 000 FCFA
Chiffre d'affaires = 440 000 FCFA
```

### Cas particuliers

- Vente non encaissée : incluse dans produit économique, non incluse dans encaissement réel tant que non payée.
- Vente à prix nul : don ou transfert, pas vente commerciale.

### Règles de validation

- Prix non négatif.
- Quantité positive.
- Date de vente renseignée pour le calendrier.

### Cas d'erreur

- Vente sans quantité : erreur bloquante.
- Vente sans prix : erreur bloquante sauf montant total connu et signalé estimé.

### Impact sur les autres indicateurs

Impact sur chiffre d'affaires, produit brut, résultat financier, trésorerie, rentabilité et IPE.

## 3.2 Autoconsommation

### Définition

L'autoconsommation correspond à la quantité produite consommée par le ménage, l'exploitation ou l'atelier au lieu d'être vendue.

### Objectif

Valoriser l'avantage économique réel sans le confondre avec une vente encaissée.

### Variables utilisées

- Quantité autoconsommée.
- Prix de valorisation.
- Unité.

### Formule mathématique

```text
Valeur autoconsommation = Quantité autoconsommée × Prix de valorisation
```

### Explication détaillée

L'autoconsommation est intégrée au produit brut économique, mais présentée séparément des ventes. Elle n'est pas un encaissement financier.

### Exemple numérique complet

```text
Quantité autoconsommée = 100 kg
Prix de valorisation = 250 FCFA/kg
Valeur autoconsommation = 100 × 250 = 25 000 FCFA
```

### Cas particuliers

- Don familial hors ménage : à classer selon règles de projet.
- Produit transformé autoconsommé : utiliser valeur du produit brut ou transformé selon règle validée.

### Règles de validation

- Quantité non négative.
- Prix de valorisation documenté.

### Cas d'erreur

- Autoconsommation supérieure à production disponible : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur produit brut économique, résultat économique, ratio bénéfice/coût économique et IQD.

## 3.3 Produits secondaires

### Définition

Les produits secondaires sont des produits issus de la spéculation mais non considérés comme produit principal : paille, fumier, œufs secondaires, sous-produits, résidus valorisés.

### Objectif

Ne pas sous-estimer le revenu ou la valeur économique de l'activité.

### Variables utilisées

- Quantité de produit secondaire.
- Prix unitaire ou valeur de remplacement.
- Type de produit secondaire.

### Formule mathématique

```text
Valeur produits secondaires = Somme(Quantité produit secondaire i × Prix unitaire i)
```

### Explication détaillée

Les produits secondaires doivent être séparés du produit principal pour conserver la lisibilité économique.

### Exemple numérique complet

```text
Fumier vendu = 30 sacs × 500 FCFA = 15 000 FCFA
Paille valorisée = 100 bottes × 100 FCFA = 10 000 FCFA
Produits secondaires = 25 000 FCFA
```

### Cas particuliers

- Produit secondaire autoconsommé : intégrer comme valeur économique non encaissée.
- Produit secondaire sans marché : valeur nulle ou estimée selon référentiel.

### Règles de validation

- Type de produit secondaire obligatoire.
- Prix ou méthode de valorisation obligatoire si montant non nul.

### Cas d'erreur

- Produit secondaire sans unité : alerte.

### Impact sur les autres indicateurs

Impact sur produit brut, marge, résultat et comparaison.

## 3.4 Variation de stocks

### Définition

La variation de stocks mesure l'augmentation ou la diminution de la valeur des stocks entre le début et la fin d'un cycle ou d'une période.

### Objectif

Intégrer dans l'analyse économique les produits non vendus mais conservés.

### Variables utilisées

- Stock entrant.
- Stock sortant.
- Prix de valorisation stock entrant.
- Prix de valorisation stock sortant.

### Formule mathématique

```text
Valeur stock entrant = Stock entrant × Prix entrant
Valeur stock sortant = Stock sortant × Prix sortant
Variation de stocks = Valeur stock sortant - Valeur stock entrant
```

### Explication détaillée

Une variation positive augmente le produit économique. Une variation négative peut traduire une consommation de stock ou une baisse de stock disponible.

### Exemple numérique complet

```text
Stock entrant = 100 kg × 200 FCFA = 20 000 FCFA
Stock sortant = 220 kg × 250 FCFA = 55 000 FCFA
Variation de stocks = 55 000 - 20 000 = 35 000 FCFA
```

### Cas particuliers

- Stock périssable : appliquer décote si validée.
- Stock de qualité différente : valoriser par qualité.

### Règles de validation

- Stocks non négatifs.
- Prix de valorisation renseignés.

### Cas d'erreur

- Stock sortant négatif : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur produit brut économique et résultat économique.

## 3.5 Produits exceptionnels

### Définition

Les produits exceptionnels sont des ressources non récurrentes liées indirectement à l'activité, par exemple indemnités, compensations ou aides ponctuelles.

### Objectif

Les identifier sans masquer la performance normale de la spéculation.

### Variables utilisées

- Montant du produit exceptionnel.
- Nature.
- Date.
- Lien avec la spéculation.

### Formule mathématique

```text
Produits exceptionnels = Somme(Montants exceptionnels retenus)
```

### Explication détaillée

Les produits exceptionnels doivent être présentés séparément du produit brut d'exploitation, sauf règle spécifique validée.

### Exemple numérique complet

```text
Indemnité climatique = 75 000 FCFA
Produit brut d'exploitation = 600 000 FCFA
Produit total après élément exceptionnel = 675 000 FCFA
```

### Cas particuliers

- Subvention : traitée dans une rubrique séparée.
- Crédit : jamais produit exceptionnel.

### Règles de validation

- Nature obligatoire.
- Inclusion ou exclusion du résultat principal explicitée.

### Cas d'erreur

- Crédit classé en produit exceptionnel : erreur de classification.

### Impact sur les autres indicateurs

Peut affecter un résultat après éléments exceptionnels, mais ne doit pas fausser la rentabilité intrinsèque.

---

# 4. Charges

## 4.1 Charges variables

### Définition

Les charges variables sont les charges qui varient avec la production, la superficie, l'effectif, le cycle ou le niveau d'activité.

### Objectif

Calculer la marge brute et analyser les coûts directement liés à la spéculation.

### Variables utilisées

- Quantité consommée.
- Coût unitaire.
- Nombre de cycles.
- Taux d'application.

### Formule mathématique

```text
Montant charge variable = Quantité × Coût unitaire
Charges variables totales = Somme(Montants charges variables)
```

### Explication détaillée

Les semences, engrais, pesticides, aliments, médicaments, main-d'œuvre journalière et transport de production sont généralement variables.

### Exemple numérique complet

```text
Semences = 40 kg × 500 FCFA = 20 000 FCFA
Engrais = 8 sacs × 18 000 FCFA = 144 000 FCFA
Main-d'œuvre = 20 jours × 3 000 FCFA = 60 000 FCFA
Charges variables = 224 000 FCFA
```

### Cas particuliers

- Une charge peut être variable pour une spéculation et fixe pour une autre selon le contexte.
- Les charges propres au premier cycle doivent être isolées.

### Règles de validation

- Quantité non négative.
- Coût unitaire non négatif.
- Catégorie obligatoire.

### Cas d'erreur

- Charge sans catégorie : alerte bloquante à la validation.

### Impact sur les autres indicateurs

Impact sur marge brute, charges totales, coût de production, résultat et seuil de rentabilité.

## 4.2 Charges fixes

### Définition

Les charges fixes sont les charges qui ne varient pas directement avec le volume produit à court terme.

### Objectif

Calculer les charges totales, la marge nette et le seuil de rentabilité.

### Variables utilisées

- Montant annuel ou périodique.
- Période d'affectation.
- Clé de répartition si charge partagée.

### Formule mathématique

```text
Charges fixes affectées = Somme(Charges fixes × Taux affectation)
```

### Explication détaillée

Les bâtiments, certains équipements, loyers, assurances ou frais administratifs peuvent être fixes.

### Exemple numérique complet

```text
Loyer annuel = 120 000 FCFA
Part affectée à la spéculation = 50 %
Charge fixe affectée = 120 000 × 50 % = 60 000 FCFA
```

### Cas particuliers

- Charge fixe annuelle répartie sur plusieurs cycles.
- Charge fixe partagée entre plusieurs spéculations.

### Règles de validation

- Montant non négatif.
- Période de rattachement obligatoire.
- Clé d'affectation obligatoire si partagée.

### Cas d'erreur

- Total des clés d'affectation supérieur à 100 % sans justification : erreur.

### Impact sur les autres indicateurs

Impact sur charges totales, marge nette, bénéfice, seuil de rentabilité et ROI.

## 4.3 Investissements

### Définition

Un investissement est une dépense destinée à produire des avantages sur plusieurs cycles ou années.

### Objectif

Distinguer les dépenses d'équipement ou d'installation des charges consommées immédiatement.

### Variables utilisées

- Valeur d'acquisition.
- Date d'acquisition.
- Durée d'utilisation.
- Valeur résiduelle éventuelle.
- Part affectée.

### Formule mathématique

```text
Investissement total affecté = Valeur acquisition × Taux affectation
```

### Explication détaillée

L'investissement n'est pas entièrement une charge de production du cycle. Il est pris en compte par amortissement, sauf analyse de trésorerie où le décaissement peut être total à l'achat.

### Exemple numérique complet

```text
Motopompe = 300 000 FCFA
Part utilisée pour tomate = 40 %
Investissement affecté = 300 000 × 40 % = 120 000 FCFA
```

### Cas particuliers

- Investissement financé par crédit : décaissement et financement traités séparément.
- Investissement offert ou subventionné : valorisation séparée.

### Règles de validation

- Valeur d'acquisition positive.
- Durée d'utilisation renseignée si amortissement.

### Cas d'erreur

- Investissement sans durée et devant être amorti : calcul d'amortissement bloqué.

### Impact sur les autres indicateurs

Impact sur amortissement, ROI, besoin de trésorerie, délai de récupération et cultures pérennes.

## 4.4 Amortissements

### Définition

L'amortissement est la répartition du coût d'un investissement sur sa durée d'utilisation économique.

### Objectif

Intégrer progressivement le coût des actifs durables dans le résultat économique.

### Variables utilisées

- Valeur d'acquisition.
- Valeur résiduelle.
- Durée d'amortissement.
- Période d'utilisation.
- Taux d'affectation.

### Formule mathématique

```text
Amortissement annuel = (Valeur acquisition - Valeur résiduelle) / Durée amortissement années
Amortissement affecté = Amortissement annuel × Taux affectation × Fraction période
```

### Explication détaillée

La méthode retenue est linéaire. La durée est paramétrable. L'amortissement est une charge économique, pas nécessairement un décaissement du cycle.

### Exemple numérique complet

```text
Équipement = 500 000 FCFA
Valeur résiduelle = 0 FCFA
Durée = 5 ans
Amortissement annuel = 500 000 / 5 = 100 000 FCFA/an
Part affectée = 60 %
Amortissement affecté = 100 000 × 60 % = 60 000 FCFA/an
```

### Cas particuliers

- Cycle de 6 mois : fraction période = 6 / 12 = 0,5.
- Actif partagé : appliquer la clé d'affectation.

### Règles de validation

- Durée strictement positive.
- Valeur résiduelle inférieure ou égale à valeur d'acquisition.

### Cas d'erreur

- Durée égale à zéro : erreur bloquante.
- Valeur résiduelle supérieure à acquisition : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur charges économiques, résultat économique, coût total, ROI et analyses pluriannuelles.

## 4.5 Charges valorisées

### Définition

Les charges valorisées sont des consommations de ressources non payées en monnaie mais ayant une valeur économique.

### Objectif

Calculer un résultat économique complet.

### Variables utilisées

- Quantité de ressource utilisée.
- Valeur unitaire de référence.
- Source de valorisation.

### Formule mathématique

```text
Charge valorisée = Quantité utilisée × Valeur unitaire de référence
```

### Explication détaillée

La main-d'œuvre familiale, les intrants autoproduits ou les équipements prêtés peuvent être valorisés.

### Exemple numérique complet

```text
Main-d'œuvre familiale = 30 jours
Valeur journalière = 3 000 FCFA
Charge valorisée = 30 × 3 000 = 90 000 FCFA
```

### Cas particuliers

- Ressource familiale sans barème : donnée estimée.
- Apport en nature : valorisé au prix de remplacement.

### Règles de validation

- Méthode de valorisation obligatoire.
- Signalement dans le rapport.

### Cas d'erreur

- Charge valorisée sans prix de référence : alerte ou calcul économique incomplet.

### Impact sur les autres indicateurs

Impact sur résultat économique, charges économiques, IPE et IQD.

## 4.6 Charges réellement payées

### Définition

Les charges réellement payées sont les dépenses monétaires effectuées pour la spéculation.

### Objectif

Calculer le résultat financier et le besoin de trésorerie.

### Variables utilisées

- Montant payé.
- Date de paiement.
- Moyen de paiement.
- Rubrique de charge.

### Formule mathématique

```text
Charges payées totales = Somme(Montants effectivement payés)
```

### Explication détaillée

Une charge consommée mais non encore payée est économique mais pas encore sortie de trésorerie.

### Exemple numérique complet

```text
Engrais acheté = 100 000 FCFA
Montant payé immédiatement = 60 000 FCFA
Montant à payer plus tard = 40 000 FCFA
Charge économique = 100 000 FCFA
Sortie de trésorerie immédiate = 60 000 FCFA
```

### Cas particuliers

- Achat à crédit.
- Paiement fractionné.
- Dette fournisseur.

### Règles de validation

- Date de paiement obligatoire si la charge est marquée payée.
- Montant payé inférieur ou égal au montant dû.

### Cas d'erreur

- Montant payé supérieur au montant de la charge sans justification : alerte.

### Impact sur les autres indicateurs

Impact sur résultat financier, trésorerie, capacité de remboursement et besoin maximal.

## 4.7 Main-d'œuvre familiale

### Définition

La main-d'œuvre familiale est le travail fourni par le ménage ou la famille sans paiement monétaire direct.

### Objectif

Mesurer le coût économique réel du travail familial.

### Variables utilisées

- Nombre de jours-homme.
- Valeur journalière.
- Type d'opération.
- Période.

### Formule mathématique

```text
Valeur main-d'œuvre familiale = Jours familiaux × Coût journalier de référence
```

### Explication détaillée

Elle est incluse dans le résultat économique et exclue du résultat financier si elle n'est pas payée.

### Exemple numérique complet

```text
Travail familial = 25 jours
Coût journalier = 3 000 FCFA
Valeur = 25 × 3 000 = 75 000 FCFA
```

### Cas particuliers

- Travail familial partiellement rémunéré : séparer part payée et part valorisée.
- Plusieurs catégories de tâches : utiliser des barèmes différents si validés.

### Règles de validation

- Nombre de jours non négatif.
- Barème obligatoire pour calcul économique.

### Cas d'erreur

- Jours familiaux renseignés sans barème : alerte IQD.

### Impact sur les autres indicateurs

Impact sur résultat économique, productivité du travail, coût de production économique et IPE.

## 4.8 Prestations externes

### Définition

Les prestations externes sont des services achetés à des tiers : labour, pulvérisation, récolte, transport, gardiennage, conseil technique ou location de matériel.

### Objectif

Intégrer les services achetés dans les charges de production.

### Variables utilisées

- Quantité de service.
- Prix unitaire.
- Montant forfaitaire.
- Date ou période.

### Formule mathématique

```text
Montant prestation = Quantité × Prix unitaire
```

ou :

```text
Montant prestation = Montant forfaitaire
```

### Explication détaillée

Les prestations externes sont généralement des charges payées. Elles peuvent être variables ou fixes selon leur nature.

### Exemple numérique complet

```text
Labour = 2 ha × 35 000 FCFA/ha = 70 000 FCFA
Transport = forfait 25 000 FCFA
Prestations externes = 95 000 FCFA
```

### Cas particuliers

- Prestation partagée entre plusieurs spéculations.
- Prestation payée à crédit.

### Règles de validation

- Nature de prestation obligatoire.
- Montant ou quantité/prix obligatoire.

### Cas d'erreur

- Quantité renseignée sans prix ni montant : alerte.

### Impact sur les autres indicateurs

Impact sur charges variables/fixes, résultat financier, résultat économique et trésorerie.

## 4.9 Frais financiers

### Définition

Les frais financiers sont les intérêts, commissions, frais de dossier et autres coûts liés au financement.

### Objectif

Évaluer le coût du crédit sans le confondre avec les charges techniques de production.

### Variables utilisées

- Capital emprunté.
- Taux d'intérêt.
- Durée.
- Commissions.
- Frais fixes.

### Formule mathématique

Formule simplifiée V1 si intérêt simple :

```text
Intérêt = Capital × Taux × Durée_en_années
Frais financiers = Intérêt + Commissions + Frais fixes
```

### Explication détaillée

Les frais financiers sont présentés séparément. Le capital reçu n'est pas un produit d'exploitation. Le remboursement du capital n'est pas une charge de production.

### Exemple numérique complet

```text
Capital = 500 000 FCFA
Taux = 12 % annuel
Durée = 6 mois = 0,5 an
Commission = 10 000 FCFA
Intérêt = 500 000 × 12 % × 0,5 = 30 000 FCFA
Frais financiers = 30 000 + 10 000 = 40 000 FCFA
```

### Cas particuliers

- Intérêt dégressif : préparé pour versions futures.
- Crédit en nature : valorisation séparée.

### Règles de validation

- Capital positif.
- Taux non négatif.
- Durée positive.

### Cas d'erreur

- Crédit classé comme produit brut : erreur de classification.

### Impact sur les autres indicateurs

Impact sur résultat après financement, trésorerie, capacité de remboursement et IPE si le besoin financier est pris en compte.

## 4.10 Charges partagées

### Définition

Une charge partagée est une charge utilisée par plusieurs spéculations, parcelles, ateliers, cycles ou producteurs.

### Objectif

Affecter équitablement une charge commune à chaque activité analysée.

### Variables utilisées

- Montant total de la charge.
- Méthode d'affectation.
- Base d'affectation par activité.
- Somme des bases.

### Formule mathématique

```text
Taux affectation activité = Base activité / Somme bases activités
Charge affectée activité = Charge totale × Taux affectation activité
```

### Explication détaillée

Les bases possibles sont superficie, effectif, biomasse, quantité produite, chiffre d'affaires, temps d'utilisation, jours de travail ou clé manuelle.

### Exemple numérique complet

Une motopompe coûte 90 000 FCFA à utiliser pour deux cultures :

```text
Tomate = 1 ha
Riz = 2 ha
Base totale = 3 ha
Taux tomate = 1 / 3 = 33,33 %
Taux riz = 2 / 3 = 66,67 %
Charge tomate = 90 000 × 33,33 % = 30 000 FCFA
Charge riz = 90 000 × 66,67 % = 60 000 FCFA
```

### Cas particuliers

- Clé manuelle : doit totaliser 100 %.
- Charge partagée entre culture et élevage : utiliser une base commune validée, souvent temps d'utilisation ou clé manuelle.

### Règles de validation

- Somme des clés égale à 100 % ou base totale strictement positive.
- La charge source ne doit pas être comptée en plus des charges affectées.

### Cas d'erreur

- Somme des clés supérieure à 100 % : erreur.
- Base totale nulle : affectation impossible.

### Impact sur les autres indicateurs

Impact sur charges, coûts unitaires, marges, résultats et comparaisons.

---

# 5. Calcul du coût de production

## 5.1 Coût total

### Définition

Le coût total est la somme des charges affectées à une spéculation, un cycle ou une période.

### Objectif

Mesurer le coût global nécessaire pour produire.

### Variables utilisées

- Charges variables.
- Charges fixes affectées.
- Amortissements.
- Charges valorisées.
- Frais financiers selon périmètre choisi.

### Formule mathématique

```text
Coût total financier = Charges réellement payées d'exploitation
Coût total économique = Charges payées + Charges valorisées + Amortissements économiques
```

### Explication détaillée

Le coût total financier et le coût total économique doivent être distingués.

### Exemple numérique complet

```text
Charges payées = 400 000 FCFA
Main-d'œuvre familiale valorisée = 80 000 FCFA
Amortissement = 40 000 FCFA
Coût total financier = 400 000 FCFA
Coût total économique = 400 000 + 80 000 + 40 000 = 520 000 FCFA
```

### Cas particuliers

- Frais financiers inclus seulement dans un résultat après financement si la règle le prévoit.

### Règles de validation

- Toutes les charges doivent être classées.

### Cas d'erreur

- Charge sans nature financière/économique : alerte de qualité.

### Impact sur les autres indicateurs

Impact sur coût unitaire, marge nette, résultat, ROI, seuil de rentabilité.

## 5.2 Coût unitaire

### Définition

Le coût unitaire est le coût de production par unité de produit vendable ou produit.

### Objectif

Comparer le coût au prix de vente et mesurer la compétitivité.

### Variables utilisées

- Coût total.
- Quantité produite vendable ou quantité vendue.

### Formule mathématique

```text
Coût unitaire = Coût total / Quantité de référence
```

### Explication détaillée

La quantité de référence doit être clairement indiquée : production totale, production nette, quantité vendue ou quantité vendable.

### Exemple numérique complet

```text
Coût total = 520 000 FCFA
Production nette = 2 000 kg
Coût unitaire = 520 000 / 2 000 = 260 FCFA/kg
```

### Cas particuliers

- Production nulle : coût unitaire non calculable.
- Produits multiples : coût par produit nécessite une clé d'affectation ou un produit principal.

### Règles de validation

- Quantité de référence strictement positive.

### Cas d'erreur

- Division par zéro : indicateur non applicable avec alerte.

### Impact sur les autres indicateurs

Impact sur seuil de rentabilité, alertes coût, comparaison et diagnostic.

## 5.3 Coût par hectare

### Définition

Le coût par hectare est le coût total rapporté à la superficie exploitée.

### Objectif

Comparer l'intensité des charges entre parcelles, producteurs et spéculations végétales.

### Variables utilisées

- Coût total.
- Superficie en hectares.

### Formule mathématique

```text
Coût par hectare = Coût total / Superficie_ha
```

### Exemple numérique complet

```text
Coût total = 600 000 FCFA
Superficie = 2 ha
Coût par hectare = 600 000 / 2 = 300 000 FCFA/ha
```

### Cas particuliers

- Non applicable aux élevages sauf analyse par surface d'installation.

### Règles de validation

- Superficie strictement positive.

### Cas d'erreur

- Superficie nulle : non calculable.

### Impact sur les autres indicateurs

Impact sur rentabilité par hectare et alertes de coût.

## 5.4 Coût par kilogramme

### Définition

Le coût par kilogramme est le coût total divisé par la quantité produite ou vendable en kilogrammes.

### Objectif

Comparer le coût au prix de vente par kilogramme.

### Variables utilisées

- Coût total.
- Quantité en kg.

### Formule mathématique

```text
Coût par kg = Coût total / Quantité_kg
```

### Exemple numérique complet

```text
Coût total = 750 000 FCFA
Production = 3 000 kg
Coût par kg = 750 000 / 3 000 = 250 FCFA/kg
```

### Cas particuliers

- Unité non kg : conversion requise.

### Règles de validation

- Quantité kg positive.
- Conversion disponible si nécessaire.

### Cas d'erreur

- Quantité ou conversion manquante : calcul bloqué.

### Impact sur les autres indicateurs

Impact sur seuil de rentabilité et alerte prix/coût.

## 5.5 Coût par tête

### Définition

Le coût par tête est le coût total d'un élevage rapporté à un effectif de référence.

### Objectif

Mesurer le coût moyen par animal.

### Variables utilisées

- Coût total.
- Effectif initial, vendu ou produit.

### Formule mathématique

```text
Coût par tête initiale = Coût total / Effectif initial
Coût par tête vendue = Coût total / Effectif vendu
```

### Exemple numérique complet

```text
Coût total = 900 000 FCFA
Effectif initial = 500
Effectif vendu = 460
Coût par tête initiale = 900 000 / 500 = 1 800 FCFA
Coût par tête vendue = 900 000 / 460 = 1 956,52 FCFA
```

### Cas particuliers

- Forte mortalité : coût par tête vendue augmente.
- Pisciculture : coût par kg de biomasse peut être plus pertinent.

### Règles de validation

- Effectif de référence positif.

### Cas d'erreur

- Effectif nul : non calculable.

### Impact sur les autres indicateurs

Impact sur rentabilité d'élevage, alertes mortalité et diagnostic.

## 5.6 Coût par cycle

### Définition

Le coût par cycle est le coût total rattaché à un cycle de production.

### Objectif

Comparer des cycles et annualiser les résultats.

### Variables utilisées

- Charges du cycle.
- Charges fixes réparties.
- Amortissements proratisés.
- Nombre de cycles par an.

### Formule mathématique

```text
Coût par cycle = Charges variables du cycle + Charges fixes affectées au cycle + Amortissements du cycle
Coût annuel = Somme(Coûts par cycle) ou Coût cycle moyen × Nombre de cycles
```

### Exemple numérique complet

```text
Charges variables cycle = 300 000 FCFA
Charges fixes annuelles = 120 000 FCFA
Nombre cycles = 3
Charges fixes par cycle = 120 000 / 3 = 40 000 FCFA
Coût par cycle = 300 000 + 40 000 = 340 000 FCFA
Coût annuel = 340 000 × 3 = 1 020 000 FCFA
```

### Cas particuliers

- Premier cycle plus coûteux.
- Cycles suivants avec charges réduites.

### Règles de validation

- Nombre de cycles positif.
- Règles de récurrence renseignées.

### Cas d'erreur

- Nombre de cycles égal à zéro : annualisation impossible.

### Impact sur les autres indicateurs

Impact sur bénéfice annuel, rentabilité annuelle, IPE et trésorerie.

---

# 6. Calculs économiques

## 6.1 Produit brut

### Définition

Le produit brut est la valeur totale générée par la spéculation sur une période, avant déduction des charges.

### Objectif

Mesurer la production économique totale.

### Variables utilisées

- Ventes.
- Autoconsommation valorisée.
- Produits secondaires.
- Variation positive de stocks.
- Produits exceptionnels si inclus dans une présentation séparée.

### Formule mathématique

```text
Produit brut économique = Ventes + Autoconsommation valorisée + Produits secondaires + Variation de stocks retenue
Produit brut financier = Ventes encaissées d'exploitation
```

### Exemple numérique complet

```text
Ventes = 800 000 FCFA
Autoconsommation = 50 000 FCFA
Produits secondaires = 30 000 FCFA
Variation de stocks = 20 000 FCFA
Produit brut économique = 900 000 FCFA
Produit brut financier = 800 000 FCFA si toutes les ventes sont encaissées
```

### Cas particuliers

- Vente à crédit : incluse économiquement, encaissée financièrement à la date de paiement.
- Subvention : présentée séparément.

### Règles de validation

- Toutes les composantes doivent être identifiées séparément.

### Cas d'erreur

- Crédit reçu inclus dans produit brut : erreur.

### Impact sur les autres indicateurs

Impact sur marges, résultats, ratios, IPE et tableau de bord.

## 6.2 Marge brute

### Définition

La marge brute est la différence entre produit brut et charges variables.

### Objectif

Mesurer la performance directe de la spéculation avant charges fixes.

### Variables utilisées

- Produit brut.
- Charges variables.

### Formule mathématique

```text
Marge brute = Produit brut - Charges variables
```

### Exemple numérique complet

```text
Produit brut = 900 000 FCFA
Charges variables = 500 000 FCFA
Marge brute = 400 000 FCFA
```

### Cas particuliers

- Marge brute financière et économique peuvent être séparées.

### Règles de validation

- Produit brut calculé.
- Charges variables classées.

### Cas d'erreur

- Charges non classées : marge brute provisoire uniquement.

### Impact sur les autres indicateurs

Impact sur seuil de rentabilité, marge nette, diagnostic et alertes.

## 6.3 Valeur ajoutée

### Définition

La valeur ajoutée mesure la richesse créée après déduction des consommations intermédiaires.

### Objectif

Analyser la contribution économique de la spéculation indépendamment de certains facteurs comme main-d'œuvre et amortissements.

### Variables utilisées

- Produit brut.
- Consommations intermédiaires.

### Formule mathématique

```text
Valeur ajoutée = Produit brut - Consommations intermédiaires
```

### Exemple numérique complet

```text
Produit brut = 1 000 000 FCFA
Intrants + services externes = 420 000 FCFA
Valeur ajoutée = 580 000 FCFA
```

### Cas particuliers

- La main-d'œuvre familiale peut être exclue des consommations intermédiaires selon présentation.

### Règles de validation

- Identifier clairement les consommations intermédiaires.

### Cas d'erreur

- Charge sans classification : calcul incomplet.

### Impact sur les autres indicateurs

Impact sur diagnostic économique et formation.

## 6.4 Résultat financier

### Définition

Le résultat financier mesure les entrées et sorties monétaires réelles liées à l'exploitation.

### Objectif

Mesurer la capacité de l'activité à générer de l'argent disponible.

### Variables utilisées

- Encaissements d'exploitation.
- Décaissements d'exploitation.
- Frais financiers si inclus dans résultat après financement.

### Formule mathématique

```text
Résultat financier d'exploitation = Encaissements d'exploitation - Décaissements d'exploitation
Résultat financier après financement = Résultat financier d'exploitation - Frais financiers payés
```

### Exemple numérique complet

```text
Encaissements ventes = 750 000 FCFA
Charges payées = 480 000 FCFA
Frais financiers = 30 000 FCFA
Résultat financier d'exploitation = 750 000 - 480 000 = 270 000 FCFA
Résultat financier après financement = 270 000 - 30 000 = 240 000 FCFA
```

### Cas particuliers

- Vente non encore encaissée exclue des encaissements courants.
- Achat à crédit exclu des décaissements jusqu'au paiement.

### Règles de validation

- Dates ou statuts de paiement nécessaires.

### Cas d'erreur

- Encaissement sans vente ou avance non liée : alerte de classification.

### Impact sur les autres indicateurs

Impact sur trésorerie, capacité de remboursement et diagnostic financier.

## 6.5 Résultat économique

### Définition

Le résultat économique mesure la rentabilité complète en intégrant les produits valorisés et charges valorisées.

### Objectif

Évaluer si l'activité crée réellement de la valeur économique.

### Variables utilisées

- Produit brut économique.
- Charges économiques.

### Formule mathématique

```text
Résultat économique = Produit brut économique - Charges économiques totales
```

### Exemple numérique complet

```text
Produit brut économique = 900 000 FCFA
Charges payées = 480 000 FCFA
Main-d'œuvre familiale = 90 000 FCFA
Amortissement = 50 000 FCFA
Charges économiques = 620 000 FCFA
Résultat économique = 900 000 - 620 000 = 280 000 FCFA
```

### Cas particuliers

- Charges valorisées sans barème : résultat économique provisoire.

### Règles de validation

- Toutes les valorisations doivent être documentées.

### Cas d'erreur

- Barème absent pour charge valorisée obligatoire : alerte ou calcul incomplet.

### Impact sur les autres indicateurs

Impact sur ROI économique, IPE, comparaison et recommandations.

## 6.6 Bénéfice

### Définition

Le bénéfice est le résultat positif obtenu après déduction des charges du produit. Dans AGRIECO PRO, il doit être précisé comme financier ou économique.

### Objectif

Mesurer le gain net d'une activité.

### Variables utilisées

- Résultat financier.
- Résultat économique.

### Formule mathématique

```text
Bénéfice financier = max(Résultat financier, Résultat financier) avec signe conservé
Bénéfice économique = max(Résultat économique, Résultat économique) avec signe conservé
```

En pratique, le terme bénéfice peut être négatif ; il correspond alors à une perte.

### Exemple numérique complet

```text
Produit brut = 700 000 FCFA
Charges totales = 850 000 FCFA
Bénéfice = 700 000 - 850 000 = -150 000 FCFA
```

Le résultat est une perte de 150 000 FCFA.

### Cas particuliers

- Si résultat négatif, afficher « perte » dans les rapports.

### Règles de validation

- Toujours préciser le périmètre : financier ou économique.

### Cas d'erreur

- Bénéfice affiché sans périmètre : erreur de présentation.

### Impact sur les autres indicateurs

Impact sur ROI, ratio bénéfice/coût, IPE, alertes et classements.

## 6.7 Bénéfices annualisés et spécifiques

### Définition

Les bénéfices spécifiques rapportent le bénéfice à une période ou une base physique : annuel, cycle, hectare, tête ou jour.

### Objectif

Comparer des spéculations de durées et tailles différentes.

### Variables utilisées

- Bénéfice par cycle.
- Nombre de cycles par an.
- Superficie.
- Effectif.
- Durée en jours.

### Formule mathématique

```text
Bénéfice annuel = Somme(Bénéfices des cycles de l'année)
Bénéfice annuel simplifié = Bénéfice par cycle × Nombre de cycles par an
Bénéfice par hectare = Bénéfice / Superficie_ha
Bénéfice par tête = Bénéfice / Effectif de référence
Bénéfice journalier = Bénéfice / Durée_cycle_jours
```

### Exemple numérique complet

```text
Bénéfice par cycle = 200 000 FCFA
Nombre de cycles = 3
Bénéfice annuel = 200 000 × 3 = 600 000 FCFA
Superficie = 2 ha
Bénéfice par hectare = 200 000 / 2 = 100 000 FCFA/ha/cycle
Durée = 90 jours
Bénéfice journalier = 200 000 / 90 = 2 222,22 FCFA/jour
```

### Cas particuliers

- Premier cycle différent : utiliser la somme réelle des cycles et non la multiplication simple.
- Culture pérenne : annualisation par phase ou moyenne pluriannuelle.

### Règles de validation

- Nombre de cycles positif.
- Durée positive.
- Superficie ou effectif positif si indicateur correspondant.

### Cas d'erreur

- Division par zéro : indicateur non applicable.

### Impact sur les autres indicateurs

Impact sur rentabilité annuelle, IPE, comparaison et tableaux de bord.

---

# 7. Indicateurs de rentabilité

## 7.1 ROI

### Définition

Le ROI mesure le rendement du capital investi.

### Objectif

Comparer le bénéfice généré par rapport à l'investissement nécessaire.

### Variables utilisées

- Bénéfice.
- Investissement total ou capital engagé.

### Formule mathématique

```text
ROI = Bénéfice / Investissement total × 100
```

### Exemple numérique complet

```text
Bénéfice = 250 000 FCFA
Investissement total = 1 000 000 FCFA
ROI = 250 000 / 1 000 000 × 100 = 25 %
```

### Cas particuliers

- Investissement nul : ROI non applicable.
- Pour cycles courts, utiliser capital engagé du cycle ou besoin maximal de trésorerie selon règle validée.

### Règles de validation

- Investissement strictement positif.
- Bénéfice clairement financier ou économique.

### Cas d'erreur

- Investissement nul : afficher non calculable.

### Impact sur les autres indicateurs

Impact sur IPE, comparaison et classement.

## 7.2 Ratio bénéfice/coût

### Définition

Le ratio bénéfice/coût compare le produit brut aux charges totales.

### Objectif

Mesurer combien 1 FCFA de charge génère en produit.

### Variables utilisées

- Produit brut.
- Charges totales.

### Formule mathématique

```text
Ratio bénéfice/coût = Produit brut / Charges totales
```

### Exemple numérique complet

```text
Produit brut = 900 000 FCFA
Charges totales = 600 000 FCFA
Ratio = 900 000 / 600 000 = 1,5
```

Interprétation : 1 FCFA dépensé génère 1,5 FCFA de produit brut.

### Cas particuliers

- Charges nulles : ratio non applicable.

### Règles de validation

- Charges totales strictement positives.

### Cas d'erreur

- Division par zéro.

### Impact sur les autres indicateurs

Impact sur classement, IPE et alertes.

## 7.3 Marge nette

### Définition

La marge nette correspond au produit brut diminué des charges variables et fixes.

### Objectif

Mesurer le résultat après prise en compte de toutes les charges de production.

### Variables utilisées

- Produit brut.
- Charges variables.
- Charges fixes.

### Formule mathématique

```text
Marge nette = Produit brut - Charges variables - Charges fixes
```

### Exemple numérique complet

```text
Produit brut = 1 000 000 FCFA
Charges variables = 550 000 FCFA
Charges fixes = 150 000 FCFA
Marge nette = 300 000 FCFA
```

### Cas particuliers

- Les frais financiers peuvent être présentés après marge nette.

### Règles de validation

- Charges classées en variables/fixes.

### Cas d'erreur

- Charge non classée : marge nette provisoire.

### Impact sur les autres indicateurs

Impact sur bénéfice, rentabilité, alertes et diagnostic.

## 7.4 Rentabilités spécifiques

### Définition

Les rentabilités spécifiques expriment le résultat par rapport à une base : année, hectare, FCFA investi ou cycle.

### Objectif

Comparer des activités de taille ou durée différentes.

### Variables utilisées

- Bénéfice.
- Produit brut.
- Superficie.
- Investissement.
- Nombre de cycles.

### Formule mathématique

```text
Rentabilité annuelle = Bénéfice annuel / Investissement total × 100
Rentabilité par hectare = Bénéfice / Superficie_ha
Rentabilité par FCFA investi = Bénéfice / Investissement total
```

### Exemple numérique complet

```text
Bénéfice annuel = 600 000 FCFA
Investissement = 1 500 000 FCFA
Rentabilité annuelle = 600 000 / 1 500 000 × 100 = 40 %
Bénéfice = 300 000 FCFA
Superficie = 2 ha
Rentabilité par hectare = 150 000 FCFA/ha
```

### Cas particuliers

- Investissement nul : rentabilité par FCFA investi non applicable.

### Règles de validation

- Bases de calcul positives.

### Cas d'erreur

- Base nulle : indicateur non calculable.

### Impact sur les autres indicateurs

Impact sur IPE, comparaison et tableau de bord.

## 7.5 Seuil de rentabilité

### Définition

Le seuil de rentabilité est le niveau minimal de chiffre d'affaires ou de quantité nécessaire pour couvrir les charges.

### Objectif

Identifier le point à partir duquel l'activité ne perd plus d'argent.

### Variables utilisées

- Charges fixes.
- Prix unitaire moyen.
- Coût variable unitaire.
- Taux de marge sur coûts variables.

### Formule mathématique

```text
Marge sur coût variable unitaire = Prix unitaire - Coût variable unitaire
Seuil quantité = Charges fixes / Marge sur coût variable unitaire
Seuil valeur = Seuil quantité × Prix unitaire
```

ou :

```text
Taux marge sur coûts variables = (Produit brut - Charges variables) / Produit brut
Seuil valeur = Charges fixes / Taux marge sur coûts variables
```

### Exemple numérique complet

```text
Prix = 300 FCFA/kg
Coût variable unitaire = 180 FCFA/kg
Marge unitaire = 120 FCFA/kg
Charges fixes = 120 000 FCFA
Seuil quantité = 120 000 / 120 = 1 000 kg
Seuil valeur = 1 000 × 300 = 300 000 FCFA
```

### Cas particuliers

- Marge unitaire nulle ou négative : seuil impossible à atteindre.
- Production multi-produits : seuil simplifié seulement si produit principal dominant.

### Règles de validation

- Prix supérieur au coût variable unitaire.
- Charges fixes non négatives.

### Cas d'erreur

- Marge unitaire <= 0 : alerte critique.

### Impact sur les autres indicateurs

Impact sur alertes, diagnostic, recommandations et risque.

## 7.6 Délai de récupération

### Définition

Le délai de récupération est le temps nécessaire pour récupérer l'investissement initial grâce aux flux nets générés.

### Objectif

Évaluer la vitesse de retour sur investissement.

### Variables utilisées

- Investissement initial.
- Flux net annuel ou par cycle.
- Cumul des flux nets.

### Formule mathématique

```text
Délai récupération simple = Investissement initial / Flux net moyen annuel
```

ou par cumul : première période où :

```text
Cumul flux nets >= Investissement initial
```

### Exemple numérique complet

```text
Investissement = 1 000 000 FCFA
Flux net annuel moyen = 250 000 FCFA
Délai récupération = 1 000 000 / 250 000 = 4 ans
```

### Cas particuliers

- Flux variables : utiliser cumul par période.
- Flux négatifs prolongés : délai non atteint.

### Règles de validation

- Investissement positif.
- Flux net moyen positif pour formule simple.

### Cas d'erreur

- Flux net <= 0 : délai non calculable ou non atteint.

### Impact sur les autres indicateurs

Impact sur analyse pérenne, investissement, IPE et financement.

## 7.7 VAN et TRI — préparés pour versions futures

### Définition

La VAN est la valeur actuelle nette des flux futurs. Le TRI est le taux qui annule la VAN.

### Objectif

Préparer l'analyse avancée des cultures pérennes et investissements longs.

### Variables utilisées

- Flux net par période.
- Taux d'actualisation.
- Année ou période.

### Formule mathématique

```text
VAN = Somme(Flux net t / (1 + taux)^t) - Investissement initial
TRI = taux tel que VAN = 0
```

### Exemple numérique complet

```text
Investissement initial = 1 000 000 FCFA
Flux année 1 = 0
Flux année 2 = 300 000
Flux année 3 = 500 000
Flux année 4 = 600 000
Taux = 10 %
VAN = -1 000 000 + 0/1,1 + 300 000/1,1² + 500 000/1,1³ + 600 000/1,1⁴
VAN ≈ -1 000 000 + 0 + 247 934 + 375 657 + 409 808 = 33 399 FCFA
```

### Cas particuliers

- TRI multiple possible si les flux changent plusieurs fois de signe.
- VAN et TRI ne sont pas obligatoires dans la V1.

### Règles de validation

- Taux d'actualisation non négatif.
- Flux datés.

### Cas d'erreur

- Flux non périodiques sans dates : calcul impossible.

### Impact sur les autres indicateurs

Prépare les futures analyses de cultures pérennes et investissements.

---

# 8. Trésorerie

## 8.1 Flux entrants

### Définition

Les flux entrants sont les encaissements réels : ventes encaissées, avances, crédits reçus, subventions encaissées et autres entrées de trésorerie.

### Objectif

Calculer la trésorerie disponible par période.

### Variables utilisées

- Montant encaissé.
- Date d'encaissement.
- Nature du flux.

### Formule mathématique

```text
Flux entrants période = Somme(Montants encaissés dans la période)
```

### Exemple numérique complet

```text
Vente encaissée = 300 000 FCFA
Avance reçue = 100 000 FCFA
Crédit reçu = 500 000 FCFA
Flux entrants = 900 000 FCFA
```

### Cas particuliers

- Le crédit est entrant de trésorerie, mais pas produit d'exploitation.

### Règles de validation

- Date et nature obligatoires.

### Cas d'erreur

- Flux entrant sans nature : alerte.

### Impact sur les autres indicateurs

Impact sur besoin maximal de trésorerie et capacité de remboursement.

## 8.2 Flux sortants

### Définition

Les flux sortants sont les décaissements réels : achats payés, salaires payés, transport payé, intérêts, commissions et remboursements.

### Objectif

Mesurer les besoins de liquidité.

### Variables utilisées

- Montant payé.
- Date de paiement.
- Nature du paiement.

### Formule mathématique

```text
Flux sortants période = Somme(Montants payés dans la période)
```

### Exemple numérique complet

```text
Intrants payés = 250 000 FCFA
Main-d'œuvre payée = 80 000 FCFA
Intérêts payés = 20 000 FCFA
Flux sortants = 350 000 FCFA
```

### Cas particuliers

- Remboursement de capital : flux sortant de trésorerie, mais pas charge de production.

### Règles de validation

- Date de paiement obligatoire.

### Cas d'erreur

- Paiement sans rattachement : alerte.

### Impact sur les autres indicateurs

Impact sur trésorerie, capacité de remboursement et financement.

## 8.3 Besoin maximal de trésorerie

### Définition

Le besoin maximal de trésorerie est le déficit cumulé maximal avant que les entrées ne compensent les sorties.

### Objectif

Estimer le financement minimal nécessaire pour réaliser l'activité sans rupture de trésorerie.

### Variables utilisées

- Flux entrants par période.
- Flux sortants par période.
- Solde initial.

### Formule mathématique

```text
Solde période t = Solde période t-1 + Entrées t - Sorties t
Besoin maximal = Valeur absolue du solde minimal si solde minimal < 0
```

### Exemple numérique complet

```text
Solde initial = 0
Mois 1 : entrées 0, sorties 300 000 -> solde -300 000
Mois 2 : entrées 0, sorties 200 000 -> solde -500 000
Mois 3 : entrées 700 000, sorties 50 000 -> solde 150 000
Solde minimal = -500 000
Besoin maximal = 500 000 FCFA
```

### Cas particuliers

- Crédit reçu : peut réduire le besoin de trésorerie mais doit rester séparé du produit.
- Avances reçues : entrées de trésorerie à rattacher.

### Règles de validation

- Dates ou périodes des flux obligatoires.
- Flux classés par nature.

### Cas d'erreur

- Flux sans période : exclu du calendrier avec alerte.

### Impact sur les autres indicateurs

Impact sur IPE, financement, diagnostic et comparaison.

## 8.4 Capacité de remboursement

### Définition

La capacité de remboursement mesure la part du flux net pouvant couvrir les échéances de crédit.

### Objectif

Aider les institutions financières et conseillers à évaluer la solvabilité liée à une activité.

### Variables utilisées

- Flux net disponible.
- Échéances de crédit.
- Frais financiers.

### Formule mathématique

```text
Capacité de remboursement = Flux net disponible / Échéance totale
```

### Exemple numérique complet

```text
Flux net disponible = 300 000 FCFA
Échéance crédit = 200 000 FCFA
Capacité = 300 000 / 200 000 = 1,5
```

### Cas particuliers

- Plusieurs crédits : sommer les échéances.
- Flux saisonniers : analyser par période critique.

### Règles de validation

- Échéance positive.
- Flux net calculé.

### Cas d'erreur

- Échéance nulle : indicateur non applicable.

### Impact sur les autres indicateurs

Impact sur diagnostic financier et risque.

## 8.5 Trésorerie minimale et calendrier

### Définition

La trésorerie minimale est le plus bas niveau de solde de trésorerie sur la période. Le calendrier de trésorerie présente les entrées, sorties et soldes par période.

### Objectif

Visualiser les périodes de tension.

### Variables utilisées

- Entrées par période.
- Sorties par période.
- Solde initial.

### Formule mathématique

```text
Trésorerie minimale = Min(Solde période t)
```

### Exemple numérique complet

Avec soldes mensuels :

```text
Janvier = -100 000
Février = -250 000
Mars = 50 000
Trésorerie minimale = -250 000 FCFA
```

### Cas particuliers

- Solde initial positif réduit le besoin.

### Règles de validation

- Toutes les périodes doivent être ordonnées chronologiquement.

### Cas d'erreur

- Dates incohérentes : alerte.

### Impact sur les autres indicateurs

Impact sur besoin maximal, IPE et recommandations.

---

# 9. Cultures pérennes

## 9.1 Phases de culture pérenne

### Définition

Une culture pérenne est analysée par phases : installation, années improductives, montée en production, régime de croisière et renouvellement.

### Objectif

Représenter correctement les dépenses et recettes sur plusieurs années.

### Variables utilisées

- Année.
- Phase.
- Charges annuelles.
- Produits annuels.
- Investissements.
- Rendement par âge.

### Formule mathématique

```text
Résultat annuel phase = Produits annuels - Charges annuelles - Amortissements affectés
Résultat cumulé année n = Somme(Résultats annuels de 1 à n)
```

### Exemple numérique complet

```text
Année 1 installation : produits 0, charges 600 000 -> résultat -600 000
Année 2 entretien : produits 0, charges 200 000 -> cumul -800 000
Année 3 entrée production : produits 300 000, charges 250 000 -> cumul -750 000
Année 4 croisière : produits 900 000, charges 350 000 -> cumul -200 000
Année 5 croisière : produits 1 000 000, charges 350 000 -> cumul 450 000
```

### Cas particuliers

- Production partielle en montée.
- Renouvellement partiel.
- Association temporaire avec cultures vivrières.

### Règles de validation

- Chaque année doit avoir une phase.
- Les investissements d'installation doivent être séparés des charges d'entretien.

### Cas d'erreur

- Année sans phase : alerte.

### Impact sur les autres indicateurs

Impact sur délai de récupération, résultat cumulé, VAN/TRI futurs, IPE et comparaison.

## 9.2 Annualisation

### Définition

L'annualisation exprime une performance longue en moyenne annuelle ou par année de phase.

### Objectif

Comparer une culture pérenne à une activité à cycle court.

### Variables utilisées

- Résultat cumulé.
- Nombre d'années.
- Résultats par année.

### Formule mathématique

```text
Bénéfice moyen annuel = Résultat cumulé / Nombre d'années
```

### Exemple numérique complet

```text
Résultat cumulé sur 5 ans = 450 000 FCFA
Bénéfice moyen annuel = 450 000 / 5 = 90 000 FCFA/an
```

### Cas particuliers

- Les années d'installation peuvent être présentées séparément.
- Les comparaisons doivent indiquer l'horizon utilisé.

### Règles de validation

- Nombre d'années positif.

### Cas d'erreur

- Horizon nul : calcul impossible.

### Impact sur les autres indicateurs

Impact sur rentabilité annuelle, IPE et comparaison des spéculations.

---

# 10. Comparaison des spéculations

## 10.1 Méthode générale de comparaison

### Définition

La comparaison des spéculations consiste à classer plusieurs activités selon des indicateurs homogènes.

### Objectif

Aider à choisir les activités les plus performantes selon les objectifs du producteur ou du conseiller.

### Variables utilisées

- Bénéfice.
- Rentabilité.
- Risque.
- Investissement.
- Besoin de trésorerie.
- Durée du cycle.
- Score marché.
- IPE.

### Formule mathématique

```text
Rang indicateur = Position de la spéculation après tri selon l'indicateur retenu
```

### Exemple numérique complet

```text
Maïs bénéfice annuel = 300 000 FCFA
Tomate bénéfice annuel = 900 000 FCFA
Poulet bénéfice annuel = 600 000 FCFA
Classement bénéfice : 1 Tomate, 2 Poulet, 3 Maïs
```

### Cas particuliers

- Une spéculation très rentable mais très risquée peut être moins recommandée.
- Les cultures pérennes doivent être comparées sur une base annualisée.

### Règles de validation

- Les indicateurs comparés doivent utiliser la même période ou être annualisés.

### Cas d'erreur

- Comparaison entre cycle et année sans annualisation : alerte.

### Impact sur les autres indicateurs

Impact sur tableau de bord, IPE et recommandations.

## 10.2 Critères de comparaison

### Définition

Les critères sont les axes de performance retenus : bénéfice, rentabilité, risque, investissement, trésorerie, durée et marché.

### Objectif

Éviter une comparaison basée sur un seul indicateur.

### Variables utilisées

- Valeurs normalisées ou brutes des critères.

### Formule mathématique

```text
Score multicritère = Somme(Score critère × Pondération critère)
```

### Exemple numérique complet

```text
Score rentabilité = 80, poids 30 % -> 24
Score risque = 60, poids 15 % -> 9
Score marché = 70, poids 10 % -> 7
Contribution partielle = 40 points sur les critères affichés
```

### Cas particuliers

- Pondérations différentes selon type d'utilisateur.

### Règles de validation

- Somme des pondérations égale à 100 %.

### Cas d'erreur

- Pondérations incomplètes : classement multicritère non calculable.

### Impact sur les autres indicateurs

Base du calcul IPE.

---

# 11. Calcul de l'IPE

## 11.1 Définition générale

### Définition

L'Indice de Performance Économique, ou IPE, est un score sur 100 qui synthétise la performance d'une spéculation selon des critères financiers, techniques, commerciaux et opérationnels.

### Objectif

Classer les spéculations et guider la décision sans remplacer l'analyse du conseiller.

### Variables utilisées

- Rentabilité annuelle.
- Durée du cycle.
- Investissement nécessaire.
- Niveau de risque.
- Besoin de trésorerie.
- Disponibilité du marché.
- Besoin en main-d'œuvre.
- Pondérations.

### Formule mathématique

```text
IPE = Somme(Score normalisé critère i × Pondération critère i)
```

avec :

```text
Somme pondérations = 100 %
Score normalisé entre 0 et 100
IPE entre 0 et 100
```

### Exemple numérique complet

```text
Rentabilité : score 80, poids 30 % -> 24
Durée cycle : score 70, poids 10 % -> 7
Investissement : score 60, poids 15 % -> 9
Risque : score 65, poids 15 % -> 9,75
Trésorerie : score 50, poids 10 % -> 5
Marché : score 75, poids 10 % -> 7,5
Main-d'œuvre : score 60, poids 10 % -> 6
IPE = 24 + 7 + 9 + 9,75 + 5 + 7,5 + 6 = 68,25 / 100
```

### Cas particuliers

- Critère manquant : IPE provisoire ou non calculable selon criticité.
- Pondérations changées : historiser la version.

### Règles de validation

- Pondérations totalisant 100 %.
- Scores compris entre 0 et 100.

### Cas d'erreur

- Score hors plage : erreur.
- Pondération négative : erreur.

### Impact sur les autres indicateurs

Impact sur classement, tableau de bord et recommandations.

## 11.2 Normalisation des indicateurs

### Définition

La normalisation transforme un indicateur brut en score de 0 à 100.

### Objectif

Comparer des indicateurs de nature différente.

### Variables utilisées

- Valeur observée.
- Valeur minimale de référence.
- Valeur maximale de référence.
- Sens du critère.

### Formule mathématique

Pour un critère où une valeur élevée est meilleure :

```text
Score = (Valeur - Minimum) / (Maximum - Minimum) × 100
```

Pour un critère où une valeur faible est meilleure :

```text
Score = (Maximum - Valeur) / (Maximum - Minimum) × 100
```

Le score est borné entre 0 et 100.

### Exemple numérique complet

Rentabilité observée = 30 %, minimum = 0 %, maximum = 50 % :

```text
Score = (30 - 0) / (50 - 0) × 100 = 60
```

Besoin de trésorerie observé = 400 000 FCFA, minimum = 100 000, maximum = 700 000, faible est meilleur :

```text
Score = (700 000 - 400 000) / (700 000 - 100 000) × 100 = 50
```

### Cas particuliers

- Minimum égal maximum : utiliser score neutre 50 ou bloquer selon paramètre.

### Règles de validation

- Maximum supérieur au minimum.

### Cas d'erreur

- Plage invalide : normalisation impossible.

### Impact sur les autres indicateurs

Impact direct sur IPE.

## 11.3 Classement A à E

### Définition

Le classement A à E traduit le score IPE en classe qualitative.

### Objectif

Faciliter l'interprétation du score.

### Variables utilisées

- Score IPE.
- Seuils de classe.

### Formule mathématique

```text
A : 80 à 100
B : 65 à 79,99
C : 50 à 64,99
D : 35 à 49,99
E : 0 à 34,99
```

### Exemple numérique complet

```text
IPE = 68,25
Classe = B
```

### Cas particuliers

- Les seuils doivent être paramétrables.

### Règles de validation

- Seuils continus et non chevauchants.

### Cas d'erreur

- Score IPE hors 0-100 : erreur.

### Impact sur les autres indicateurs

Impact sur alertes, recommandations et tableau de bord.

## 11.4 Historique des pondérations

### Définition

L'historique des pondérations conserve les paramètres utilisés lors de chaque calcul IPE.

### Objectif

Permettre de reproduire les classements passés même si les pondérations changent.

### Variables utilisées

- Version pondération.
- Date d'application.
- Pondérations par critère.
- Auteur de modification.

### Formule mathématique

Pas de formule financière ; règle de traçabilité :

```text
IPE historisé = Scores critères + Pondérations versionnées + Date calcul
```

### Exemple numérique complet

Un IPE calculé le 01/07 avec rentabilité pondérée à 30 % doit conserver ce 30 %, même si la pondération passe à 35 % le 01/09.

### Cas particuliers

- Recalcul volontaire avec nouvelle pondération : créer une nouvelle version de résultat.

### Règles de validation

- Toute pondération utilisée doit avoir une version.

### Cas d'erreur

- Résultat IPE sans version de pondération : alerte d'audit.

### Impact sur les autres indicateurs

Impact sur comparaisons historiques et audit.

---

# 12. Alertes

## 12.1 Règle générale des alertes

### Définition

Une alerte est un message automatique déclenché lorsqu'une valeur dépasse un seuil, est incohérente ou présente un risque économique.

### Objectif

Aider le conseiller à détecter rapidement les problèmes.

### Variables utilisées

- Indicateur surveillé.
- Seuil.
- Sens du seuil.
- Niveau d'alerte.

### Formule mathématique

```text
Alerte déclenchée si Valeur indicateur respecte Condition seuil
```

### Exemple numérique complet

```text
Résultat économique = -50 000 FCFA
Seuil alerte résultat négatif = < 0
Alerte = Résultat déficitaire
```

### Cas particuliers

- Certaines alertes peuvent être informatives et non bloquantes.

### Règles de validation

- Les seuils doivent être paramétrables.
- Chaque alerte doit avoir un message clair.

### Cas d'erreur

- Seuil absent : alerte non évaluée.

### Impact sur les autres indicateurs

Impact sur tableau de bord, diagnostic et rapport.

## 12.2 Alertes minimales

### Définition

Les alertes minimales couvrent résultat négatif, coût anormal, rendement incohérent, prix incohérent, marge faible, risque élevé et besoin de trésorerie important.

### Objectif

Garantir un contrôle économique de base.

### Variables utilisées

- Résultat.
- Coût unitaire.
- Rendement.
- Prix.
- Marge.
- Score risque.
- Besoin de trésorerie.

### Formule mathématique

Exemples :

```text
Alerte résultat négatif si Résultat < 0
Alerte marge faible si Marge nette / Produit brut < Seuil marge
Alerte risque élevé si Score risque > Seuil risque
Alerte trésorerie si Besoin maximal / Produit brut > Seuil trésorerie
```

### Exemple numérique complet

```text
Marge nette = 40 000 FCFA
Produit brut = 800 000 FCFA
Taux marge = 40 000 / 800 000 = 5 %
Seuil marge faible = 10 %
Alerte marge faible déclenchée
```

### Cas particuliers

- Produit brut nul : taux de marge non calculable, alerte critique.

### Règles de validation

- Les seuils doivent être validés par le propriétaire du projet.

### Cas d'erreur

- Division par zéro dans un taux : indicateur non applicable avec alerte.

### Impact sur les autres indicateurs

Impact sur diagnostic automatique et recommandations.

---

# 13. Cas particuliers

## 13.1 Production ou rendement nul

### Définition

Une production nulle signifie qu'aucune quantité n'a été obtenue. Un rendement nul est une production nulle rapportée à une base positive.

### Objectif

Traiter les échecs de production sans provoquer d'erreurs de calcul.

### Variables utilisées

- Production.
- Superficie ou effectif.
- Causes d'échec.

### Formule mathématique

```text
Rendement = 0 / Base positive = 0
Coût unitaire = Coût total / 0 = Non applicable
```

### Exemple numérique complet

```text
Production = 0 kg
Superficie = 1 ha
Charges = 200 000 FCFA
Rendement = 0 kg/ha
Coût par kg = non applicable
Résultat = 0 - 200 000 = -200 000 FCFA
```

### Cas particuliers

- Culture pérenne en année improductive : production nulle attendue.

### Règles de validation

- Cause obligatoire pour production nulle en phase productive.

### Cas d'erreur

- Tentative de calcul coût unitaire : non applicable.

### Impact sur les autres indicateurs

Impact sur résultat, alertes, IQD et diagnostic.

## 13.2 Superficie ou effectif nul

### Définition

Une base physique nulle signifie qu'aucune surface ou aucun effectif n'est renseigné.

### Objectif

Empêcher les divisions par zéro et les comptes invalides.

### Variables utilisées

- Superficie.
- Effectif.
- Type de spéculation.

### Formule mathématique

Aucune division par zéro n'est autorisée.

### Exemple numérique complet

```text
Charges = 100 000 FCFA
Superficie = 0 ha
Coût par hectare = non calculable
```

### Cas particuliers

- Compte brouillon : valeur manquante autorisée.
- Compte validé : valeur positive obligatoire.

### Règles de validation

- Culture validée : superficie > 0.
- Élevage validé : effectif ou biomasse > 0.

### Cas d'erreur

- Compte validé avec base nulle : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur rendements, coûts par base physique et IQD.

## 13.3 Données manquantes

### Définition

Une donnée manquante est une information nécessaire qui n'a pas été renseignée.

### Objectif

Distinguer zéro, non applicable et inconnu.

### Variables utilisées

- Statut du champ.
- Valeur.
- Criticité du champ.

### Formule mathématique

Pour la qualité des données :

```text
Pénalité données manquantes = Somme(Poids champs obligatoires manquants)
```

### Exemple numérique complet

Si le prix de vente est manquant, le chiffre d'affaires ne peut pas être calculé :

```text
Quantité vendue = 1 000 kg
Prix = manquant
Montant vente = non calculable
```

### Cas particuliers

- Donnée estimée : calcul autorisé avec pénalité IQD.

### Règles de validation

- Champs critiques obligatoires pour validation.

### Cas d'erreur

- Champ obligatoire manquant : validation bloquée.

### Impact sur les autres indicateurs

Impact sur tous les calculs dépendants et sur l'IQD.

## 13.4 Ventes sans récolte et récoltes sans vente

### Définition

Une vente sans récolte est une vente saisie sans production disponible. Une récolte sans vente est une production non vendue, à affecter à stock, autoconsommation ou pertes.

### Objectif

Maintenir la cohérence physique.

### Variables utilisées

- Récoltes.
- Ventes.
- Stocks.
- Autoconsommation.
- Pertes.

### Formule mathématique

```text
Ventes <= Stock entrant + Production - Pertes - Autoconsommation
```

### Exemple numérique complet

```text
Production = 1 000 kg
Stock entrant = 0
Ventes = 1 200 kg
Écart = 1 000 - 1 200 = -200 kg
Alerte vente sans quantité disponible
```

### Cas particuliers

- Vente issue d'un stock entrant : autorisée si stock renseigné.

### Règles de validation

- Toute vente doit être couverte par production ou stock.

### Cas d'erreur

- Vente non couverte : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur produit brut, stocks, IQD et alertes.

## 13.5 Pertes supérieures à la production et stock négatif

### Définition

Une perte supérieure à la quantité disponible ou un stock négatif indique une incohérence physique.

### Objectif

Bloquer les résultats non crédibles.

### Variables utilisées

- Production.
- Stock entrant.
- Pertes.
- Stock sortant.

### Formule mathématique

```text
Pertes <= Production + Stock entrant
Stock sortant >= 0
```

### Exemple numérique complet

```text
Production = 500 kg
Stock entrant = 0
Pertes = 600 kg
Pertes disponibles = 500 kg
Erreur = pertes supérieures de 100 kg
```

### Cas particuliers

- Correction de stock : doit être enregistrée séparément.

### Règles de validation

- Incohérence physique bloquante pour validation.

### Cas d'erreur

- Stock négatif : erreur bloquante.

### Impact sur les autres indicateurs

Impact sur IQD, rendement, produit brut, stock et alertes.

---

# 14. Moteurs de calcul

## 14.1 Moteur 1 — Compte d'exploitation

### Rôle

Calculer produits, charges, marges, résultats et coûts de production.

### Données d'entrée

- Producteur, exploitation, spéculation, campagne, cycle.
- Charges.
- Produits.
- Stocks.
- Pertes.
- Paramètres de valorisation.

### Calculs

- Produit brut.
- Charges variables et fixes.
- Amortissements.
- Résultat financier.
- Résultat économique.
- Coûts unitaires.

### Résultats

- Compte d'exploitation complet.
- Indicateurs de base.
- Alertes de cohérence.

### Interactions avec les autres moteurs

- Fournit les résultats au moteur de rentabilité.
- Fournit les flux au moteur de trésorerie.
- Fournit les indicateurs au moteur de comparaison et à l'IPE.

## 14.2 Moteur 2 — Analyse de trésorerie

### Rôle

Calculer les flux entrants, sortants, soldes et besoin maximal de trésorerie.

### Données d'entrée

- Dates d'encaissement.
- Dates de paiement.
- Avances.
- Crédits.
- Subventions.
- Remboursements.

### Calculs

- Solde par période.
- Trésorerie minimale.
- Besoin maximal.
- Capacité de remboursement.

### Résultats

- Calendrier de trésorerie.
- Besoin de financement.
- Alertes de tension.

### Interactions

- Reçoit charges et produits du compte d'exploitation.
- Alimente l'IPE et le diagnostic financier.

## 14.3 Moteur 3 — Rentabilité

### Rôle

Calculer ROI, ratios, marges, seuils et délais de récupération.

### Données d'entrée

- Résultats économiques et financiers.
- Investissements.
- Charges.
- Produits.
- Durée.

### Calculs

- ROI.
- Ratio bénéfice/coût.
- Rentabilité annuelle.
- Seuil de rentabilité.
- Délai de récupération.
- VAN/TRI pour versions futures.

### Résultats

- Indicateurs de rentabilité.
- Alertes de faible performance.

### Interactions

- Alimente comparaison, IPE et diagnostic.

## 14.4 Moteur 4 — Comparaison des spéculations

### Rôle

Comparer plusieurs spéculations sur une base homogène.

### Données d'entrée

- Indicateurs annualisés.
- Risque.
- Marché.
- Durée.
- Trésorerie.
- Investissement.

### Calculs

- Classements par indicateur.
- Classement multicritère.
- Écarts entre spéculations.

### Résultats

- Rang des spéculations.
- Tableaux comparatifs.

### Interactions

- Utilise les résultats des moteurs compte, trésorerie et rentabilité.
- Alimente le tableau de bord et l'assistant de diagnostic.

## 14.5 Moteur 5 — Simulation économique

### Rôle

Simuler l'effet de variations de prix, rendement, charges, pertes, cycles ou investissements.

### Données d'entrée

- Compte de référence.
- Variables à modifier.
- Scénarios.

### Calculs

- Recalcul des indicateurs sous hypothèses modifiées.
- Écarts scénario/base.

### Résultats

- Scénarios optimiste, central, prudent.
- Sensibilité des résultats.

### Interactions

- Utilise tous les moteurs économiques de base.
- Alimente le diagnostic et les recommandations.

## 14.6 Moteur 6 — Prévision économique

### Rôle

Estimer automatiquement les besoins, coûts, productions et bénéfices à partir de référentiels.

### Données d'entrée

- Spéculation.
- Superficie ou effectif.
- Rendements de référence.
- Charges types.
- Prix de référence.

### Calculs

- Compte prévisionnel.
- Besoins d'intrants.
- Coûts prévus.
- Produits attendus.

### Résultats

- Budget prévisionnel.
- Besoin de trésorerie prévisionnel.
- Rentabilité prévisionnelle.

### Interactions

- Fournit une base au compte réel.
- Alimente la comparaison prévision/réalisation.

## 14.7 Moteur 7 — Assistant de diagnostic

### Rôle

Générer des commentaires, points forts, points faibles et recommandations à partir des indicateurs.

### Données d'entrée

- Résultats économiques.
- Alertes.
- IPE.
- IQD.
- Comparaison.

### Calculs

- Identification des postes critiques.
- Sélection de messages selon règles métier.
- Priorisation des recommandations.

### Résultats

- Commentaires automatiques.
- Points forts.
- Points faibles.
- Recommandations.

### Interactions

- Utilise les résultats de tous les moteurs.
- Alimente rapports et tableau de bord.

---

# 15. Référentiel métier

## 15.1 Chaîne logique d'une exploitation agricole

### Définition

AGRIECO PRO modélise l'exploitation agricole selon une chaîne métier commune au mode individuel et au mode conseiller/organisation. En mode individuel, le producteur voit cette chaîne comme son propre parcours de gestion ; en mode organisation, elle devient la base de suivi et de consolidation de plusieurs producteurs.

La chaîne métier de référence est la suivante :

```text
Producteur
↓
Exploitation
↓
Parcelle ou atelier
↓
Spéculation
↓
Cycle
↓
Opérations
↓
Charges
↓
Récoltes ou productions
↓
Ventes, autoconsommation, stocks et pertes
↓
Résultats
↓
Indicateurs
↓
Rapports
```

### Objectif

Clarifier les relations métier indépendamment de toute technologie.

### Variables utilisées

- Producteur.
- Exploitation.
- Parcelle.
- Atelier.
- Spéculation.
- Cycle.
- Opérations.
- Flux physiques et financiers.

### Formule mathématique

La chaîne métier est relationnelle :

```text
Un producteur possède une ou plusieurs exploitations.
Une exploitation contient une ou plusieurs parcelles ou ateliers.
Une parcelle ou un atelier porte une ou plusieurs spéculations selon les périodes.
Une spéculation comporte un ou plusieurs cycles.
Un cycle contient des opérations, charges, productions et ventes.
```

### Exemple numérique complet

Un producteur possède une exploitation de 5 ha. Sur 2 ha, il cultive du maïs pendant la campagne 2026. Le cycle contient des opérations de préparation, semis, fertilisation, entretien, récolte et vente. Les charges totalisent 400 000 FCFA et les produits 700 000 FCFA. Le résultat est 300 000 FCFA.

### Cas particuliers

- Une parcelle peut accueillir plusieurs cultures dans l'année.
- Une charge peut être partagée entre plusieurs spéculations.
- Un atelier d'élevage peut ne pas être lié à une parcelle.

### Règles de validation

- Un compte doit être rattaché à une spéculation et une période.
- Les relations doivent être cohérentes dans le temps.

### Cas d'erreur

- Cycle sans spéculation : erreur.
- Vente sans cycle : erreur.

### Impact sur les autres indicateurs

Structure tous les calculs, rapports, comparaisons et historiques.

---

# 16. Qualité des données

## 16.1 Indice de Qualité des Données — IQD

### Définition

L'Indice de Qualité des Données, ou IQD, est un score en pourcentage évaluant la complétude, la précision, la cohérence et la traçabilité des données saisies.

### Objectif

Indiquer le niveau de confiance à accorder aux résultats.

### Variables utilisées

- Nombre ou poids des champs obligatoires renseignés.
- Nombre ou poids des données estimées.
- Nombre ou poids des incohérences.
- Nombre ou poids des données non traçables.

### Formule mathématique

Méthode proposée :

```text
IQD = 100 - Pénalité données manquantes - Pénalité données estimées - Pénalité incohérences - Pénalité absence traçabilité
```

Le score est borné entre 0 et 100.

Pondération proposée :

```text
Complétude : 40 points
Cohérence : 30 points
Précision/estimation : 20 points
Traçabilité : 10 points
```

### Explication détaillée

Un compte peut être calculé avec des données estimées, mais son IQD doit diminuer. Les rapports doivent afficher le score et les principaux motifs de pénalité.

### Exemple numérique complet

Un compte présente :

- un prix de vente estimé : pénalité 5 points ;
- une charge de transport manquante mais non critique : pénalité 4 points ;
- une incohérence de stock corrigée : pénalité 6 points ;
- une source de prix non renseignée : pénalité 3 points.

```text
IQD = 100 - 5 - 4 - 6 - 3 = 82 %
```

Interprétation : données de bonne qualité, mais résultats à commenter.

### Cas particuliers

- Erreur bloquante non corrigée : le compte ne doit pas être validé, quel que soit l'IQD.
- Donnée non applicable : aucune pénalité.
- Donnée estimée validée par conseiller : pénalité réduite possible si règle validée.

### Règles de validation

- Les données obligatoires doivent être définies par module et par spéculation.
- Les données estimées doivent être signalées.
- Les incohérences physiques majeures bloquent la validation.
- L'IQD doit être affiché dans les rapports.

### Cas d'erreur

- IQD inférieur à un seuil minimal : rapport affiché avec avertissement.
- Données critiques manquantes : calcul bloqué ou résultat provisoire.

### Impact sur les autres indicateurs

L'IQD n'entre pas directement dans le produit brut ou les marges, mais influence la confiance, les alertes, le diagnostic et la décision.

## 16.2 Effet du niveau de saisie sur l'IQD

### Définition

Le niveau de saisie indique si le compte a été établi avec un jeu minimal de données ou avec une description complète des opérations, flux, stocks, dettes, créances et trésorerie.

### Objectif

Permettre à l'entrepreneur agricole d'obtenir rapidement une première analyse tout en signalant que les résultats simplifiés sont moins précis qu'une saisie détaillée.

### Variables utilisées

- Niveau de saisie : simplifié ou détaillé.
- Nombre de données estimées par référentiel.
- Nombre de rubriques non détaillées.
- Présence ou absence de flux de trésorerie datés.

### Formule mathématique

```text
IQD ajusté = IQD de base - Pénalité niveau simplifié - Pénalités estimations supplémentaires
```

### Explication détaillée

La saisie simplifiée peut suffire pour une décision rapide, mais elle ne doit pas donner la même confiance qu'un compte détaillé. Les rapports doivent donc afficher le niveau de saisie et le score IQD.

### Exemple numérique complet

```text
IQD de base après contrôles = 88 %
Pénalité saisie simplifiée = 8 points
Pénalité charges estimées = 5 points
IQD ajusté = 88 - 8 - 5 = 75 %
```

### Cas particuliers

- Si toutes les données simplifiées obligatoires sont renseignées mais que les charges secondaires sont estimées, le compte peut rester exploitable avec avertissement.
- Un compte détaillé peut avoir un mauvais IQD si ses données sont incohérentes.

### Règles de validation

- Le niveau de saisie doit être enregistré.
- Les données estimées doivent être identifiées.
- Les recommandations automatiques doivent tenir compte de l'IQD.

### Cas d'erreur

- Niveau de saisie inconnu : alerte de qualité.
- Recommandation forte produite avec IQD inférieur au seuil minimal : alerte ou blocage selon paramètre.

### Impact sur les autres indicateurs

Impact sur rapports, recommandations, comparaisons et assistant de diagnostic.

## 16.2 Données obligatoires, estimées et incohérences

### Définition

- **Donnée obligatoire** : donnée indispensable à la validation d'un compte.
- **Donnée estimée** : donnée approximative acceptée mais signalée.
- **Incohérence** : contradiction logique entre données.

### Objectif

Structurer le contrôle qualité.

### Variables utilisées

- Statut de chaque champ.
- Poids qualité.
- Criticité.

### Formule mathématique

```text
Score complétude = Points champs obligatoires renseignés / Points champs obligatoires totaux × 100
Score cohérence = 100 - Pénalités incohérences
```

### Exemple numérique complet

```text
10 champs obligatoires de 4 points chacun = 40 points
8 renseignés = 32 points
Score complétude partiel = 32 / 40 × 100 = 80 %
```

### Cas particuliers

- Champ obligatoire pour culture mais non applicable à élevage.

### Règles de validation

- Les champs obligatoires doivent dépendre de la spéculation.
- Les estimations doivent être visibles dans le rapport.

### Cas d'erreur

- Donnée obligatoire absente dans compte validé : erreur.

### Impact sur les autres indicateurs

Impact sur IQD, alertes et diagnostic.

---

# 17. Règles d'arrondi et de présentation

## 17.1 Arrondis

### Définition

Les arrondis déterminent la présentation des résultats sans modifier les calculs internes de référence.

### Objectif

Obtenir des rapports lisibles et cohérents.

### Variables utilisées

- Montants monétaires.
- Quantités.
- Pourcentages.
- Scores.

### Formule mathématique

```text
Montants FCFA présentés = arrondis à l'unité FCFA ou à la dizaine selon paramètre
Pourcentages = arrondis à 2 décimales
Scores IPE/IQD = arrondis à 1 ou 2 décimales selon paramètre
```

### Exemple numérique complet

```text
ROI brut = 33,333333 %
ROI présenté = 33,33 %
IPE brut = 68,246
IPE présenté = 68,25 ou 68,2 selon paramètre
```

### Cas particuliers

- Les calculs cumulés doivent utiliser les valeurs non arrondies si disponibles.

### Règles de validation

- Ne pas recalculer un indicateur à partir de valeurs déjà arrondies si cela crée un écart.

### Cas d'erreur

- Écart d'arrondi significatif dans un rapport : afficher note ou recalculer depuis sources.

### Impact sur les autres indicateurs

Impact sur présentation, audit et comparaison fine.

---

# 18. Synthèse des règles préparées pour versions futures

Les règles suivantes sont préparées mais non obligatoires dans la V1 :

- VAN pour cultures pérennes et investissements longs.
- TRI pour cultures pérennes et investissements longs.
- Intérêts dégressifs et échéanciers détaillés de crédit.
- Simulations avancées multi-scénarios.
- Prévision automatique complète à partir de référentiels technico-économiques.
- Assistant de diagnostic avancé générant des recommandations détaillées.
- Analyse multi-pays et multi-devise.
- Consolidation multi-utilisateur ou synchronisation serveur.

Ces règles doivent être prises en compte dans l'architecture afin d'éviter une refonte majeure lors des évolutions Web, Android ou SQL.
