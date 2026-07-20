# Modèle conceptuel de données — AGRIECO PRO

## Statut du document

Ce document définit le **modèle conceptuel officiel des données métier d'AGRIECO PRO**. Il est indépendant de toute technologie et ne décrit aucune implémentation VBA, SQL, Android, Web ou Excel.

Il sert de référence pour les futures implémentations et doit rester cohérent avec les documents officiels de conception métier, d'architecture, de règles de calcul et de matrice d'arbitrages.

---

# 1. Principes généraux

## 1.1 Objectifs du modèle

Le modèle conceptuel de données a pour objectifs de :

- représenter les objets métier nécessaires à la gestion technico-économique d'une exploitation agricole ou d'élevage ;
- garantir la cohérence entre mode individuel et mode conseiller/organisation ;
- séparer clairement les données saisies, calculées et paramétrables ;
- préparer la traçabilité des comptes d'exploitation, résultats, simulations, alertes et rapports ;
- permettre une future implémentation dans Excel/VBA, SQL, Windows, Android ou Web sans refonte majeure ;
- fournir une base commune aux développeurs, experts métier, conseillers agricoles et propriétaires du produit.

## 1.2 Indépendance technologique

Le modèle décrit uniquement les concepts métier, leurs attributs et leurs relations. Les termes table, feuille, classe, formulaire, API ou écran ne sont pas utilisés comme objets d'implémentation dans ce document.

Une entité conceptuelle pourra plus tard devenir :

- une table structurée dans Excel ;
- une table SQL ;
- une ressource API ;
- un objet applicatif ;
- un écran ou une section de formulaire.

Cette transformation relève des documents techniques ultérieurs.

## 1.3 Séparation des données

| Catégorie | Définition | Exemples |
|---|---|---|
| Données saisies | Données fournies par l'utilisateur ou importées | Quantité vendue, prix, superficie, charge, date d'opération |
| Données calculées | Données produites par les règles de calcul | Produit brut, marge, ROI, IPE, IQD, coût unitaire |
| Données paramétrables | Données modifiables par un administrateur métier sans changer le moteur | Seuils d'alerte, pondérations IPE, unités, durées d'amortissement |
| Données référentielles | Données de classification ou de référence | Types de spéculations, catégories de charges, unités, zones |
| Données d'audit | Données décrivant l'historique et les modifications | Date de validation, auteur, ancien statut, nouvelle valeur |

## 1.4 Règles de normalisation conceptuelle

- Chaque entité métier doit avoir un identifiant unique stable.
- Les noms visibles ne doivent pas servir d'identifiants uniques.
- Les données calculées doivent être traçables vers leurs données sources et paramètres utilisés.
- Les entités principales ne doivent pas être supprimées physiquement ; elles doivent pouvoir être désactivées, archivées ou historisées.
- Les valeurs paramétrables doivent être versionnables lorsque leur modification peut changer un résultat.
- Les flux physiques, financiers et économiques doivent être distingués.
- Les données communes aux cultures et élevages doivent être mutualisées dans des concepts génériques lorsque possible.
- Les spécificités par spéculation doivent être portées par des paramètres ou attributs spécialisés, sans dupliquer tout le modèle.

---

# 2. Catalogue complet des entités métier

## 2.1 Vue d'ensemble des familles d'entités

| Famille | Entités principales |
|---|---|
| Acteurs | Organisation, Producteur, Utilisateur, Rôle, Client, Fournisseur |
| Structure agricole | Exploitation, Parcelle, Atelier d'élevage |
| Activités | Spéculation, Cycle de production, Opération culturale ou zootechnique |
| Flux physiques | Intrant, Récolte, Produit, Stock, Perte |
| Flux économiques et financiers | Charge, Vente, Trésorerie, Dette, Créance, Investissement, Immobilisation, Amortissement |
| Résultats | Compte d'exploitation, Indicateur, IPE, IQD, Alerte, Rapport |
| Paramétrage | Paramètre métier, Référentiel, Unité, Conversion, Prix de référence, Rendement de référence, Pondération |
| Audit | Historique, Version de paramètres, Journal de validation |

---

# 3. Dictionnaire conceptuel des entités

> Convention des types conceptuels : `Texte`, `Identifiant`, `Nombre`, `Montant`, `Pourcentage`, `Date`, `DateHeure`, `Booléen`, `Enumération`, `Liste`, `Objet lié`, `Calculé`.

## 3.1 Organisation

### Objectif

Représenter une structure qui utilise AGRIECO PRO pour accompagner plusieurs producteurs ou exploitations.

### Description

Une organisation peut être une coopérative, ONG, projet, cabinet de conseil, institution financière, centre de formation ou grande institution agricole.

### Identifiant unique

`OrganisationID` — identifiant stable et obligatoire.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| OrganisationID | Identifiant | Oui | Saisie/généré | Unique et non vide | Oui |
| NomOrganisation | Texte | Oui | Saisie | Non vide | Oui |
| TypeOrganisation | Enumération | Oui | Saisie | Coopérative, ONG, projet, cabinet, institution, formation, autre | Oui |
| ContactPrincipal | Texte | Non | Saisie | Format libre contrôlé | Non |
| ZoneIntervention | Texte/Liste | Non | Saisie | Cohérent avec référentiel zones | Oui |
| Statut | Enumération | Oui | Saisie | Active, inactive, archivée | Oui |

## 3.2 Producteur

### Objectif

Identifier un entrepreneur agricole ou producteur accompagné.

### Description

Le producteur peut utiliser AGRIECO PRO en mode individuel ou être suivi par une organisation.

### Identifiant unique

`ProducteurID` — identifiant stable et obligatoire.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ProducteurID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| OrganisationID | Objet lié | Non | Saisie | Obligatoire si mode organisation | Oui |
| NomProducteur | Texte | Oui | Saisie | Non vide | Oui |
| Contact | Texte | Non | Saisie | Format téléphone recommandé | Non |
| Localisation | Texte/Objet lié | Oui | Saisie | Zone connue ou saisie libre validée | Oui |
| TypeUtilisateurMetier | Enumération | Oui | Saisie | Entrepreneur, producteur structuré, bénéficiaire | Oui |
| Statut | Enumération | Oui | Saisie | Actif, inactif, archivé | Oui |

## 3.3 Exploitation

### Objectif

Représenter l'unité agricole gérée ou analysée.

### Description

Une exploitation regroupe des parcelles, ateliers, activités, ressources et historiques de campagnes.

### Identifiant unique

`ExploitationID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ExploitationID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| ProducteurID | Objet lié | Oui | Saisie | Producteur existant | Oui |
| NomExploitation | Texte | Non | Saisie | Recommandé | Oui |
| SuperficieTotale | Nombre | Non | Saisie | >= 0 | Oui |
| UniteSuperficie | Objet lié | Non | Saisie | Convertible en hectare | Oui |
| Localisation | Texte/Objet lié | Oui | Saisie | Cohérente | Oui |
| Statut | Enumération | Oui | Saisie | Active, inactive, archivée | Oui |

## 3.4 Parcelle

### Objectif

Décrire une surface agricole utilisée pour une ou plusieurs spéculations végétales.

### Description

Une parcelle appartient à une exploitation et peut porter plusieurs cycles ou cultures dans le temps.

### Identifiant unique

`ParcelleID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ParcelleID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| ExploitationID | Objet lié | Oui | Saisie | Exploitation existante | Oui |
| NomParcelle | Texte | Non | Saisie | Recommandé | Oui |
| Superficie | Nombre | Oui | Saisie | > 0 pour parcelle active | Oui |
| UniteSuperficie | Objet lié | Oui | Saisie | Convertible en hectare | Oui |
| TypeSol | Texte/Référentiel | Non | Saisie | Valeur référentielle si disponible | Non |
| Statut | Enumération | Oui | Saisie | Active, repos, archivée | Oui |

## 3.5 Atelier d'élevage

### Objectif

Représenter une unité de production animale ou aquacole.

### Description

Un atelier d'élevage peut être un poulailler, une porcherie, un bassin piscicole ou autre unité zootechnique.

### Identifiant unique

`AtelierID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| AtelierID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| ExploitationID | Objet lié | Oui | Saisie | Exploitation existante | Oui |
| NomAtelier | Texte | Non | Saisie | Recommandé | Oui |
| TypeAtelier | Enumération | Oui | Saisie | Volaille, porc, bovin, ovin, caprin, pisciculture, autre | Oui |
| Capacité | Nombre | Non | Saisie | >= 0 | Oui |
| UniteCapacite | Objet lié | Non | Saisie | Tête, kg, bassin, m², autre | Oui |
| Statut | Enumération | Oui | Saisie | Actif, inactif, archivé | Oui |

## 3.6 Spéculation

### Objectif

Définir une activité productive agricole ou d'élevage analysable.

### Description

La spéculation regroupe cultures et élevages dans un concept commun afin d'utiliser un moteur économique unique.

### Identifiant unique

`SpeculationID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| SpeculationID | Identifiant | Oui | Référentiel | Unique | Oui |
| NomSpeculation | Texte | Oui | Référentiel | Non vide | Oui |
| TypeSpeculation | Enumération | Oui | Référentiel | Culture, élevage, mixte | Oui |
| Categorie | Enumération | Oui | Référentiel | Annuelle, pérenne, maraîchage, volaille, porc, pisciculture, autre | Oui |
| UniteProductionReference | Objet lié | Oui | Référentiel | Unité active | Oui |
| DureeCycleDefaut | Nombre | Non | Référentiel | > 0 si cycle standard | Oui |
| NombreCyclesAnDefaut | Nombre | Non | Référentiel | >= 0 | Oui |
| Statut | Enumération | Oui | Référentiel | Active, inactive, future | Oui |

## 3.7 Cycle de production

### Objectif

Délimiter une période technique de production pour une spéculation.

### Description

Un cycle peut être prévisionnel, réel ou corrigé. Il rattache les opérations, charges, produits, stocks et résultats à une période.

### Identifiant unique

`CycleID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| CycleID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| SpeculationID | Objet lié | Oui | Saisie | Spéculation active | Oui |
| ExploitationID | Objet lié | Oui | Saisie | Exploitation existante | Oui |
| ParcelleID | Objet lié | Non | Saisie | Obligatoire pour culture sauf exception | Oui |
| AtelierID | Objet lié | Non | Saisie | Obligatoire pour élevage sauf exception | Oui |
| DateDebut | Date | Oui | Saisie | <= DateFin si DateFin connue | Oui |
| DateFin | Date | Non | Saisie | >= DateDebut | Oui |
| TypeCompte | Enumération | Oui | Saisie | Prévisionnel, réel, corrigé | Oui |
| StatutCycle | Enumération | Oui | Saisie | Brouillon, validé, archivé | Oui |

## 3.8 Opération culturale ou zootechnique

### Objectif

Décrire une action technique réalisée pendant un cycle.

### Description

Une opération peut être préparation, semis, fertilisation, traitement, alimentation, vaccination, récolte, transport ou autre.

### Identifiant unique

`OperationID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| OperationID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| TypeOperation | Référentiel | Oui | Saisie | Valeur active | Oui |
| DateOperation | Date | Non | Saisie | Dans période du cycle ou justification | Oui |
| Description | Texte | Non | Saisie | Libre | Non |
| Statut | Enumération | Oui | Saisie | Prévue, réalisée, annulée | Oui |

## 3.9 Intrant

### Objectif

Décrire une ressource consommée ou utilisée dans une opération.

### Description

Un intrant peut être semence, plant, engrais, pesticide, aliment, médicament, eau, énergie ou emballage.

### Identifiant unique

`IntrantID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| IntrantID | Identifiant | Oui | Référentiel/saisie | Unique | Oui |
| NomIntrant | Texte | Oui | Référentiel/saisie | Non vide | Oui |
| CategorieIntrant | Référentiel | Oui | Référentiel | Active | Oui |
| UniteReference | Objet lié | Oui | Référentiel | Convertible | Oui |
| Statut | Enumération | Oui | Référentiel | Actif, inactif | Oui |

## 3.10 Charge

### Objectif

Représenter une consommation de ressource ou une dépense affectée à un cycle, une opération ou une spéculation.

### Description

Une charge peut être payée, valorisée, fixe, variable, partagée, récurrente ou propre au premier cycle.

### Identifiant unique

`ChargeID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ChargeID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| OperationID | Objet lié | Non | Saisie | Cohérent avec cycle | Oui |
| CategorieCharge | Référentiel | Oui | Saisie | Valeur active | Oui |
| NatureCharge | Enumération | Oui | Saisie | Fixe, variable, investissement, amortissement, frais financier | Oui |
| TypePaiement | Enumération | Oui | Saisie | Payée, à payer, valorisée, apport familial | Oui |
| Quantite | Nombre | Non | Saisie | >= 0 | Oui |
| Unite | Objet lié | Non | Saisie | Convertible si quantité renseignée | Oui |
| CoutUnitaire | Montant | Non | Saisie | >= 0 | Oui |
| Montant | Montant/Calculé | Oui | Calculé/saisie contrôlée | >= 0 | Oui |
| Recurrente | Booléen | Non | Saisie | Oui/non | Oui |
| Partagee | Booléen | Non | Saisie | Oui/non | Oui |

## 3.11 Produit

### Objectif

Représenter une valeur générée par la production.

### Description

Un produit peut provenir de ventes, autoconsommation, produits secondaires, variation de stocks ou produits exceptionnels.

### Identifiant unique

`ProduitID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ProduitID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| NatureProduit | Enumération | Oui | Saisie | Vente, autoconsommation, secondaire, stock, exceptionnel | Oui |
| Designation | Texte | Oui | Saisie | Non vide | Oui |
| Quantite | Nombre | Non | Saisie | >= 0 | Oui |
| Unite | Objet lié | Non | Saisie | Convertible | Oui |
| PrixUnitaire | Montant | Non | Saisie/référence | >= 0 | Oui |
| Montant | Montant/Calculé | Oui | Calculé | >= 0 | Oui |

## 3.12 Récolte

### Objectif

Représenter une production physique obtenue pendant un cycle.

### Description

Une récolte peut être unique ou multiple, totale ou partielle, et peut alimenter ventes, autoconsommation, stocks ou pertes.

### Identifiant unique

`RecolteID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| RecolteID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| DateRecolte | Date | Non | Saisie | Dans période du cycle ou justification | Oui |
| QuantiteRecoltee | Nombre | Oui | Saisie | >= 0 | Oui |
| Unite | Objet lié | Oui | Saisie | Convertible | Oui |
| Qualite | Texte/Référentiel | Non | Saisie | Libre ou référentiel | Non |
| PertesAssociees | Nombre | Non | Saisie | >= 0 et <= quantité disponible | Oui |

## 3.13 Vente

### Objectif

Représenter une transaction de vente liée à un produit agricole ou animal.

### Description

Une vente peut être encaissée immédiatement ou générer une créance.

### Identifiant unique

`VenteID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| VenteID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| ClientID | Objet lié | Non | Saisie | Client existant si renseigné | Oui |
| DateVente | Date | Oui | Saisie | Cohérente avec cycle ou stock | Oui |
| QuantiteVendue | Nombre | Oui | Saisie | > 0 | Oui |
| Unite | Objet lié | Oui | Saisie | Convertible | Oui |
| PrixUnitaire | Montant | Oui | Saisie | >= 0 | Oui |
| MontantVente | Montant/Calculé | Oui | Calculé | Quantité × prix | Oui |
| StatutEncaissement | Enumération | Oui | Saisie | Encaissée, partielle, non encaissée | Oui |

## 3.14 Client

### Objectif

Identifier l'acheteur d'un produit.

### Description

Le client peut être un particulier, commerçant, coopérative, transformateur, institution ou autre acheteur.

### Identifiant unique

`ClientID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ClientID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| NomClient | Texte | Oui | Saisie | Non vide | Oui |
| TypeClient | Enumération | Non | Saisie | Particulier, commerçant, coopérative, institution, autre | Non |
| Contact | Texte | Non | Saisie | Format libre | Non |

## 3.15 Fournisseur

### Objectif

Identifier la source d'un intrant, service, équipement ou crédit fournisseur.

### Description

Le fournisseur peut générer une charge payée immédiatement ou une dette.

### Identifiant unique

`FournisseurID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| FournisseurID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| NomFournisseur | Texte | Oui | Saisie | Non vide | Oui |
| TypeFournisseur | Enumération | Non | Saisie | Intrants, service, équipement, financement, autre | Non |
| Contact | Texte | Non | Saisie | Format libre | Non |

## 3.16 Stock

### Objectif

Suivre les quantités disponibles d'intrants ou de produits agricoles.

### Description

Un stock peut concerner les intrants, produits récoltés, produits secondaires ou aliments d'élevage.

### Identifiant unique

`StockID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| StockID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| TypeStock | Enumération | Oui | Saisie | Intrant, produit agricole, produit secondaire | Oui |
| ElementStockeID | Objet lié | Oui | Saisie | Intrant ou produit existant | Oui |
| QuantiteInitiale | Nombre | Non | Saisie | >= 0 | Oui |
| Entrees | Nombre/Calculé | Non | Saisie/calculé | >= 0 | Oui |
| Sorties | Nombre/Calculé | Non | Saisie/calculé | >= 0 | Oui |
| QuantiteFinale | Nombre/Calculé | Oui | Calculé | >= 0 | Oui |
| Unite | Objet lié | Oui | Saisie | Convertible | Oui |

## 3.17 Trésorerie

### Objectif

Représenter les flux monétaires entrants et sortants.

### Description

La trésorerie suit les encaissements, décaissements, avances, crédits, remboursements, frais financiers et soldes par période.

### Identifiant unique

`FluxTresorerieID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| FluxTresorerieID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Non | Saisie | Obligatoire si flux lié à un cycle | Oui |
| NatureFlux | Enumération | Oui | Saisie | Entrée, sortie | Oui |
| CategorieFlux | Référentiel | Oui | Saisie | Vente, achat, avance, crédit, remboursement, intérêt, autre | Oui |
| Montant | Montant | Oui | Saisie | >= 0 | Oui |
| DateFlux | Date | Oui | Saisie | Date valide | Oui |
| StatutFlux | Enumération | Oui | Saisie | Prévu, réalisé, annulé | Oui |

## 3.18 Dette

### Objectif

Représenter un montant dû à un fournisseur, prêteur ou tiers.

### Description

Une dette peut provenir d'un achat à crédit, d'un emprunt, d'une charge non encore payée ou d'un engagement financier.

### Identifiant unique

`DetteID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| DetteID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| FournisseurID | Objet lié | Non | Saisie | Recommandé | Oui |
| CycleID | Objet lié | Non | Saisie | Si dette liée à activité | Oui |
| MontantInitial | Montant | Oui | Saisie | > 0 | Oui |
| MontantRembourse | Montant/Calculé | Non | Saisie/calculé | >= 0 et <= initial | Oui |
| SoldeDette | Montant/Calculé | Oui | Calculé | Initial - remboursé | Oui |
| DateEcheance | Date | Non | Saisie | >= date création | Oui |
| StatutDette | Enumération | Oui | Saisie | Ouverte, partielle, réglée, annulée | Oui |

## 3.19 Créance

### Objectif

Représenter un montant à recevoir d'un client ou tiers.

### Description

Une créance peut provenir d'une vente non encaissée ou partiellement encaissée.

### Identifiant unique

`CreanceID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| CreanceID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| ClientID | Objet lié | Non | Saisie | Recommandé | Oui |
| VenteID | Objet lié | Non | Saisie/calculé | Recommandé si créance liée à vente | Oui |
| MontantInitial | Montant | Oui | Saisie/calculé | > 0 | Oui |
| MontantEncaisse | Montant/Calculé | Non | Saisie/calculé | >= 0 et <= initial | Oui |
| SoldeCreance | Montant/Calculé | Oui | Calculé | Initial - encaissé | Oui |
| DateEcheance | Date | Non | Saisie | Date valide | Oui |
| StatutCreance | Enumération | Oui | Saisie | Ouverte, partielle, encaissée, annulée | Oui |

## 3.20 Investissement

### Objectif

Représenter une dépense durable destinée à générer des avantages sur plusieurs périodes.

### Description

Un investissement peut devenir une immobilisation amortissable.

### Identifiant unique

`InvestissementID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| InvestissementID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| Designation | Texte | Oui | Saisie | Non vide | Oui |
| DateAcquisition | Date | Oui | Saisie | Date valide | Oui |
| ValeurAcquisition | Montant | Oui | Saisie | > 0 | Oui |
| SourceFinancement | Enumération | Non | Saisie | Fonds propres, crédit, subvention, mixte | Oui |
| Affectation | Texte/Objet lié | Non | Saisie | Spéculation, exploitation, partagé | Oui |

## 3.21 Immobilisation

### Objectif

Représenter un actif durable utilisé par l'exploitation.

### Description

Une immobilisation porte les informations nécessaires au calcul d'amortissement.

### Identifiant unique

`ImmobilisationID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ImmobilisationID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| InvestissementID | Objet lié | Non | Saisie | Recommandé si issu d'un investissement | Oui |
| CategorieImmobilisation | Référentiel | Oui | Saisie | Active | Oui |
| ValeurAmortissable | Montant | Oui | Calculé/saisie | > 0 | Oui |
| ValeurResiduelle | Montant | Non | Saisie | >= 0 et <= valeur acquisition | Oui |
| DureeAmortissement | Nombre | Oui | Saisie/référence | > 0 | Oui |
| Statut | Enumération | Oui | Saisie | En service, sorti, archivé | Oui |

## 3.22 Amortissement

### Objectif

Représenter la charge économique calculée liée à une immobilisation.

### Description

L'amortissement est calculé selon une méthode linéaire avec durée paramétrable.

### Identifiant unique

`AmortissementID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| AmortissementID | Identifiant | Oui | Calculé/généré | Unique | Oui |
| ImmobilisationID | Objet lié | Oui | Calculé | Immobilisation existante | Oui |
| Periode | Date/Période | Oui | Calculé | Période valide | Oui |
| MontantAmortissement | Montant/Calculé | Oui | Calculé | >= 0 | Oui |
| Methode | Enumération | Oui | Paramètre | Linéaire pour V1 | Oui |
| TauxAffectation | Pourcentage | Non | Saisie | 0 à 100 % | Oui |

## 3.23 Compte d'exploitation

### Objectif

Regrouper les données et résultats économiques d'une spéculation sur un cycle ou une période.

### Description

Le compte d'exploitation est le pivot des calculs produits, charges, marges, résultats, rentabilité, IPE et IQD.

### Identifiant unique

`CompteID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| CompteID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| CycleID | Objet lié | Oui | Saisie | Cycle existant | Oui |
| TypeCompte | Enumération | Oui | Saisie | Prévisionnel, réel, corrigé | Oui |
| NiveauSaisie | Enumération | Oui | Saisie | Simplifié, détaillé | Oui |
| StatutCompte | Enumération | Oui | Saisie | Brouillon, validé, corrigé, archivé | Oui |
| DateValidation | DateHeure | Non | Calculé/saisie | Obligatoire si validé | Oui |
| VersionReglesCalcul | Texte | Oui | Paramètre | Version connue | Oui |

## 3.24 Indicateur

### Objectif

Stocker ou représenter un résultat calculé à partir d'un compte d'exploitation.

### Description

Les indicateurs incluent produit brut, charges totales, marges, bénéfices, coûts, ROI, seuils, IPE et IQD.

### Identifiant unique

`IndicateurID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| IndicateurID | Identifiant | Oui | Calculé/généré | Unique | Oui |
| CompteID | Objet lié | Oui | Calculé | Compte existant | Oui |
| NomIndicateur | Référentiel | Oui | Calculé | Valeur active | Oui |
| Valeur | Nombre/Montant/Pourcentage | Non | Calculé | Selon formule | Oui |
| Unite | Objet lié | Non | Calculé | Cohérente avec indicateur | Oui |
| StatutCalcul | Enumération | Oui | Calculé | Calculé, non applicable, erreur, provisoire | Oui |
| MessageCalcul | Texte | Non | Calculé | Obligatoire si erreur | Oui |

## 3.25 IPE

### Objectif

Représenter l'Indice de Performance Économique d'un compte ou d'une spéculation.

### Description

L'IPE est un score sur 100 calculé à partir de critères pondérés et historisés.

### Identifiant unique

`IPEID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| IPEID | Identifiant | Oui | Calculé/généré | Unique | Oui |
| CompteID | Objet lié | Oui | Calculé | Compte existant | Oui |
| VersionPonderation | Objet lié | Oui | Paramètre | Version active ou historisée | Oui |
| ScoreGlobal | Pourcentage/Nombre | Oui | Calculé | 0 à 100 | Oui |
| ClasseIPE | Enumération | Oui | Calculé | A, B, C, D, E | Oui |
| DetailScores | Liste | Oui | Calculé | Scores par critère | Oui |

## 3.26 IQD

### Objectif

Représenter l'Indice de Qualité des Données d'un compte.

### Description

L'IQD mesure la complétude, cohérence, précision et traçabilité des données.

### Identifiant unique

`IQDID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| IQDID | Identifiant | Oui | Calculé/généré | Unique | Oui |
| CompteID | Objet lié | Oui | Calculé | Compte existant | Oui |
| ScoreIQD | Pourcentage | Oui | Calculé | 0 à 100 | Oui |
| Penalites | Liste | Non | Calculé | Détail par motif | Oui |
| NiveauConfiance | Enumération | Oui | Calculé | Élevé, moyen, faible | Oui |
| DonneesCritiquesManquantes | Booléen | Oui | Calculé | Oui/non | Oui |

## 3.27 Alerte

### Objectif

Signaler automatiquement une anomalie, un risque ou une situation nécessitant attention.

### Description

Une alerte peut porter sur rendement, prix, marge, coût, trésorerie, risque, IQD ou résultat.

### Identifiant unique

`AlerteID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| AlerteID | Identifiant | Oui | Calculé/généré | Unique | Oui |
| CompteID | Objet lié | Non | Calculé | Obligatoire si alerte liée à compte | Oui |
| TypeAlerte | Référentiel | Oui | Calculé | Valeur active | Oui |
| NiveauAlerte | Enumération | Oui | Calculé | Info, vigilance, critique | Oui |
| Message | Texte | Oui | Calculé | Non vide | Oui |
| StatutAlerte | Enumération | Oui | Saisie/calculé | Ouverte, traitée, ignorée | Oui |

## 3.28 Rapport

### Objectif

Représenter une restitution structurée des données et indicateurs.

### Description

Un rapport peut concerner un producteur, une exploitation, un compte, une comparaison ou une consolidation.

### Identifiant unique

`RapportID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| RapportID | Identifiant | Oui | Généré | Unique | Oui |
| TypeRapport | Référentiel | Oui | Saisie | Compte, producteur, comparaison, consolidation, autre | Oui |
| Perimetre | Objet lié/Liste | Oui | Saisie | Au moins un objet source | Oui |
| DateGeneration | DateHeure | Oui | Calculé | Date valide | Oui |
| VersionDonnees | Texte | Non | Calculé | Recommandé | Oui |
| StatutRapport | Enumération | Oui | Calculé | Généré, archivé, annulé | Oui |

## 3.29 Paramètre métier

### Objectif

Définir une valeur modifiable sans changer les règles du moteur.

### Description

Un paramètre peut concerner seuils d'alerte, pondérations, durées, barèmes, unités, prix ou rendements.

### Identifiant unique

`ParametreID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ParametreID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| NomParametre | Texte | Oui | Saisie | Non vide | Oui |
| Domaine | Référentiel | Oui | Saisie | IPE, IQD, alerte, amortissement, unité, prix, rendement, autre | Oui |
| Valeur | Texte/Nombre/Montant/Pourcentage | Oui | Saisie | Selon domaine | Oui |
| DateDebutValidite | Date | Non | Saisie | <= date fin si fin connue | Oui |
| DateFinValidite | Date | Non | Saisie | >= début | Oui |
| Statut | Enumération | Oui | Saisie | Actif, inactif, archivé | Oui |

## 3.30 Référentiel

### Objectif

Définir les listes contrôlées utilisées par le modèle.

### Description

Les référentiels normalisent les catégories, types, unités, statuts, risques, marchés et familles de charges.

### Identifiant unique

`ReferentielID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| ReferentielID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| DomaineReferentiel | Texte | Oui | Saisie | Non vide | Oui |
| Code | Texte | Oui | Saisie | Unique dans domaine | Oui |
| Libelle | Texte | Oui | Saisie | Non vide | Oui |
| OrdreAffichage | Nombre | Non | Saisie | >= 0 | Non |
| Statut | Enumération | Oui | Saisie | Actif, inactif | Oui |

## 3.31 Utilisateur

### Objectif

Identifier une personne utilisant AGRIECO PRO.

### Description

Un utilisateur peut être entrepreneur agricole, conseiller, administrateur, superviseur ou lecteur.

### Identifiant unique

`UtilisateurID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| UtilisateurID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| NomUtilisateur | Texte | Oui | Saisie | Unique dans périmètre | Oui |
| NomComplet | Texte | Oui | Saisie | Non vide | Oui |
| RoleID | Objet lié | Oui | Saisie | Rôle actif | Oui |
| OrganisationID | Objet lié | Non | Saisie | Selon mode | Oui |
| ProducteurID | Objet lié | Non | Saisie | Obligatoire si utilisateur entrepreneur lié à producteur | Oui |
| Statut | Enumération | Oui | Saisie | Actif, inactif | Oui |

## 3.32 Rôle

### Objectif

Définir les droits et responsabilités d'un utilisateur.

### Description

Un rôle contrôle l'accès aux données, validations, rapports, paramètres et consolidations.

### Identifiant unique

`RoleID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| RoleID | Identifiant | Oui | Saisie/généré | Unique | Oui |
| NomRole | Texte | Oui | Saisie | Non vide | Oui |
| Portee | Enumération | Oui | Saisie | Individuel, organisation, administration | Oui |
| Droits | Liste | Oui | Saisie | Liste cohérente | Oui |
| Statut | Enumération | Oui | Saisie | Actif, inactif | Oui |

## 3.33 Historique

### Objectif

Conserver la trace des modifications, validations et corrections importantes.

### Description

L'historique garantit audit, traçabilité et reproductibilité.

### Identifiant unique

`HistoriqueID`.

### Attributs

| Attribut | Type | Obligatoire | Source | Règles de validation | Historique |
|---|---|---:|---|---|---|
| HistoriqueID | Identifiant | Oui | Généré | Unique | Oui |
| EntiteConcernee | Texte | Oui | Généré | Nom d'entité connu | Oui |
| IdentifiantEntite | Identifiant | Oui | Généré | Objet existant au moment de l'action | Oui |
| Action | Enumération | Oui | Généré | Création, modification, validation, correction, archivage | Oui |
| AncienneValeur | Texte | Non | Généré | Si applicable | Oui |
| NouvelleValeur | Texte | Non | Généré | Si applicable | Oui |
| UtilisateurID | Objet lié | Non | Généré | Utilisateur connu si connecté | Oui |
| DateAction | DateHeure | Oui | Généré | Date valide | Oui |

---

# 4. Relations entre les entités

## 4.1 Relations principales et cardinalités

| Relation | Cardinalité | Contrainte d'intégrité |
|---|---|---|
| Organisation — Producteur | Une organisation peut suivre 0 à N producteurs ; un producteur peut être lié à 0 ou 1 organisation en mode individuel, ou 1 organisation en mode organisation | Si mode organisation, `OrganisationID` obligatoire pour producteur suivi |
| Producteur — Exploitation | Un producteur possède 1 à N exploitations ; une exploitation appartient à 1 producteur | Une exploitation ne peut pas exister sans producteur |
| Exploitation — Parcelle | Une exploitation possède 0 à N parcelles ; une parcelle appartient à 1 exploitation | Parcelle obligatoire pour cultures sauf cas spécifique |
| Exploitation — Atelier | Une exploitation possède 0 à N ateliers ; un atelier appartient à 1 exploitation | Atelier obligatoire pour élevages sauf cas spécifique |
| Spéculation — Cycle | Une spéculation peut être utilisée dans 0 à N cycles ; un cycle concerne 1 spéculation | Spéculation active requise pour cycle validé |
| Cycle — Opération | Un cycle contient 0 à N opérations ; une opération appartient à 1 cycle | Opérations facultatives en saisie simplifiée |
| Cycle — Charge | Un cycle contient 0 à N charges ; une charge appartient à 1 cycle | Charge validée liée à un cycle |
| Cycle — Produit | Un cycle contient 0 à N produits ; un produit appartient à 1 cycle | Produit validé lié à un cycle |
| Cycle — Récolte | Un cycle contient 0 à N récoltes ; une récolte appartient à 1 cycle | Récolte facultative pour années improductives |
| Récolte — Vente | Une récolte peut alimenter 0 à N ventes ; une vente peut être liée à 0 ou 1 récolte | Vente doit être couverte par production ou stock |
| Vente — Client | Un client peut avoir 0 à N ventes ; une vente peut avoir 0 ou 1 client identifié | Client facultatif en V1 |
| Fournisseur — Charge | Un fournisseur peut être lié à 0 à N charges ; une charge peut avoir 0 ou 1 fournisseur | Fournisseur recommandé pour dettes |
| Vente — Créance | Une vente peut générer 0 ou 1 créance ; une créance peut être liée à 0 ou 1 vente | Créance si vente non encaissée |
| Charge — Dette | Une charge peut générer 0 ou 1 dette ; une dette peut être liée à 0 ou 1 charge | Dette si charge non payée |
| Investissement — Immobilisation | Un investissement peut créer 0 à N immobilisations ; une immobilisation peut provenir de 0 ou 1 investissement | Immobilisation amortissable si durée définie |
| Immobilisation — Amortissement | Une immobilisation génère 0 à N amortissements ; un amortissement concerne 1 immobilisation | Durée > 0 obligatoire |
| Cycle — Compte d'exploitation | Un cycle peut avoir 1 à N comptes ; un compte appartient à 1 cycle | Prévisionnel et réel peuvent coexister |
| Compte — Indicateur | Un compte génère 0 à N indicateurs ; un indicateur appartient à 1 compte | Indicateurs recalculables et traçables |
| Compte — IPE | Un compte peut générer 0 ou 1 IPE courant et N versions historisées | Pondérations historisées nécessaires |
| Compte — IQD | Un compte génère 1 IQD | IQD obligatoire pour validation finale |
| Compte — Alerte | Un compte peut générer 0 à N alertes | Alertes liées aux règles paramétrées |
| Rapport — Compte/Producteur/Organisation | Un rapport peut porter sur 1 ou N objets métier | Périmètre du rapport obligatoire |
| Paramètre métier — Indicateur/Alerte/IPE/IQD | Un paramètre peut influencer 0 à N calculs | Version requise si impact calculatoire |
| Utilisateur — Historique | Un utilisateur peut générer 0 à N événements historiques | Action sensible historisée |
| Rôle — Utilisateur | Un rôle peut être attribué à 0 à N utilisateurs ; un utilisateur a 1 rôle principal | Rôle actif obligatoire |

## 4.2 Contraintes transversales d'intégrité

- Un compte validé doit être rattaché à un producteur, une exploitation, une spéculation, un cycle et une version de règles de calcul.
- Une culture validée doit avoir une superficie positive ou une justification validée.
- Un élevage validé doit avoir un effectif ou une biomasse positive, sauf exception métier documentée.
- Une vente ne peut pas dépasser la quantité disponible issue des récoltes, stocks ou corrections validées.
- Un stock final ne peut pas être négatif.
- Une dette réglée doit avoir un solde nul.
- Une créance encaissée doit avoir un solde nul.
- Une immobilisation amortissable doit avoir une durée d'amortissement strictement positive.
- Un indicateur calculé doit conserver son statut : calculé, provisoire, non applicable ou erreur.
- Les paramètres influençant les calculs doivent être versionnés lorsque leur modification peut changer un résultat historique.

---

# 5. Diagramme conceptuel textuel MCD

```text
ORGANISATION 0,N ── suit ── 0,1 PRODUCTEUR
PRODUCTEUR 1,1 ── possède ── 1,N EXPLOITATION
EXPLOITATION 1,1 ── contient ── 0,N PARCELLE
EXPLOITATION 1,1 ── contient ── 0,N ATELIER_ELEVAGE

SPECULATION 1,1 ── est analysée dans ── 0,N CYCLE_PRODUCTION
EXPLOITATION 1,1 ── accueille ── 0,N CYCLE_PRODUCTION
PARCELLE 0,1 ── supporte ── 0,N CYCLE_PRODUCTION
ATELIER_ELEVAGE 0,1 ── supporte ── 0,N CYCLE_PRODUCTION

CYCLE_PRODUCTION 1,1 ── comprend ── 0,N OPERATION
OPERATION 0,1 ── consomme ── 0,N INTRANT
CYCLE_PRODUCTION 1,1 ── génère ── 0,N CHARGE
CYCLE_PRODUCTION 1,1 ── génère ── 0,N PRODUIT
CYCLE_PRODUCTION 1,1 ── produit ── 0,N RECOLTE
RECOLTE 0,1 ── alimente ── 0,N VENTE
VENTE 0,N ── est faite à ── 0,1 CLIENT
CHARGE 0,N ── est fournie par ── 0,1 FOURNISSEUR

VENTE 0,1 ── génère ── 0,1 CREANCE
CHARGE 0,1 ── génère ── 0,1 DETTE
CYCLE_PRODUCTION 0,1 ── mouvemente ── 0,N FLUX_TRESORERIE
INTRANT/PRODUIT 1,1 ── est suivi dans ── 0,N STOCK

INVESTISSEMENT 0,1 ── crée ── 0,N IMMOBILISATION
IMMOBILISATION 1,1 ── génère ── 0,N AMORTISSEMENT

CYCLE_PRODUCTION 1,1 ── donne lieu à ── 1,N COMPTE_EXPLOITATION
COMPTE_EXPLOITATION 1,1 ── produit ── 0,N INDICATEUR
COMPTE_EXPLOITATION 1,1 ── produit ── 0,N ALERTE
COMPTE_EXPLOITATION 1,1 ── produit ── 1,1 IQD
COMPTE_EXPLOITATION 0,1 ── produit ── 0,1 IPE courant

PARAMETRE_METIER 1,1 ── configure ── 0,N INDICATEUR/ALERTE/IPE/IQD
REFERENTIEL 1,1 ── normalise ── 0,N ENTITES_METIER
UTILISATEUR 1,1 ── possède ── 1,1 ROLE
UTILISATEUR 0,1 ── déclenche ── 0,N HISTORIQUE
RAPPORT 1,1 ── restitue ── 1,N OBJETS_METIER
```

---

# 6. Paramètres métier modifiables sans changer le code

| Domaine | Paramètres modifiables | Impact |
|---|---|---|
| IPE | Pondérations, critères, seuils A-E, bornes de normalisation | Classement et score de performance |
| IQD | Pondérations, pénalités, seuil minimal, niveau de confiance | Fiabilité des résultats et recommandations |
| Alertes | Seuils, messages, niveaux, couleurs, statut bloquant ou non | Diagnostic et tableau de bord |
| Amortissement | Durées par catégorie, valeur résiduelle par défaut, méthode active | Charges économiques et ROI |
| Main-d'œuvre | Barème journalier, type de tâche, valorisation familiale | Résultat économique |
| Autoconsommation | Méthode de valorisation, prix de référence | Produit brut économique |
| Prix | Prix de référence par spéculation, zone, campagne ou marché | Prévision, alertes, valorisations |
| Rendements | Rendements de référence par spéculation, zone ou système | Prévisions et alertes rendement |
| Unités | Unités autorisées, coefficients de conversion, unités locales | Quantités, rendements, coûts unitaires |
| Trésorerie | Périodicité, seuils de tension, traitement avances/crédits | Besoin maximal et capacité de remboursement |
| Simulations | Variables autorisées, scénarios par défaut, bornes de variation | Analyse de sensibilité |
| Spéculations | Liste active, catégories, cycles par défaut, unités de production | Paramétrabilité du produit |
| Rapports | Types de rapports, périmètres, niveaux de détail | Restitution utilisateur |

---

# 7. Données V1 et données prévues pour versions futures

## 7.1 Données spécifiques à la V1

| Entité ou donnée | Justification V1 |
|---|---|
| Producteur | Gestion individuelle et accompagnement |
| Organisation | Mode conseiller/organisation |
| Exploitation | Base de gestion agricole |
| Parcelle | Cultures pilotes |
| Atelier d'élevage | Élevages pilotes |
| Spéculation | Maïs, riz, manioc, tomate, cacao, poulet de chair, porc, pisciculture |
| Cycle de production | Calcul par cycle et annualisation |
| Charge | Compte d'exploitation |
| Produit | Produit brut et résultats |
| Récolte | Rendements, pertes, ventes multiples |
| Vente | Chiffre d'affaires, créances, prix moyen |
| Stock simple | Intrants et produits agricoles |
| Trésorerie simple | Besoin maximal, flux entrants/sortants |
| Dette simple | Achats à crédit et dépenses restant à payer |
| Créance simple | Ventes non encaissées |
| Investissement/Immobilisation/Amortissement | Coût économique et ROI |
| Indicateur/IPE/IQD/Alerte | Décision et reporting |
| Paramètres métier/Référentiels | Paramétrabilité sans code |
| Utilisateur/Rôle | Droits et modes d'usage |

## 7.2 Données préparées pour versions futures

| Entité ou donnée | Usage futur |
|---|---|
| Échéancier détaillé de crédit | Gestion financière avancée |
| Lots de stock | Traçabilité fine et dates de péremption |
| Contrats de vente | Relations commerciales avancées |
| Données météo | Analyse climatique et risque |
| Données de marché historisées | Prix dynamiques et prévisions |
| Données multi-pays/multi-devise | Extension régionale |
| Synchronisation multi-utilisateur | Version Web/Android/SQL |
| Indicateurs VAN/TRI détaillés | Analyse avancée cultures pérennes |
| Diagnostic avancé | Recommandations automatiques enrichies |
| Données environnementales/sociales | Extensions durabilité et impact |

---

# 8. Évolutivité du modèle

Le modèle est conçu pour évoluer sans modification majeure vers plusieurs plateformes :

| Destination future | Exigence conceptuelle |
|---|---|
| Excel/VBA | Entités transformables en tableaux structurés et référentiels |
| SQL | Identifiants uniques, relations normalisées, cardinalités explicites |
| Windows | Séparation données, calculs et navigation |
| Android | Modèle compatible saisie terrain et synchronisation future |
| Web | Modèle multi-utilisateur, organisations, rôles et consolidation |

Principes d'évolutivité :

- conserver des identifiants stables ;
- éviter les champs dépendants d'une interface ;
- historiser les paramètres qui influencent les calculs ;
- séparer les entités de base des résultats calculés ;
- conserver la compatibilité entre mode individuel et mode organisation ;
- prévoir les entités futures sans les rendre obligatoires en V1.

---

# 9. Tableau de traçabilité entités — modules — règles de calcul

| Entité | Modules fonctionnels utilisateurs | Règles de calcul liées |
|---|---|---|
| Organisation | Mode conseiller/organisation, consolidation, rapports | Agrégations, statistiques, comparaisons |
| Producteur | Gestion producteurs, mode individuel, rapports | Résultats par producteur, historique, IQD |
| Exploitation | Gestion exploitation, tableau de bord | Coût par exploitation, consolidation |
| Parcelle | Gestion parcelles, cultures | Superficie, rendement, coût/ha, revenu/ha |
| Atelier d'élevage | Gestion élevages | Effectif, coût/tête, mortalité, production animale |
| Spéculation | Sélection activité, comparaison | IPE, rentabilité, cycles, référentiels |
| Cycle de production | Campagnes/cycles, comptes | Bénéfice par cycle, annualisation, trésorerie |
| Opération | Journal des opérations | Charges par opération, diagnostic technique |
| Intrant | Dépenses, stocks intrants | Charges variables, stocks, coût de production |
| Charge | Compte d'exploitation, dépenses | Charges totales, marges, coût, résultat |
| Produit | Produits, autoconsommation, stocks | Produit brut, résultat économique |
| Récolte | Suivi récoltes | Rendement, pertes, disponibilité physique |
| Vente | Suivi ventes, créances | Chiffre d'affaires, prix moyen, trésorerie |
| Client | Ventes, créances | Encaissements, créances clients |
| Fournisseur | Dépenses, dettes | Charges, dettes fournisseurs, trésorerie |
| Stock | Stocks intrants et produits | Variation de stocks, cohérence physique, IQD |
| Trésorerie | Ma trésorerie, financement | Flux entrants/sortants, besoin maximal, capacité remboursement |
| Dette | Dettes fournisseurs, achats à crédit | Résultat financier, trésorerie, capacité remboursement |
| Créance | Ventes non encaissées | Trésorerie, résultat financier, encaissements |
| Investissement | Projet agricole, immobilisations | ROI, besoin financement, délai récupération |
| Immobilisation | Actifs durables | Amortissement, coût économique |
| Amortissement | Charges économiques | Résultat économique, coût total, ROI |
| Compte d'exploitation | Prévisionnel, réel, comparaison | Tous les indicateurs économiques |
| Indicateur | Résultats, tableau de bord | Produit brut, marges, coûts, ROI, seuils |
| IPE | Comparaison, aide à décision | Normalisation, pondérations, classement A-E |
| IQD | Qualité données, rapports | Pénalités, confiance, recommandations |
| Alerte | Tableau de bord, diagnostic | Seuils économiques, incohérences, risque |
| Rapport | Rapports imprimables | Restitution des résultats et hypothèses |
| Paramètre métier | Administration métier | IPE, IQD, alertes, conversions, amortissements |
| Référentiel | Paramétrage | Catégories, unités, types, statuts |
| Utilisateur | Authentification, droits | Traçabilité, périmètre d'accès |
| Rôle | Gestion des droits | Autorisations, validation, modification paramètres |
| Historique | Audit, corrections | Reproductibilité et traçabilité |

---

# 10. Points de vigilance avant implémentation

- Ne pas confondre le modèle conceptuel avec une structure de feuilles Excel ou de tables SQL.
- Ne pas dupliquer le modèle entre cultures et élevages si une entité générique suffit.
- Ne pas intégrer les données calculées comme des données saisies sans statut clair.
- Ne pas rendre obligatoires en V1 les entités prévues pour versions futures.
- Ne pas perdre la distinction entre résultat financier et résultat économique.
- Ne pas intégrer les crédits reçus comme produits d'exploitation.
- Ne pas intégrer le remboursement du capital comme charge de production.
- Ne pas permettre de résultats validés sans IQD et sans version des règles de calcul.
- Ne pas figer les pondérations IPE/IQD ni les seuils d'alerte dans une logique non paramétrable.
