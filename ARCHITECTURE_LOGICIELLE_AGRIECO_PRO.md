# Architecture logicielle — AGRIECO PRO

## Statut du document

Ce document définit l'**architecture logicielle générale d'AGRIECO PRO**. Il est indépendant de tout langage de programmation et de toute plateforme d'implémentation.

Il ne contient aucun code, aucune implémentation VBA, aucun UserForm, aucune structure Excel physique, aucune table SQL et aucune architecture Android ou Web détaillée. Il décrit les principes, modules, moteurs, flux et services qui devront rester communs aux futures implémentations Excel/VBA, Windows, Android et Web.

---

# 1. Principes d'architecture

## 1.1 Séparation entre présentation, logique métier et données

AGRIECO PRO doit être conçu selon une séparation claire entre :

| Couche | Responsabilité | Exemples conceptuels |
|---|---|---|
| Présentation | Interaction avec l'utilisateur, affichage, navigation, saisie et restitution | Menus, tableaux de bord, formulaires, rapports |
| Logique applicative | Orchestration des cas d'utilisation et coordination des modules | Créer un compte, valider un cycle, générer un rapport |
| Logique métier | Règles agricoles, économiques, financières et de décision | Produit brut, charges, ROI, IPE, IQD, alertes |
| Données | Conservation des entités métier et résultats | Producteurs, exploitations, cycles, charges, produits, indicateurs |
| Services transversaux | Fonctions communes à tous les modules | Validation, référentiels, historique, sécurité, exports, sauvegarde |

Aucune couche de présentation ne doit contenir une formule économique officielle. Les calculs doivent appartenir aux moteurs métier.

## 1.2 Moteur économique unique

Le logiciel doit utiliser un **moteur économique unique** pour :

- le mode individuel ;
- le mode conseiller ou organisation ;
- les cultures ;
- les élevages ;
- les comptes prévisionnels ;
- les comptes réels ;
- les simulations ;
- les rapports ;
- les futures plateformes.

La différence entre les contextes d'utilisation doit porter sur le parcours utilisateur, les droits, les vues, les agrégations et les rapports, mais jamais sur une duplication des formules.

## 1.3 Modularité

L'application doit être découpée en modules indépendants ayant chacun une responsabilité claire. Un module ne doit pas connaître les détails internes d'un autre module ; il doit interagir avec lui par des données validées et des résultats formalisés.

## 1.4 Extensibilité

L'architecture doit permettre d'ajouter :

- de nouvelles cultures ;
- de nouveaux élevages ;
- de nouveaux indicateurs ;
- de nouvelles règles d'alerte ;
- de nouveaux rapports ;
- de nouvelles plateformes ;
- de nouveaux profils utilisateurs ;
- de nouveaux référentiels métier.

Ces ajouts doivent se faire principalement par paramétrage ou extension contrôlée, sans réécriture du moteur économique.

## 1.5 Maintenabilité

Les règles métier doivent être centralisées, documentées et versionnées. Toute modification d'une règle de calcul doit pouvoir être tracée, testée et expliquée.

## 1.6 Réutilisabilité

Les moteurs métier doivent être réutilisables par plusieurs interfaces et plateformes. Le moteur de calcul d'un compte d'exploitation doit produire le même résultat qu'il soit appelé par une interface Excel, une application Windows, une application Android ou une application Web.

---

# 2. Découpage en modules

## 2.1 Vue d'ensemble

| Module | Responsabilité principale | Entrées principales | Sorties principales |
|---|---|---|---|
| Gestion des utilisateurs et droits | Gérer profils, accès et périmètres | Utilisateur, rôle, organisation | Session, droits, périmètre |
| Gestion des organisations | Gérer coopératives, ONG, cabinets, institutions | Organisation, zones | Portefeuille et consolidation |
| Gestion des producteurs | Créer et suivre producteurs/entrepreneurs | Identité, localisation, organisation | Fiche producteur, historique |
| Gestion des exploitations | Décrire exploitations, parcelles et ateliers | Producteur, exploitation, surfaces, capacités | Structure productive |
| Gestion des spéculations | Administrer cultures et élevages | Référentiels spéculations | Activités disponibles |
| Gestion des campagnes et cycles | Planifier et suivre périodes de production | Dates, spéculation, parcelle/atelier | Cycle prévisionnel ou réel |
| Gestion des opérations | Suivre opérations agricoles ou zootechniques | Cycle, opération, date | Journal des opérations |
| Gestion des charges | Saisir, classer et affecter les charges | Charges, intrants, paiements, affectations | Charges validées |
| Gestion des produits | Saisir produits, récoltes, ventes, autoconsommation | Récoltes, ventes, stocks | Produits validés |
| Gestion des stocks | Suivre intrants et produits agricoles | Entrées, sorties, pertes | Stock disponible |
| Gestion de trésorerie | Suivre encaissements, décaissements, dettes, créances | Flux, échéances, paiements | Besoin de trésorerie |
| Comptes d'exploitation | Consolider charges, produits et résultats | Cycle, charges, produits, paramètres | Compte prévisionnel/réel |
| Indicateurs et diagnostic | Calculer indicateurs, IPE, IQD, alertes | Compte, paramètres | Résultats et recommandations |
| Simulation | Tester variations de prix, rendements, charges, superficies | Compte de référence, scénarios | Résultats simulés |
| Rapports | Produire restitutions personnelles ou institutionnelles | Résultats, périmètre, modèle | Rapport imprimable/exportable |
| Administration métier | Gérer paramètres et référentiels | Seuils, unités, pondérations, barèmes | Paramètres actifs |
| Import/export | Échanger des données | Données internes/externes | Fichiers ou jeux de données |
| Sauvegarde/restauration | Préserver et restaurer les données | État applicatif | Version sauvegardée/restaurée |
| Historique et audit | Tracer actions et changements | Actions utilisateurs | Journal historique |

## 2.2 Modules du mode individuel

Le mode individuel doit organiser les modules autour du vocabulaire personnel de l'entrepreneur agricole :

| Rubrique utilisateur | Modules mobilisés |
|---|---|
| Mon exploitation | Producteurs, exploitations, parcelles, ateliers |
| Mes parcelles | Parcelles, spéculations, cycles |
| Mes activités | Spéculations, cycles, comptes |
| Mon projet agricole | Prévision, investissement, simulation, trésorerie |
| Mes dépenses | Charges, dettes, trésorerie |
| Mes récoltes | Récoltes, stocks, pertes |
| Mes ventes | Ventes, créances, trésorerie |
| Mes stocks | Stocks d'intrants et produits |
| Ma trésorerie | Flux, dettes, créances, besoin de financement |
| Mes résultats | Comptes, indicateurs, IPE, IQD, alertes |
| Mes simulations | Moteur de simulation |
| Mes objectifs | Prévision, comparaison, diagnostic |
| Mes rapports | Rapports personnels |

## 2.3 Modules du mode conseiller ou organisation

Le mode conseiller ou organisation doit organiser les modules autour du suivi multi-producteurs :

| Rubrique organisationnelle | Modules mobilisés |
|---|---|
| Producteurs suivis | Organisations, producteurs, exploitations |
| Exploitations suivies | Exploitations, parcelles, ateliers |
| Campagnes et cycles | Cycles, spéculations, comptes |
| Comptes d'exploitation | Charges, produits, indicateurs |
| Consolidation | Comptes, indicateurs, statistiques |
| Comparaisons | IPE, rentabilité, risque, trésorerie |
| Statistiques | Agrégations par zone, spéculation, campagne |
| Rapports institutionnels | Rapports, exports, historiques |
| Administration métier | Paramètres, référentiels, droits |

---

# 3. Moteurs métier indépendants

## 3.1 Moteur économique

### Rôle

Calculer le compte d'exploitation : produits, charges, marges, résultats financiers et économiques, coûts de production et bénéfices.

### Données d'entrée

- Cycle de production.
- Charges validées.
- Produits validés.
- Stocks et pertes.
- Paramètres économiques.
- Règles de valorisation.

### Calculs

- Produit brut financier et économique.
- Charges variables, fixes, valorisées, payées.
- Amortissements.
- Marges brute et nette.
- Résultat financier.
- Résultat économique.
- Coûts unitaires et par base physique.

### Résultats

- Compte d'exploitation calculé.
- Données sources des indicateurs.
- Statuts de calcul.
- Messages d'erreur ou d'incomplétude.

### Interactions

- Alimente les moteurs de rentabilité, trésorerie, indicateurs, alertes, simulation et rapports.

## 3.2 Moteur de trésorerie

### Rôle

Calculer les flux monétaires, le calendrier de trésorerie, le besoin maximal de trésorerie, les dettes, créances et capacité de remboursement.

### Données d'entrée

- Ventes encaissées ou à encaisser.
- Charges payées ou à payer.
- Avances.
- Crédits.
- Remboursements.
- Frais financiers.
- Dates ou périodes.

### Calculs

- Flux entrants.
- Flux sortants.
- Solde par période.
- Trésorerie minimale.
- Besoin maximal de trésorerie.
- Capacité de remboursement.

### Résultats

- Calendrier de trésorerie.
- Alertes de tension.
- Besoin de financement.

### Interactions

- Reçoit des données du module ventes, charges, dettes et créances.
- Alimente IPE, alertes, simulation et rapports.

## 3.3 Moteur de rentabilité

### Rôle

Calculer les indicateurs de rentabilité et d'investissement.

### Données d'entrée

- Résultats économiques.
- Résultats financiers.
- Investissements.
- Amortissements.
- Durée de cycle ou horizon pluriannuel.

### Calculs

- ROI.
- Ratio bénéfice/coût.
- Rentabilité annuelle.
- Rentabilité par hectare, tête, kilogramme ou FCFA investi.
- Seuil de rentabilité.
- Délai de récupération.
- VAN et TRI préparés pour versions futures.

### Résultats

- Indicateurs de rentabilité.
- Classements par performance.
- Données de diagnostic.

### Interactions

- Alimente moteur d'indicateurs, comparaison, simulation, alertes et rapports.

## 3.4 Moteur de simulation

### Rôle

Évaluer l'impact de variations d'hypothèses sur les résultats.

### Données d'entrée

- Compte de référence.
- Variables de simulation : prix, rendement, superficie, effectif, charges, pertes, cycles.
- Scénarios.

### Calculs

- Recalcul du compte sous hypothèses simulées.
- Comparaison scénario/base.
- Sensibilité des indicateurs.

### Résultats

- Résultats par scénario.
- Écarts de rentabilité.
- Recommandations de prudence.

### Interactions

- Appelle le moteur économique et le moteur de rentabilité sans modifier les données validées.

## 3.5 Moteur d'indicateurs

### Rôle

Calculer et consolider les indicateurs de performance, y compris IPE et IQD.

### Données d'entrée

- Résultats du moteur économique.
- Résultats de trésorerie.
- Résultats de rentabilité.
- Données de qualité.
- Pondérations et paramètres.

### Calculs

- IPE.
- IQD.
- Productivité.
- Indicateurs de comparaison.
- Scores normalisés.

### Résultats

- Scores.
- Classes.
- Statuts de confiance.
- Données prêtes pour tableau de bord.

### Interactions

- Alimente alertes, diagnostic, rapports et tableaux de bord.

## 3.6 Moteur d'alertes

### Rôle

Détecter automatiquement les anomalies, incohérences et risques.

### Données d'entrée

- Données saisies.
- Résultats calculés.
- Seuils d'alerte.
- Paramètres de qualité.

### Calculs

- Comparaison aux seuils.
- Détection d'incohérences physiques.
- Détection de résultat négatif, coût anormal, rendement incohérent, prix incohérent, marge faible, risque élevé, trésorerie insuffisante.

### Résultats

- Alertes informatives.
- Alertes de vigilance.
- Alertes critiques.
- Messages de diagnostic.

### Interactions

- Utilise les résultats de tous les moteurs.
- Alimente tableau de bord, rapports et assistant de diagnostic.

---

# 4. Flux de données

## 4.1 Flux général depuis la saisie jusqu'au rapport

```text
Saisie ou import
   ↓
Validation des données saisies
   ↓
Normalisation des unités et référentiels
   ↓
Enregistrement conceptuel des entités métier
   ↓
Calcul du compte d'exploitation
   ↓
Calcul de trésorerie
   ↓
Calcul de rentabilité
   ↓
Calcul IPE, IQD et autres indicateurs
   ↓
Déclenchement des alertes
   ↓
Diagnostic et recommandations
   ↓
Rapports, tableaux de bord, exports
   ↓
Historisation et sauvegarde
```

## 4.2 Flux des données saisies

Les données saisies entrent par les modules de gestion : producteurs, exploitations, parcelles, ateliers, cycles, opérations, charges, produits, récoltes, ventes, stocks, dettes, créances et trésorerie.

Avant d'être utilisées par les moteurs, elles doivent passer par :

1. validation de présence des champs obligatoires ;
2. validation de cohérence métier ;
3. conversion des unités ;
4. rattachement aux référentiels ;
5. classification comme donnée saisie, estimée, calculée ou paramétrable.

## 4.3 Flux des données calculées

Les moteurs produisent des résultats qui ne doivent pas être saisis manuellement : montants calculés, marges, résultats, coûts, IPE, IQD, alertes, soldes de trésorerie et classements.

Ces données calculées doivent conserver :

- leur source ;
- leur règle de calcul ;
- leur version de paramètres ;
- leur statut de calcul ;
- leur date de calcul.

## 4.4 Flux des paramètres

Les paramètres alimentent les moteurs mais ne doivent pas être mélangés avec les données d'exploitation. Toute modification d'un paramètre ayant un impact sur un résultat doit être historisée.

## 4.5 Flux des rapports

Les rapports consomment des résultats déjà calculés. Ils ne doivent pas recalculer les indicateurs selon des formules propres. Ils doivent restituer les hypothèses, paramètres, alertes et niveaux de qualité.

---

# 5. Services transversaux

## 5.1 Service des paramètres

Responsabilités :

- fournir les pondérations IPE et IQD ;
- fournir les seuils d'alerte ;
- fournir les barèmes de main-d'œuvre ;
- fournir les durées d'amortissement ;
- fournir les prix et rendements de référence ;
- fournir les règles de simulation ;
- historiser les versions de paramètres.

## 5.2 Service des référentiels

Responsabilités :

- gérer les unités ;
- gérer les conversions ;
- gérer les catégories de charges ;
- gérer les types de produits ;
- gérer les statuts ;
- gérer les types de spéculations ;
- gérer les zones, marchés et risques.

## 5.3 Service de validation des données

Responsabilités :

- contrôler les champs obligatoires ;
- distinguer zéro, vide et non applicable ;
- contrôler les quantités négatives ;
- contrôler les divisions par zéro ;
- contrôler la cohérence stocks/récoltes/ventes/pertes ;
- contrôler les statuts avant validation d'un compte.

## 5.4 Service de journalisation et historique

Responsabilités :

- tracer les créations, modifications, validations, corrections et archivages ;
- conserver l'auteur et la date des actions ;
- historiser les paramètres critiques ;
- permettre la reproductibilité des résultats.

## 5.5 Service de sécurité et droits

Responsabilités :

- identifier l'utilisateur ;
- appliquer le rôle ;
- limiter l'accès au périmètre autorisé ;
- distinguer mode individuel, conseiller, organisation et administration ;
- protéger la modification des paramètres sensibles.

## 5.6 Service d'import/export

Responsabilités :

- importer des données de référence ;
- exporter des rapports ;
- exporter des données normalisées ;
- préparer les échanges futurs entre plateformes ;
- préserver les identifiants uniques.

## 5.7 Service de sauvegarde et restauration

Responsabilités :

- créer des sauvegardes datées ;
- restaurer un état antérieur ;
- signaler les versions sauvegardées ;
- éviter la perte de données locales.

---

# 6. Gestion des paramètres configurables

Les éléments suivants doivent être configurables sans modification du code :

| Domaine | Éléments configurables |
|---|---|
| IPE | Critères, pondérations, seuils de classes, bornes de normalisation |
| IQD | Pondérations, pénalités, seuils de confiance, pénalités de saisie simplifiée |
| Alertes | Seuils, messages, niveaux de gravité, activation/désactivation |
| Charges | Catégories, nature fixe/variable, charges types, règles de récurrence |
| Produits | Types de produits, produits secondaires, méthodes de valorisation |
| Main-d'œuvre | Barèmes, types de main-d'œuvre, règles de valorisation familiale |
| Amortissements | Durées par catégorie, valeur résiduelle par défaut, méthode active |
| Unités | Unités autorisées, coefficients de conversion, unités locales |
| Prix | Prix de référence par spéculation, zone, campagne ou marché |
| Rendements | Rendements de référence par spéculation et contexte |
| Trésorerie | Périodicité, seuils de tension, règles dettes/créances |
| Simulations | Variables simulables, scénarios par défaut, bornes de variation |
| Spéculations | Cultures et élevages actifs, cycles par défaut, unités de production |
| Rapports | Types, périmètres, niveaux de détail, modèles de restitution |
| Droits | Rôles, accès, permissions de validation et modification |

---

# 7. Architecture cible multi-plateforme

## 7.1 Principe commun

Toutes les plateformes doivent partager les mêmes concepts : entités métier, règles de calcul, paramètres, statuts et moteurs. La plateforme ne doit pas modifier le résultat économique.

## 7.2 Excel/VBA

Dans une implémentation Excel/VBA, l'architecture devra respecter :

- séparation entre feuilles de données, écrans de saisie, calculs et rapports ;
- centralisation des règles dans des modules de calcul ;
- données stockées selon le modèle conceptuel ;
- paramètres modifiables sans réécriture des règles ;
- exports et sauvegardes locaux.

## 7.3 Windows

Dans une application Windows, l'architecture devra reprendre :

- les mêmes moteurs métier ;
- une base locale ou synchronisable ;
- des vues adaptées au mode individuel et organisation ;
- des rapports imprimables ;
- une gestion renforcée de la sécurité et de la sauvegarde.

## 7.4 Android

Dans une application Android, l'architecture devra permettre :

- la saisie terrain hors connexion ;
- la synchronisation future ;
- la saisie simplifiée ;
- l'accès aux référentiels ;
- la consultation de résultats et alertes essentiels ;
- la réutilisation des mêmes règles métier.

## 7.5 Web

Dans une application Web, l'architecture devra permettre :

- le multi-utilisateur ;
- le multi-organisation ;
- la consolidation ;
- les tableaux de bord avancés ;
- les exports institutionnels ;
- la gestion centralisée des paramètres ;
- l'utilisation du même moteur économique.

---

# 8. Contraintes de conception

## 8.1 Préserver le moteur unique

- Ne jamais créer un moteur différent pour mode individuel et mode organisation.
- Ne jamais créer un moteur différent par culture ou élevage si un paramétrage suffit.
- Ne pas recalculer les indicateurs dans les rapports selon des formules séparées.
- Ne pas dupliquer les règles entre prévisionnel, réel et simulation.

## 8.2 Éviter les duplications

- Mutualiser les entités communes aux cultures et élevages.
- Mutualiser les charges, produits, stocks, ventes et flux de trésorerie.
- Utiliser les référentiels pour les spécificités.
- Utiliser les paramètres pour les seuils, pondérations et barèmes.

## 8.3 Garantir la cohérence entre plateformes

- Chaque plateforme doit utiliser les mêmes règles officielles.
- Les paramètres doivent être versionnés.
- Les résultats calculés doivent être reproductibles.
- Les exports doivent conserver les identifiants et statuts.
- Les différences de présentation ne doivent pas modifier les calculs.

## 8.4 Maîtriser le périmètre de la V1

- Les fonctions préparées pour versions futures ne doivent pas bloquer la V1.
- Les simulations doivent rester limitées à un périmètre validé.
- La gestion des dettes, créances et stocks doit rester simple en V1 si cela est arbitré.
- Les cultures pérennes doivent être supportées sans transformer la V1 en outil complet d'analyse financière avancée.

## 8.5 Maintenir la traçabilité

- Toute validation de compte doit conserver la version des règles et paramètres.
- Toute correction doit être historisée.
- Toute alerte doit pouvoir être reliée à sa règle de déclenchement.
- Tout rapport doit indiquer son périmètre, sa date et son niveau de qualité des données.

---

# 9. Synthèse de l'architecture logicielle cible

```text
Présentation
  ├─ Mode individuel
  └─ Mode conseiller / organisation
        ↓
Logique applicative
  ├─ Producteurs / exploitations / parcelles / ateliers
  ├─ Cycles / opérations
  ├─ Charges / produits / stocks / ventes
  ├─ Trésorerie / dettes / créances
  ├─ Simulations
  └─ Rapports
        ↓
Services transversaux
  ├─ Validation
  ├─ Paramètres
  ├─ Référentiels
  ├─ Sécurité
  ├─ Historique
  ├─ Import / export
  └─ Sauvegarde
        ↓
Moteurs métier
  ├─ Moteur économique
  ├─ Moteur de trésorerie
  ├─ Moteur de rentabilité
  ├─ Moteur de simulation
  ├─ Moteur d'indicateurs
  └─ Moteur d'alertes
        ↓
Données métier
  ├─ Entités saisies
  ├─ Entités calculées
  ├─ Paramètres
  ├─ Référentiels
  └─ Historique
```

Cette architecture doit servir de référence commune pour toute future implémentation d'AGRIECO PRO.
