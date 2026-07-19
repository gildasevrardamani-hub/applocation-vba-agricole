# Architecture fonctionnelle et technique — AGRIECO PRO

## 1. Vision générale

AGRIECO PRO est une application professionnelle d'aide à la décision technico-économique pour l'analyse des exploitations agricoles et d'élevage. La première version est prévue sous Microsoft Excel avec VBA, interface graphique professionnelle et base de données interne Excel, tout en conservant une architecture évolutive vers Android, Web et SQL.

### 1.1 Objectifs métier

L'application doit permettre de :

- enregistrer les caractéristiques d'une exploitation agricole ou d'élevage ;
- élaborer automatiquement les comptes d'exploitation ;
- calculer les coûts de production, revenus et marges ;
- mesurer la rentabilité économique par spéculation, producteur, exploitation et campagne ;
- comparer plusieurs spéculations agricoles et d'élevage ;
- classer les activités selon leur performance économique ;
- fournir une aide structurée à la prise de décision pour les producteurs, conseillers, coopératives, ONG, projets de développement, institutions financières et centres de formation.

### 1.2 Principes d'architecture

- **Séparation des couches** : interface utilisateur, logique métier, calculs économiques, stockage, sécurité et reporting.
- **Traçabilité** : chaque compte d'exploitation est lié à un producteur, une exploitation, une campagne, une spéculation et une version de calcul.
- **Paramétrage** : les cultures, élevages, unités, types de charges, coefficients, risques et licences doivent être administrables sans modifier le code.
- **Évolutivité** : les feuilles Excel jouent le rôle de tables normalisées afin de préparer une future migration SQL.
- **Contrôle qualité** : validations de données, journal des erreurs, journal des actions sensibles et verrouillage des formules critiques.

## 2. Architecture applicative cible

### 2.1 Couches logicielles

| Couche | Rôle | Implémentation Excel/VBA |
|---|---|---|
| Présentation | Écrans, formulaires, navigation, tableaux de bord | UserForms VBA, ruban personnalisé, feuilles protégées |
| Services applicatifs | Orchestration des cas d'utilisation | Modules `modAppController`, `modWorkflow`, `modDashboard` |
| Domaine métier | Producteurs, exploitations, cultures, élevages, comptes, indicateurs | Modules métier et classes VBA |
| Moteur économique | Calcul des charges, produits, marges, ROI, IPE | `modCalculCompteExploitation`, `modCalculIPE` |
| Données | Tables internes normalisées | Feuilles préfixées `T_` avec tableaux structurés Excel |
| Sécurité | Authentification, rôles, licence, audit | `modAuth`, `modLicence`, `modAudit` |
| Reporting | États imprimables, exports PDF/Excel | `modReport`, modèles de feuilles `R_` |
| Intégration future | Import/export, migration SQL/API | `modImportExport`, schéma d'identifiants stables |

### 2.2 Organisation logique des données

Les données doivent être structurées autour de six noyaux :

1. **Référentiels** : cultures, élevages, unités, types de charges, risques, marchés, campagnes.
2. **Acteurs** : utilisateurs, clients, producteurs, coopératives, institutions.
3. **Exploitations** : localisation, superficies, systèmes de production, historique.
4. **Activités productives** : spéculations culturales et animales.
5. **Comptes d'exploitation** : charges, produits, cycles et indicateurs calculés.
6. **Décision/reporting** : classements, IPE, tableaux de bord, recommandations.

## 3. Modules fonctionnels

## Module 1 — Authentification, droits et licence

### Fonctions

- Connexion utilisateur avec identifiant et mot de passe.
- Gestion des profils : administrateur, conseiller, superviseur, lecteur, formateur.
- Association de l'utilisateur à un client, une organisation ou une coopérative.
- Gestion d'une licence locale : clé, date d'activation, date d'expiration, type d'offre, nombre maximal d'utilisateurs ou d'exploitations.
- Journalisation des connexions, échecs, modifications sensibles et exports.

### Règles de conception

- Les mots de passe ne doivent pas être stockés en clair ; prévoir un hachage avec sel dans la limite des possibilités VBA.
- Les feuilles sensibles doivent être masquées en `xlVeryHidden` et protégées.
- Les actions critiques doivent être conditionnées par les rôles.
- La licence doit être conçue comme un service indépendant pour faciliter une future activation en ligne.

## Module 2 — Tableau de bord

### Indicateurs affichés

- Nombre d'exploitations analysées.
- Nombre de cultures enregistrées.
- Nombre d'élevages enregistrés.
- Bénéfice moyen.
- Rentabilité moyenne.
- Classement des spéculations par bénéfice, marge nette, ROI, ratio bénéfice/coût et IPE.

### Vues recommandées

- Vue synthèse globale.
- Vue par campagne agricole.
- Vue par zone géographique.
- Vue par producteur ou coopérative.
- Vue cultures uniquement.
- Vue élevages uniquement.
- Top 10 des spéculations les plus performantes.
- Alertes : activité déficitaire, coût de production élevé, seuil de rentabilité non atteint, risque élevé.

## Module 3 — Gestion des producteurs

### Données principales

- Nom et prénom ou raison sociale.
- Contact téléphonique et adresse.
- Localisation : pays, région, département, commune, village, coordonnées GPS.
- Coopérative ou organisation d'appartenance.
- Superficie totale, superficie exploitée, superficie par spéculation.
- Historique des campagnes et comptes d'exploitation.

### Fonctions

- Création, modification, désactivation et recherche de producteurs.
- Consultation de l'historique économique.
- Association de plusieurs exploitations à un même producteur.
- Fiche producteur imprimable.

## Module 4 — Gestion des cultures

### Catégories prises en charge

- Cultures annuelles : maïs, riz, manioc, arachide, coton, soja.
- Cultures pérennes : cacao, anacarde, café, palmier à huile, hévéa.
- Maraîchage : tomate, piment, chou, oignon, laitue, gombo et autres légumes.

### Fonctions

- Référentiel des cultures et variétés.
- Paramètres agronomiques : durée du cycle, unité de rendement, densité, période de production.
- Paramètres économiques par défaut : charges types, rendement de référence, prix indicatif, niveau de risque, besoin en main-d'œuvre.
- Gestion des produits principaux et secondaires.

## Module 5 — Gestion des élevages

### Filières prises en charge

- Volaille.
- Porc.
- Bovin.
- Ovin.
- Caprin.
- Pisciculture.

### Fonctions

- Référentiel des espèces, races et ateliers.
- Gestion des cycles d'élevage : durée, effectif initial, mortalité, indice de consommation, poids final.
- Charges spécifiques : aliments, médicaments, vaccins, litière, eau, énergie, main-d'œuvre, transport.
- Produits spécifiques : animaux vendus, œufs, lait, fumier, alevins, produits secondaires.

## Module principal — Compte d'exploitation

### 6.1 Charges fixes

Le module doit calculer et consolider les charges fixes suivantes :

- bâtiments ;
- équipements ;
- matériels ;
- amortissements ;
- assurances ;
- frais administratifs ;
- loyers ou fermages, si applicables.

### 6.2 Charges variables

Le module doit calculer les charges variables suivantes :

- semences ;
- plants ;
- engrais ;
- pesticides ;
- aliments ;
- médicaments ;
- main-d'œuvre ;
- transport ;
- énergie ;
- eau ;
- emballage ;
- frais de commercialisation ;
- pertes et imprévus paramétrables.

### 6.3 Produits

Le module doit enregistrer et calculer :

- quantité produite ;
- unité de production ;
- prix unitaire ;
- chiffre d'affaires ;
- produits secondaires ;
- autoconsommation valorisée ;
- stocks de fin de cycle, si nécessaire.

### 6.4 Indicateurs économiques obligatoires

| Indicateur | Formule de référence |
|---|---|
| Produit brut | Somme des ventes, produits secondaires, autoconsommation valorisée et variation de stock positive |
| Charges totales | Charges fixes + charges variables |
| Marge brute | Produit brut - charges variables |
| Marge nette | Produit brut - charges totales |
| Bénéfice | Produit brut - charges totales, après ajustements éventuels |
| Coût de production | Charges totales / quantité produite vendable |
| Coût par hectare | Charges totales / superficie exploitée |
| Revenu par hectare | Produit brut / superficie exploitée |
| ROI | Bénéfice / investissement total x 100 |
| Ratio bénéfice/coût | Produit brut / charges totales |
| Seuil de rentabilité en valeur | Charges fixes / taux de marge sur coûts variables |
| Seuil de rentabilité en quantité | Seuil de rentabilité en valeur / prix unitaire moyen |
| Productivité de la main-d'œuvre | Produit brut ou quantité produite / nombre de jours-homme |

## 7. Module d'intelligence économique

### 7.1 Gestion des cycles

Le module doit permettre de :

- saisir le nombre de cycles par an ;
- calculer le bénéfice par cycle ;
- annualiser les produits et charges variables ;
- répartir les charges fixes annuelles sur les cycles ;
- comparer une culture longue avec une activité à cycles courts ;
- distinguer le bénéfice par cycle, le bénéfice annuel et le besoin de trésorerie maximal.

### 7.2 Indice de Performance Économique — IPE

L'IPE est un score sur 100 permettant de comparer les spéculations. Il doit combiner des facteurs financiers, opérationnels et commerciaux.

| Critère | Pondération proposée | Sens d'interprétation |
|---|---:|---|
| Rentabilité annuelle | 30 % | Plus elle est élevée, meilleur est le score |
| Durée du cycle | 10 % | Plus le cycle est court, meilleur est le score |
| Investissement nécessaire | 15 % | Plus il est faible, meilleur est le score |
| Niveau de risque | 15 % | Plus le risque est faible, meilleur est le score |
| Besoin de trésorerie | 10 % | Plus le besoin est faible, meilleur est le score |
| Disponibilité du marché | 10 % | Plus le marché est disponible, meilleur est le score |
| Besoin en main-d'œuvre | 10 % | Plus le besoin est compatible avec l'exploitation, meilleur est le score |
| **Total** | **100 %** | Score final sur 100 |

### 7.3 Méthode de calcul recommandée

1. Normaliser chaque critère sur une échelle de 0 à 100.
2. Appliquer les pondérations configurées dans le référentiel `T_Param_IPE`.
3. Calculer le score final :

```text
IPE = Somme(ScoreCritere x PondérationCritere)
```

4. Classer les activités :

| Score IPE | Classe | Interprétation |
|---:|---|---|
| 80 à 100 | A | Très performante |
| 65 à 79 | B | Performante |
| 50 à 64 | C | Moyenne, à surveiller |
| 35 à 49 | D | Faible performance |
| 0 à 34 | E | Non recommandée ou très risquée |

## 8. Structure des feuilles Excel

### 8.1 Conventions

- `T_` : tables de données internes.
- `R_` : rapports imprimables.
- `D_` : tableaux de bord.
- `P_` : paramètres.
- `TMP_` : feuilles temporaires.
- `LOG_` : journaux techniques et fonctionnels.

### 8.2 Feuilles principales

| Feuille | Rôle |
|---|---|
| `D_Accueil` | Écran d'accueil et navigation principale |
| `D_TableauBord` | Indicateurs, graphiques et classements |
| `T_Utilisateurs` | Comptes utilisateurs et rôles |
| `T_Roles` | Droits d'accès par profil |
| `T_Licence` | Informations de licence |
| `T_Clients` | Organisations clientes, coopératives ou institutions |
| `T_Producteurs` | Fiches producteurs |
| `T_Exploitations` | Informations physiques et géographiques des exploitations |
| `T_Campagnes` | Campagnes agricoles et périodes d'analyse |
| `T_Cultures` | Référentiel cultures |
| `T_Elevages` | Référentiel élevages |
| `T_Speculations` | Liste unifiée des activités agricoles et animales |
| `T_Comptes` | En-têtes des comptes d'exploitation |
| `T_Charges` | Lignes de charges fixes et variables |
| `T_Produits` | Lignes de produits principaux et secondaires |
| `T_Indicateurs` | Résultats calculés des comptes d'exploitation |
| `T_Cycles` | Nombre de cycles, durées et annualisation |
| `T_Param_IPE` | Pondérations et seuils de l'IPE |
| `T_Resultats_IPE` | Scores IPE et classements |
| `T_Unites` | Unités physiques et monétaires |
| `T_Prix_Reference` | Prix de référence par zone, campagne et spéculation |
| `T_Risques` | Grille des risques techniques, climatiques, sanitaires et commerciaux |
| `R_FicheProducteur` | Rapport imprimable producteur |
| `R_CompteExploitation` | Rapport imprimable du compte d'exploitation |
| `R_Comparaison` | Rapport comparatif des spéculations |
| `LOG_Audit` | Journal des actions utilisateur |
| `LOG_Erreurs` | Journal des erreurs VBA |

## 9. Dictionnaire des données

### 9.1 Tables d'identité et sécurité

#### `T_Utilisateurs`

| Champ | Type | Description |
|---|---|---|
| `UtilisateurID` | Texte UUID | Identifiant unique |
| `ClientID` | Texte UUID | Organisation rattachée |
| `NomUtilisateur` | Texte | Identifiant de connexion |
| `NomComplet` | Texte | Nom affiché |
| `RoleID` | Texte UUID | Profil de droits |
| `MotDePasseHash` | Texte | Empreinte du mot de passe |
| `SelMotDePasse` | Texte | Sel cryptographique |
| `Actif` | Booléen | Compte activé ou non |
| `DerniereConnexion` | DateHeure | Dernière connexion réussie |

#### `T_Licence`

| Champ | Type | Description |
|---|---|---|
| `LicenceID` | Texte UUID | Identifiant de licence |
| `ClientID` | Texte UUID | Client titulaire |
| `CleLicence` | Texte | Clé de licence |
| `TypeLicence` | Texte | Démo, standard, pro, institution |
| `DateActivation` | Date | Date d'activation |
| `DateExpiration` | Date | Date d'expiration |
| `MaxUtilisateurs` | Numérique | Limite utilisateurs |
| `MaxExploitations` | Numérique | Limite exploitations |
| `Statut` | Texte | Active, expirée, suspendue |

### 9.2 Tables producteurs et exploitations

#### `T_Producteurs`

| Champ | Type | Description |
|---|---|---|
| `ProducteurID` | Texte UUID | Identifiant unique |
| `ClientID` | Texte UUID | Organisation gestionnaire |
| `CodeProducteur` | Texte | Code métier lisible |
| `Nom` | Texte | Nom ou raison sociale |
| `Contact` | Texte | Téléphone principal |
| `LocalisationID` | Texte UUID | Lien vers localisation |
| `Cooperative` | Texte | Coopérative ou groupement |
| `SuperficieTotaleHa` | Décimal | Superficie totale |
| `DateCreation` | DateHeure | Date d'enregistrement |
| `Actif` | Booléen | Producteur actif |

#### `T_Exploitations`

| Champ | Type | Description |
|---|---|---|
| `ExploitationID` | Texte UUID | Identifiant unique |
| `ProducteurID` | Texte UUID | Producteur propriétaire ou gestionnaire |
| `NomExploitation` | Texte | Nom de l'exploitation |
| `Pays` | Texte | Pays |
| `Region` | Texte | Région |
| `Commune` | Texte | Commune |
| `Village` | Texte | Village |
| `Latitude` | Décimal | Coordonnée GPS |
| `Longitude` | Décimal | Coordonnée GPS |
| `SuperficieHa` | Décimal | Superficie exploitée |

### 9.3 Tables spéculations

#### `T_Speculations`

| Champ | Type | Description |
|---|---|---|
| `SpeculationID` | Texte UUID | Identifiant unique |
| `TypeSpeculation` | Texte | Culture ou élevage |
| `Categorie` | Texte | Annuelle, pérenne, maraîchage, volaille, porc, etc. |
| `NomSpeculation` | Texte | Nom de l'activité |
| `UniteProduction` | Texte | kg, tonne, tête, litre, plateau, etc. |
| `DureeCycleJoursDefaut` | Entier | Durée standard du cycle |
| `CyclesParAnDefaut` | Décimal | Nombre de cycles standard |
| `RisqueDefaut` | Entier | Score de risque 1 à 5 |
| `MarcheDefaut` | Entier | Disponibilité du marché 1 à 5 |

### 9.4 Tables comptes d'exploitation

#### `T_Comptes`

| Champ | Type | Description |
|---|---|---|
| `CompteID` | Texte UUID | Identifiant unique |
| `ProducteurID` | Texte UUID | Producteur analysé |
| `ExploitationID` | Texte UUID | Exploitation concernée |
| `CampagneID` | Texte UUID | Campagne agricole |
| `SpeculationID` | Texte UUID | Culture ou élevage analysé |
| `DateAnalyse` | Date | Date du compte |
| `SuperficieHa` | Décimal | Surface utilisée pour les cultures |
| `EffectifInitial` | Décimal | Effectif de départ pour élevages |
| `DureeCycleJours` | Entier | Durée du cycle réel |
| `CyclesParAn` | Décimal | Nombre de cycles annuels |
| `StatutCompte` | Texte | Brouillon, validé, archivé |

#### `T_Charges`

| Champ | Type | Description |
|---|---|---|
| `ChargeID` | Texte UUID | Identifiant unique |
| `CompteID` | Texte UUID | Compte rattaché |
| `TypeCharge` | Texte | Fixe ou variable |
| `CategorieCharge` | Texte | Semences, engrais, main-d'œuvre, bâtiment, etc. |
| `Designation` | Texte | Libellé détaillé |
| `Quantite` | Décimal | Quantité consommée |
| `Unite` | Texte | Unité |
| `CoutUnitaire` | Monétaire | Prix unitaire |
| `Montant` | Monétaire | Quantité x coût unitaire |
| `Annualisable` | Booléen | Charge à annualiser ou non |

#### `T_Produits`

| Champ | Type | Description |
|---|---|---|
| `ProduitID` | Texte UUID | Identifiant unique |
| `CompteID` | Texte UUID | Compte rattaché |
| `TypeProduit` | Texte | Principal, secondaire, autoconsommation, stock |
| `Designation` | Texte | Produit vendu ou valorisé |
| `Quantite` | Décimal | Quantité produite |
| `Unite` | Texte | Unité |
| `PrixUnitaire` | Monétaire | Prix de vente ou valorisation |
| `Montant` | Monétaire | Quantité x prix unitaire |

#### `T_Indicateurs`

| Champ | Type | Description |
|---|---|---|
| `IndicateurID` | Texte UUID | Identifiant unique |
| `CompteID` | Texte UUID | Compte rattaché |
| `ProduitBrut` | Monétaire | Total des produits |
| `ChargesFixes` | Monétaire | Total charges fixes |
| `ChargesVariables` | Monétaire | Total charges variables |
| `ChargesTotales` | Monétaire | Total charges |
| `MargeBrute` | Monétaire | Produit brut - charges variables |
| `MargeNette` | Monétaire | Produit brut - charges totales |
| `Benefice` | Monétaire | Résultat économique |
| `CoutProduction` | Monétaire | Coût par unité produite |
| `CoutParHa` | Monétaire | Charges totales / ha |
| `RevenuParHa` | Monétaire | Produit brut / ha |
| `ROI` | Pourcentage | Bénéfice / investissement |
| `RatioBeneficeCout` | Décimal | Produit brut / charges totales |
| `SeuilRentabiliteValeur` | Monétaire | Point mort en valeur |
| `SeuilRentabiliteQuantite` | Décimal | Point mort en volume |
| `ProductiviteMainOeuvre` | Décimal | Produit ou valeur par jour-homme |
| `BeneficeAnnuel` | Monétaire | Bénéfice annualisé |

#### `T_Resultats_IPE`

| Champ | Type | Description |
|---|---|---|
| `IPEID` | Texte UUID | Identifiant unique |
| `CompteID` | Texte UUID | Compte évalué |
| `ScoreRentabilite` | Décimal | Score normalisé |
| `ScoreCycle` | Décimal | Score normalisé |
| `ScoreInvestissement` | Décimal | Score normalisé |
| `ScoreRisque` | Décimal | Score normalisé |
| `ScoreTresorerie` | Décimal | Score normalisé |
| `ScoreMarche` | Décimal | Score normalisé |
| `ScoreMainOeuvre` | Décimal | Score normalisé |
| `ScoreIPE` | Décimal | Score final sur 100 |
| `ClasseIPE` | Texte | A, B, C, D ou E |
| `Rang` | Entier | Position dans le classement |

## 10. Liste des modules VBA nécessaires

### 10.1 Modules standards

| Module VBA | Responsabilité |
|---|---|
| `modApp` | Initialisation, fermeture, constantes globales |
| `modNavigation` | Gestion des menus, affichage des écrans et retours |
| `modAuth` | Connexion, session utilisateur, vérification des droits |
| `modLicence` | Validation de licence et limites d'utilisation |
| `modSecurity` | Protection des feuilles, masquage, contrôle d'accès |
| `modProducteurs` | Services de gestion des producteurs |
| `modExploitations` | Services de gestion des exploitations |
| `modCultures` | Gestion du référentiel cultures |
| `modElevages` | Gestion du référentiel élevages |
| `modSpeculations` | Référentiel unifié des activités |
| `modCompteExploitation` | Création, validation et chargement des comptes |
| `modCalculCharges` | Calcul des charges fixes, variables et amortissements |
| `modCalculProduits` | Calcul des produits principaux et secondaires |
| `modCalculIndicateurs` | Calcul des indicateurs économiques obligatoires |
| `modCalculIPE` | Calcul de l'Indice de Performance Économique |
| `modComparaison` | Comparaison et classement des spéculations |
| `modDashboard` | Alimentation des indicateurs et graphiques |
| `modReports` | Génération des rapports et exports PDF |
| `modImportExport` | Import/export CSV, Excel et préparation migration SQL |
| `modValidation` | Contrôles de saisie et règles métier |
| `modDataAccess` | Lecture/écriture dans les tableaux structurés |
| `modAudit` | Journalisation des actions utilisateur |
| `modErrorHandler` | Gestion centralisée des erreurs |
| `modUtils` | Fonctions utilitaires : UUID, dates, formats, conversions |

### 10.2 Classes VBA recommandées

| Classe | Rôle |
|---|---|
| `clsUtilisateur` | Représentation d'un utilisateur connecté |
| `clsSession` | Session active, droits et contexte client |
| `clsProducteur` | Données et méthodes liées au producteur |
| `clsExploitation` | Données et méthodes liées à l'exploitation |
| `clsSpeculation` | Culture ou élevage analysé |
| `clsCompteExploitation` | Agrégat principal d'analyse économique |
| `clsCharge` | Ligne de charge |
| `clsProduit` | Ligne de produit |
| `clsIndicateursEconomiques` | Résultats calculés |
| `clsIPE` | Scores de performance économique |
| `clsRepository` | Accès générique aux tables Excel |

### 10.3 UserForms recommandés

| UserForm | Usage |
|---|---|
| `frmLogin` | Connexion utilisateur |
| `frmAccueil` | Navigation principale |
| `frmProducteur` | Création et modification producteur |
| `frmExploitation` | Gestion de l'exploitation |
| `frmSpeculation` | Sélection culture ou élevage |
| `frmCompteExploitation` | Saisie guidée des charges et produits |
| `frmCharges` | Saisie détaillée des charges |
| `frmProduits` | Saisie détaillée des produits |
| `frmResultats` | Consultation des indicateurs |
| `frmComparaison` | Comparaison multicritère |
| `frmParametres` | Administration des référentiels |
| `frmLicence` | Activation et contrôle de licence |

## 11. Schéma de fonctionnement général

```text
Ouverture du fichier
        |
        v
Contrôle licence -> Licence invalide : écran activation / arrêt contrôlé
        |
        v
Authentification utilisateur
        |
        v
Chargement session, droits et client
        |
        v
Tableau de bord / menu principal
        |
        +--> Gestion producteurs -> Exploitations -> Historique campagnes
        |
        +--> Référentiels cultures / élevages / prix / risques
        |
        +--> Nouveau compte d'exploitation
                |
                v
        Sélection producteur + exploitation + campagne + spéculation
                |
                v
        Saisie charges fixes, charges variables et produits
                |
                v
        Validation métier des données
                |
                v
        Calcul indicateurs économiques
                |
                v
        Calcul cycles annuels et IPE
                |
                v
        Classement et comparaison des spéculations
                |
                v
        Rapport, export PDF et mise à jour tableau de bord
```

## 12. Relations entre les modules

| Module source | Module cible | Nature de la relation |
|---|---|---|
| Authentification | Tous les modules | Filtre l'accès selon les droits |
| Licence | Authentification, Producteurs, Comptes | Vérifie l'autorisation et les limites |
| Producteurs | Exploitations | Un producteur possède une ou plusieurs exploitations |
| Exploitations | Comptes | Une exploitation peut avoir plusieurs comptes |
| Cultures/Élevages | Spéculations | Alimentent le référentiel unifié |
| Spéculations | Comptes | Chaque compte analyse une spéculation |
| Comptes | Charges/Produits | Un compte contient plusieurs lignes de charges et produits |
| Charges/Produits | Indicateurs | Les montants alimentent les calculs économiques |
| Indicateurs | IPE | Les résultats financiers alimentent le score de performance |
| IPE | Tableau de bord/Comparaison | Les scores alimentent les classements |
| Reporting | Producteurs/Comptes/IPE | Génère les fiches et rapports |
| Audit | Tous les modules sensibles | Trace les créations, modifications, suppressions et exports |

## 13. Recommandations techniques professionnelles

### 13.1 Qualité des données

- Utiliser des tableaux structurés Excel pour toutes les tables `T_`.
- Générer des identifiants uniques stables de type UUID pour éviter les conflits lors de migrations ou consolidations.
- Interdire les suppressions physiques des données métier ; préférer un champ `Actif` ou `Statut`.
- Mettre en place des listes déroulantes liées aux référentiels.
- Valider les champs numériques : aucune quantité, superficie, prix ou durée négative sans justification métier.
- Contrôler les divisions par zéro dans tous les calculs.

### 13.2 Sécurité et robustesse

- Protéger les feuilles de données et masquer les feuilles techniques.
- Centraliser la gestion des erreurs dans `modErrorHandler`.
- Journaliser les actions critiques dans `LOG_Audit`.
- Prévoir une sauvegarde automatique horodatée avant les opérations massives.
- Séparer les constantes, paramètres et formules métier du code VBA.

### 13.3 Ergonomie

- Utiliser un parcours de saisie guidé : producteur, exploitation, spéculation, charges, produits, résultats.
- Afficher des messages métier clairs plutôt que des erreurs techniques VBA.
- Prévoir des boutons d'action visibles : Nouveau, Enregistrer, Valider, Calculer, Exporter, Retour.
- Utiliser des codes couleur sobres : vert pour rentable, orange pour vigilance, rouge pour déficitaire.
- Prévoir des fiches imprimables standardisées pour les conseillers et institutions financières.

### 13.4 Performance Excel/VBA

- Désactiver temporairement `ScreenUpdating`, `EnableEvents` et le recalcul automatique lors des traitements lourds.
- Charger les plages de données en tableaux VBA en mémoire pour les traitements comparatifs.
- Éviter les boucles cellule par cellule sur de grands volumes.
- Indexer en mémoire les tables clés par identifiant avec `Scripting.Dictionary`.
- Nettoyer régulièrement les feuilles temporaires.

### 13.5 Préparation à la migration SQL/Web/Android

- Maintenir une structure relationnelle stricte dans les feuilles Excel.
- Éviter de lier la logique métier aux coordonnées fixes des cellules.
- Documenter toutes les tables et champs.
- Prévoir des exports CSV normalisés par table.
- Conserver les mêmes noms d'entités dans VBA et dans le futur modèle SQL.
- Prévoir une future API avec les ressources : `/producteurs`, `/exploitations`, `/speculations`, `/comptes`, `/indicateurs`, `/ipe`.

## 14. Feuille de route de réalisation

### Phase 1 — Socle

- Créer la structure des feuilles et référentiels.
- Développer l'authentification locale.
- Mettre en place la navigation et la sécurité des feuilles.
- Créer la gestion des producteurs et exploitations.

### Phase 2 — Calcul économique

- Développer la saisie des comptes d'exploitation.
- Implémenter les charges, produits et indicateurs économiques.
- Générer le rapport de compte d'exploitation.

### Phase 3 — Intelligence économique

- Ajouter la gestion des cycles annuels.
- Implémenter l'IPE.
- Développer la comparaison des spéculations et les classements.

### Phase 4 — Professionnalisation

- Finaliser le tableau de bord.
- Ajouter les exports PDF/Excel.
- Renforcer l'audit, les sauvegardes et la validation.
- Préparer les formats d'export pour migration SQL.
