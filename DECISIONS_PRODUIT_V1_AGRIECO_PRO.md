# Décisions produit V1 — AGRIECO PRO

## Statut de validation

Ce document est **validé comme document officiel de cadrage de la V1**. Les décisions relatives à l'intégration des entrepreneurs agricoles comme cible principale, aux deux modes d'utilisation, au moteur économique unique, aux profils utilisateurs et à l'impact sur l'IQD sont approuvées.

Les questions restant à arbitrer seront traitées progressivement lors de la conception détaillée des règles de calcul et de la préparation de la V1. Elles ne doivent pas entraîner de dérive de périmètre : la priorité est de stabiliser l'architecture fonctionnelle avant d'ajouter de nouvelles fonctionnalités.

## 1. Objet du document

Ce document formalise les décisions produit, métier, économiques, fonctionnelles et techniques retenues pour la version 1 d'**AGRIECO PRO** avant tout développement VBA.

Il sert de référence de cadrage pour :

- éviter de commencer le développement sur des hypothèses implicites ;
- distinguer les décisions déjà validées des sujets encore ouverts ;
- préparer le futur document `REGLES_CALCUL_AGRIECO_PRO.md` ;
- garantir que la première version Excel/VBA reste compatible avec une migration future vers SQL, Web et Android.

> Aucun module VBA, aucun UserForm et aucun classeur Excel n'est généré dans ce document.

## 2. Méthode de présentation des décisions

Chaque décision importante est présentée selon la structure suivante :

- **Décision retenue** : choix officiellement adopté pour la V1 ou orientation à documenter.
- **Justification** : raison métier, économique ou technique.
- **Conséquences fonctionnelles** : effets visibles pour les utilisateurs et les processus.
- **Conséquences techniques** : effets sur les données, l'architecture Excel/VBA ou la future migration.
- **Risques éventuels** : limites, ambiguïtés ou risques d'implémentation.
- **Points restant à confirmer** : éléments à valider avant développement ou avant paramétrage final.

---

# 3. Décisions d'identité produit

## 3.1 Nom provisoire du produit

### Décision retenue

Le nom provisoire de l'application est **AGRIECO PRO**.

### Justification

Le nom exprime clairement la finalité agricole et économique de l'application. Il reste suffisamment générique pour couvrir les cultures, les élevages, les activités de conseil, les coopératives, les ONG, les institutions financières et les centres de formation.

### Conséquences fonctionnelles

- Le nom **AGRIECO PRO** sera utilisé dans les documents de conception.
- Les rapports imprimables, tableaux de bord et écrans pourront utiliser ce nom à titre provisoire.
- Le nom pourra être remplacé ultérieurement sans changer les règles économiques.

### Conséquences techniques

- Les noms de fichiers de documentation peuvent intégrer `AGRIECO_PRO`.
- Le nom ne doit pas être codé en dur dans plusieurs modules VBA ; il devra être centralisé dans une table de paramètres ou une constante d'application.

### Risques éventuels

- Le nom peut déjà être utilisé commercialement par un autre acteur.
- Un changement tardif de nom peut entraîner des corrections dans les rapports, menus, messages et exports.

### Points restant à confirmer

- Vérifier la disponibilité commerciale et juridique du nom.
- Valider le logo et l'identité visuelle.

## 3.2 Pays initial

### Décision retenue

Le pays initial de déploiement de la V1 est la **Côte d'Ivoire**.

### Justification

La Côte d'Ivoire dispose d'une forte diversité de spéculations agricoles et d'élevage : cultures vivrières, maraîchage, cultures pérennes, élevages courts et pisciculture. Elle constitue donc un terrain pertinent pour tester le modèle économique d'AGRIECO PRO.

### Conséquences fonctionnelles

- Les listes géographiques initiales devront être adaptées à la Côte d'Ivoire.
- Les référentiels de prix, campagnes agricoles, risques et marchés devront être contextualisés.
- Les rapports pourront afficher par défaut des libellés compatibles avec le contexte ivoirien.

### Conséquences techniques

- Prévoir des champs `Pays`, `Region`, `Departement`, `Commune`, `Village` et éventuellement `CoordonneesGPS`.
- Ne pas figer le pays dans le code ; conserver une table de localisations paramétrable.
- Préparer les tables pour supporter ultérieurement plusieurs pays.

### Risques éventuels

- Les calendriers agricoles varient fortement selon les zones agroécologiques.
- Les prix de référence peuvent varier par région, saison, qualité et circuit de commercialisation.

### Points restant à confirmer

- Liste administrative exacte à intégrer en V1.
- Niveau de détail géographique souhaité : région uniquement, département, sous-préfecture, commune, village, GPS.

## 3.3 Langue initiale

### Décision retenue

La langue initiale de la V1 est le **français**.

### Justification

Le français est la langue administrative, professionnelle et de formation dominante pour les utilisateurs visés en Côte d'Ivoire.

### Conséquences fonctionnelles

- Tous les menus, formulaires, messages, rapports et tableaux de bord seront en français.
- Les libellés économiques doivent être pédagogiques pour les conseillers et techniciens.

### Conséquences techniques

- Les libellés doivent idéalement être centralisés dans une table ou une feuille de paramètres pour faciliter une future traduction.
- Les formats de date, décimaux et monnaie doivent être compatibles avec les usages francophones.

### Risques éventuels

- Certains producteurs ou agents de terrain peuvent préférer des langues locales ou des pictogrammes.
- La traduction future sera plus coûteuse si les libellés sont dispersés dans le code VBA.

### Points restant à confirmer

- Nécessité d'une version bilingue à moyen terme.
- Niveau de langage : technique, simplifié ou formation.

## 3.4 Devise par défaut

### Décision retenue

La devise par défaut est le **franc CFA — FCFA**.

### Justification

Le FCFA est la monnaie utilisée en Côte d'Ivoire et dans plusieurs pays de l'UEMOA, ce qui facilite une extension régionale ultérieure.

### Conséquences fonctionnelles

- Les montants financiers seront affichés par défaut en FCFA.
- Les rapports devront préciser la devise utilisée.
- Les exports devront conserver la devise ou au minimum l'information de devise par compte.

### Conséquences techniques

- Prévoir une table `T_Devises` ou intégrer la devise dans `T_Parametres`.
- Prévoir le stockage des montants sans symbole dans les tables, avec format d'affichage séparé.
- Prévoir un champ `Devise` au niveau des comptes ou du client si l'application devient multi-pays.

### Risques éventuels

- Les exports CSV peuvent perdre le format monétaire.
- Une future extension hors zone FCFA nécessitera une gestion multi-devise.

### Points restant à confirmer

- Format exact d'affichage : `1 000 FCFA`, `1 000 XOF`, ou `FCFA 1 000`.
- Besoin futur de taux de change.

---

# 4. Décisions d'environnement technique

## 4.1 Fonctionnement hors connexion

### Décision retenue

AGRIECO PRO V1 doit fonctionner entièrement **hors connexion**.

### Justification

Les utilisateurs cibles peuvent intervenir dans des zones rurales où l'accès à Internet est limité, instable ou coûteux.

### Conséquences fonctionnelles

- Toutes les saisies, calculs, rapports, exports et sauvegardes doivent fonctionner localement.
- Aucune fonctionnalité essentielle ne doit dépendre d'un serveur distant.
- Les utilisateurs doivent pouvoir travailler sur le terrain sans connexion.

### Conséquences techniques

- Stockage local dans le classeur Excel ou dans des fichiers associés.
- Sauvegarde locale horodatée.
- Pas d'activation de licence en ligne obligatoire pour le moteur économique de la V1.
- Prévoir des exports CSV/Excel pour synchronisation manuelle future.

### Risques éventuels

- Risque de duplication ou divergence de fichiers si plusieurs conseillers travaillent séparément.
- Consolidation multi-utilisateur plus difficile sans base centralisée.
- Sécurité limitée par les capacités natives d'Excel.

### Points restant à confirmer

- Besoin de consolidation périodique entre plusieurs fichiers.
- Emplacement recommandé des sauvegardes locales.

## 4.2 Système d'exploitation

### Décision retenue

La V1 est prévue pour **Windows uniquement**.

### Justification

Excel VBA est plus stable et plus complet sur Windows, notamment pour les UserForms, les contrôles ActiveX, les exports et certaines interactions système.

### Conséquences fonctionnelles

- Les utilisateurs devront disposer d'un ordinateur Windows.
- Les fonctionnalités ne seront pas testées ni garanties sur macOS.

### Conséquences techniques

- Les développements peuvent utiliser les capacités VBA disponibles sous Windows.
- Éviter toutefois les dépendances excessives à des bibliothèques non standard pour faciliter la maintenance.

### Risques éventuels

- Exclusion des utilisateurs Mac.
- Dépendance au parc informatique Windows des organisations clientes.

### Points restant à confirmer

- Versions Windows minimales à supporter.
- Politique d'installation sur postes verrouillés par des organisations.

## 4.3 Version minimale Excel

### Décision retenue

La version minimale supportée est **Microsoft Excel 2016**.

### Justification

Excel 2016 est suffisamment répandu dans les organisations et permet l'utilisation des tableaux structurés, graphiques, macros VBA et fonctionnalités nécessaires à la V1.

### Conséquences fonctionnelles

- Les fichiers doivent rester compatibles avec Excel 2016.
- Les fonctions Excel plus récentes ne doivent pas être requises pour les calculs essentiels.

### Conséquences techniques

- Éviter les fonctions dynamiques récentes comme `FILTER`, `UNIQUE`, `XLOOKUP` si elles ne sont pas disponibles.
- Préférer les calculs VBA et les tableaux structurés compatibles.
- Tester le classeur sur Excel 2016 avant diffusion.

### Risques éventuels

- Des utilisateurs peuvent disposer de versions antérieures.
- Certaines fonctionnalités d'export ou d'interface peuvent se comporter différemment selon les installations Office.

### Points restant à confirmer

- Compatibilité attendue avec Office 365.
- Besoin de prise en charge des versions 32 bits et 64 bits d'Excel.

## 4.4 Interface mixte

### Décision retenue

L'interface de la V1 sera mixte :

- **UserForms** pour les saisies ;
- **feuilles Excel** pour les tableaux de bord, résultats et rapports.

### Justification

Les UserForms offrent une saisie contrôlée et professionnelle, tandis que les feuilles Excel restent performantes pour afficher des tableaux, graphiques et états imprimables.

### Conséquences fonctionnelles

- Les utilisateurs saisissent les données via des formulaires guidés.
- Les résultats, comparaisons et rapports sont consultés dans des feuilles formatées.
- Les rapports peuvent être imprimés ou exportés facilement.

### Conséquences techniques

- Développer un contrôleur de navigation entre UserForms et feuilles.
- Protéger les feuilles de résultats contre les modifications accidentelles.
- Centraliser les validations pour éviter des règles différentes entre formulaires et feuilles.

### Risques éventuels

- Complexité de synchronisation entre formulaires et tableaux Excel.
- Risque d'accès direct aux feuilles de données si la protection est insuffisante.

### Points restant à confirmer

- Niveau exact de personnalisation graphique des UserForms.
- Liste finale des écrans de saisie et de consultation.

## 4.5 Stockage local dans des tableaux structurés Excel

### Décision retenue

Les données de la V1 seront stockées localement dans des **tableaux structurés Excel**.

### Justification

Les tableaux structurés permettent de simuler une base relationnelle simple, d'améliorer la lisibilité et de préparer une migration future vers SQL.

### Conséquences fonctionnelles

- Les données sont incluses dans le classeur ou dans des fichiers Excel locaux associés.
- L'utilisateur peut sauvegarder et transporter l'application sans serveur.

### Conséquences techniques

- Chaque table doit disposer d'une clé primaire stable.
- Les relations doivent être gérées par identifiants et non par noms visibles.
- Les suppressions physiques doivent être évitées au profit de statuts.

### Risques éventuels

- Excel n'est pas une base de données transactionnelle.
- Risque de corruption ou de lenteur si le volume devient trop important.

### Points restant à confirmer

- Capacité maximale de producteurs, comptes et lignes de charges attendue pour la V1.
- Stratégie de compactage ou archivage.

## 4.6 Compatibilité migration SQL, Web et Android

### Décision retenue

L'architecture doit rester compatible avec une migration future vers **SQL, Web et Android**.

### Justification

La V1 Excel/VBA sert de version opérationnelle rapide, mais l'objectif professionnel impose une trajectoire d'évolution vers des plateformes plus robustes et multi-utilisateurs.

### Conséquences fonctionnelles

- Les concepts métier doivent rester stables : producteurs, exploitations, campagnes, spéculations, comptes, charges, produits, indicateurs, IPE.
- Les exports doivent permettre la reprise ou consolidation des données.

### Conséquences techniques

- Tables normalisées.
- Identifiants uniques stables.
- Logique métier documentée en dehors du code VBA.
- Préparation de formats CSV exploitables.

### Risques éventuels

- Une conception trop liée à Excel compliquera la migration.
- Les formules non documentées dans les cellules seront difficiles à reproduire dans SQL ou Web.

### Points restant à confirmer

- Niveau de normalisation cible pour la V1.
- Date ou jalon envisagé pour une migration.

## 4.7 Exports PDF, Excel et CSV

### Décision retenue

La V1 doit prévoir des exports en **PDF, Excel et CSV**.

### Justification

Les utilisateurs ont des besoins différents : impression et partage en PDF, analyse complémentaire en Excel, migration ou consolidation en CSV.

### Conséquences fonctionnelles

- Les rapports doivent être exportables en PDF.
- Les tableaux de données ou résultats doivent pouvoir être exportés en Excel et CSV.
- Les exports doivent être compréhensibles par des conseillers, organisations et institutions financières.

### Conséquences techniques

- Prévoir un module `modImportExport`.
- Définir un format CSV par table.
- Prévoir des noms de fichiers normalisés avec date, producteur, campagne et spéculation.

### Risques éventuels

- Les fichiers CSV peuvent poser des problèmes d'encodage et de séparateur décimal.
- Les exports Excel peuvent être modifiés manuellement et réimportés avec erreurs.

### Points restant à confirmer

- Séparateur CSV : point-virgule recommandé en contexte francophone.
- Encodage cible : UTF-8 recommandé.
- Modèles exacts de rapports PDF.

## 4.8 Sauvegarde automatique locale

### Décision retenue

La V1 doit intégrer une sauvegarde automatique locale, horodatée et restaurable.

### Justification

Le stockage local dans Excel expose l'utilisateur aux risques de suppression, corruption, mauvaise manipulation ou panne.

### Conséquences fonctionnelles

- L'utilisateur doit pouvoir restaurer une version précédente.
- Les opérations sensibles doivent déclencher ou proposer une sauvegarde.

### Conséquences techniques

- Prévoir un dossier local de sauvegarde.
- Nommer les sauvegardes avec date, heure et version.
- Prévoir un mécanisme de restauration contrôlé.

### Risques éventuels

- Accumulation de sauvegardes volumineuses.
- Restauration incorrecte si plusieurs versions du fichier existent.

### Points restant à confirmer

- Emplacement par défaut des sauvegardes.
- Nombre de sauvegardes à conserver.
- Politique de suppression automatique des anciennes sauvegardes.

## 4.9 Licence reportée après validation du moteur économique

### Décision retenue

Le système de licence sera développé seulement après validation du moteur économique.

### Justification

La priorité est de fiabiliser les calculs économiques et les processus métier avant d'ajouter une couche de restriction commerciale.

### Conséquences fonctionnelles

- La V1 de cadrage peut fonctionner sans blocage de licence.
- Les règles économiques peuvent être testées librement par les parties prenantes.

### Conséquences techniques

- Prévoir l'architecture du module licence, mais ne pas le rendre bloquant au démarrage dans les premières itérations.
- Éviter de coupler les calculs économiques au mécanisme de licence.

### Risques éventuels

- Si la licence est ajoutée tardivement sans anticipation, elle peut nécessiter une refonte.
- Risque de diffusion incontrôlée des versions de test.

### Points restant à confirmer

- Type de licence future : annuelle, perpétuelle, par poste, par organisation, par volume de producteurs.
- Besoin d'activation hors ligne ou en ligne.

---

# 5. Décisions sur les utilisateurs prioritaires

## 5.1 Cibles principales

### Décision retenue

Les cibles principales de la V1 sont :

- entrepreneurs agricoles ;
- producteurs agricoles structurés ;
- conseillers agricoles ;
- techniciens agricoles et agents d'encadrement.

### Justification

AGRIECO PRO doit être à la fois un outil personnel de gestion, de planification et d'aide à la décision pour un entrepreneur agricole, et un outil professionnel d'accompagnement pour les conseillers et agents d'encadrement. Cette orientation élargit l'utilité du produit sans changer le moteur économique.

### Conséquences fonctionnelles

- Prévoir un **mode individuel** pour gérer sa propre exploitation, ses parcelles, ses projets, ses campagnes, ses dépenses, ses récoltes, ses ventes, sa trésorerie, ses dettes, ses créances, ses stocks, ses objectifs, ses simulations et ses recommandations.
- Prévoir un **mode conseiller ou organisation** pour gérer plusieurs producteurs, plusieurs exploitations, consolider les résultats, comparer des bénéficiaires, produire des statistiques et éditer des rapports professionnels.
- Permettre deux niveaux de saisie : saisie simplifiée et saisie détaillée.
- Adapter les menus, tableaux de bord et rapports au profil sans dupliquer les règles de calcul.

### Conséquences techniques

- Le moteur économique doit rester commun aux deux modes.
- La différence entre les modes doit concerner principalement les parcours utilisateurs, menus, tableaux de bord, droits d'accès, rapports, niveau de complexité d'interface et volume de producteurs ou exploitations gérés.
- Prévoir dans les tables des champs permettant de distinguer un usage individuel d'un usage organisationnel sans modifier les formules.
- Préparer un paramètre de profil d'utilisation : `Individuel`, `Conseiller`, `Organisation` ou équivalent.

### Risques éventuels

- Risque de surcharge fonctionnelle si la V1 tente de satisfaire tous les profils avec le même niveau de détail.
- Risque d'interface trop complexe pour un entrepreneur agricole peu formé à la gestion.
- Risque de confusion si les tableaux de bord individuels et institutionnels ne sont pas séparés clairement.

### Points restant à confirmer

- Profil à privilégier lors des premiers tests utilisateurs.
- Niveau minimal de gestion des dettes, créances et stocks en mode individuel V1.
- Niveau de consolidation attendu en mode organisation V1.

## 5.2 Cibles secondaires

### Décision retenue

Les cibles secondaires sont :

- coopératives ;
- organisations professionnelles agricoles ;
- cabinets de conseil ;
- ONG ;
- projets de développement ;
- institutions financières ;
- investisseurs agricoles ;
- centres de formation ;
- grandes institutions agricoles.

### Justification

Ces acteurs ont des besoins complémentaires : consolidation, financement, formation, suivi de performance, conseil économique et reporting.

### Conséquences fonctionnelles

- Les rapports doivent être suffisamment professionnels pour être partagés.
- Les indicateurs doivent être exploitables pour le conseil, la formation et l'analyse de financement.
- Le tableau de bord doit pouvoir résumer plusieurs exploitations.

### Conséquences techniques

- Prévoir des champs organisationnels : client, coopérative, projet, conseiller, zone.
- Prévoir des exports propres pour consolidation externe.

### Risques éventuels

- Les institutions financières peuvent demander des indicateurs plus détaillés sur le risque et le crédit.
- Les ONG peuvent demander des indicateurs sociaux ou environnementaux non prévus en V1.

### Points restant à confirmer

- Priorité relative des cibles secondaires.
- Rapports spécifiques attendus par les institutions financières ou projets.

## 5.3 Double mode d'utilisation : individuel et conseiller/organisation

### Décision retenue

AGRIECO PRO doit répondre à deux usages principaux :

1. **Mode individuel** : destiné à un entrepreneur agricole ou producteur structuré qui utilise le logiciel pour gérer sa propre activité.
2. **Mode conseiller ou organisation** : destiné à un conseiller, cabinet, coopérative, ONG, projet ou institution qui accompagne plusieurs producteurs ou exploitations.

### Justification

Les entrepreneurs agricoles ont besoin d'un outil personnel de pilotage, tandis que les organisations ont besoin d'un outil de suivi, consolidation et reporting. Les deux usages reposent sur les mêmes calculs économiques ; il faut donc éviter deux produits séparés.

### Conséquences fonctionnelles

Le mode individuel devra couvrir les fonctions suivantes :

- tableau de bord personnel ;
- module « Mon projet agricole » ;
- compte d'exploitation prévisionnel ;
- compte d'exploitation réel ;
- comparaison prévision/réalisation ;
- suivi des écarts ;
- planification de campagne ;
- journal des opérations ;
- suivi des dépenses payées ;
- suivi des dépenses restant à payer ;
- suivi des achats à crédit ;
- suivi des dettes fournisseurs ;
- suivi des ventes encaissées ;
- suivi des créances clients ;
- gestion des stocks d'intrants ;
- gestion des stocks de produits agricoles ;
- suivi de la trésorerie ;
- suivi des objectifs de production ;
- suivi des objectifs financiers ;
- estimation du besoin de financement ;
- simulation du prix, du rendement, de la superficie et des charges ;
- aide à la décision pour développer, maintenir, réduire ou abandonner une activité.

Le mode conseiller ou organisation devra couvrir :

- gestion de plusieurs producteurs ;
- gestion de plusieurs exploitations ;
- consolidation des résultats ;
- comparaisons entre bénéficiaires ;
- statistiques ;
- rapports professionnels et institutionnels.

### Conséquences techniques

- Un seul moteur de calcul doit servir les deux modes.
- Les menus et tableaux de bord doivent être configurables selon le mode.
- Les droits d'accès doivent permettre de limiter un entrepreneur à ses propres données et une organisation à son portefeuille.
- Les rapports doivent partager les mêmes indicateurs mais présenter des niveaux d'agrégation différents.

### Risques éventuels

- Risque de glissement vers un logiciel de comptabilité complet si dettes, créances et trésorerie sont trop détaillées en V1.
- Risque de surcharge des écrans si les fonctions organisationnelles apparaissent dans le mode individuel.
- Risque de performance si une organisation gère trop de producteurs dans un seul classeur Excel.

### Points restant à confirmer

- Limites fonctionnelles de la gestion des dettes et créances en V1.
- Profondeur du journal des opérations : simple suivi agricole ou journal quasi-comptable.
- Capacité maximale par mode : nombre de producteurs, exploitations, activités et lignes de flux.

## 5.4 Organisation simplifiée du mode individuel

### Décision retenue

Le mode individuel doit prévoir une organisation fonctionnelle simple autour des rubriques suivantes, sans que cette liste implique encore la création de feuilles Excel ou de UserForms :

- Mon exploitation ;
- Mes parcelles ;
- Mes activités ;
- Mon projet agricole ;
- Mes dépenses ;
- Mes récoltes ;
- Mes ventes ;
- Mes stocks ;
- Ma trésorerie ;
- Mes résultats ;
- Mes simulations ;
- Mes objectifs ;
- Mes rapports.

### Justification

Ces rubriques utilisent un vocabulaire personnel et concret, mieux adapté à un entrepreneur agricole qu'une terminologie institutionnelle.

### Conséquences fonctionnelles

- Le parcours individuel doit être orienté action et décision.
- Le tableau de bord personnel doit afficher les objectifs, dépenses, ventes, trésorerie, résultats et alertes de l'utilisateur.
- Les rubriques institutionnelles doivent être masquées ou simplifiées en mode individuel.

### Conséquences techniques

- Prévoir une couche de navigation configurable par profil.
- Ne pas créer de tables séparées pour le mode individuel si les mêmes entités métier suffisent.
- Prévoir des libellés d'interface distincts sans dupliquer les calculs.

### Risques éventuels

- Des rubriques trop nombreuses peuvent alourdir la V1.
- Les objectifs et simulations peuvent demander des règles supplémentaires si leur périmètre n'est pas borné.

### Points restant à confirmer

- Rubriques à retenir dans la toute première version utilisable.
- Présentation du tableau de bord personnel.

## 5.5 Niveaux de saisie

### Décision retenue

AGRIECO PRO doit prévoir deux niveaux d'utilisation :

1. **Saisie simplifiée**, pour obtenir rapidement une analyse avec peu de données.
2. **Saisie détaillée**, pour les entrepreneurs expérimentés et les professionnels.

### Justification

Tous les utilisateurs n'ont pas le même niveau de maîtrise en gestion. La saisie simplifiée réduit la barrière d'entrée, tandis que la saisie détaillée garantit une analyse complète lorsque les données sont disponibles.

### Conséquences fonctionnelles

La saisie simplifiée doit couvrir :

- activité ;
- superficie ou effectif ;
- rendement attendu ;
- prix de vente ;
- principales charges ;
- nombre de cycles ;
- durée du cycle.

La saisie détaillée peut intégrer :

- opérations par date ;
- charges détaillées ;
- main-d'œuvre familiale et salariée ;
- amortissements ;
- investissements ;
- crédits ;
- dettes ;
- créances ;
- stocks ;
- pertes ;
- récoltes multiples ;
- ventes multiples ;
- flux de trésorerie.

### Conséquences techniques

- Les calculs doivent accepter des comptes simplifiés avec un Indice de Qualité des Données plus faible.
- Le niveau de saisie doit être enregistré pour interpréter correctement les résultats.
- Les données manquantes en saisie simplifiée doivent être soit estimées à partir de référentiels, soit signalées dans les rapports.

### Risques éventuels

- Les résultats issus d'une saisie simplifiée peuvent être moins précis.
- Les utilisateurs peuvent comparer des comptes simplifiés et détaillés sans tenir compte de la qualité des données.

### Points restant à confirmer

- Champs strictement obligatoires en saisie simplifiée.
- Méthode d'estimation par défaut des charges non saisies.
- Seuil minimal d'IQD acceptable pour afficher une recommandation.

---

# 6. Décisions sur les spéculations pilotes

## 6.1 Cultures pilotes

### Décision retenue

Les cultures pilotes de la V1 sont :

- maïs ;
- riz ;
- manioc ;
- tomate ;
- cacao.

### Justification

Cette sélection couvre plusieurs familles de cas : cultures annuelles, vivrières, maraîchères et pérennes. Elle permet de tester des cycles courts, des cultures saisonnières, des produits périssables et une culture pérenne nécessitant une analyse pluriannuelle.

### Conséquences fonctionnelles

- Les référentiels initiaux devront contenir ces cinq cultures.
- Les formulaires devront gérer des paramètres différents selon la culture sélectionnée.
- Le cacao impose une logique pluriannuelle distincte des cultures annuelles.

### Conséquences techniques

- Prévoir un modèle générique de culture avec paramètres activables.
- Ne pas créer un formulaire spécifique pour chaque culture.
- Prévoir des tables de paramètres par spéculation.

### Risques éventuels

- Le cacao complexifie la V1 en raison de la phase d'installation et de l'entrée progressive en production.
- La tomate impose de bien gérer pertes, ventes multiples et prix variables.

### Points restant à confirmer

- Variétés ou systèmes de production à représenter pour chaque culture.
- Calendriers techniques de référence par zone.

## 6.2 Élevages pilotes

### Décision retenue

Les élevages pilotes de la V1 sont :

- poulet de chair ;
- porc ;
- pisciculture.

### Justification

Ces trois spéculations couvrent des cycles courts, moyens et aquacoles, avec des charges d'alimentation, de santé, d'effectifs et de mortalité significatives.

### Conséquences fonctionnelles

- Les formulaires d'élevage devront gérer effectifs, mortalité, alimentation, médicaments, poids ou biomasse et ventes.
- La pisciculture nécessitera des paramètres spécifiques : alevins, alimentation, bassins, biomasse, captures ou récoltes.

### Conséquences techniques

- Prévoir un modèle générique d'élevage avec champs activables selon la filière.
- Paramétrer les unités : tête, kg vif, kg poisson, sac d'aliment, dose, litre, bassin.

### Risques éventuels

- Les indicateurs zootechniques peuvent varier fortement selon les pratiques.
- La pisciculture peut impliquer plusieurs récoltes ou ventes partielles.

### Points restant à confirmer

- Niveau de détail zootechnique attendu en V1.
- Distinction entre engraissement, reproduction et production mixte.

## 6.3 Paramétrabilité des spéculations

### Décision retenue

La conception doit rester paramétrable afin d'ajouter ultérieurement d'autres cultures et élevages sans modifier profondément le code.

### Justification

L'application doit pouvoir évoluer selon les zones, organisations et filières ciblées.

### Conséquences fonctionnelles

- Les administrateurs devront pouvoir ajouter ou modifier des spéculations dans des référentiels.
- Les rubriques de charges, produits et paramètres doivent être configurables.

### Conséquences techniques

- Les formulaires doivent lire les paramètres de la spéculation sélectionnée.
- Les champs spécifiques doivent être activés ou masqués selon des règles de configuration.
- Les calculs communs doivent être séparés des paramètres propres à chaque spéculation.

### Risques éventuels

- Une paramétrabilité excessive peut rendre la V1 plus longue à développer.
- Des règles très spécifiques à certaines filières peuvent nécessiter des extensions futures.

### Points restant à confirmer

- Niveau d'administration accessible à l'utilisateur final.
- Liste des paramètres configurables en V1.

---

# 7. Décisions économiques retenues

## 7.1 Amortissement linéaire

### Décision retenue

L'amortissement sera calculé selon la méthode **linéaire**.

### Justification

La méthode linéaire est simple, compréhensible et adaptée à un outil de conseil technico-économique en V1.

### Conséquences fonctionnelles

- L'utilisateur saisira la valeur d'acquisition, la durée d'amortissement et éventuellement la valeur résiduelle.
- Le rapport présentera l'amortissement séparément des dépenses réellement payées.

### Conséquences techniques

- Prévoir une table des investissements et amortissements.
- Calcul de référence : `(ValeurAcquisition - ValeurResiduelle) / DureeAmortissement`.
- La durée doit être paramétrable par équipement ou type d'investissement.

### Risques éventuels

- Les règles comptables réelles peuvent différer selon les organisations.
- L'amortissement économique peut ne pas correspondre à l'amortissement fiscal.

### Points restant à confirmer

- Utilisation ou non d'une valeur résiduelle.
- Proratisation de l'amortissement pour les cycles inférieurs à un an.

## 7.2 Durée d'amortissement paramétrable

### Décision retenue

La durée d'amortissement sera **paramétrable**.

### Justification

Les durées d'usage varient selon les bâtiments, équipements, matériels, bassins, infrastructures et outils.

### Conséquences fonctionnelles

- L'utilisateur ou l'administrateur pourra définir des durées par type d'actif.
- Les rapports devront afficher les hypothèses d'amortissement utilisées.

### Conséquences techniques

- Prévoir un référentiel des catégories d'investissements et durées par défaut.
- Stocker la durée réellement utilisée dans le compte pour assurer la traçabilité.

### Risques éventuels

- Des durées incohérentes peuvent fausser fortement les résultats.
- Des utilisateurs non formés peuvent modifier des durées sans justification.

### Points restant à confirmer

- Qui peut modifier les durées : administrateur uniquement ou conseiller ?
- Durées par défaut pour les actifs agricoles courants.

## 7.3 Autoconsommation valorisée

### Décision retenue

L'autoconsommation sera valorisée et intégrée au produit brut, mais présentée séparément des ventes.

### Justification

L'autoconsommation représente une valeur économique réelle pour le producteur, même si elle ne génère pas d'encaissement.

### Conséquences fonctionnelles

- Les rapports distingueront ventes, autoconsommation, produits secondaires et stocks.
- Le résultat financier pourra exclure l'autoconsommation des encaissements réels.
- Le résultat économique intégrera l'autoconsommation valorisée.

### Conséquences techniques

- Prévoir un type de produit `Autoconsommation` dans `T_Produits`.
- Stocker quantité, unité, prix de valorisation et montant.
- Séparer les agrégats financiers et économiques.

### Risques éventuels

- Le prix de valorisation peut être contesté.
- Risque de double comptage si une quantité autoconsommée est aussi saisie en vente ou stock.

### Points restant à confirmer

- Prix de valorisation à utiliser : prix marché local, prix moyen de vente, prix administrateur.
- Traitement des dons et consommations familiales non quantifiées.

## 7.4 Main-d'œuvre familiale valorisable

### Décision retenue

La main-d'œuvre familiale devra pouvoir être valorisée.

### Justification

La main-d'œuvre familiale est souvent centrale dans les exploitations agricoles. La valoriser permet de mesurer la rentabilité économique réelle, au-delà des seules dépenses monétaires.

### Conséquences fonctionnelles

- L'utilisateur pourra saisir des jours-homme familiaux.
- Le logiciel pourra afficher un résultat financier et un résultat économique.
- Les rapports devront distinguer main-d'œuvre payée et main-d'œuvre familiale valorisée.

### Conséquences techniques

- Prévoir un type de charge `MainOeuvreFamilialeValorisee`.
- Prévoir des coûts journaliers de référence paramétrables.
- Exclure cette charge du résultat financier si elle n'est pas effectivement payée.

### Risques éventuels

- La valorisation peut réduire fortement le bénéfice économique et être mal comprise.
- Le coût journalier de référence peut varier selon zone, saison, tâche et genre.

### Points restant à confirmer

- Barème de valorisation de la journée de travail.
- Distinction homme/femme/jeune ou type de tâche.

## 7.5 Résultat financier et résultat économique

### Décision retenue

Le logiciel devra produire deux résultats distincts :

- **résultat financier**, basé principalement sur les entrées et sorties monétaires réelles ;
- **résultat économique**, intégrant les charges valorisées, notamment la main-d'œuvre familiale.

### Justification

Cette distinction permet de répondre à deux usages : gestion de trésorerie réelle et analyse économique complète.

### Conséquences fonctionnelles

- Les rapports afficheront deux niveaux de résultat.
- Les utilisateurs pourront expliquer pourquoi une activité semble rentable financièrement mais moins rentable économiquement.
- Les institutions financières pourront mieux distinguer cash-flow et rentabilité économique.

### Conséquences techniques

- Chaque charge et produit doit indiquer s'il est monétaire, valorisé ou non monétaire.
- Les indicateurs doivent être calculés en variante financière et économique lorsque pertinent.
- Prévoir des colonnes distinctes dans les tables de résultats.

### Risques éventuels

- Complexité pédagogique accrue.
- Risque de confusion si les deux résultats ne sont pas clairement nommés.

### Points restant à confirmer

- Liste exacte des indicateurs à décliner en financier et économique.
- Format de présentation dans le tableau de bord.

## 7.6 Classification détaillée des charges

### Décision retenue

Les charges devront être distinguées en :

- charges réellement payées ;
- charges valorisées ;
- apports familiaux ;
- charges fixes ;
- charges variables ;
- investissements ;
- amortissements ;
- frais financiers.

### Justification

Cette classification permet de calculer correctement les marges, le besoin de trésorerie, le résultat financier, le résultat économique et la rentabilité.

### Conséquences fonctionnelles

- La saisie des charges devra guider l'utilisateur dans le choix de la nature de charge.
- Les rapports devront regrouper les charges par catégories.
- Les frais financiers devront être présentés séparément.

### Conséquences techniques

- Prévoir plusieurs attributs sur chaque ligne de charge : type économique, type de paiement, fixité, récurrence, cycle, spéculation, source familiale ou externe.
- Ne pas se limiter à un seul champ `TypeCharge`.

### Risques éventuels

- Trop de catégories peuvent ralentir la saisie.
- Des erreurs de classification peuvent fausser les indicateurs.

### Points restant à confirmer

- Liste finale des catégories de charges.
- Valeurs par défaut selon culture ou élevage.

## 7.7 Traitement des crédits

### Décision retenue

Les crédits reçus ne doivent pas être considérés comme des produits d'exploitation. Le remboursement du capital d'un crédit ne doit pas être considéré comme une charge de production. Les intérêts, commissions et frais financiers doivent être présentés séparément.

### Justification

Un crédit est une source de financement, pas un produit issu de l'activité agricole. Le remboursement du capital rembourse une dette, tandis que les intérêts représentent le coût du financement.

### Conséquences fonctionnelles

- Les rapports doivent distinguer exploitation et financement.
- Les intérêts peuvent affecter le résultat net après financement ou un indicateur financier séparé.
- Les crédits peuvent être utiles pour analyser le besoin de trésorerie.

### Conséquences techniques

- Prévoir une table ou rubrique `Financements`.
- Ne pas intégrer le montant du crédit dans le produit brut.
- Ne pas intégrer le capital remboursé dans le coût de production.
- Stocker séparément intérêts, commissions et frais.

### Risques éventuels

- Les utilisateurs peuvent vouloir intégrer les remboursements dans leur vision de trésorerie.
- Confusion possible entre rentabilité de production et capacité de remboursement.

### Points restant à confirmer

- Présenter ou non un tableau de cash-flow incluant les remboursements de capital.
- Niveau de détail des échéanciers de crédit en V1.

## 7.8 Produits secondaires, stocks, pertes et autoconsommations identifiables

### Décision retenue

Les produits secondaires, stocks, pertes et autoconsommations doivent être identifiables séparément.

### Justification

Ces éléments ont des impacts différents sur les ventes, le produit brut, le stockage, les pertes économiques et la disponibilité alimentaire.

### Conséquences fonctionnelles

- Les formulaires de produits devront prévoir plusieurs types de lignes.
- Les rapports devront détailler les ventes, produits secondaires, stocks, pertes et autoconsommations.

### Conséquences techniques

- Prévoir un champ `TypeProduit` ou `NatureFluxProduit`.
- Prévoir des tables ou colonnes pour stocks entrants, stocks sortants et pertes.
- Éviter le double comptage entre production, vente, autoconsommation et stock.

### Risques éventuels

- Saisie plus complexe.
- Nécessité de règles strictes pour équilibrer production, ventes, stocks, pertes et autoconsommation.

### Points restant à confirmer

- Formule de contrôle de cohérence des flux physiques.
- Traitement des pertes avant récolte et post-récolte.

## 7.9 IPE paramétrable

### Décision retenue

Les pondérations de l'IPE doivent être paramétrables et non codées définitivement dans le VBA. Les pondérations actuellement proposées restent provisoires.

### Justification

Les priorités économiques varient selon les utilisateurs : rentabilité, risque, besoin de trésorerie, marché ou main-d'œuvre peuvent avoir des poids différents selon les projets.

### Conséquences fonctionnelles

- Un administrateur pourra ajuster les pondérations.
- Les rapports devront afficher les pondérations utilisées pour assurer la transparence.
- Les comparaisons devront être recalculables après modification des pondérations.

### Conséquences techniques

- Utiliser une table `T_Param_IPE`.
- Stocker les pondérations utilisées lors du calcul pour historiser les résultats.
- Contrôler que la somme des pondérations est égale à 100 %.

### Risques éventuels

- Modifier les pondérations peut changer fortement le classement des spéculations.
- Les comparaisons historiques peuvent devenir incohérentes si les pondérations ne sont pas historisées.

### Points restant à confirmer

- Pondérations finales de la V1.
- Droits nécessaires pour modifier les pondérations.

## 7.10 Cultures pérennes analysées sur plusieurs années

### Décision retenue

Les cultures pérennes doivent être analysées sur plusieurs années en distinguant :

- installation ;
- entretien avant production ;
- entrée en production ;
- régime de croisière ;
- renouvellement éventuel.

### Justification

Les cultures pérennes comme le cacao ne peuvent pas être analysées correctement sur un seul cycle annuel standard. Les dépenses d'installation précèdent les revenus et doivent être réparties ou analysées sur plusieurs années.

### Conséquences fonctionnelles

- Le compte d'exploitation des cultures pérennes devra afficher une vision pluriannuelle.
- Les années improductives doivent être visibles.
- Les indicateurs de rentabilité devront être adaptés au cycle long.

### Conséquences techniques

- Prévoir une table de phases pérennes ou une colonne `PhaseCulturePerenne`.
- Prévoir des comptes liés par projet pluriannuel.
- Préparer des indicateurs comme cumul des flux, année de retour sur investissement, bénéfice moyen annuel ou valeur actualisée si retenue plus tard.

### Risques éventuels

- Complexité supérieure aux cultures annuelles.
- Risque de mélanger analyse de campagne et analyse d'investissement pluriannuel.

### Points restant à confirmer

- Durée d'analyse standard du cacao.
- Utilisation ou non de taux d'actualisation en V1.
- Indicateurs pluriannuels exacts à retenir.

---

# 8. Principes de conception des formulaires

## 8.1 Formulaires génériques par famille d'activité

### Décision retenue

Il ne faut pas créer un formulaire entièrement différent pour chaque culture ou chaque élevage. La V1 doit prévoir :

- un modèle générique de saisie pour les cultures ;
- un modèle générique de saisie pour les élevages ;
- des champs, paramètres et rubriques spécifiques activés selon la spéculation choisie ;
- des référentiels configurables par spéculation.

### Justification

Cette approche réduit le volume de code, facilite la maintenance et permet d'ajouter de nouvelles spéculations sans refonte majeure.

### Conséquences fonctionnelles

- L'utilisateur choisit d'abord la spéculation.
- Le formulaire affiche ensuite uniquement les rubriques pertinentes.
- Les charges et produits types peuvent être proposés automatiquement selon la spéculation.

### Conséquences techniques

- Les UserForms doivent être pilotés par des référentiels.
- Les contrôles doivent être activés, masqués ou rendus obligatoires selon la configuration.
- Les validations doivent être génériques mais paramétrables.

### Risques éventuels

- Un formulaire trop générique peut devenir complexe à utiliser.
- Les exceptions métier peuvent nécessiter des extensions ciblées.

### Points restant à confirmer

- Champs obligatoires communs aux cultures.
- Champs obligatoires communs aux élevages.
- Champs spécifiques par spéculation pilote.

---

# 9. Fonctions prioritaires de la V1

## 9.1 Socle de gestion

### Décision retenue

La V1 doit couvrir au minimum :

- gestion des producteurs ;
- gestion des exploitations ;
- gestion des campagnes et cycles ;
- sélection de la spéculation.

### Justification

Ces fonctions constituent la base relationnelle nécessaire à tout compte d'exploitation fiable.

### Conséquences fonctionnelles

- Chaque analyse doit être rattachée à un producteur, une exploitation, une campagne et une spéculation.
- L'historique doit être consultable.

### Conséquences techniques

- Tables nécessaires : producteurs, exploitations, campagnes, cycles, spéculations.
- Identifiants uniques nécessaires pour toutes les entités.

### Risques éventuels

- Des données de base mal saisies compromettent toutes les analyses.

### Points restant à confirmer

- Champs obligatoires exacts pour producteur, exploitation et campagne.

## 9.2 Comptes prévisionnels, réels et comparaison

### Décision retenue

La V1 doit gérer :

- comptes prévisionnels ;
- comptes réels ;
- comparaison entre prévisions et réalisations.

### Justification

Le conseil agricole nécessite de planifier puis de comparer les résultats obtenus avec les hypothèses initiales.

### Conséquences fonctionnelles

- Un compte pourra être marqué comme prévisionnel ou réel.
- Le rapport affichera écarts en valeur et en pourcentage.
- Les conseillers pourront identifier les postes responsables des écarts.

### Conséquences techniques

- Prévoir un champ `TypeCompte` : Prévisionnel, Réel, Ajusté.
- Prévoir un lien entre un compte prévisionnel et son compte réel.
- Prévoir des calculs d'écarts par charge, produit et indicateur.

### Risques éventuels

- Les prévisions et réalisations peuvent ne pas avoir les mêmes rubriques, rendant la comparaison difficile.

### Points restant à confirmer

- Niveau de détail de comparaison attendu.
- Possibilité d'avoir plusieurs versions prévisionnelles.

## 9.3 Gestion détaillée des charges et produits

### Décision retenue

La V1 doit gérer :

- saisie des charges ;
- saisie des produits ;
- ventes multiples à des dates et prix différents ;
- autoconsommation ;
- produits secondaires ;
- stocks entrants et sortants ;
- pertes de production et pertes post-récolte ;
- main-d'œuvre familiale ;
- amortissements ;
- charges partagées entre plusieurs spéculations.

### Justification

La rentabilité agricole dépend de flux multiples et hétérogènes. Une saisie trop simplifiée produirait des résultats peu fiables.

### Conséquences fonctionnelles

- Les formulaires devront supporter plusieurs lignes par type de flux.
- Les utilisateurs pourront détailler des ventes à prix différents.
- Les pertes et stocks seront visibles séparément.

### Conséquences techniques

- Prévoir des tables de lignes de charges, produits, stocks, pertes, ventes et affectations.
- Prévoir des contrôles de cohérence des quantités physiques.
- Prévoir des catégories économiques et financières sur chaque ligne.

### Risques éventuels

- Charge de saisie importante pour les utilisateurs.
- Risque d'erreurs si les rubriques ne sont pas bien guidées.

### Points restant à confirmer

- Niveau de détail minimal obligatoire.
- Possibilité de saisie simplifiée et saisie détaillée.

## 9.4 Cycles, charges récurrentes et charges propres au premier cycle

### Décision retenue

La V1 doit gérer :

- calcul par cycle ;
- calcul annuel ;
- charges propres au premier cycle ;
- charges récurrentes ;
- charges réduites ou supprimées lors des cycles suivants.

### Justification

Certaines spéculations à cycles courts supportent des charges initiales différentes des charges récurrentes, par exemple préparation, installation, équipement léger ou nettoyage.

### Conséquences fonctionnelles

- L'utilisateur devra pouvoir indiquer si une charge s'applique au premier cycle seulement ou à tous les cycles.
- Le résultat annuel devra ajuster correctement les charges selon le nombre de cycles.

### Conséquences techniques

- Ajouter des attributs `CycleApplication`, `Recurrente`, `PremierCycleSeulement`, `TauxReductionCyclesSuivants`.
- Le moteur de calcul devra distinguer charges fixes annuelles, charges variables par cycle et charges spécifiques.

### Risques éventuels

- Mauvaise annualisation si les règles ne sont pas claires.
- Complexité pour les utilisateurs non formés.

### Points restant à confirmer

- Interface de saisie la plus simple pour ces règles.
- Règles par défaut par spéculation.

## 9.5 Calculs, classement, tableau de bord et rapports

### Décision retenue

La V1 doit gérer :

- calcul des indicateurs économiques ;
- calcul du besoin de trésorerie ;
- comparaison et classement des spéculations ;
- IPE paramétrable ;
- tableau de bord simple ;
- rapports imprimables ;
- export PDF, Excel et CSV ;
- sauvegarde et restauration.

### Justification

Ces éléments concrétisent la valeur principale d'AGRIECO PRO : transformer les données de terrain en décision économique.

### Conséquences fonctionnelles

- L'utilisateur obtient des résultats exploitables après saisie.
- Les spéculations peuvent être comparées pour orienter les choix.
- Les rapports peuvent être partagés avec producteurs, coopératives ou financeurs.

### Conséquences techniques

- Prévoir des modules de calcul séparés des modules de saisie.
- Prévoir des tables de résultats calculés.
- Prévoir des modèles de rapports Excel imprimables.

### Risques éventuels

- Des calculs mal validés nuiraient à la crédibilité du logiciel.
- Les rapports doivent éviter une surcharge d'information.

### Points restant à confirmer

- Liste finale des indicateurs affichés dans le tableau de bord V1.
- Format exact des rapports imprimables.

---

# 10. Questions encore à documenter avant développement

## 10.1 Méthode d'affectation d'une charge partagée entre plusieurs spéculations

### Décision retenue

Le sujet doit être documenté avant développement. La V1 doit permettre d'affecter une charge partagée selon une clé explicite.

### Justification

Certains coûts, comme transport, matériel, main-d'œuvre, énergie ou bâtiment, peuvent servir à plusieurs spéculations.

### Conséquences fonctionnelles

- L'utilisateur devra choisir une méthode d'affectation.
- Le rapport devra indiquer les charges réparties et leur clé.

### Conséquences techniques

- Prévoir une table `T_AffectationsCharges`.
- Lier une charge source à plusieurs comptes ou spéculations.
- Stocker la clé utilisée : superficie, quantité produite, chiffre d'affaires, temps de travail, effectif, saisie manuelle.

### Risques éventuels

- Répartition arbitraire ou contestable.
- Double comptage si la charge source et ses affectations sont toutes incluses.

### Points restant à confirmer

- Méthodes d'affectation autorisées en V1.
- Méthode par défaut recommandée.

## 10.2 Gestion des unités et conversions

### Décision retenue

Le sujet doit être documenté. La V1 doit gérer les unités et conversions via un référentiel.

### Justification

Les données agricoles utilisent des unités variées : kg, tonne, sac, panier, hectare, are, tête, litre, jour-homme, bassin, cycle.

### Conséquences fonctionnelles

- L'utilisateur pourra saisir dans des unités courantes.
- Le logiciel convertira vers une unité de référence pour les calculs.

### Conséquences techniques

- Prévoir `T_Unites` et `T_ConversionsUnites`.
- Stocker l'unité saisie et l'unité de calcul.
- Interdire les conversions sans coefficient validé.

### Risques éventuels

- Les unités locales peuvent varier selon zone ou produit.
- Un sac de 50 kg et un sac de 100 kg ne doivent pas être confondus.

### Points restant à confirmer

- Liste des unités locales à intégrer.
- Coefficients de conversion par produit.

## 10.3 Traitement des achats à crédit

### Décision retenue

Le sujet doit être documenté. Les achats à crédit doivent être distingués des charges payées immédiatement.

### Justification

Un intrant reçu à crédit est une charge économique, mais son paiement peut intervenir plus tard.

### Conséquences fonctionnelles

- Le résultat économique inclut la charge au moment de son utilisation.
- La trésorerie tient compte du paiement différé.

### Conséquences techniques

- Prévoir des champs `ModePaiement`, `DatePaiementPrevue`, `DatePaiementEffective`, `MontantRestantDu`.
- Prévoir un suivi simple des dettes fournisseurs si retenu en V1.

### Risques éventuels

- Confusion entre charge de production et sortie de trésorerie.
- Complexité de suivi si les paiements sont fractionnés.

### Points restant à confirmer

- Gestion détaillée ou simplifiée des dettes fournisseurs en V1.

## 10.4 Traitement des paiements différés

### Décision retenue

Le sujet doit être documenté. Les paiements différés doivent affecter la trésorerie à leur date réelle ou prévue, sans modifier la date économique de consommation de la charge.

### Justification

La rentabilité et la trésorerie ne répondent pas à la même logique temporelle.

### Conséquences fonctionnelles

- Les tableaux de trésorerie afficheront les sorties selon les dates de paiement.
- Les comptes d'exploitation afficheront les charges selon leur consommation ou rattachement au cycle.

### Conséquences techniques

- Ajouter une dimension temporelle aux flux financiers.
- Prévoir une table d'échéances si nécessaire.

### Risques éventuels

- Saisie plus longue.
- Difficulté si l'utilisateur ne connaît pas les dates exactes.

### Points restant à confirmer

- Niveau de précision : date exacte, mois, période ou cycle.

## 10.5 Traitement des avances reçues

### Décision retenue

Le sujet doit être documenté. Une avance reçue ne doit pas être automatiquement considérée comme produit d'exploitation tant que la vente correspondante n'est pas réalisée ou reconnue.

### Justification

Une avance est souvent un financement ou un acompte, pas nécessairement un revenu définitif.

### Conséquences fonctionnelles

- Les avances doivent être visibles dans la trésorerie.
- Les ventes doivent être reconnues séparément.

### Conséquences techniques

- Prévoir une rubrique `AvancesRecues` ou `FluxTresorerie`.
- Lier éventuellement une avance à une vente future.

### Risques éventuels

- Double comptage si l'avance et la vente sont toutes deux comptées comme produit.

### Points restant à confirmer

- Besoin de gérer les avances acheteurs, coopératives ou contrats.

## 10.6 Traitement des subventions

### Décision retenue

Le sujet doit être documenté. Les subventions doivent être identifiées séparément des produits d'exploitation.

### Justification

Une subvention améliore les ressources du producteur mais ne provient pas directement de la production vendue.

### Conséquences fonctionnelles

- Les rapports doivent pouvoir afficher un résultat avant et après subvention si retenu.
- Les subventions ne doivent pas masquer la performance intrinsèque de la spéculation.

### Conséquences techniques

- Prévoir un type de flux `Subvention` séparé.
- Ne pas l'inclure automatiquement dans le produit brut agricole sans règle validée.

### Risques éventuels

- Les financeurs peuvent vouloir des traitements différents selon le type de subvention.

### Points restant à confirmer

- Subventions d'exploitation, d'investissement ou en nature : traitement différencié ou non.

## 10.7 Distinction campagne, exercice annuel et cycle de production

### Décision retenue

Le sujet doit être documenté. La V1 doit distinguer campagne, exercice annuel et cycle de production.

### Justification

Une campagne agricole, un exercice comptable annuel et un cycle technique ne se recouvrent pas toujours.

### Conséquences fonctionnelles

- L'utilisateur devra rattacher un compte à une campagne et à un cycle.
- Les rapports pourront présenter des résultats par cycle, par campagne et par année.

### Conséquences techniques

- Tables nécessaires : `T_Campagnes`, `T_Cycles`, éventuellement `T_Exercices`.
- Les dates de début et de fin doivent être explicites.

### Risques éventuels

- Mauvaise comparaison d'activités si les périodes ne sont pas harmonisées.

### Points restant à confirmer

- Définition officielle de la campagne agricole par spéculation et zone.

## 10.8 Gestion de plusieurs récoltes pour un même cycle

### Décision retenue

Le sujet doit être documenté. La V1 doit permettre plusieurs récoltes pour un même cycle lorsque la spéculation le nécessite.

### Justification

Certaines cultures maraîchères ou productions animales/aquacoles peuvent générer plusieurs récoltes ou sorties partielles.

### Conséquences fonctionnelles

- L'utilisateur pourra saisir plusieurs récoltes avec dates et quantités.
- Les ventes pourront être liées ou non à des récoltes spécifiques.

### Conséquences techniques

- Prévoir une table `T_Recoltes` ou une sous-catégorie dans `T_Produits`.
- Contrôler la cohérence entre récoltes, ventes, pertes, autoconsommation et stocks.

### Risques éventuels

- Complexité supplémentaire pour les cultures simples.

### Points restant à confirmer

- Spéculations pilotes concernées dès la V1.

## 10.9 Gestion de plusieurs ventes à différents prix

### Décision retenue

La V1 doit gérer plusieurs ventes à différents prix.

### Justification

Les ventes agricoles sont souvent fractionnées et réalisées à des prix variables selon la date, la qualité, le client ou le marché.

### Conséquences fonctionnelles

- L'utilisateur pourra saisir plusieurs lignes de vente.
- Le prix moyen pondéré pourra être calculé.

### Conséquences techniques

- `T_Produits` ou `T_Ventes` doit permettre plusieurs lignes par compte.
- Calcul du chiffre d'affaires total et du prix moyen pondéré.

### Risques éventuels

- Saisie plus longue si les ventes sont nombreuses.

### Points restant à confirmer

- Limite pratique du nombre de ventes par compte.
- Champs nécessaires : date, client, marché, qualité, prix, quantité.

## 10.10 Gestion des données manquantes

### Décision retenue

Le sujet doit être documenté. La V1 doit distinguer donnée absente, donnée non applicable et valeur zéro.

### Justification

Une cellule vide ne signifie pas nécessairement zéro. Les calculs économiques doivent éviter des interprétations erronées.

### Conséquences fonctionnelles

- Les formulaires devront signaler les données obligatoires manquantes.
- Les rapports devront éventuellement afficher les hypothèses manquantes.

### Conséquences techniques

- Prévoir des statuts de complétude.
- Bloquer la validation d'un compte si les données critiques sont absentes.
- Autoriser les brouillons incomplets.

### Risques éventuels

- Trop de blocages peuvent gêner la saisie terrain.
- Trop de tolérance peut produire des résultats faux.

### Points restant à confirmer

- Liste des champs obligatoires par module.
- Règles minimales pour calculer un résultat provisoire.

## 10.11 Traitement d'une production nulle

### Décision retenue

Le sujet doit être documenté. Une production nulle doit être autorisée mais traitée comme un cas d'alerte.

### Justification

Un échec de production peut se produire et doit être analysé économiquement.

### Conséquences fonctionnelles

- Le compte doit pouvoir être validé avec production nulle si les causes sont renseignées.
- Les indicateurs dépendant de la quantité produite doivent afficher une alerte plutôt qu'une division par zéro.

### Conséquences techniques

- Contrôler toutes les divisions par quantité.
- Prévoir un champ `CauseProductionNulle` ou une rubrique d'observation.

### Risques éventuels

- Coût de production unitaire impossible à calculer.

### Points restant à confirmer

- Liste des causes d'échec : climatique, sanitaire, inondation, sécheresse, ravageurs, abandon, autre.

## 10.12 Traitement d'une superficie ou d'un effectif nul

### Décision retenue

Le sujet doit être documenté. Une superficie ou un effectif nul ne doit pas être accepté pour un compte validé, sauf cas explicitement non applicable.

### Justification

Les indicateurs par hectare ou par effectif nécessitent une base physique positive.

### Conséquences fonctionnelles

- Les comptes brouillons peuvent contenir des valeurs manquantes.
- Les comptes validés doivent avoir une superficie ou un effectif cohérent selon le type de spéculation.

### Conséquences techniques

- Règles de validation différentes pour cultures et élevages.
- Contrôles anti-division par zéro.

### Risques éventuels

- Certains élevages peuvent être suivis en biomasse plutôt qu'en effectif.

### Points restant à confirmer

- Champs physiques obligatoires par spéculation.

## 10.13 Statuts des comptes

### Décision retenue

La V1 doit gérer des comptes brouillons, validés, corrigés et archivés.

### Justification

La saisie de terrain se fait souvent progressivement. Les corrections et l'archivage sont nécessaires pour la traçabilité.

### Conséquences fonctionnelles

- Un compte brouillon peut être incomplet.
- Un compte validé sert de référence pour rapports et comparaisons.
- Un compte corrigé conserve l'historique de correction.
- Un compte archivé reste consultable mais non modifiable.

### Conséquences techniques

- Champ `StatutCompte` obligatoire.
- Prévoir une table d'historique des statuts et corrections.
- Verrouiller les comptes validés ou archivés selon les droits.

### Risques éventuels

- Trop de statuts peuvent compliquer l'utilisation.

### Points restant à confirmer

- Qui peut valider, corriger ou archiver un compte ?

## 10.14 Historique des modifications

### Décision retenue

La V1 doit conserver l'historique des modifications importantes.

### Justification

Les résultats économiques peuvent servir à des décisions de conseil ou de financement ; la traçabilité est donc essentielle.

### Conséquences fonctionnelles

- Les utilisateurs autorisés peuvent consulter les changements majeurs.
- Les rapports peuvent indiquer la date de validation ou correction.

### Conséquences techniques

- Prévoir `LOG_Audit` et éventuellement `T_HistoriqueComptes`.
- Journaliser utilisateur, date, action, objet, ancienne valeur et nouvelle valeur pour les champs critiques.

### Risques éventuels

- Historique volumineux dans Excel.
- Protection limitée contre une modification directe des feuilles si le classeur est déverrouillé.

### Points restant à confirmer

- Niveau de détail de l'historique à conserver.
- Durée de conservation.

## 10.15 Capacité maximale envisagée pour la V1

### Décision retenue

Le sujet doit être confirmé avant développement. La V1 doit définir une capacité maximale réaliste pour rester performante sous Excel.

### Justification

Excel peut devenir lent si le nombre de producteurs, comptes et lignes devient trop élevé.

### Conséquences fonctionnelles

- Les utilisateurs doivent connaître les limites recommandées.
- Les organisations importantes devront prévoir une consolidation ou une version future SQL.

### Conséquences techniques

- Optimisation nécessaire des lectures/écritures en mémoire.
- Archivage ou séparation par fichier possible.

### Risques éventuels

- Dégradation des performances.
- Risque de corruption de classeur volumineux.

### Points restant à confirmer

- Nombre maximal recommandé de producteurs.
- Nombre maximal de comptes.
- Nombre maximal de lignes de charges/produits.

## 10.16 Analyse pluriannuelle des cultures pérennes

### Décision retenue

Le sujet doit être détaillé dans les règles de calcul. Les cultures pérennes doivent être analysées par phases et années.

### Justification

La rentabilité d'une culture pérenne dépend des investissements initiaux et de revenus étalés sur plusieurs années.

### Conséquences fonctionnelles

- Les rapports devront présenter une vision annuelle et cumulée.
- Les années sans production doivent être visibles.

### Conséquences techniques

- Lier plusieurs comptes annuels à un même projet pérenne.
- Prévoir une table de phases et une table de flux annuels.

### Risques éventuels

- Comparaison difficile avec les cycles courts sans indicateur annualisé.

### Points restant à confirmer

- Durée d'horizon de calcul.
- Taux d'actualisation ou non.
- Indicateurs pluriannuels de V1.

## 10.17 Besoin maximal de trésorerie

### Décision retenue

Le sujet doit être documenté. Le besoin maximal de trésorerie doit être calculé à partir du cumul temporel des sorties et entrées monétaires.

### Justification

Une spéculation rentable peut nécessiter une trésorerie initiale importante, ce qui influence la décision du producteur ou du financeur.

### Conséquences fonctionnelles

- Le rapport affichera le montant maximal à financer avant encaissement suffisant.
- L'IPE pourra intégrer ce besoin.

### Conséquences techniques

- Dates ou périodes obligatoires pour les flux de trésorerie.
- Calcul d'un solde cumulé par période.
- Identification du solde négatif maximal.

### Risques éventuels

- Les dates approximatives peuvent réduire la précision.
- Complexité accrue de saisie.

### Points restant à confirmer

- Granularité : jour, semaine, mois, cycle.
- Inclusion ou non des crédits, avances et subventions dans la trésorerie.

## 10.18 Méthode de notation du risque

### Décision retenue

Le sujet doit être validé. La notation du risque doit être paramétrable.

### Justification

Le risque dépend de facteurs climatiques, techniques, sanitaires, de marché, de prix, de conservation et de financement.

### Conséquences fonctionnelles

- L'utilisateur ou l'administrateur pourra noter le risque selon des critères.
- L'IPE utilisera un score de risque normalisé.

### Conséquences techniques

- Prévoir `T_Risques` et éventuellement `T_CriteresRisque`.
- Stocker les scores par critère et le score agrégé.

### Risques éventuels

- Notation subjective.
- Comparabilité limitée entre zones si les critères ne sont pas standardisés.

### Points restant à confirmer

- Échelle de notation : 1 à 5, 1 à 10 ou 0 à 100.
- Critères exacts de risque.

## 10.19 Méthode de notation du marché

### Décision retenue

Le sujet doit être validé. La disponibilité du marché doit être notée selon une méthode explicite et paramétrable.

### Justification

Une activité techniquement rentable peut être difficile à vendre si le marché est éloigné, saturé ou instable.

### Conséquences fonctionnelles

- L'IPE intégrera un score de marché.
- Les rapports pourront afficher une alerte de marché faible.

### Conséquences techniques

- Prévoir des critères : demande locale, accès au marché, volatilité des prix, conservation, nombre d'acheteurs, existence de contrats.
- Stocker les scores et pondérations.

### Risques éventuels

- Données de marché difficiles à obtenir.
- Scores rapidement obsolètes.

### Points restant à confirmer

- Critères exacts.
- Source des données de marché.

## 10.20 Seuils déclenchant les alertes économiques

### Décision retenue

Le sujet doit être validé. Les seuils d'alerte doivent être paramétrables.

### Justification

Les alertes aident les conseillers à identifier rapidement les activités déficitaires, risquées ou peu performantes.

### Conséquences fonctionnelles

- Le tableau de bord affichera des alertes simples.
- Les rapports pourront signaler les points critiques.

### Conséquences techniques

- Prévoir une table `T_Param_Alertes`.
- Associer chaque alerte à un indicateur, un seuil, une couleur et un message.

### Risques éventuels

- Trop d'alertes peuvent réduire la lisibilité.
- Des seuils mal définis peuvent générer de fausses alertes.

### Points restant à confirmer

- Liste des alertes V1.
- Seuils par défaut.

---

# 11. Décisions définitivement validées

Les décisions suivantes sont considérées comme arrêtées pour la V1 :

## 11.1 Identité et contexte

- Nom provisoire : **AGRIECO PRO**.
- Pays initial : **Côte d'Ivoire**.
- Langue initiale : **français**.
- Devise par défaut : **franc CFA — FCFA**.
- Fonctionnement entièrement hors connexion.

## 11.2 Environnement technique

- Système d'exploitation : **Windows uniquement pour la V1**.
- Version minimale : **Microsoft Excel 2016**.
- Interface mixte : UserForms pour les saisies ; feuilles Excel pour tableaux de bord, résultats et rapports.
- Stockage local dans des tableaux structurés Excel.
- Architecture compatible avec une migration future vers SQL, Web et Android.
- Exports prévus : PDF, Excel et CSV.
- Sauvegarde automatique locale, horodatée et restaurable.
- Système de licence reporté après validation du moteur économique.

## 11.3 Utilisateurs

- Cibles principales : entrepreneurs agricoles, producteurs agricoles structurés, conseillers agricoles, techniciens agricoles et agents d'encadrement.
- Cibles secondaires : coopératives, organisations professionnelles agricoles, cabinets de conseil, ONG, projets de développement, institutions financières, investisseurs agricoles, centres de formation et grandes institutions agricoles.
- Deux modes d'utilisation : mode individuel et mode conseiller/organisation, avec moteur économique commun.

## 11.4 Spéculations pilotes

- Cultures pilotes : maïs, riz, manioc, tomate, cacao.
- Élevages pilotes : poulet de chair, porc, pisciculture.
- Paramétrabilité requise pour ajouter d'autres cultures et élevages.

## 11.5 Décisions économiques

- Amortissement linéaire.
- Durée d'amortissement paramétrable.
- Autoconsommation valorisée, intégrée au produit brut et présentée séparément des ventes.
- Main-d'œuvre familiale valorisable.
- Deux résultats distincts : résultat financier et résultat économique.
- Distinction des charges réellement payées, valorisées, apports familiaux, charges fixes, charges variables, investissements, amortissements et frais financiers.
- Crédits reçus exclus des produits d'exploitation.
- Remboursement du capital de crédit exclu des charges de production.
- Intérêts, commissions et frais financiers présentés séparément.
- Produits secondaires, stocks, pertes et autoconsommations identifiables séparément.
- Pondérations IPE paramétrables et non codées définitivement dans le VBA.
- Pondérations IPE actuellement proposées considérées comme provisoires.
- Cultures pérennes analysées sur plusieurs années avec phases distinctes.

## 11.6 Formulaires et fonctions V1

- Pas de formulaire différent pour chaque culture ou élevage.
- Modèle générique de saisie pour cultures.
- Modèle générique de saisie pour élevages.
- Champs et rubriques spécifiques activés selon la spéculation.
- Référentiels configurables par spéculation.
- Fonctions prioritaires V1 telles que gestion producteurs, exploitations, campagnes, cycles, comptes prévisionnels, comptes réels, comparaison, charges, produits, ventes multiples, autoconsommation, stocks, pertes, main-d'œuvre familiale, amortissements, charges partagées, calculs, IPE, tableau de bord, rapports, exports, sauvegarde et restauration.

---

# 12. Décisions nécessitant encore une validation

Les décisions suivantes nécessitent encore une validation du propriétaire du projet avant développement :

## 12.1 Paramétrage métier

- Liste finale des champs obligatoires pour producteurs, exploitations, campagnes et comptes.
- Référentiels détaillés par spéculation pilote.
- Unités locales et coefficients de conversion.
- Calendriers agricoles par zone et spéculation.
- Niveau de détail attendu pour les élevages.

## 12.2 Règles économiques

- Barème de valorisation de la main-d'œuvre familiale.
- Prix de valorisation de l'autoconsommation.
- Traitement détaillé des stocks entrants et sortants.
- Traitement détaillé des pertes avant récolte et post-récolte.
- Gestion détaillée ou simplifiée des achats à crédit.
- Gestion des paiements différés.
- Traitement des avances reçues.
- Traitement des subventions.
- Présentation des crédits dans un tableau de trésorerie ou de financement.

## 12.3 Règles de calcul avancées

- Méthode d'affectation des charges partagées.
- Méthode de calcul du besoin maximal de trésorerie.
- Granularité temporelle des flux de trésorerie.
- Méthode d'analyse pluriannuelle des cultures pérennes.
- Utilisation ou non d'un taux d'actualisation.
- Indicateurs pluriannuels exacts à intégrer.

## 12.4 IPE et alertes

- Pondérations finales de l'IPE.
- Méthode de notation du risque.
- Méthode de notation du marché.
- Score minimal recommandé pour conseiller une spéculation.
- Classes finales de performance.
- Seuils déclenchant les alertes économiques.
- Droits nécessaires pour modifier les paramètres IPE et alertes.

## 12.5 Technique, UX et exploitation

- Capacité maximale envisagée pour la V1.
- Emplacement et politique de rétention des sauvegardes.
- Format exact des exports CSV.
- Modèles de rapports PDF.
- Charte graphique et ergonomie des formulaires.
- Rôles et droits détaillés.
- Politique de protection des feuilles et du classeur.

---

# 13. Recommandations de Codex

Ces recommandations sont proposées pour sécuriser la conception, mais elles ne sont pas automatiquement validées.

## 13.1 Recommandations fonctionnelles

1. **Prévoir deux modes de saisie** : un mode simplifié pour le terrain et un mode détaillé pour l'analyse approfondie.
2. **Afficher systématiquement les hypothèses utilisées** dans les rapports : prix, rendements, durée, cycles, taux de valorisation, amortissements et pondérations IPE.
3. **Séparer clairement résultat financier et résultat économique** avec une explication pédagogique dans les rapports.
4. **Prévoir des avertissements plutôt que des blocages excessifs** pendant la saisie, sauf pour les données indispensables au calcul.
5. **Créer des fiches de référence par spéculation pilote** avant développement pour stabiliser les champs activés.

## 13.2 Recommandations économiques

1. **Utiliser un prix moyen pondéré** pour les ventes multiples.
2. **Calculer le besoin maximal de trésorerie par solde cumulé périodique**, idéalement au mois pour simplifier la V1.
3. **Utiliser une clé d'affectation par défaut basée sur la superficie pour les cultures** et sur l'effectif ou la biomasse pour les élevages, tout en autorisant une clé manuelle.
4. **Conserver l'IPE comme indicateur d'aide à la décision et non comme verdict automatique**.
5. **Présenter les subventions séparément** afin de distinguer performance intrinsèque et appui externe.

## 13.3 Recommandations techniques

1. **Créer un dictionnaire de données définitif avant VBA**, avec nom, type, obligation, règle de validation et relation de chaque champ.
2. **Éviter toute formule métier cachée uniquement dans les cellules Excel** ; les règles doivent être documentées puis implémentées dans des modules de calcul.
3. **Utiliser des identifiants uniques stables** pour toutes les entités métier.
4. **Prévoir une table d'audit dès la V1**, même si elle reste simple.
5. **Limiter les dépendances Windows non standard** pour faciliter l'installation dans les organisations.
6. **Prévoir des jeux de données de test** pour maïs, riz, manioc, tomate, cacao, poulet de chair, porc et pisciculture.

## 13.4 Recommandations UX

1. **Utiliser un assistant de saisie en étapes** : producteur, exploitation, campagne, spéculation, charges, produits, résultats.
2. **Utiliser des codes couleur prudents** : vert pour rentable, orange pour vigilance, rouge pour déficitaire, bleu pour information.
3. **Limiter le nombre de champs visibles en même temps** grâce aux rubriques activées selon la spéculation.
4. **Prévoir une page de synthèse imprimable en une page** pour les producteurs et financeurs.
5. **Ajouter une zone d'observations/commentaires** dans les comptes pour contextualiser les résultats.

---

# 14. Prochaine étape recommandée

La prochaine étape recommandée est la création du document :

```text
REGLES_CALCUL_AGRIECO_PRO.md
```

Ce document devra détailler les règles de calcul officielles avant tout développement VBA, notamment :

- produit brut financier et économique ;
- charges monétaires et charges valorisées ;
- marges brute et nette ;
- bénéfice financier et économique ;
- amortissement linéaire ;
- main-d'œuvre familiale ;
- autoconsommation ;
- stocks et pertes ;
- ventes multiples ;
- crédits, intérêts, avances et subventions ;
- charges partagées ;
- cycles et annualisation ;
- cultures pérennes ;
- besoin de trésorerie ;
- IPE ;
- alertes économiques ;
- cas particuliers comme production nulle, données manquantes, superficie nulle ou effectif nul.

> Ne pas commencer `REGLES_CALCUL_AGRIECO_PRO.md` avant validation du présent document `DECISIONS_PRODUIT_V1_AGRIECO_PRO.md`.
