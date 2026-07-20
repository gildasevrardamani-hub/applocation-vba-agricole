# Sprint 2 — Gestion des producteurs

Ce document décrit l'installation et le protocole de test du module **Gestion des producteurs** d'AGRIECO PRO.

## Objectif du sprint

Le Sprint 2 rend utilisable le premier module métier du classeur : création, modification, suppression contrôlée, recherche et affichage des producteurs. Les données sont enregistrées dans la feuille `PRODUCTEURS`, dans le tableau structuré généré par le classeur Sprint 1.

## Fichiers livrés

- `src/vba/modProducteurs.bas` : fonctions métier du module Producteurs.
- `src/vba/frmProducteur.frm` : formulaire de saisie, recherche et gestion des producteurs.
- `src/vba/modApp.bas` : version applicative portée à `V1-Sprint2`.

## Installation dans Excel

1. Générer le classeur avec `build.bat` ou avec la commande `python3 scripts/generate_workbook.py`.
2. Ouvrir `AGRIECO_PRO_V1.xlsm` dans Microsoft Excel.
3. Activer les macros si Excel affiche un avertissement de sécurité.
4. Ouvrir l'éditeur VBA avec `ALT + F11`.
5. Importer tous les fichiers `.bas` du dossier `src/vba`.
6. Importer le formulaire `src/vba/frmProducteur.frm`.
7. Enregistrer le classeur au format `.xlsm`.
8. Lancer la macro `OuvrirGestionProducteurs` pour ouvrir le formulaire Producteurs.

## Protocole de test fonctionnel

### 1. Ouverture du formulaire

- Exécuter `OuvrirGestionProducteurs`.
- Vérifier que le formulaire **AGRIECO PRO - Gestion des producteurs** s'affiche.
- Vérifier que la liste est vide si aucun producteur n'a encore été créé.

### 2. Création d'un producteur

- Cliquer sur **Nouveau**.
- Saisir au minimum :
  - `Nom` : `Producteur Test` ;
  - `Localisation` : `Bouaké` ;
  - `Type` : `Entrepreneur agricole`.
- Cliquer sur **Enregistrer**.
- Vérifier qu'un identifiant commençant par `PROD-` est généré.
- Vérifier que la ligne est ajoutée dans le tableau de la feuille `PRODUCTEURS`.
- Vérifier qu'une action `Creation` est ajoutée dans la feuille `JOURNAL`.

### 3. Validation des champs obligatoires

- Cliquer sur **Nouveau**.
- Laisser le champ `Nom` vide.
- Cliquer sur **Enregistrer**.
- Vérifier qu'un message d'erreur indique que le nom est obligatoire.
- Vérifier qu'aucune ligne incomplète n'est enregistrée dans `PRODUCTEURS`.

### 4. Recherche d'un producteur

- Créer deux producteurs avec des noms ou contacts différents.
- Saisir une partie du nom, du code ou du contact dans la zone de recherche.
- Cliquer sur **Rechercher**.
- Vérifier que la liste affiche uniquement les producteurs correspondants.

### 5. Modification d'un producteur

- Sélectionner un producteur dans la liste.
- Modifier le contact ou la localisation.
- Cliquer sur **Modifier**.
- Vérifier que la ligne correspondante est mise à jour dans `PRODUCTEURS`.
- Vérifier qu'une action `Modification` est ajoutée dans `JOURNAL`.

### 6. Suppression sans dépendance

- Sélectionner un producteur sans exploitation rattachée.
- Cliquer sur **Supprimer** et confirmer.
- Vérifier que la ligne est supprimée de `PRODUCTEURS`.
- Vérifier qu'une action `Suppression` est ajoutée dans `JOURNAL`.

### 7. Suppression avec dépendance

- Créer un producteur.
- Dans `EXPLOITATIONS`, ajouter manuellement une ligne dont `ProducteurID` correspond à ce producteur.
- Revenir au formulaire Producteurs, sélectionner le producteur et cliquer sur **Supprimer**.
- Vérifier que la suppression est bloquée.
- Vérifier qu'un message indique qu'une exploitation est rattachée au producteur.
- Vérifier qu'une action `Erreur` est ajoutée dans `JOURNAL`.

## Limites volontairement conservées pour le Sprint 2

- Le module ne crée pas encore d'exploitation, de parcelle, de cycle ou de compte d'exploitation.
- Le formulaire Producteurs utilise une liste simple et ne constitue pas encore l'interface finale complète.
- Les référentiels de types de producteurs sont initialisés dans le formulaire pour ce sprint ; ils pourront être externalisés dans `REFERENTIELS` lors d'un sprint de paramétrage avancé.
