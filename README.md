# applocation-vba-agricole

Ce dépôt contient la conception fonctionnelle et technique de l'application professionnelle **AGRIECO PRO**, un outil Excel/VBA d'aide à la décision technico-économique pour les exploitations agricoles et d'élevage.

## Documentation

- [Architecture fonctionnelle et technique — AGRIECO PRO](ARCHITECTURE_AGRIECO_PRO.md)
- [Décisions produit V1 — AGRIECO PRO](DECISIONS_PRODUIT_V1_AGRIECO_PRO.md)
- [Règles de calcul — AGRIECO PRO](REGLES_CALCUL_AGRIECO_PRO.md)
- [Matrice des arbitrages métier — AGRIECO PRO](MATRICE_ARBITRAGES_METIER_AGRIECO_PRO.md)
- [Modèle conceptuel de données — AGRIECO PRO](MODELE_DONNEES_AGRIECO_PRO.md)
- [Architecture logicielle — AGRIECO PRO](ARCHITECTURE_LOGICIELLE_AGRIECO_PRO.md)
- [Sprint 1 — Installation et protocole de test](docs/SPRINT1_INSTALLATION_TESTS.md)
- [Sprint 2 — Gestion des producteurs et protocole de test](docs/SPRINT2_PRODUCTEURS_TESTS.md)
- [Sprint 2.1 — Stabilisation technique transverse](docs/SPRINT2_1_STABILISATION_TECHNIQUE.md)
- [Sprint 3 — Gestion des exploitations et protocole de test](docs/SPRINT3_EXPLOITATIONS_TESTS.md)
- [Guide unique — Installation locale et tests Excel](docs/GUIDE_UNIQUE_INSTALLATION_TEST_EXCEL.md)

## Classeur Sprint 1

Les fichiers Excel binaires ne sont pas versionnés dans la Pull Request. Ils doivent être générés localement depuis le dépôt afin d'éviter les limites de prise en charge des fichiers binaires.

- `scripts/generate_workbook.py` : script Python standard pour générer `AGRIECO_PRO_V1.xlsx` et `AGRIECO_PRO_V1.xlsm` localement.
- `AGRIECO_PRO_V1.xlsx` : classeur sans macros généré localement pour ouverture/test.
- `AGRIECO_PRO_V1.xlsm` : variante macro-enabled générée localement pour importer les modules VBA `.bas`.

Commande de génération :

```bash
python3 scripts/generate_workbook.py
```

Vérification statique du package VBA avant import Excel :

```bash
python3 scripts/verify_vba_package.py
```

## Installation Windows simplifiée

Pour générer le classeur sans connaissance technique sur Windows :

1. Télécharger et décompresser le dépôt GitHub.
2. Double-cliquer sur `build.bat`.
3. Ouvrir le classeur généré `AGRIECO_PRO_V1.xlsm`.

Guide détaillé : [Guide utilisateur — Générer le classeur AGRIECO PRO Sprint 1](docs/GUIDE_UTILISATEUR_INSTALLATION_SPRINT1.md).

## Sprint 2 — Gestion des producteurs

Le Sprint 2 ajoute le module VBA `modProducteurs` et le UserForm `frmProducteur` pour créer, modifier, supprimer avec contrôle des dépendances, rechercher et lister les producteurs. Les instructions d'importation et de test sont disponibles dans `docs/SPRINT2_PRODUCTEURS_TESTS.md`.

## Sprint 3 — Gestion des exploitations

Le Sprint 3 ajoute le module VBA `modExploitations` et le UserForm `frmExploitation` pour gérer les exploitations rattachées aux producteurs actifs. Les tests manuels sont décrits dans `docs/SPRINT3_EXPLOITATIONS_TESTS.md`.
