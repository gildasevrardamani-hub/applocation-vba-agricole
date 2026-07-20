# Sprint 1 — Installation et protocole de test

## Livrables

- `scripts/generate_workbook.py` : script standard Python permettant de générer localement les fichiers `.xlsx` et `.xlsm`.
- `AGRIECO_PRO_V1.xlsx` : classeur sans macros généré localement pour ouverture/test.
- `AGRIECO_PRO_V1.xlsm` : classeur macro-enabled généré localement pour importer les modules VBA `.bas`.

> Les fichiers Excel binaires ne sont pas stockés dans la Pull Request afin d'éviter le blocage « Les fichiers binaires ne sont pas pris en charge ».
- `src/vba/modApp.bas` : démarrage de l'application.
- `src/vba/modNavigation.bas` : navigation simple entre feuilles.
- `src/vba/modIdentifiants.bas` : génération d'identifiants.
- `src/vba/modValidation.bas` : validations communes.
- `src/vba/modErreurs.bas` : signalement centralisé des erreurs.
- `src/vba/modJournal.bas` : journalisation.
- `src/vba/modParametres.bas` : lecture des paramètres.

## Génération locale du classeur

### Option recommandée sur Windows

Double-cliquer sur `build.bat`, puis suivre les messages affichés. Ce script vérifie la présence de Python, exécute `scripts/generate_workbook.py` et confirme la création des classeurs.

### Option en ligne de commande

Depuis la racine du dépôt, exécuter :

```bash
python3 scripts/generate_workbook.py
```

Cette commande crée localement `AGRIECO_PRO_V1.xlsx` et `AGRIECO_PRO_V1.xlsm` avec les feuilles et tableaux structurés du Sprint 1.

## Importation manuelle des modules VBA

1. Générer le classeur avec `python3 scripts/generate_workbook.py`.
2. Ouvrir `AGRIECO_PRO_V1.xlsm` dans Microsoft Excel 2016 ou supérieur sur Windows. Si Excel bloque le fichier `.xlsm`, ouvrir `AGRIECO_PRO_V1.xlsx`, puis l'enregistrer localement au format `.xlsm` avant import des modules VBA.
3. Ouvrir l'éditeur VBA avec `ALT + F11`.
4. Importer chaque fichier `.bas` du dossier `src/vba` via `File > Import File`.
5. Enregistrer le classeur au format `.xlsm`.
6. Lancer la macro `DemarrerApplication` pour tester le démarrage et la navigation vers `ACCUEIL`.

## Protocole de test simple

1. Générer localement le classeur avec `python3 scripts/generate_workbook.py`.
2. Ouvrir le classeur et vérifier la présence des feuilles attendues.
3. Vérifier que chaque feuille technique, sauf `ACCUEIL`, contient un tableau structuré nommé `tbl_NOMFEUILLE`.
4. Importer les modules VBA.
5. Exécuter `DemarrerApplication`.
6. Vérifier qu'une ligne est ajoutée dans le journal après démarrage.
7. Exécuter `OuvrirModule "PRODUCTEURS"` depuis la fenêtre d'exécution VBA.
8. Vérifier que la feuille `PRODUCTEURS` devient active.
9. Exécuter `NouvelIdentifiant("TEST")` et vérifier qu'un identifiant non vide est retourné.

## Limites connues du Sprint 1

- Le moteur économique complet n'est pas développé.
- L'IPE, l'IQD, les simulations et rapports avancés ne sont pas développés.
- Les modules VBA sont fournis comme fichiers importables ; l'environnement d'exécution automatisé ne permet pas d'injecter un projet VBA signé directement dans le classeur.
- Le classeur est généré localement comme ossature fonctionnelle de départ destinée à être ouverte et finalisée dans Excel sur Windows.
