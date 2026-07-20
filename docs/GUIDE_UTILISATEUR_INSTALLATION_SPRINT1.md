# Guide utilisateur — Générer le classeur AGRIECO PRO Sprint 1

Ce guide s'adresse à un utilisateur Windows non développeur.

## Objectif

Créer localement le classeur Excel AGRIECO PRO à partir du dépôt téléchargé depuis GitHub.

Le dépôt ne contient pas directement les fichiers Excel binaires afin d'éviter les blocages de Pull Request. Le classeur est généré automatiquement sur votre ordinateur.

## Prérequis

- Un ordinateur Windows.
- Microsoft Excel 2016 ou supérieur.
- Python installé sur Windows.

Si Python n'est pas installé :

1. Aller sur https://www.python.org/downloads/windows/.
2. Télécharger la dernière version stable de Python.
3. Lancer l'installation.
4. Cocher l'option **Add python.exe to PATH**.
5. Terminer l'installation.

## Étapes simples

1. Télécharger le dépôt GitHub sous forme de fichier ZIP.
2. Décompresser le ZIP dans un dossier local, par exemple `Documents\AGRIECO_PRO`.
3. Ouvrir le dossier décompressé.
4. Double-cliquer sur le fichier `build.bat`.
5. Attendre le message de confirmation.

Si tout se passe bien, deux fichiers sont créés dans le dossier :

- `AGRIECO_PRO_V1.xlsx` ;
- `AGRIECO_PRO_V1.xlsm`.

## Quel fichier ouvrir ?

Ouvrez en priorité :

```text
AGRIECO_PRO_V1.xlsm
```

Ce fichier est destiné à recevoir les modules VBA.

Si Excel bloque le fichier `.xlsm`, ouvrez d'abord :

```text
AGRIECO_PRO_V1.xlsx
```

puis enregistrez-le depuis Excel au format **Classeur Excel prenant en charge les macros (*.xlsm)**.

## Importer les modules VBA

Cette étape est destinée au test Sprint 1.

1. Ouvrir `AGRIECO_PRO_V1.xlsm` dans Excel.
2. Appuyer sur `ALT + F11` pour ouvrir l'éditeur VBA.
3. Dans le menu, choisir `File > Import File`.
4. Importer un par un les fichiers `.bas` du dossier `src\vba`.
5. Enregistrer le classeur.
6. Lancer la macro `DemarrerApplication`.

## Message de réussite attendu

À la fin de `build.bat`, vous devez voir un message indiquant que les classeurs ont été créés avec succès.

## En cas de problème

### Python n'est pas détecté

Réinstallez Python et cochez l'option **Add python.exe to PATH**.

### Le fichier generate_workbook.py est introuvable

Vérifiez que vous avez décompressé tout le dépôt GitHub et que le dossier `scripts` est présent.

### Les fichiers Excel ne sont pas créés

Relancez `build.bat` et lisez le message d'erreur affiché dans la fenêtre.

## Limites du Sprint 1

- Le moteur économique complet n'est pas encore développé.
- L'IPE, l'IQD, les simulations et les rapports avancés ne sont pas encore développés.
- Le classeur généré est une ossature technique permettant de tester la structure de départ.
