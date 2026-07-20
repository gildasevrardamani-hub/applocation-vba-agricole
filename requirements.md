# Prérequis d'installation — AGRIECO PRO V1

Ce document liste les prérequis nécessaires pour générer et tester localement l'ossature Sprint 1 du classeur AGRIECO PRO.

## 1. Système d'exploitation

- Windows 10 ou Windows 11 recommandé.
- Windows est requis pour l'utilisation cible de Microsoft Excel VBA.
- Les scripts de génération peuvent être lus sur d'autres systèmes, mais le test utilisateur final doit être réalisé sous Windows avec Excel.

## 2. Microsoft Excel

- Microsoft Excel 2016 ou version supérieure.
- Excel doit pouvoir ouvrir les classeurs `.xlsx` et `.xlsm`.
- Pour tester les modules VBA, Excel doit autoriser l'ouverture de classeurs prenant en charge les macros.

## 3. Macros VBA

Pour importer et tester les modules VBA :

1. Ouvrir le fichier généré `AGRIECO_PRO_V1.xlsm`.
2. Activer les macros si Excel affiche une barre d'avertissement de sécurité.
3. Ouvrir l'éditeur VBA avec `ALT + F11`.
4. Importer les fichiers `.bas` situés dans `src/vba`.
5. Enregistrer le classeur au format `.xlsm`.

> Important : les macros peuvent être bloquées par la politique de sécurité de Windows, d'Office ou de l'organisation. Dans ce cas, demander à l'administrateur informatique d'autoriser le fichier ou son emplacement.

## 4. Python

Python est nécessaire uniquement pour générer localement les classeurs Excel à partir du dépôt.

### Version recommandée

- Python 3.10 ou supérieur.

### Installation sous Windows

1. Télécharger Python depuis https://www.python.org/downloads/windows/.
2. Lancer l'installation.
3. Cocher l'option **Add python.exe to PATH**.
4. Terminer l'installation.
5. Fermer puis rouvrir l'invite de commandes si elle était déjà ouverte.

### Vérification de Python

Dans une invite de commandes Windows, exécuter :

```bat
python --version
```

Si cette commande ne fonctionne pas, essayer :

```bat
py --version
```

Au moins une des deux commandes doit afficher une version de Python 3.

## 5. Génération du classeur

Après téléchargement et décompression du dépôt GitHub, deux méthodes sont possibles.

### Méthode recommandée pour utilisateur Windows non technique

1. Ouvrir le dossier du projet.
2. Double-cliquer sur `build.bat`.
3. Attendre le message de confirmation.

Le script vérifie la présence de Python, exécute `scripts/generate_workbook.py` et confirme la création des classeurs.

### Méthode en ligne de commande

Depuis la racine du projet, exécuter :

```bat
python scripts\generate_workbook.py
```

ou :

```bat
py scripts\generate_workbook.py
```

## 6. Fichiers générés attendus

Après génération, les fichiers suivants doivent apparaître à la racine du projet :

- `AGRIECO_PRO_V1.xlsx` ;
- `AGRIECO_PRO_V1.xlsm`.

Ces fichiers ne sont pas versionnés dans Git afin d'éviter les blocages liés aux fichiers binaires dans les Pull Requests.

## 7. Vérification de l'environnement

Avant de commencer le test Sprint 1, vérifier que :

- le dépôt contient le dossier `scripts` ;
- le fichier `scripts/generate_workbook.py` est présent ;
- le dossier `src/vba` contient les modules `.bas` ;
- Python est disponible via `python --version` ou `py --version` ;
- `build.bat` se lance sans erreur ;
- les fichiers `AGRIECO_PRO_V1.xlsx` et `AGRIECO_PRO_V1.xlsm` sont créés ;
- Microsoft Excel peut ouvrir au moins le fichier `.xlsx` ;
- Microsoft Excel peut ouvrir le fichier `.xlsm` si les macros sont autorisées.

## 8. Dépendances Python

Le script de génération utilise uniquement la bibliothèque standard Python.

Aucune installation de paquet externe n'est requise :

```bat
pip install ...
```

n'est pas nécessaire pour le Sprint 1.

## 9. Problèmes fréquents

### Python n'est pas reconnu

Réinstaller Python en cochant **Add python.exe to PATH**, puis relancer `build.bat`.

### Le script generate_workbook.py est introuvable

Vérifier que le dépôt GitHub a été téléchargé entièrement et que le dossier `scripts` est présent.

### Excel bloque les macros

Ouvrir d'abord `AGRIECO_PRO_V1.xlsx` pour vérifier le classeur. Pour tester VBA, enregistrer le fichier au format `.xlsm` et autoriser les macros selon la politique de sécurité de votre environnement.

### Les classeurs ne sont pas créés

Relancer `build.bat` et lire le message affiché. Si l'erreur persiste, exécuter manuellement :

```bat
python scripts\generate_workbook.py
```

pour voir le message d'erreur détaillé.

## 10. Limites du Sprint 1

- Le moteur économique complet n'est pas encore développé.
- L'IPE, l'IQD, les simulations et les rapports avancés ne sont pas encore développés.
- Le classeur généré est une ossature technique destinée à tester les feuilles, tableaux structurés et modules VBA de base.
