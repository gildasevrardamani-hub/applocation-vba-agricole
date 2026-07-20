# Guide unique — Installation locale et tests Excel AGRIECO PRO

Ce guide explique comment générer et tester localement AGRIECO PRO sous **Windows + Microsoft Excel** après téléchargement du dépôt GitHub.

Il couvre les Sprints 1, 2, 2.1 et 3 : ossature du classeur, module Producteurs, services techniques transversaux et module Exploitations.

## 1. Public visé

Ce guide est destiné à un utilisateur non-développeur qui souhaite obtenir un classeur Excel testable sans modifier le code source.

## 2. Prérequis

- Windows 10 ou Windows 11.
- Microsoft Excel 2016 ou version plus récente.
- Python 3 installé et disponible dans le `PATH` Windows.
- Le dépôt AGRIECO PRO téléchargé et décompressé localement.
- Macros Excel autorisées pour le classeur généré.

## 3. Contenu VBA attendu dans le dépôt

Les fichiers suivants doivent être présents dans `src/vba` avant l'installation :

### Modules communs

1. `modValidation.bas`
2. `modTables.bas`
3. `modIdentifiants.bas`
4. `modJournal.bas`
5. `modErreurs.bas`
6. `modParametres.bas`
7. `modReferentiels.bas`
8. `modFormulaires.bas`
9. `modNavigation.bas`
10. `modApp.bas`

### Modules métier

11. `modProducteurs.bas`
12. `modExploitations.bas`

### Formulaires

13. `frmProducteur.frm`
14. `frmExploitation.frm`

Aucun fichier `.frx` n'est nécessaire pour les formulaires actuels, car ils ne contiennent pas d'image, d'icône ou d'objet binaire embarqué. Si un fichier `.frx` apparaît dans un sprint futur, il devra être conservé dans le même dossier que le fichier `.frm` correspondant avant import dans Excel.

## 4. Vérification automatique du package VBA

Avant d'ouvrir Excel, vous pouvez lancer une vérification statique depuis le dossier du dépôt :

```bash
python3 scripts/verify_vba_package.py
```

Sur Windows, si la commande `python3` n'est pas reconnue, essayer :

```bat
python scripts\verify_vba_package.py
```

Le script vérifie :

- la présence de tous les modules `.bas` et formulaires `.frm` nécessaires ;
- la présence de `Option Explicit` ;
- la présence des procédures appelées par les formulaires et la navigation ;
- l'absence de dépendance binaire `.frx` pour les formulaires actuels ;
- l'ordre recommandé d'importation dans Excel.

## 5. Générer le classeur AGRIECO PRO

### Méthode simple sous Windows

1. Ouvrir le dossier du dépôt.
2. Double-cliquer sur `build.bat`.
3. Attendre le message confirmant la création des fichiers.
4. Vérifier que les fichiers suivants sont créés à la racine du dépôt :
   - `AGRIECO_PRO_V1.xlsx` ;
   - `AGRIECO_PRO_V1.xlsm`.

### Méthode par terminal

Depuis le dossier du dépôt :

```bash
python3 scripts/generate_workbook.py
```

ou sous Windows :

```bat
python scripts\generate_workbook.py
```

## 6. Ouvrir et enregistrer le classeur dans Excel

1. Ouvrir `AGRIECO_PRO_V1.xlsm` avec Microsoft Excel.
2. Si Excel affiche un bandeau de sécurité, cliquer sur **Activer le contenu**.
3. Enregistrer immédiatement le fichier avec `Fichier > Enregistrer sous`.
4. Vérifier que le format sélectionné est **Classeur Excel prenant en charge les macros (*.xlsm)**.

## 7. Activer l'accès au projet VBA si nécessaire

Cette étape n'est généralement pas nécessaire pour un import manuel. Elle peut être utile si Excel bloque certaines opérations VBA.

1. Ouvrir Excel.
2. Aller dans `Fichier > Options`.
3. Ouvrir `Centre de gestion de la confidentialité`.
4. Cliquer sur `Paramètres du Centre de gestion de la confidentialité`.
5. Aller dans `Paramètres des macros`.
6. Cocher **Faire confiance à l'accès au modèle d'objet du projet VBA** si l'environnement l'exige.
7. Valider, fermer puis rouvrir Excel.

## 8. Importer les modules VBA

1. Dans Excel, appuyer sur `ALT + F11` pour ouvrir l'éditeur VBA.
2. Dans l'explorateur de projet, sélectionner le classeur `AGRIECO_PRO_V1.xlsm`.
3. Pour chaque fichier `.bas`, faire `Fichier > Importer un fichier...`.
4. Importer les modules dans l'ordre recommandé :
   1. `modValidation.bas`
   2. `modTables.bas`
   3. `modIdentifiants.bas`
   4. `modJournal.bas`
   5. `modErreurs.bas`
   6. `modParametres.bas`
   7. `modReferentiels.bas`
   8. `modFormulaires.bas`
   9. `modNavigation.bas`
   10. `modApp.bas`
   11. `modProducteurs.bas`
   12. `modExploitations.bas`

## 9. Importer les UserForms

Toujours dans l'éditeur VBA :

1. Faire `Fichier > Importer un fichier...`.
2. Importer `frmProducteur.frm`.
3. Importer `frmExploitation.frm`.
4. Vérifier que les deux formulaires apparaissent dans l'explorateur de projet.

Pour les Sprints 1 à 3, aucun fichier `.frx` n'est attendu. Si Excel demande un fichier `.frx`, vérifier que le dépôt est complet et que le fichier `.frm` n'a pas été modifié manuellement.

## 10. Enregistrer après import

1. Revenir dans Excel.
2. Enregistrer le classeur avec `CTRL + S`.
3. Fermer puis rouvrir le classeur pour confirmer que les macros et formulaires sont conservés.

## 11. Lancer AGRIECO PRO

### Démarrage général

Dans Excel :

1. Appuyer sur `ALT + F8`.
2. Sélectionner `DemarrerApplication`.
3. Cliquer sur **Exécuter**.

### Gestion des Producteurs

1. Appuyer sur `ALT + F8`.
2. Sélectionner `OuvrirGestionProducteurs`.
3. Cliquer sur **Exécuter**.

### Gestion des Exploitations

1. Créer au moins un producteur actif dans le module Producteurs.
2. Appuyer sur `ALT + F8`.
3. Sélectionner `OuvrirGestionExploitations`.
4. Cliquer sur **Exécuter**.

## 12. Checklist de validation fonctionnelle minimale

### Producteurs

- [ ] Le formulaire Producteurs s'ouvre sans erreur.
- [ ] Le bouton **Nouveau** vide le formulaire.
- [ ] Un producteur peut être créé avec nom, localisation et type.
- [ ] L'identifiant généré commence par `PROD-`.
- [ ] La ligne apparaît dans `tbl_PRODUCTEURS`.
- [ ] Une création apparaît dans `JOURNAL`.
- [ ] Une recherche par nom, code ou contact filtre la liste.
- [ ] Un producteur sélectionné peut être modifié.
- [ ] Une suppression sans exploitation rattachée fonctionne après confirmation.
- [ ] Une suppression avec exploitation rattachée est bloquée.

### Exploitations

- [ ] Le formulaire Exploitations s'ouvre sans erreur.
- [ ] La liste Producteur contient les producteurs actifs.
- [ ] Les listes Type d'exploitation et Unité de superficie sont alimentées depuis `REFERENTIELS`.
- [ ] Une exploitation peut être créée pour un producteur actif.
- [ ] L'identifiant généré commence par `EXP-`.
- [ ] La ligne apparaît dans `tbl_EXPLOITATIONS`.
- [ ] Une exploitation peut être recherchée par code, nom ou localité.
- [ ] Les exploitations sont filtrées selon le producteur sélectionné.
- [ ] Une exploitation sélectionnée peut être modifiée.
- [ ] Une suppression sans parcelle, atelier ou cycle rattaché fonctionne après confirmation.
- [ ] Une suppression avec parcelle, atelier ou cycle rattaché est bloquée.
- [ ] Les créations, modifications et suppressions alimentent `JOURNAL` avec entité et identifiant.

## 13. Erreurs fréquentes et solutions

| Problème | Cause probable | Solution |
|---|---|---|
| `python` ou `python3` n'est pas reconnu | Python n'est pas installé ou pas dans le `PATH`. | Installer Python 3 depuis python.org et cocher l'option d'ajout au `PATH`, puis relancer `build.bat`. |
| Le fichier `.xlsm` n'existe pas | Le générateur n'a pas été exécuté ou a échoué. | Relancer `build.bat` et lire le message d'erreur affiché. |
| Excel bloque les macros | Le fichier provient d'Internet ou les macros sont désactivées. | Débloquer le fichier dans les propriétés Windows, puis activer le contenu dans Excel. |
| Les macros n'apparaissent pas dans `ALT + F8` | Les modules `.bas` n'ont pas été importés dans le bon classeur. | Vérifier que le projet sélectionné est `AGRIECO_PRO_V1.xlsm` et réimporter les modules. |
| Erreur `Sub ou Function non définie` | Un module dépendant n'a pas été importé. | Importer tous les modules dans l'ordre recommandé, puis enregistrer. |
| Le formulaire ne s'ouvre pas | Le fichier `.frm` n'a pas été importé ou a été importé dans un autre classeur. | Importer `frmProducteur.frm` et `frmExploitation.frm` dans `AGRIECO_PRO_V1.xlsm`. |
| Les listes Type ou Unité sont vides | La feuille `REFERENTIELS` ne contient pas les lignes générées ou le classeur est ancien. | Regénérer le classeur avec le script actuel ou compléter `REFERENTIELS`. |
| La suppression d'une exploitation est refusée | Une parcelle, un atelier ou un cycle est rattaché. | Supprimer d'abord la dépendance de test ou conserver l'exploitation. |
| Le journal ne se remplit pas | La table `tbl_JOURNAL` est absente ou le module `modJournal` n'est pas importé. | Regénérer le classeur et réimporter `modJournal.bas`. |

## 14. Vérification finale du package

Le package est prêt pour test Excel lorsque :

- [ ] `scripts/verify_vba_package.py` indique que la vérification est OK ;
- [ ] le classeur `.xlsm` est généré ;
- [ ] tous les modules `.bas` sont importés ;
- [ ] les deux formulaires `.frm` sont importés ;
- [ ] `DemarrerApplication` fonctionne ;
- [ ] `OuvrirGestionProducteurs` fonctionne ;
- [ ] `OuvrirGestionExploitations` fonctionne ;
- [ ] les tests essentiels Producteurs et Exploitations sont validés.

## 15. Limites actuelles

- Les fichiers Excel binaires ne sont pas versionnés dans GitHub ; ils sont générés localement.
- Les tests VBA automatisés ne sont pas encore disponibles.
- Les contrôles Codex sont statiques et ne remplacent pas une exécution réelle sous Microsoft Excel.
- Le Sprint 4 n'est pas inclus dans ce package.
