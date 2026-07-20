#!/usr/bin/env python3
"""Verify the AGRIECO PRO VBA package before local Excel import.

The check is intentionally static: it validates that the files, entry points and
UserForm dependencies required by Sprints 1, 2, 2.1 and 3 are present in the
repository. It does not execute Excel or VBA.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA_DIR = ROOT / "src" / "vba"

REQUIRED_FILES = [
    "modApp.bas",
    "modNavigation.bas",
    "modTables.bas",
    "modIdentifiants.bas",
    "modParametres.bas",
    "modReferentiels.bas",
    "modValidation.bas",
    "modErreurs.bas",
    "modJournal.bas",
    "modFormulaires.bas",
    "modProducteurs.bas",
    "modExploitations.bas",
    "frmProducteur.frm",
    "frmExploitation.frm",
]

REQUIRED_PROCEDURES = [
    "DemarrerApplication",
    "InitialiserApplication",
    "AllerAccueil",
    "OuvrirModule",
    "ObtenirTableau",
    "LigneDisponible",
    "LigneParIdentifiant",
    "ValeurLigne",
    "EcrireValeurLigne",
    "IdentifiantExiste",
    "NouvelIdentifiant",
    "NouvelIdentifiantUnique",
    "LireParametre",
    "ChargerListeDepuisReferentiel",
    "EstValeurActive",
    "EstTexteRenseigne",
    "EstNombrePositif",
    "EstDateValide",
    "SignalerErreur",
    "GererErreur",
    "Journaliser",
    "AppliquerEtatBoutonsCrud",
    "OuvrirGestionProducteurs",
    "AjouterProducteur",
    "ModifierProducteur",
    "SupprimerProducteur",
    "ChargerProducteursDansListe",
    "ChargerProducteurDansFormulaire",
    "OuvrirGestionExploitations",
    "ChargerProducteursActifsDansCombo",
    "AjouterExploitation",
    "ModifierExploitation",
    "SupprimerExploitation",
    "ChargerExploitationsDansListe",
    "ChargerExploitationDansFormulaire",
]

FORM_REQUIRED_CALLS = {
    "frmProducteur.frm": [
        "InitialiserTypesProducteur",
        "ViderFormulaire",
        "ChargerProducteursDansListe",
        "AjouterProducteur",
        "ModifierProducteur",
        "SupprimerProducteur",
        "ChargerProducteurDansFormulaire",
        "AppliquerEtatBoutonsCrud",
        "ChargerListeDepuisReferentiel",
    ],
    "frmExploitation.frm": [
        "ChargerProducteursActifsDansCombo",
        "ChargerListeDepuisReferentiel",
        "ViderFormulaire",
        "RechargerListeExploitations",
        "AppliquerEtatBoutonsCrud",
        "AjouterExploitation",
        "ModifierExploitation",
        "SupprimerExploitation",
        "ChargerExploitationsDansListe",
        "ChargerExploitationDansFormulaire",
        "ProducteurSelectionneID",
    ],
}

IMPORT_ORDER = [
    "modValidation.bas",
    "modTables.bas",
    "modIdentifiants.bas",
    "modJournal.bas",
    "modErreurs.bas",
    "modParametres.bas",
    "modReferentiels.bas",
    "modFormulaires.bas",
    "modNavigation.bas",
    "modApp.bas",
    "modProducteurs.bas",
    "modExploitations.bas",
    "frmProducteur.frm",
    "frmExploitation.frm",
]

PROC_RE = re.compile(
    r"^\s*(?:Public|Private)\s+(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE | re.MULTILINE,
)
USERFORM_OBJECT = 'Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"'
USERFORM_BEGIN_RE = re.compile(
    r"^Begin\s+\{C62A69F0-16A6-11CE-9E98-00AA00574A4F\}\s+([A-Za-z_][A-Za-z0-9_]*)\s*$",
    re.MULTILINE,
)
CONTROL_RE = re.compile(
    r"^\s*Begin\s+MSForms\.[A-Za-z0-9_]+\s+([A-Za-z_][A-Za-z0-9_]*)\s*$",
    re.MULTILINE,
)
EVENT_RE = re.compile(
    r"^\s*Private\s+Sub\s+([A-Za-z_][A-Za-z0-9_]*)_(?:Click|Change|Initialize|DblClick|Enter|Exit|KeyDown|KeyPress|KeyUp)\s*\(",
    re.IGNORECASE | re.MULTILINE,
)


def read_vba_files() -> dict[str, str]:
    return {path.name: path.read_text(encoding="utf-8") for path in VBA_DIR.glob("*.bas")} | {
        path.name: path.read_text(encoding="utf-8") for path in VBA_DIR.glob("*.frm")
    }


def collect_procedures(files: dict[str, str]) -> set[str]:
    procedures: set[str] = set()
    for content in files.values():
        procedures.update(match.group(1) for match in PROC_RE.finditer(content))
    return procedures


def validate_userforms(files: dict[str, str], errors: list[str]) -> None:
    for filename in ["frmProducteur.frm", "frmExploitation.frm"]:
        content = files.get(filename, "")
        if not content:
            continue

        if "Begin VB.UserForm" in content:
            errors.append(
                f"UserForm invalide importe comme module standard : {filename} contient Begin VB.UserForm"
            )

        if USERFORM_OBJECT not in content or not USERFORM_BEGIN_RE.search(content):
            errors.append(f"UserForm sans en-tete MSForms/FM20.DLL valide : {filename}")

        declared_controls = set(CONTROL_RE.findall(content))
        for source in EVENT_RE.findall(content):
            if source.lower() == "userform":
                continue
            if source not in declared_controls:
                errors.append(
                    f"Evenement rattache a un controle non declare : {filename} -> {source}"
                )

        if re.search(r"OleObjectBlob|Picture\s*=|Icon\s*=|\.frx", content, re.IGNORECASE):
            errors.append(f"Dependance binaire .frx probable a verifier : {filename}")


def main() -> int:
    errors: list[str] = []
    files = read_vba_files()
    procedures = collect_procedures(files)

    for filename in REQUIRED_FILES:
        if filename not in files:
            errors.append(f"Fichier VBA/Form manquant : {filename}")

    for filename in REQUIRED_FILES:
        if filename in files and "Option Explicit" not in files[filename]:
            errors.append(f"Option Explicit absent : {filename}")

    for procedure in REQUIRED_PROCEDURES:
        if procedure not in procedures:
            errors.append(f"Procedure attendue absente : {procedure}")

    for filename, calls in FORM_REQUIRED_CALLS.items():
        content = files.get(filename, "")
        for call in calls:
            if call not in content and call not in procedures:
                errors.append(f"Appel formulaire non resolu : {filename} -> {call}")

    validate_userforms(files, errors)

    print("Ordre recommande d'import VBA :")
    for index, filename in enumerate(IMPORT_ORDER, 1):
        print(f"{index:02d}. {filename}")

    if errors:
        print("\nECHEC verification package VBA :", file=sys.stderr)
        for error in errors:
            print(f"- {error}", file=sys.stderr)
        return 1

    print("\nVerification package VBA OK : fichiers, procedures et formulaires attendus presents.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
