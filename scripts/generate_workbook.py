#!/usr/bin/env python3
"""Generate the AGRIECO PRO Excel workbook skeleton for Sprint 1 and Sprint 2.

This script uses only the Python standard library. It creates a real OpenXML
workbook with the validated Sprint 1 worksheets, structured tables and Sprint 2 entry points. By default it
writes both:
- AGRIECO_PRO_V1.xlsx: workbook skeleton for local opening/testing.
- AGRIECO_PRO_V1.xlsm: macro-enabled extension variant for importing VBA modules.

The VBA modules and UserForms remain versioned separately in src/vba/ and can be imported
manually into the workbook from the Excel VBA editor.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

SHEETS = [
    "ACCUEIL",
    "PARAMETRES",
    "UTILISATEURS",
    "PRODUCTEURS",
    "EXPLOITATIONS",
    "PARCELLES_ATELIERS",
    "SPECULATIONS",
    "CYCLES",
    "CHARGES",
    "PRODUITS",
    "RECOLTES",
    "VENTES",
    "TRESORERIE",
    "RESULTATS",
    "JOURNAL",
    "REFERENTIELS",
]

HEADERS = {
    "PARAMETRES": ["ParametreID", "Domaine", "Cle", "Valeur", "TypeValeur", "Description", "Actif", "DateMaj"],
    "UTILISATEURS": ["UtilisateurID", "NomUtilisateur", "NomComplet", "Role", "ModeUtilisation", "Actif", "DateCreation"],
    "PRODUCTEURS": ["ProducteurID", "OrganisationID", "CodeProducteur", "NomProducteur", "Contact", "Localisation", "TypeProducteur", "Actif", "DateCreation"],
    "EXPLOITATIONS": ["ExploitationID", "ProducteurID", "CodeExploitation", "NomExploitation", "TypeExploitation", "Localite", "VillageQuartier", "SuperficieTotale", "UniteSuperficie", "CoordonneesGPS", "Actif", "Observations", "DateCreation"],
    "PARCELLES_ATELIERS": ["ParcelleAtelierID", "ExploitationID", "TypeUnite", "Nom", "SuperficieOuCapacite", "Unite", "Statut", "DateCreation"],
    "SPECULATIONS": ["SpeculationID", "NomSpeculation", "TypeSpeculation", "Categorie", "UniteProduction", "DureeCycleDefaut", "CyclesParAnDefaut", "Actif"],
    "CYCLES": ["CycleID", "ExploitationID", "ParcelleAtelierID", "SpeculationID", "TypeCompte", "DateDebut", "DateFin", "StatutCycle", "NiveauSaisie"],
    "CHARGES": ["ChargeID", "CycleID", "Operation", "CategorieCharge", "NatureCharge", "TypePaiement", "Quantite", "Unite", "CoutUnitaire", "Montant", "DateCharge"],
    "PRODUITS": ["ProduitID", "CycleID", "NatureProduit", "Designation", "Quantite", "Unite", "PrixUnitaire", "Montant", "DateProduit"],
    "RECOLTES": ["RecolteID", "CycleID", "DateRecolte", "QuantiteRecoltee", "Unite", "Qualite", "PertesAssociees", "Observation"],
    "VENTES": ["VenteID", "CycleID", "Client", "DateVente", "QuantiteVendue", "Unite", "PrixUnitaire", "MontantVente", "StatutEncaissement"],
    "TRESORERIE": ["FluxTresorerieID", "CycleID", "NatureFlux", "CategorieFlux", "Montant", "DateFlux", "StatutFlux", "Observation"],
    "RESULTATS": ["ResultatID", "CompteID", "CycleID", "NomIndicateur", "Valeur", "Unite", "StatutCalcul", "DateCalcul", "VersionRegles"],
    "JOURNAL": ["JournalID", "DateAction", "Utilisateur", "Module", "Action", "Entite", "IdentifiantEntite", "Message", "Niveau", "AncienEtat", "NouvelEtat", "DateHeure"],
    "REFERENTIELS": ["ReferentielID", "Domaine", "Code", "Libelle", "Valeur1", "Valeur2", "Actif", "OrdreAffichage"],
}

DEFAULT_ROWS = {
    "REFERENTIELS": [
        ["REF-TYPE-PROD-001", "TYPE_PRODUCTEUR", "ENTREPRENEUR_AGRICOLE", "Entrepreneur agricole", "", "", "TRUE", "1"],
        ["REF-TYPE-PROD-002", "TYPE_PRODUCTEUR", "PRODUCTEUR_STRUCTURE", "Producteur structure", "", "", "TRUE", "2"],
        ["REF-TYPE-PROD-003", "TYPE_PRODUCTEUR", "BENEFICIAIRE_ACCOMPAGNE", "Beneficiaire accompagne", "", "", "TRUE", "3"],
        ["REF-TYPE-EXP-001", "TYPE_EXPLOITATION", "INDIVIDUELLE", "Individuelle", "", "", "TRUE", "1"],
        ["REF-TYPE-EXP-002", "TYPE_EXPLOITATION", "FAMILIALE", "Familiale", "", "", "TRUE", "2"],
        ["REF-TYPE-EXP-003", "TYPE_EXPLOITATION", "COOPERATIVE", "Cooperative", "", "", "TRUE", "3"],
        ["REF-UNITE-SUP-001", "UNITE_SUPERFICIE", "HA", "ha", "1", "hectare", "TRUE", "1"],
        ["REF-UNITE-SUP-002", "UNITE_SUPERFICIE", "M2", "m2", "0.0001", "metre carre", "TRUE", "2"],
        ["REF-UNITE-SUP-003", "UNITE_SUPERFICIE", "ACRE", "acre", "0.404686", "acre", "TRUE", "3"],
    ]
}


def column_name(index: int) -> str:
    value = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        value = chr(65 + remainder) + value
    return value


def worksheet_xml(name: str) -> str:
    if name == "ACCUEIL":
        rows = [
            ["AGRIECO PRO - V1", ""],
            ["Ossature technique fonctionnelle", "Sprint 1"],
            ["Navigation", "Utiliser les onglets pour accéder aux modules."],
            ["Producteurs", "Importer les modules VBA puis lancer OuvrirGestionProducteurs."],
            ["Exploitations", "Importer les modules VBA puis lancer OuvrirGestionExploitations."],
            ["Statut", "Classeur préparé sans moteur économique complet."],
        ]
        row_xml = []
        for row_index, values in enumerate(rows, 1):
            cells = []
            for col_index, value in enumerate(values, 1):
                cells.append(
                    f'<c r="{column_name(col_index)}{row_index}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
                )
            row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheetData>{"".join(row_xml)}</sheetData></worksheet>'
        )

    headers = HEADERS[name]
    header_row = "".join(
        f'<c r="{column_name(index)}1" t="inlineStr"><is><t>{escape(value)}</t></is></c>'
        for index, value in enumerate(headers, 1)
    )
    data_rows = DEFAULT_ROWS.get(name, [[]])
    row_xml = [f'<row r="1">{header_row}</row>']
    for row_offset, values in enumerate(data_rows, 2):
        cells = []
        for index in range(1, len(headers) + 1):
            value = values[index - 1] if index <= len(values) else ""
            cells.append(
                f'<c r="{column_name(index)}{row_offset}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'
            )
        row_xml.append(f'<row r="{row_offset}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>'
    )


def table_xml(name: str, table_id: int) -> str:
    headers = HEADERS[name]
    last_row = max(2, 1 + len(DEFAULT_ROWS.get(name, [[]])))
    ref = f'A1:{column_name(len(headers))}{last_row}'
    columns = "".join(
        f'<tableColumn id="{index}" name="{escape(value)}"/>'
        for index, value in enumerate(headers, 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" '
        f'name="tbl_{name}" displayName="tbl_{name}" ref="{ref}" totalsRowShown="0">'
        f'<autoFilter ref="{ref}"/><tableColumns count="{len(headers)}">{columns}</tableColumns>'
        '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
        'showRowStripes="1" showColumnStripes="0"/></table>'
    )


def generate_workbook(output_path: Path) -> None:
    extension = output_path.suffix.lower()
    if extension == ".xlsm":
        workbook_content_type = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    else:
        workbook_content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

    with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
        worksheet_overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(SHEETS) + 1)
        )
        table_overrides = "".join(
            f'<Override PartName="/xl/tables/table{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
            for index in range(1, len(SHEETS))
        )
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{workbook_content_type}"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            f'{worksheet_overrides}{table_overrides}</Types>',
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>',
        )
        workbook_sheets = "".join(
            f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
            for index, name in enumerate(SHEETS, 1)
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets>{workbook_sheets}</sheets></workbook>',
        )
        workbook_relationships = "".join(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
            for index in range(1, len(SHEETS) + 1)
        )
        workbook_relationships += (
            '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{workbook_relationships}</Relationships>',
        )
        archive.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf/></cellXfs></styleSheet>',
        )

        table_id = 1
        for sheet_index, sheet_name in enumerate(SHEETS, 1):
            archive.writestr(f"xl/worksheets/sheet{sheet_index}.xml", worksheet_xml(sheet_name))
            if sheet_name == "ACCUEIL":
                continue
            archive.writestr(
                f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
                f'Target="../tables/table{table_id}.xml"/></Relationships>',
            )
            archive.writestr(f"xl/tables/table{table_id}.xml", table_xml(sheet_name, table_id))
            table_id += 1


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate AGRIECO PRO workbook skeletons.")
    parser.add_argument("--xlsx", default="AGRIECO_PRO_V1.xlsx", help="Output .xlsx path")
    parser.add_argument("--xlsm", default="AGRIECO_PRO_V1.xlsm", help="Output .xlsm path")
    parser.add_argument("--only", choices=["xlsx", "xlsm", "both"], default="both", help="Which workbook(s) to generate")
    args = parser.parse_args()

    if args.only in {"xlsx", "both"}:
        generate_workbook(Path(args.xlsx))
        print(args.xlsx)
    if args.only in {"xlsm", "both"}:
        generate_workbook(Path(args.xlsm))
        print(args.xlsm)


if __name__ == "__main__":
    main()
