"""
analyse_consolidee.py — Assemble parser_bo_offrem.py + bdpm.py pour produire
un classeur Excel consolidé : CA par génériqueur, par nature (répertoire /
remboursement), et analyse de la "fuite" vers d'autres génériqueurs pour les
molécules disponibles chez Biogaran.

Usage :
    python3 analyse_consolidee.py fichier1.pdf [fichier2.pdf ...] --data data/ --out rapport.xlsx

Peut aussi être importé et utilisé programmatiquement (voir build_report()).
"""
import argparse
import re
import sys
from pathlib import Path
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from parser_bo_offrem import parse_bo_offrem
from bdpm import BDPM

# ---------------------------------------------------------------------------
# Classification des offres BO-OFFREM -> génériqueur / catégorie
# ---------------------------------------------------------------------------
# Construit à partir de la table "Total Remise en % par fournisseur et par
# offre" du BO-OFFREM (pages de récapitulatif en fin de document), qui donne
# le nom de laboratoire officiel pour chaque libellé d'offre.
# ⚠️ Si de nouvelles offres apparaissent dans de futurs BO-OFFREM et ne sont
# pas dans cette table, elles ressortiront comme "À CLASSIFIER" dans le
# rapport — il suffit alors d'ajouter une ligne ci-dessous.
OFFRE_TO_GENERIQUEUR = {
    "BIOGARAN Génériques RSF": ("BIOGARAN", True),
    "EG LABO Génériques RSF": ("EG LABO", True),
    "EG LABO non remise": ("EG LABO", True),
    "SANDOZ Génériques RSF": ("SANDOZ", True),
    "SAWIS Génériques RSF": ("SAWIS", True),
    "MEDAC Génériques RSF": ("MEDAC", True),
    "ARROW Génériques RSF": ("ARROW", True),
    "ARROW non remise": ("ARROW", True),
    "ARROW NR RSF": ("ARROW", True),
    "SUBSTIPHARM non remise": ("SUBSTIPHARM", True),
    "SUN PHARMA Génériques RSF": ("SUN PHARMA", True),
    "SUN PHARMA non remise": ("SUN PHARMA", True),
    "TEVA Génériques RSF": ("TEVA", True),
    "TEVA non remise": ("TEVA", True),
    "VIATRIS Génériques RSF": ("VIATRIS", True),
    "ZENTIVA Génériques RSF": ("ZENTIVA", True),
    "ZENTIVA non remise": ("ZENTIVA", True),
    "SA NR ZENTIVA SORBITOL RSF": ("ZENTIVA", True),
    "OASE NR ACCORD HEALTHCARE RSF": ("ACCORD HEALTHCARE", True),
    "OASE ZYDUS RSF": ("ZYDUS", True),
    # Offres non-génériqueurs (dispositifs, nutrition, parapharmacie...) :
    # exclues de l'analyse génériqueur/fuite mais gardées dans le détail.
    "MARQUE CONSEIL PREMIUM RSF": ("DEPOTRADE (dispositifs)", False),
    "NUTRISENS OFFRE_1 GRP RSF": ("NUTRISENS (nutrition)", False),
    "OASE CONVATEC RSF": ("CONVATEC (dispositifs)", False),
    "OASE MOVICOL OTC RSF": ("NORGINE (OTC)", False),
    "DEAL EMBECTA OF3_27,35% RSF": ("EMBECTA (dispositifs)", False),
    "PROMO DAYANG RSF": ("DAYANG (parapharmacie)", False),
    "NESTLE PRIVILEGE AMBASS 2026 OFC RSF": ("NESTLE (nutrition)", False),
    "SIGVARIS CIBLEE 38% RSF": ("SIGVARIS (dispositifs)", False),
    "SIGVARIS DYNAVEN CIBLEE 56% RSF": ("SIGVARIS (dispositifs)", False),
    "COLOPLAST Offre Ciblée 30-15% RSF": ("COLOPLAST (dispositifs)", False),
}

MOIS_FR = {
    "JANVIER": 1, "FEVRIER": 2, "FÉVRIER": 2, "MARS": 3, "AVRIL": 4, "MAI": 5,
    "JUIN": 6, "JUILLET": 7, "AOUT": 8, "AOÛT": 8, "SEPTEMBRE": 9,
    "OCTOBRE": 10, "NOVEMBRE": 11, "DECEMBRE": 12, "DÉCEMBRE": 12,
}


def classify_offre(offre_name):
    if offre_name in OFFRE_TO_GENERIQUEUR:
        return OFFRE_TO_GENERIQUEUR[offre_name]
    upper = (offre_name or "").upper()
    for known_offre, (brand, is_gen) in OFFRE_TO_GENERIQUEUR.items():
        prefix = known_offre.split(" Gén")[0].split(" non")[0].split(" NR")[0].split(" RSF")[0]
        if upper.startswith(prefix.upper()):
            return (brand, is_gen)
    return ("À CLASSIFIER", None)


def detect_periode(pdf_path):
    """Extrait le mois/année depuis le nom de fichier ou le contenu du PDF."""
    name = Path(pdf_path).stem.upper()
    for mois_nom, mois_num in MOIS_FR.items():
        m = re.search(rf"{mois_nom}\D*(\d{{4}})", name)
        if m:
            return f"{mois_num:02d}/{m.group(1)}"
    import pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    for mois_nom, mois_num in MOIS_FR.items():
        m = re.search(rf"{mois_nom}\s+(\d{{4}})", text.upper())
        if m:
            return f"{mois_num:02d}/{m.group(1)}"
    return "Inconnu"


def enrich_rows(pdf_paths, bdpm):
    all_rows = []
    for pdf_path in pdf_paths:
        periode = detect_periode(pdf_path)
        rows, anomalies, _ = parse_bo_offrem(pdf_path)
        if anomalies:
            print(f"⚠️  {len(anomalies)} anomalie(s) sur {pdf_path} :", file=sys.stderr)
            for a in anomalies[:10]:
                print("   ", a, file=sys.stderr)
        for r in rows:
            brand, is_gen = classify_offre(r["offre"])
            nat = bdpm.nature(r["cip13"])
            groupe_id, groupe_libelle = bdpm.groupe_generique(r["cip13"])
            all_rows.append({
                **r,
                "periode": periode,
                "genericqueur": brand,
                "est_genericqueur": is_gen,
                "repertoire": nat["repertoire"],
                "remboursement": nat["remboursement"],
                "groupe_id": groupe_id,
                "groupe_libelle": groupe_libelle,
                "fabricant_titulaire": bdpm.fabricant(r["cip13"]),
            })
    return all_rows


def compute_biogaran_catalog(all_rows):
    catalog = set()
    for r in all_rows:
        if r["genericqueur"] == "BIOGARAN" and r["groupe_id"]:
            catalog.add(r["groupe_id"])
    return catalog


def compute_fuite(all_rows, catalog_biogaran):
    fuite = []
    for r in all_rows:
        if (
            r["est_genericqueur"] is True
            and r["genericqueur"] != "BIOGARAN"
            and r["groupe_id"]
            and r["groupe_id"] in catalog_biogaran
        ):
            fuite.append(r)
    return fuite


# ---------------------------------------------------------------------------
# Construction du classeur Excel
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
NAVY = "1F3864"
RED_LIGHT = "FCE4E4"
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color=NAVY)
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, size=10, color="666666")
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, bold=True, size=10)
THIN_BORDER = Border(*(Side(style="thin", color="CCCCCC"),) * 4)
EUR_FMT = '#,##0.00 €;(#,##0.00 €);"-"'
PCT_FMT = '0.0%'


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def autosize(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_sheet_detail(wb, all_rows):
    ws = wb.create_sheet("Détail lignes")
    headers = [
        "Période", "CIP13", "Libellé produit", "Génériqueur", "Offre (libellé complet)",
        "Catégorie", "Statut répertoire", "Remboursement",
        "Qté Cdée", "Qté fact", "PPHT Unit (€)", "CA PPHT (€)", "Taux Remise",
        "Remise HT (€)", "Groupe générique (libellé BDPM)", "Fabricant (titulaire AMM)",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    for r in all_rows:
        categorie = r["genericqueur"] if r["est_genericqueur"] else (r["genericqueur"] or "Autre")
        ws.append([
            r["periode"], r["cip13"], r["libelle"], r["genericqueur"] if r["est_genericqueur"] else "",
            r["offre"], categorie, r["repertoire"], r["remboursement"],
            r["qte_cdee"], r["qte_fact"], r["ppht_unit"], r["ca_ppht"],
            r["taux_remise"] / 100, r["remise_ht"], r["groupe_libelle"] or "",
            r["fabricant_titulaire"] or "",
        ])

    n = len(all_rows) + 1
    for row in range(2, n + 1):
        ws.cell(row=row, column=11).number_format = EUR_FMT
        ws.cell(row=row, column=12).number_format = EUR_FMT
        ws.cell(row=row, column=13).number_format = PCT_FMT
        ws.cell(row=row, column=14).number_format = EUR_FMT
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).font = NORMAL_FONT

    tab = Table(displayName="DetailLignes", ref=f"A1:{get_column_letter(len(headers))}{n}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    widths = {"A": 10, "B": 15, "C": 32, "D": 14, "E": 30, "F": 20, "G": 22,
              "H": 16, "I": 9, "J": 9, "K": 12, "L": 12, "M": 10, "N": 12,
              "O": 35, "P": 26}
    autosize(ws, widths)
    ws.freeze_panes = "A2"
    return ws, n


def build_sheet_genericqueur(wb, n_detail):
    ws = wb.create_sheet("CA par génériqueur")
    ws["A1"] = "CA par génériqueur — vue consolidée"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Calculé par formules SUMIF/COUNTIF depuis l'onglet 'Détail lignes'"
    ws["A2"].font = SUBTITLE_FONT

    genericqueurs = sorted({v[0] for v in OFFRE_TO_GENERIQUEUR.values() if v[1] is True})

    headers = ["Génériqueur", "Nb lignes", "Qté fact totale", "CA PPHT (€)",
               "Remise HT (€)", "% du CA génériqueurs", "CA Répertoire (€)",
               "CA Hors répertoire (€)", "CA Non remboursé (€)"]
    start_row = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start_row, column=i, value=h)
    style_header_row(ws, start_row, len(headers))

    detail_sheet = "'Détail lignes'"
    col_categorie, col_qte, col_ca, col_remise = "F", "J", "L", "N"
    col_repertoire, col_remb = "G", "H"

    first_data_row = start_row + 1
    for i, gq in enumerate(genericqueurs):
        row = first_data_row + i
        rng_cat = f"{detail_sheet}!${col_categorie}$2:${col_categorie}${n_detail}"
        rng_qte = f"{detail_sheet}!${col_qte}$2:${col_qte}${n_detail}"
        rng_ca = f"{detail_sheet}!${col_ca}$2:${col_ca}${n_detail}"
        rng_remise = f"{detail_sheet}!${col_remise}$2:${col_remise}${n_detail}"
        rng_rep = f"{detail_sheet}!${col_repertoire}$2:${col_repertoire}${n_detail}"
        rng_remb = f"{detail_sheet}!${col_remb}$2:${col_remb}${n_detail}"

        ws.cell(row=row, column=1, value=gq)
        ws.cell(row=row, column=2, value=f"=COUNTIF({rng_cat},A{row})")
        ws.cell(row=row, column=3, value=f"=SUMIF({rng_cat},A{row},{rng_qte})")
        ws.cell(row=row, column=4, value=f"=SUMIF({rng_cat},A{row},{rng_ca})")
        ws.cell(row=row, column=5, value=f"=SUMIF({rng_cat},A{row},{rng_remise})")
        ws.cell(row=row, column=7, value=f'=SUMIFS({rng_ca},{rng_cat},A{row},{rng_rep},"Générique (répertoire)")')
        ws.cell(row=row, column=8, value=f'=SUMIFS({rng_ca},{rng_cat},A{row},{rng_rep},"Hors répertoire")')
        ws.cell(row=row, column=9, value=f'=SUMIFS({rng_ca},{rng_cat},A{row},{rng_remb},"Non remboursé (ou non renseigné)")')

    last_row = first_data_row + len(genericqueurs) - 1
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = BOLD_FONT
    for col in (2, 3, 4, 5, 7, 8, 9):
        letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUM({letter}{first_data_row}:{letter}{last_row})")
        c.font = BOLD_FONT

    for row in range(first_data_row, total_row + 1):
        for col in (4, 5, 7, 8, 9):
            ws.cell(row=row, column=col).number_format = EUR_FMT
        ws.cell(row=row, column=6, value=f"=D{row}/D${total_row}")
        ws.cell(row=row, column=6).number_format = PCT_FMT
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=c)
            if cell.font != BOLD_FONT:
                cell.font = NORMAL_FONT

    widths = {"A": 20, "B": 11, "C": 15, "D": 14, "E": 14, "F": 12, "G": 16, "H": 18, "I": 16}
    autosize(ws, widths)
    return ws


def build_sheet_fuite(wb, all_rows, catalog_biogaran):
    ws = wb.create_sheet("Fuite Biogaran")
    ws["A1"] = "Analyse de fuite — génériques disponibles chez Biogaran achetés ailleurs"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("« Disponible chez Biogaran » = constaté par un achat réel sous l'offre "
                "BIOGARAN dans les périodes analysées (voir onglet Méthodologie).")
    ws["A2"].font = SUBTITLE_FONT

    fuite_rows = compute_fuite(all_rows, catalog_biogaran)
    fuite_rows.sort(key=lambda r: -r["ca_ppht"])

    headers = ["Période", "CIP13", "Libellé produit", "Acheté chez", "Groupe générique (BDPM)",
               "Qté fact", "CA PPHT (€)", "Remise HT (€)"]
    start_row = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start_row, column=i, value=h)
    style_header_row(ws, start_row, len(headers))

    row = start_row + 1
    for r in fuite_rows:
        ws.cell(row=row, column=1, value=r["periode"])
        ws.cell(row=row, column=2, value=r["cip13"])
        ws.cell(row=row, column=3, value=r["libelle"])
        ws.cell(row=row, column=4, value=r["genericqueur"])
        ws.cell(row=row, column=5, value=r["groupe_libelle"] or "")
        ws.cell(row=row, column=6, value=r["qte_fact"])
        ws.cell(row=row, column=7, value=r["ca_ppht"])
        ws.cell(row=row, column=7).number_format = EUR_FMT
        ws.cell(row=row, column=8, value=r["remise_ht"])
        ws.cell(row=row, column=8).number_format = EUR_FMT
        for c in range(1, len(headers) + 1):
            ws.cell(row=row, column=c).font = NORMAL_FONT
            ws.cell(row=row, column=c).fill = PatternFill("solid", start_color=RED_LIGHT, end_color=RED_LIGHT)
        row += 1

    last_row = row - 1
    n_fuite = max(0, last_row - start_row)
    if n_fuite > 0:
        tab = Table(displayName="FuiteBiogaran", ref=f"A{start_row}:{get_column_letter(len(headers))}{last_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium11", showRowStripes=True)
        ws.add_table(tab)

        total_row = last_row + 2
        ws.cell(row=total_row, column=3, value="TOTAL FUITE").font = BOLD_FONT
        ws.cell(row=total_row, column=6, value=f"=SUM(F{start_row+1}:F{last_row})").font = BOLD_FONT
        ws.cell(row=total_row, column=7, value=f"=SUM(G{start_row+1}:G{last_row})").font = BOLD_FONT
        ws.cell(row=total_row, column=7).number_format = EUR_FMT
    else:
        ws.cell(row=start_row + 1, column=1, value="Aucune fuite détectée sur la période analysée.")

    widths = {"A": 10, "B": 15, "C": 34, "D": 18, "E": 40, "F": 9, "G": 12, "H": 12}
    autosize(ws, widths)
    ws.freeze_panes = f"A{start_row+1}"
    return ws, n_fuite


def build_sheet_resume(wb, all_rows, n_fuite):
    ws = wb.create_sheet("Résumé", 0)
    ws["A1"] = "Analyse Factures Pharmacie — Récapitulatif"
    ws["A1"].font = TITLE_FONT
    periodes = sorted({r["periode"] for r in all_rows})
    ws["A2"] = f"Période(s) analysée(s) : {', '.join(periodes)}"
    ws["A2"].font = SUBTITLE_FONT

    total_ca = sum(r["ca_ppht"] for r in all_rows)
    total_remise = sum(r["remise_ht"] for r in all_rows)
    total_lignes = len(all_rows)

    kpis = [
        ("Nombre de lignes produit analysées", total_lignes),
        ("CA PPHT total (€)", total_ca),
        ("Remise HT totale (€)", total_remise),
        ("Nombre de lignes en fuite Biogaran", n_fuite),
    ]
    row = 4
    for label, val in kpis:
        ws.cell(row=row, column=1, value=label).font = BOLD_FONT
        c = ws.cell(row=row, column=2, value=val)
        if "€" in label:
            c.number_format = EUR_FMT
        c.font = NORMAL_FONT
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Voir les onglets :").font = BOLD_FONT
    row += 1
    for txt in [
        "→ 'Détail lignes' : chaque ligne produit enrichie (nature, remboursement, fabricant)",
        "→ 'CA par génériqueur' : synthèse CA/remise par génériqueur et par nature",
        "→ 'Fuite Biogaran' : achats faits ailleurs alors que Biogaran couvre le même générique",
        "→ 'Méthodologie' : règles et limites de l'analyse (à lire avant interprétation)",
    ]:
        ws.cell(row=row, column=1, value=txt).font = NORMAL_FONT
        row += 1

    autosize(ws, {"A": 55, "B": 18})
    return ws


def build_sheet_methodologie(wb):
    ws = wb.create_sheet("Méthodologie")
    ws["A1"] = "Méthodologie & limites"
    ws["A1"].font = TITLE_FONT

    paragraphs = [
        "",
        "SOURCE DES DONNÉES",
        "- Lignes produit extraites des BO-OFFREM OCP (PDF) via parser_bo_offrem.py, "
        "validé au centime sur plusieurs mois (0 anomalie).",
        "- Enrichissement via les fichiers officiels ANSM (BDPM) : CIS_bdpm.txt, "
        "CIS_CIP_bdpm.txt, CIS_GENER_bdpm.txt.",
        "",
        "GÉNÉRIQUEUR (fournisseur commercial)",
        "- Déterminé par le nom de l'offre BO-OFFREM (ex : 'BIOGARAN Génériques RSF' "
        "-> BIOGARAN), PAS par le titulaire d'AMM officiel.",
        "- Raison : un même produit peut être commercialisé sous une marque (ex. Biogaran) "
        "alors que le titulaire légal de l'AMM est un autre laboratoire (co-exploitation). "
        "Exemple vérifié sur ce jeu de données : MACROGOL 4000 BIOG a pour titulaire BDPM "
        "'MAYOLY PHARMA FRANCE', pas Biogaran — alors qu'il est bien acheté et facturé "
        "sous marque Biogaran.",
        "",
        "STATUT RÉPERTOIRE / REMBOURSEMENT",
        "- 'Générique (répertoire)' / 'Princeps' / 'Hors répertoire' : déterminé via le "
        "fichier CIS_GENER_bdpm.txt (groupes génériques officiels ANSM).",
        "- Remboursement : taux issu de CIS_CIP_bdpm.txt (colonne officielle). "
        "'Non remboursé (ou non renseigné)' regroupe les deux cas faute de pouvoir les "
        "distinguer de façon fiable avec ces seuls fichiers.",
        "",
        "CATÉGORIES NON DISPONIBLES : HYBRIDES ET BIOSIMILAIRES",
        "- Les 3 fichiers BDPM utilisés ne codent PAS de façon fiable les statuts "
        "'hybride' et 'biosimilaire'. Ces catégories, demandées initialement, ne sont "
        "donc PAS distinguées dans ce rapport pour éviter une classification erronée. "
        "Une source complémentaire (liste ANSM/EMA des biosimilaires) serait nécessaire "
        "pour les intégrer proprement.",
        "",
        "ANALYSE DE FUITE BIOGARAN",
        "- Principe : un groupe générique (molécule + dosage + forme, code officiel "
        "ANSM) est considéré 'disponible chez Biogaran' s'il a été RÉELLEMENT ACHETÉ "
        "sous l'offre BIOGARAN au moins une fois dans les périodes fournies à ce rapport.",
        "- Ce choix (achat réel plutôt que titulaire BDPM) est déterminant : le titulaire "
        "sous-estime fortement le catalogue Biogaran réel à cause des cas de "
        "co-exploitation (cf. exemple Macrogol ci-dessus).",
        "- Limite : si un mois donné n'a pas été inclus dans l'analyse, un générique "
        "acheté chez Biogaran uniquement ce mois-là n'apparaîtra pas dans le catalogue "
        "de référence. Plus vous fournissez de mois d'historique, plus la détection de "
        "fuite est complète et fiable.",
        "- Sont exclues de l'analyse : les lignes dont l'offre n'est pas un génériqueur "
        "(dispositifs médicaux, nutrition, parapharmacie — ex. Coloplast, Sigvaris, "
        "Nestlé, Convatec) et les produits hors répertoire générique (pas de groupe "
        "générique officiel, donc pas de base de comparaison).",
        "",
        "CLASSIFICATION DES OFFRES",
        "- Le mapping offre -> génériqueur est construit à partir de la table "
        "'Total Remise en % par fournisseur et par offre' imprimée en fin de "
        "BO-OFFREM. Si une offre inédite apparaît dans un futur BO-OFFREM sans être "
        "dans ce mapping, elle ressort comme 'À CLASSIFIER' dans le détail — il suffit "
        "d'ajouter une ligne dans OFFRE_TO_GENERIQUEUR (scripts/analyse_consolidee.py).",
    ]
    row = 3
    for p in paragraphs:
        cell = ws.cell(row=row, column=1, value=p)
        if p and p == p.upper() and not p.startswith("-"):
            cell.font = Font(name=FONT_NAME, bold=True, size=11, color=NAVY)
        else:
            cell.font = NORMAL_FONT
        ws.row_dimensions[row].height = 15 if p else 6
        row += 1

    ws.column_dimensions["A"].width = 110
    for r in range(3, row):
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def build_report(pdf_paths, data_dir, out_path):
    bdpm = BDPM(data_dir)
    all_rows = enrich_rows(pdf_paths, bdpm)
    catalog_biogaran = compute_biogaran_catalog(all_rows)
    fuite_rows = compute_fuite(all_rows, catalog_biogaran)

    wb = Workbook()
    wb.remove(wb.active)

    _, n_detail = build_sheet_detail(wb, all_rows)
    build_sheet_genericqueur(wb, n_detail)
    _, n_fuite = build_sheet_fuite(wb, all_rows, catalog_biogaran)
    build_sheet_resume(wb, all_rows, n_fuite)
    build_sheet_methodologie(wb)

    wb.save(out_path)
    print(f"Rapport généré : {out_path}")
    print(f"  {len(all_rows)} lignes produit, {len(fuite_rows)} lignes de fuite détectées")
    print(f"  CA PPHT total : {sum(r['ca_ppht'] for r in all_rows):,.2f} €")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdfs", nargs="+", help="Un ou plusieurs BO-OFFREM PDF")
    ap.add_argument("--data", default="data", help="Dossier des fichiers BDPM")
    ap.add_argument("--out", default="rapport_pharmacie.xlsx", help="Fichier Excel de sortie")
    args = ap.parse_args()
    build_report(args.pdfs, args.data, args.out)
