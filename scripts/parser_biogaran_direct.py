# -*- coding: utf-8 -*-
"""
Parser pour les factures Biogaran en direct (via Centre Spécialités
Pharmaceutiques / Movianto, "AU NOM ET POUR LE COMPTE DE BIOGARAN").

Contrairement à OCP (récap mensuel par offre) et Alliance (facture grossiste
multi-génériqueurs), une facture Biogaran direct ne concerne QUE des produits
Biogaran : le génériqueur est donc toujours "BIOGARAN" pour toutes les lignes
d'un fichier de ce canal, sans ambiguïté à lever.

Deux formats coexistent :

FORMAT 1 — "CSP" (factures jusqu'à juin 2026, "FACTURE DE SPÉCIALITÉS",
une ligne produit = une ligne de texte propre) :
    <CIP13> <DESIGNATION> <LOT> <QTE> <PU_HT> <REMISE%> <PU_APRES_REMISE> <MONTANT_HT> <TVA%>
Ex:
    3400930055953 BETAMETHASONE 0.05% CREME 30G BGN 0017 30 1.49 2.50 1.453 43.58 2.10

Le numéro de facture + sa date apparaissent sous la forme d'un code du type
"4L60622658 22/06/2026" (répété 2 fois dans le PDF : dans l'entête et dans
le pavé paiement en bas de page) — bien plus fiable à cibler que les libellés
"N° facture :" qui sont sur une ligne séparée de leur valeur.

FORMAT 2 — "Movianto" (factures à partir de juillet 2026, plus un exemplaire
isolé le 07/04/2026 ; en-tête "Dénomination sociale : BIOGARAN", "FACTURE
ETABLIE PAR MOVIANTO POUR ORDRE ET POUR COMPTE DE BIOGARAN"). Un produit =
un bloc de plusieurs lignes :
    <DÉSIGNATION>
    Code Article : <CIP13>
    Remise unitaire : RSP remise simple (Montant X EUR, Taux <REMISE%>
    <N°> <QTE> PCE <PU_NET> <TVA> % <MONTANT_NET>
    % )
    Code CIP : <CIP7>
    Prix brut <PU_BRUT>
N° et date : "N° de facture : 4L..." / "Date de facturation : JJ/MM/AAAA".

Les "Récapitulatifs des remises" Biogaran (avoirs RDP, fichiers
content-guid-…-BIOGARAN_TECH_IN-….pdf) ne sont pas des factures : ils sont
reconnus et renvoyés avec type_document = "avoir_rdp" et aucune ligne.

Contrôle : le "Montant Total net HT" imprimé est extrait (total_net_ht) pour
vérifier que la somme des lignes retombe dessus au centime.
"""

import os
import re
import pdfplumber

LIGNE_PRODUIT = re.compile(
    r"^(?P<cip>\d{13})\s+"
    r"(?P<designation>.+?)\s+"
    r"(?P<lot>\S+)\s+"
    r"(?P<qte>\d+)\s+"
    r"(?P<pu_ht>\d+[.,]\d+)\s+"
    r"(?P<remise>\d+[.,]\d+)\s+"
    r"(?P<pu_apres>\d+[.,]\d+)\s+"
    r"(?P<montant>\d+[.,]\d+)\s+"
    r"(?P<tva>\d+[.,]\d+)$"
)

BLOC_MOVIANTO_RE = re.compile(
    r"(?P<designation>[^\n]+)\n"
    r"Code Article\s*:\s*(?P<cip>\d+)\n"
    r"(?:Remise unitaire\s*:[^\n]*?Taux\s+(?P<remise>-?[\d.]+)[^\n]*\n)?"
    r"(?P<numero>\d+)\s+(?P<qte>-?[\d.]+)\s+(?P<unite>\S+)\s+(?P<pu_net>-?[\d.]+)\s+"
    r"(?P<tva>[\d.]+)\s*%\s+(?P<montant>-?[\d.]+)\n"
    r"(?:%\s*\)\s*\n)?"
    r"(?:Code CIP\s*:\s*\d+\s*\n)?"
    r"Prix brut\s+(?P<prix_brut>-?[\d.]+)"
)
NUM_FACTURE_MOVIANTO_RE = re.compile(r"N°\s*de facture\s*:\s*(\S+)")
DATE_FACTURE_MOVIANTO_RE = re.compile(r"Date de facturation\s*:\s*(\d{2}/\d{2}/\d{4})")
TOTAL_NET_HT_RE = re.compile(r"Montant Total net HT\s*:\s*(-?[\d.,]+)")

NUM_DATE_FACTURE_RE = re.compile(r"(?P<num>\d[A-Z]\d{8})\s+(?P<date>\d{2}/\d{2}/\d{4})")
CLIENT_RE = re.compile(r"N°\s*CLIENT\s*(\d+)")

# Sections vues sur les factures (le libellé de section précède un bloc de
# lignes produit ; utile en filet de sécurité méthodologique, pas exploité
# pour l'instant car le remboursement est déjà déterminé via bdpm.py).
SECTION_RE = re.compile(r"^SPÉCIALITÉS (REMBOURSÉES|NON REMBOURSÉES)\s*$")


def to_float(x):
    if x is None:
        return None
    try:
        return float(str(x).strip().replace(",", "."))
    except ValueError:
        return None


def extraire_facture_biogaran_direct(path_pdf):
    """
    Extrait l'entête + toutes les lignes produit d'une facture Biogaran
    direct (1 page). Retourne un dict :
        {"num_facture": str, "date_facture": str, "lignes": [ {...}, ... ]}

    Chaque ligne produit contient les champs alignés sur le schéma de sortie
    de parser_bo_offrem.py (canal OCP), pour rester compatible avec
    analyse_consolidee.py :
        cip13, libelle, offre, tva, qte_cdee, qte_fact, ppht_unit,
        prix_net_unitaire, ca_ppht, ca_pght, taux_remise, remise_ht, canal
    """
    with pdfplumber.open(path_pdf) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    if "RECAPITULATIF DES REMISES" in text:
        return {"num_facture": "", "date_facture": "", "total_net_ht": None,
                "type_document": "avoir_rdp", "lignes": []}
    if "Code Article" in text and "Prix brut" in text:
        return _extraire_format_movianto(text)

    num_dates = NUM_DATE_FACTURE_RE.findall(text)
    num_facture = num_dates[0][0] if num_dates else ""
    date_facture = num_dates[0][1] if num_dates else ""

    lignes = []
    for brute in text.split("\n"):
        m = LIGNE_PRODUIT.match(brute.strip())
        if not m:
            continue
        qte = int(m.group("qte"))
        pu_ht = to_float(m.group("pu_ht"))          # prix catalogue HT (brut, avant remise)
        taux_remise = to_float(m.group("remise"))    # % donné explicitement sur la facture
        prix_net_unitaire = to_float(m.group("pu_apres"))
        montant_net = to_float(m.group("montant"))   # = ca net après remise (donnée facture)
        tva = to_float(m.group("tva"))

        ca_ppht = round(qte * pu_ht, 2)               # CA brut reconstitué (avant remise)
        remise_ht = round(ca_ppht - montant_net, 2)    # remise en valeur, cohérente avec le taux

        lignes.append({
            "cip13": m.group("cip"),
            "libelle": re.sub(r"\s+", " ", m.group("designation")).strip(),
            "offre": "BIOGARAN",       # pas d'ambiguïté : ce canal = Biogaran uniquement
            "tva": tva,
            "qte_cdee": qte,           # pas de distinction commandée/facturée sur ce format
            "qte_fact": qte,
            "ppht_unit": pu_ht,
            "prix_net_unitaire": prix_net_unitaire,
            "ca_ppht": ca_ppht,
            "ca_pght": montant_net,    # CA net (après remise), donnée fiable de la facture
            "taux_remise": taux_remise,
            "remise_ht": remise_ht,
            "canal": "Biogaran Direct",
            "num_facture": num_facture,
            "date_facture": date_facture,
        })

    return {"num_facture": num_facture, "date_facture": date_facture,
            "total_net_ht": None, "type_document": "facture", "lignes": lignes}


def _extraire_format_movianto(text):
    """Format 2 (Movianto, à partir de juillet 2026) — voir docstring du module."""
    num_m = NUM_FACTURE_MOVIANTO_RE.search(text)
    date_m = DATE_FACTURE_MOVIANTO_RE.search(text)
    num_facture = num_m.group(1) if num_m else ""
    date_facture = date_m.group(1) if date_m else ""
    total_m = TOTAL_NET_HT_RE.search(text)
    total_net_ht = to_float(total_m.group(1)) if total_m else None

    lignes = []
    for m in BLOC_MOVIANTO_RE.finditer(text):
        qte_f = to_float(m.group("qte"))           # imprimée "16.000"
        qte = int(round(qte_f))
        prix_brut = to_float(m.group("prix_brut"))
        prix_net_unitaire = to_float(m.group("pu_net"))
        montant_net = to_float(m.group("montant"))
        taux_remise = to_float(m.group("remise")) if m.group("remise") else 0.0
        tva = to_float(m.group("tva"))

        ca_ppht = round(qte * prix_brut, 2)
        remise_ht = round(ca_ppht - montant_net, 2)

        lignes.append({
            "cip13": m.group("cip"),
            "libelle": re.sub(r"\s+", " ", m.group("designation")).strip(),
            "offre": "BIOGARAN",
            "tva": tva,
            "qte_cdee": qte,
            "qte_fact": qte,
            "ppht_unit": prix_brut,
            "prix_net_unitaire": prix_net_unitaire,
            "ca_ppht": ca_ppht,
            "ca_pght": montant_net,
            "taux_remise": taux_remise,
            "remise_ht": remise_ht,
            "canal": "Biogaran Direct",
            "num_facture": num_facture,
            "date_facture": date_facture,
        })

    return {"num_facture": num_facture, "date_facture": date_facture,
            "total_net_ht": total_net_ht, "type_document": "facture",
            "lignes": lignes}


def extraire_dossier(dossier_pdfs):
    """Parcourt tous les PDF d'un dossier et retourne la liste consolidée des
    lignes. Déduplique par n° de facture : si la même facture apparaît deux
    fois (fichier déposé deux fois sous des noms différents, PDF présent à la
    fois seul et dans un dossier zippé...), elle n'est comptée qu'une fois."""
    toutes_lignes = []
    factures_vues = {}  # num_facture -> nom du fichier déjà retenu
    for nom in sorted(os.listdir(dossier_pdfs)):
        if not nom.lower().endswith(".pdf"):
            continue
        chemin = os.path.join(dossier_pdfs, nom)
        facture = extraire_facture_biogaran_direct(chemin)
        if facture.get("type_document") == "avoir_rdp":
            print(f"ℹ️  {nom} : avoir RDP (récapitulatif des remises) — ignoré")
            continue
        if not facture["lignes"]:
            print(f"⚠️  Aucune ligne produit trouvée dans {nom}")
            continue
        num_facture = facture["num_facture"]
        if num_facture and num_facture in factures_vues:
            print(f"⚠️  Facture {num_facture} ({nom}) déjà comptabilisée via "
                  f"{factures_vues[num_facture]} — doublon ignoré")
            continue
        if num_facture:
            factures_vues[num_facture] = nom
        toutes_lignes.extend(facture["lignes"])
    return toutes_lignes


if __name__ == "__main__":
    import sys
    for path in sys.argv[1:]:
        f = extraire_facture_biogaran_direct(path)
        print(f"{path}: facture {f['num_facture']} du {f['date_facture']}, {len(f['lignes'])} lignes")
