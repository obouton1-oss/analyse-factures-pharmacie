"""
analyse_consolidee.py — Assemble parser_bo_offrem.py + bdpm.py pour produire
un classeur Excel consolidé : CA par génériqueur, par nature (répertoire /
hybride / remboursement), croisé avec le laboratoire et le mois, et analyse
de la "fuite" vers d'autres génériqueurs pour les molécules disponibles chez
Biogaran.

Usage :
    python3 analyse_consolidee.py fichier1.pdf [fichier2.pdf ...] --data data/ --out rapport.xlsx

Peut aussi être importé et utilisé programmatiquement (voir build_report()).
"""
import argparse
import re
import os
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
    # Canal Biogaran Direct (factures unitaires CSP/Movianto) : le parser
    # met toujours offre="BIOGARAN" (pas d'ambiguïté possible sur ce canal),
    # entrée explicite pour ne pas dépendre du filtre heuristique ci-dessous.
    "BIOGARAN": ("BIOGARAN", True),
    # Clé virtuelle, jamais produite par un parser : sert uniquement à faire
    # apparaître "Générique (labo non identifié)" (filet de sécurité BDPM du
    # canal Alliance, cf. _resoudre_marque_alliance) dans les onglots de
    # synthèse (CA par génériqueur, CA type x laboratoire), qui énumèrent
    # les marques à partir de ce dictionnaire.
    "__ALLIANCE_GENERIQUE_NON_IDENTIFIE__": ("Générique (labo non identifié)", True),
    "BIOGARAN Génériques RSF": ("BIOGARAN", True),
    "OFFRE TRINARA RSF": ("EXELTIS", True),
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
    "ACCORD HEALTHCARE Génériques RSF": ("ACCORD HEALTHCARE", True),
    "OASE ZYDUS RSF": ("ZYDUS", True),
    "ZYDUS Génériques RSF": ("ZYDUS", True),
    "VENIPHARM Génériques RSF": ("VENIPHARM", True),
    "OASE NR VIATRIS RSF": ("VIATRIS", True),
    "SA NR ZENTIVA RSF": ("ZENTIVA", True),
    "CRISTERS Génériques RSF": ("CRISTERS", True),
    "KRKA Génériques RSF": ("KRKA", True),
    # Offres biosimilaires : gardées comme génériqueur (True) pour ne pas
    # perdre le CA dans le rapport, mais avec une marque distincte pour ne
    # pas les confondre avec les génériques classiques dans les filtres.
    "ACCORD BIOSIMILAIRE RSF": ("ACCORD HEALTHCARE (biosimilaire)", True),
    "FRESENIUS BIOSIMILAIRE RSF": ("FRESENIUS (biosimilaire)", True),
    # Offres non-génériqueurs (dispositifs, nutrition, parapharmacie...) :
    # exclues de l'analyse génériqueur/fuite mais gardées dans le détail.
    "MARQUE CONSEIL PREMIUM RSF": ("DEPOTRADE (dispositifs)", False),
    "NUTRISENS OFFRE_1 GRP RSF": ("NUTRISENS (nutrition)", False),
    "old_NUTRISENS GRP OFFRE 2 RSF": ("NUTRISENS (nutrition)", False),
    "OASE CONVATEC RSF": ("CONVATEC (dispositifs)", False),
    "OASE MOVICOL OTC RSF": ("NORGINE (OTC)", False),
    "DEAL EMBECTA OF3_27,35% RSF": ("EMBECTA (dispositifs)", False),
    "PROMO DAYANG RSF": ("DAYANG (parapharmacie)", False),
    # ISN = Ineldea Santé Naturelle (vu sur INELDEA ENDOMELIA, recap juin 2026)
    "PROMO ISN RSF": ("INELDEA (parapharmacie)", False),
    "NESTLE PRIVILEGE AMBASS 2026 OFC RSF": ("NESTLE (nutrition)", False),
    "SIGVARIS CIBLEE 38% RSF": ("SIGVARIS (dispositifs)", False),
    "SIGVARIS DYNAVEN CIBLEE 56% RSF": ("SIGVARIS (dispositifs)", False),
    "COLOPLAST Offre Ciblée 30-15% RSF": ("COLOPLAST (dispositifs)", False),
    "COLOPLAST Offre Ciblée 30% RSF": ("COLOPLAST (dispositifs)", False),
    "COLOPLAST Offre Ciblée 23-10% RSF": ("COLOPLAST (dispositifs)", False),
    # Offres "OASE" non génériqueurs (observées sur les BO-OFFREM 2026).
    # Avant le 21/11/2026, elles étaient rattachées à tort à ACCORD HEALTHCARE
    # (cf. classify_offre).
    "OASE DAYANG RSF": ("DAYANG (parapharmacie)", False),
    "OASE ESSITY TENA RSF": ("ESSITY (dispositifs)", False),
    "OASE EVABIOTE RSF": ("EVABIOTE (parapharmacie)", False),
    "OASE LNS DELICAL 2026 RSF": ("DELICAL (nutrition)", False),
    "OASE THUASNE RSF": ("THUASNE (dispositifs)", False),
    "OFFRE DROVELIS RSF": ("DROVELIS (princeps)", False),
}

# ---------------------------------------------------------------------------
# Exclusions de l'assiette de remontée de remise (RDP) Biogaran
# ---------------------------------------------------------------------------
# Molécules Biogaran entrées au répertoire des génériques EN COURS D'ANNÉE.
# La BDPM étant toujours à jour, ces références ressortent "Générique
# (répertoire)" sur l'ensemble de l'année, y compris sur les mois où elles
# n'étaient pas encore au répertoire. Elles sont donc exclues de l'assiette
# servant au calcul de la remontée de remise (plafond cumulé 40 %,
# art. L138-9 CSS), sur toute l'année — décision pharmacien.
#
# Les lignes restent présentes dans le rapport et dans le CA total : seule
# l'assiette remise les déduit. Le marquage est tracé ligne à ligne dans
# l'onglet "Détail lignes" (colonnes "Exclu base remise" / "Motif exclusion").
#
# Clé = CIP13 (13 chiffres). Pour ajouter une référence : une ligne ici suffit,
# tous les onglets et formules suivent automatiquement.
EXCLUSIONS_BASE_REMISE = {
    "3400930302750": "OXAZEPAM BIOGARAN 10 mg cp (groupe 1913 — Seresta 10)",
    "3400930302774": "OXAZEPAM BIOGARAN 50 mg cp séc. (groupe 1914 — Seresta 50)",
    "3400930265109": "AMITRIPTYLINE BIOGARAN 40 mg/mL sol. buv. (groupe 1866 — Laroxyl)",
    "3400930336922": "NEFOPAM BIOGARAN 30 mg cp pell. (groupe 1829 — Acupan cp)",
    "3400930039540": "NEFOPAM BIOGARAN 20 mg/2 mL inj. (groupe 908 — Acupan inj.)",
}

# Motif commun affiché dans le rapport pour toutes les références ci-dessus.
MOTIF_EXCLUSION_REMISE = "Entrée au répertoire en cours d'année"

# Libellés utilisés dans la colonne "Exclu base remise" de l'onglet Détail
# lignes. Ce sont aussi les critères des SUMIFS des onglets de synthèse :
# ne pas les modifier sans mettre à jour les formules.
EXCL_OUI, EXCL_NON = "Oui", "Non"

# ---------------------------------------------------------------------------
# Assiette de remise EG LABO (décision pharmacien du 23/11/2026)
# ---------------------------------------------------------------------------
# Assiette = produits EG REMBOURSABLES (génériques du répertoire, hybrides,
# hors répertoire remboursés), tous canaux, MOINS les catégories ci-dessous
# lorsqu'elles sont remboursables. Règle par mots-clés sur le libellé facture
# et la dénomination BDPM (accents retirés, majuscules) : une nouvelle
# présentation est donc couverte sans mise à jour de liste.
GENERIQUEUR_EG = "EG LABO"
EXCLUSIONS_ASSIETTE_EG = [
    # (motif affiché, motifs qui doivent TOUS être présents)
    ("Calcium vitamine D3 remboursable", (r"CALCIUM", r"D3|COLECALCIF|CHOLECALCIF")),
    ("Nicotine remboursable", (r"NICOTIN",)),
]


# ---------------------------------------------------------------------------
# Rattrapage de remise TEVA (décision pharmacien du 23/11/2026)
# ---------------------------------------------------------------------------
# Génériques TEVA du répertoire, remboursables, tous canaux (Teva Direct et
# indirect : OCP, Alliance). Selon le taux de remise facturé, Teva verse une
# compensation trimestrielle pour atteindre le taux cible :
#   remise facturée 10 / 20 / 30 % -> cible 39 %
#   remise facturée 2,5 %          -> cible 12,5 %
# Compensation attendue = CA brut x taux cible - remise facturée.
GENERIQUEUR_TEVA = "TEVA"
TEVA_TAUX_CIBLE = {2.5: 0.125, 10.0: 0.39, 20.0: 0.39, 30.0: 0.39}


def palier_teva(r):
    """Palier de remise facturée (« 2,5 % », « 10 % »...) d'une ligne TEVA
    éligible au rattrapage, « autre taux » si hors grille, "" si non éligible."""
    if (r.get("genericqueur") != GENERIQUEUR_TEVA or r.get("type_medicament") != "Générique (répertoire)"
            or r.get("remboursement") == "Non remboursé (ou non renseigné)"):
        return ""
    taux = r.get("taux_remise")
    if taux is None:
        return "autre taux"
    taux = float(taux) * (100 if float(taux) <= 1 else 1)   # 0,025 ou 2,5 selon le parser
    for palier in TEVA_TAUX_CIBLE:
        if abs(taux - palier) < 0.05:
            return f"{palier:g} %".replace(".", ",")
    return "autre taux"


def trimestre_de(periode):
    if "/" not in (periode or ""):
        return ""
    mm, aaaa = periode.split("/")
    return f"{aaaa} T{(int(mm) - 1) // 3 + 1}"


def _sans_accents_maj(t):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", t or "")
                   if not unicodedata.combining(c)).upper()


def exclusion_assiette_eg(libelle, denomination, remboursement):
    """Motif d'exclusion de l'assiette EG, ou "" (produit conservé)."""
    if remboursement == "Non remboursé (ou non renseigné)":
        return ""
    texte = _sans_accents_maj(f"{libelle} {denomination}")
    for motif, regles in EXCLUSIONS_ASSIETTE_EG:
        if all(re.search(r, texte) for r in regles):
            return motif
    return ""


def cle_cip13(cip13):
    """Normalise un CIP13 en clé de comparaison : chiffres uniquement.

    Les trois parsers ne produisent pas tous le CIP sous la même forme
    (espaces, séparateurs selon le gabarit de facture), donc on compare sur
    les chiffres seuls pour éviter une exclusion silencieusement inopérante.
    """
    return re.sub(r"\D", "", str(cip13 or ""))


def exclusion_base_remise(cip13):
    """Renvoie (exclu: bool, detail: str) pour un CIP13 donné."""
    detail = EXCLUSIONS_BASE_REMISE.get(cle_cip13(cip13))
    return (detail is not None), (detail or "")


MOIS_FR = {
    "JANVIER": 1, "FEVRIER": 2, "FÉVRIER": 2, "MARS": 3, "AVRIL": 4, "MAI": 5,
    "JUIN": 6, "JUILLET": 7, "AOUT": 8, "AOÛT": 8, "SEPTEMBRE": 9,
    "OCTOBRE": 10, "NOVEMBRE": 11, "DECEMBRE": 12, "DÉCEMBRE": 12,
}


def classify_offre(offre_name):
    """Offre BO-OFFREM -> (marque, est_génériqueur).

    1. Correspondance exacte dans OFFRE_TO_GENERIQUEUR.
    2. Sinon, recherche du NOM DE LABORATOIRE en mot entier dans le libellé
       de l'offre (ex. une future "TEVA Génériques 2027 RSF" -> TEVA). Le plus
       long nom trouvé l'emporte.
    3. Sinon "À CLASSIFIER" (signalé dans le journal du rapport).

    L'ancienne règle comparait le DÉBUT du libellé au début d'une offre
    connue : "OASE NR ACCORD HEALTHCARE RSF" donnait le préfixe "OASE", si
    bien que toutes les offres "OASE ..." (Thuasne, Dayang, Delical...)
    partaient chez ACCORD HEALTHCARE, et "ACCORD HEALTHCARE Génériques RSF"
    restait non classée. Corrigé le 21/11/2026."""
    if offre_name in OFFRE_TO_GENERIQUEUR:
        return OFFRE_TO_GENERIQUEUR[offre_name]
    upper = (offre_name or "").upper()
    candidats = {}
    for brand, is_gen in OFFRE_TO_GENERIQUEUR.values():
        if brand.startswith("__") or "(" in brand and is_gen:
            continue  # clés virtuelles, biosimilaires : correspondance exacte seulement
        nom = brand.split(" (")[0].upper()
        if re.search(r"(?<![A-Z0-9])" + re.escape(nom) + r"(?![A-Z0-9])", upper):
            candidats.setdefault(nom, (brand, is_gen))
    if candidats:
        return candidats[max(candidats, key=len)]
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


def periode_depuis_date_facture(date_facture):
    """Convertit une date 'JJ/MM/AAAA' (ex: facture Biogaran Direct/Alliance)
    en période 'MM/AAAA', cohérente avec detect_periode() (canal OCP)."""
    if not date_facture or "/" not in date_facture:
        return "Inconnu"
    parts = date_facture.split("/")
    if len(parts) != 3:
        return "Inconnu"
    _, mm, aaaa = parts
    return f"{int(mm):02d}/{aaaa}"


def periode_sort_key(periode):
    """Clé de tri chronologique pour une période 'MM/AAAA' (les périodes
    'Inconnu' sont reléguées à la fin)."""
    if periode == "Inconnu" or "/" not in periode:
        return (9999, 99)
    mm, aaaa = periode.split("/")
    return (int(aaaa), int(mm))


def type_medicament(repertoire):
    """
    Normalise le champ `repertoire` retourné par BDPM.nature() en un type de
    médicament pour la vue croisée type x laboratoire :
      - "Générique (répertoire)" -> inchangé
      - "Princeps" -> inchangé
      - "Hors répertoire (Hybride)" -> "Hybride"
      - "Hors répertoire" -> inchangé
    NB : les biosimilaires ne sont pas distingués (cf. Méthodologie) et
    tombent dans "Hors répertoire" comme le reste des produits hors
    répertoire non hybrides.
    """
    if repertoire == "Hors répertoire (Hybride)":
        return "Hybride"
    return repertoire or "Hors répertoire"


def _enrichir_ligne(r, periode, canal, bdpm, brand, is_gen):
    """Enrichissement commun à tous les canaux : BDPM (répertoire, hybride,
    remboursement, fabricant titulaire). La résolution marque/génériqueur
    (brand, is_gen) est calculée en amont, différemment selon le canal :
    - OCP : classify_offre(offre) — traduit un nom d'offre commerciale
    - Biogaran Direct : toujours ("BIOGARAN", True), sans ambiguïté
    - Alliance : _resoudre_marque_alliance() — mot-clé désignation déjà
      résolu par le parser, filet de sécurité BDPM en second recours
    """
    nat = bdpm.nature(r["cip13"])
    groupe_id, groupe_libelle = bdpm.groupe_generique(r["cip13"])
    hyb = nat.get("hybride") or {}
    exclu, detail_exclu = exclusion_base_remise(r["cip13"])
    repertoire = nat["repertoire"]
    repertoire_force = False
    if exclu and repertoire == "Hors répertoire":
        # Une référence de EXCLUSIONS_BASE_REMISE est, par définition, entrée
        # au répertoire en 2026. Si la BDPM ne l'a pas encore enregistrée
        # (cas NEFOPAM BIOGARAN 30 mg, CIS 62781958 absent de CIS_GENER au
        # 19/11/2026), on la classe quand même Générique (répertoire) pour
        # qu'elle apparaisse au bon endroit — la règle reste de la garder
        # dans le CA répertoire et de ne la retirer que de l'assiette.
        repertoire = "Générique (répertoire)"
        repertoire_force = True
        if not groupe_libelle:
            m_grp = re.search(r"\(([^)]*)\)\s*$", detail_exclu)
            groupe_libelle = (m_grp.group(1) if m_grp else "") + \
                " [classement forcé : BDPM pas encore à jour]"
    remboursement = nat["remboursement"]
    if (remboursement == "Non remboursé (ou non renseigné)"
            and repertoire == "Générique (répertoire)" and r.get("tva") == 2.1):
        # Taux absent de la BDPM pour cette présentation : on reprend celui du
        # groupe générique s'il est unique, mais seulement si la ligne est
        # facturée à 2,1 % (TVA des médicaments remboursables). Les deux
        # conditions ensemble évitent de déclarer remboursé un produit vendu
        # sans ordonnance (ex. IBUPROFENE ARROW 200 mg, TVA 10 %, dont le
        # groupe compte des présentations remboursées).
        taux_grp = bdpm.taux_groupe_generique(r["cip13"])
        if taux_grp:
            remboursement = f"{taux_grp} (déduit du groupe générique)"
    return {
        **r,
        "periode": periode,
        "canal": canal,
        "genericqueur": brand,
        "est_genericqueur": is_gen,
        "repertoire": repertoire,
        "type_medicament": type_medicament(repertoire),
        "repertoire_force": repertoire_force,
        "remboursement": remboursement,
        "hybride_role": hyb.get("role") or "",
        "groupe_id": groupe_id,
        "groupe_libelle": groupe_libelle,
        "fabricant_titulaire": bdpm.fabricant(r["cip13"]),
        # Exclusion de l'assiette de remontée de remise (cf.
        # EXCLUSIONS_BASE_REMISE). N'affecte ni le CA, ni l'analyse de
        # fuite ; le type n'est corrigé que si la BDPM est en retard.
        "exclu_base_remise": exclu,
        "motif_exclusion_remise": MOTIF_EXCLUSION_REMISE if exclu else "",
        "detail_exclusion_remise": detail_exclu,
        "motif_exclusion_eg": (exclusion_assiette_eg(r.get("libelle"), bdpm.denomination(r["cip13"]),
                                                     remboursement)
                               if brand == GENERIQUEUR_EG else ""),
    }


def enrich_rows_ocp(pdf_paths, bdpm):
    """Canal OCP : un PDF = un récap mensuel BO-OFFREM (multi-offres)."""
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
            all_rows.append(_enrichir_ligne(r, periode, "OCP", bdpm, brand, is_gen))
    return all_rows


def _controler_total(facture, pdf_path):
    """Contrôle au centime : somme des montants nets des lignes extraites vs
    "Montant Total net HT" imprimé sur la facture (quand le parser le fournit).
    Un écart signale des lignes non reconnues (changement de mise en page)."""
    total = facture.get("total_net_ht")
    if total is None or not facture["lignes"]:
        return
    somme = round(sum(l["ca_pght"] for l in facture["lignes"]), 2)
    if abs(somme - total) > 0.005:
        print(f"⚠️  Écart de total sur {os.path.basename(pdf_path)} : lignes "
              f"{somme:.2f} € vs facture {total:.2f} € net HT "
              f"(écart {total - somme:.2f} €)", file=sys.stderr)


def enrich_rows_biogaran_direct(pdf_paths, bdpm):
    """Canal Biogaran Direct : un PDF = une facture de livraison unitaire.
    Génériqueur toujours BIOGARAN (pas d'ambiguïté sur ce canal)."""
    from parser_biogaran_direct import extraire_facture_biogaran_direct
    all_rows = []
    for pdf_path in pdf_paths:
        facture = extraire_facture_biogaran_direct(pdf_path)
        if facture.get("type_document") == "avoir_rdp":
            # Récapitulatif des remises Biogaran (avoir RDP) : pas une facture
            # d'achat, aucune ligne produit à intégrer — ignoré sans alerte.
            print(f"ℹ️  Avoir RDP ignoré : {os.path.basename(pdf_path)}", file=sys.stderr)
            continue
        if not facture["lignes"]:
            print(f"⚠️  Aucune ligne produit trouvée dans {pdf_path}", file=sys.stderr)
        _controler_total(facture, pdf_path)
        periode = periode_depuis_date_facture(facture["date_facture"])
        for r in facture["lignes"]:
            all_rows.append(_enrichir_ligne(r, periode, "Biogaran Direct", bdpm, "BIOGARAN", True))
    return all_rows


def lire_factures_direct_labo(pdf_paths, bdpm):
    """Canaux « <labo> Direct » (EG LABO, SANDOZ, TEVA — cf. parser_direct_labo.py).
    Renvoie (lignes enrichies, avoirs). Les avoirs sont ignorés pour l'instant
    (liste toujours vide). Le génériqueur est le laboratoire émetteur ; type de
    médicament et remboursement viennent de la BDPM comme pour les autres canaux."""
    from parser_direct_labo import extraire_facture_direct_labo
    rows, avoirs, vues = [], [], {}
    for pdf_path in pdf_paths or []:
        nom = os.path.basename(pdf_path)
        f = extraire_facture_direct_labo(pdf_path)
        for c in f["controles"]:
            print(f"⚠️  {nom} : {c}", file=sys.stderr)
        if f["type_document"] == "inconnu":
            continue
        if f["type_document"] == "facture_sans_produit":
            print(f"  Facture {f['laboratoire']} {f['num_facture']} ({nom}) sans produit facturé "
                  "(documentation gratuite...) — ignorée")
            continue
        if f["type_document"] == "avoir_ignore":
            print(f"  Avoir {f['laboratoire']} {f['num_facture']} ({nom}) ignoré")
            continue
        if f.get("pfht_non_communique"):
            print(f"  ℹ️  {nom} : PFHT non communiqué par {f['laboratoire']} — CA = montant net facturé")
        if f["num_facture"] and f["num_facture"] in vues:
            print(f"⚠️  Facture {f['num_facture']} ({nom}) déjà comptée via {vues[f['num_facture']]} "
                  "— doublon ignoré", file=sys.stderr)
            continue
        if f["num_facture"]:
            vues[f["num_facture"]] = nom
        if f["type_document"] == "avoir_remise":
            # Avoirs de remise additionnelle : ignorés pour l'instant (décision
            # pharmacien du 22/11/2026), comme les avoirs Sandoz.
            print(f"  Avoir {f['laboratoire']} {f['num_facture']} ({nom}) ignoré")
            continue
        periode = periode_depuis_date_facture(f["date_facture"])
        for r in f["lignes"]:
            rows.append(_enrichir_ligne(r, periode, f["canal"], bdpm, f["laboratoire"], True))
    return rows, avoirs


# Certaines lignes Alliance n'ont aucun mot-clé de marque dans leur
# désignation (ex. "BISOPROLOL REF CPR 1,25MG 30" pour un Biogaran vendu sans
# suffixe de marque sur cette référence) : parser_alliance.py ne peut alors
# rien détecter dans le texte. Le titulaire officiel BDPM (bdpm.fabricant())
# permet de trancher ces cas avec certitude quand il correspond à un
# génériqueur déjà suivi ailleurs dans le pipeline (OCP, Biogaran Direct) —
# construit à partir des titulaires réels observés sur les lignes "Générique
# (labo non identifié)" du rapport du 09/07/2026 (102 lignes, remontées par
# le pharmacien qui a identifié plusieurs cas Biogaran mal classés qui
# gonflaient à tort l'onglet "Fuite Biogaran"). ⚠️ Le titulaire de l'AMM
# n'est pas toujours strictement le fournisseur commercial (cf. bdpm.py) —
# ne mapper que les libellés confirmés sur des cas réels, pas de règle
# générique automatique (risque de mal classer un cas de co-exploitation).
TITULAIRE_BDPM_TO_GENERIQUEUR = {
    "BIOGARAN": "BIOGARAN",
    "CRISTERS": "CRISTERS",
    "EG LABO - LABORATOIRES EUROGENERICS": "EG LABO",
    "SUN PHARMA FRANCE": "SUN PHARMA",
    "SUN PHARMACEUTICAL INDUSTRIES EUROPE (PAYS BAS)": "SUN PHARMA",
    "VIATRIS": "VIATRIS",
    "EXELTIS SANTE": "EXELTIS",
    "TEVA": "TEVA",
    "TEVA (PAYS-BAS)": "TEVA",
    "ARROW GENERIQUES": "ARROW",
    "SUBSTIPHARM": "SUBSTIPHARM",
}


def _resoudre_marque_alliance(r, bdpm):
    """Résout la marque génériqueur d'une ligne Alliance, par ordre de priorité :
    1. Mot-clé de désignation déjà détecté par parser_alliance.py (reflète la
       marque réellement imprimée sur le produit — la source la plus fiable).
    2. Titulaire officiel BDPM (bdpm.fabricant()), s'il correspond à un
       génériqueur connu (TITULAIRE_BDPM_TO_GENERIQUEUR ci-dessus).
    3. Filet de sécurité final : si le CIP13 est officiellement au répertoire
       générique (source ANSM, indépendante de la marque) sans titulaire
       reconnu, on le signale comme "Générique (labo non identifié)" plutôt
       que de le perdre silencieusement parmi les princeps/dispositifs."""
    if r["offre"]:
        return r["offre"], True
    marque_bdpm = TITULAIRE_BDPM_TO_GENERIQUEUR.get(bdpm.fabricant(r["cip13"]))
    if marque_bdpm:
        return marque_bdpm, True
    nat = bdpm.nature(r["cip13"])
    if nat["repertoire"] == "Générique (répertoire)":
        return "Générique (labo non identifié)", True
    return None, False


def enrich_rows_alliance(pdf_paths, bdpm):
    """Canal Alliance : un PDF = une facture grossiste multi-génériqueurs."""
    from parser_alliance import extraire_facture_alliance
    all_rows = []
    for pdf_path in pdf_paths:
        facture = extraire_facture_alliance(pdf_path)
        if not facture["lignes"] and facture.get("total_net_ht") == 0:
            # Facture à 0 € (ex. produit gratuit seul) : rien à intégrer.
            continue
        if not facture["lignes"]:
            print(f"⚠️  Aucune ligne produit trouvée dans {pdf_path}", file=sys.stderr)
        _controler_total(facture, pdf_path)
        periode = periode_depuis_date_facture(facture["date_facture"])
        for r in facture["lignes"]:
            brand, is_gen = _resoudre_marque_alliance(r, bdpm)
            all_rows.append(_enrichir_ligne(r, periode, "Alliance", bdpm, brand, is_gen))
    return all_rows


def enrich_rows(ocp_pdfs, bdpm, biogaran_direct_pdfs=None, alliance_pdfs=None):
    """Point d'entrée multi-canal : agrège les lignes enrichies de chaque
    canal fourni. `ocp_pdfs` reste le paramètre historique (positionnel,
    compatible avec les appels existants) ; les autres canaux sont optionnels."""
    all_rows = enrich_rows_ocp(ocp_pdfs, bdpm)
    if biogaran_direct_pdfs:
        all_rows.extend(enrich_rows_biogaran_direct(biogaran_direct_pdfs, bdpm))
    if alliance_pdfs:
        all_rows.extend(enrich_rows_alliance(alliance_pdfs, bdpm))
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
GREY_LIGHT = "F2F2F2"
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
SUBHEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=9)
SUBHEADER_FILL = PatternFill("solid", start_color="2E5395", end_color="2E5395")
TOTAL_FILL = PatternFill("solid", start_color=GREY_LIGHT, end_color=GREY_LIGHT)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color=NAVY)
SUBTITLE_FONT = Font(name=FONT_NAME, italic=True, size=10, color="666666")
NORMAL_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, bold=True, size=10)
THIN_BORDER = Border(*(Side(style="thin", color="CCCCCC"),) * 4)
EUR_FMT = '#,##0.00 €;(#,##0.00 €);"-"'
PCT_FMT = '0.0%'

# Palette de couleurs claires, une par génériqueur, réutilisée de façon cohérente
# sur les onglets "CA par génériqueur" et "CA type x laboratoire" (le même
# génériqueur a toujours la même couleur d'un onglet à l'autre).
LABO_PALETTE = [
    "FDEBD3",  # orange clair
    "D6EAF8",  # bleu clair
    "D5F5E3",  # vert clair
    "FCF3CF",  # jaune clair
    "F5EEF8",  # violet clair
    "FADBD8",  # rose clair
    "D1F2EB",  # turquoise clair
    "EBDEF0",  # mauve clair
    "E8F8F5",  # cyan clair
    "FEF9E7",  # crème
    "D4E6F1",  # bleu ciel
    "E9F7EF",  # menthe
    "FDF2E9",  # pêche
    "F4ECF7",  # lavande
]


def build_labo_fill_map(genericqueurs):
    """Associe une PatternFill claire à chaque génériqueur, de façon stable
    (même génériqueur -> même couleur, quel que soit l'onglet)."""
    fills = {}
    for i, gq in enumerate(genericqueurs):
        color = LABO_PALETTE[i % len(LABO_PALETTE)]
        fills[gq] = PatternFill("solid", start_color=color, end_color=color)
    return fills


def style_header_row(ws, row, ncols, col_start=1):
    for c in range(col_start, col_start + ncols):
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
        "Période", "CIP13", "Libellé produit", "Laboratoire (génériqueur)",
        "Offre (libellé complet)", "Type de médicament", "Statut répertoire (détail)",
        "Remboursement", "Qté Cdée", "Qté fact", "PPHT Unit (€)", "CA PPHT (€)",
        "Taux Remise", "Remise HT (€)", "Groupe générique (libellé BDPM)",
        "Fabricant (titulaire AMM)", "Rôle hybride (R/H)", "Canal",
        # Colonnes S et T : ajoutées EN FIN de tableau à dessein, pour ne pas
        # décaler les colonnes A/D/F/H/J/L/N/R référencées par les formules
        # des onglets de synthèse.
        "Exclu base remise", "Motif exclusion base remise",
        # Colonnes U / V : exclusion de l'assiette de remise EG LABO.
        "Exclu assiette EG", "Motif exclusion EG",
        # Colonnes W / X : rattrapage de remise TEVA.
        "Palier rattrapage Teva", "Trimestre",
    ]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))

    for r in all_rows:
        laboratoire = r["genericqueur"] if r["est_genericqueur"] else ""
        ws.append([
            r["periode"], r["cip13"], r["libelle"], laboratoire,
            r["offre"], r["type_medicament"], r["repertoire"], r["remboursement"],
            r["qte_cdee"], r["qte_fact"], r["ppht_unit"], r["ca_ppht"],
            r["taux_remise"] / 100, r["remise_ht"], r["groupe_libelle"] or "",
            r["fabricant_titulaire"] or "", r["hybride_role"], r["canal"],
            EXCL_OUI if r.get("exclu_base_remise") else EXCL_NON,
            r.get("motif_exclusion_remise", ""),
            EXCL_OUI if r.get("motif_exclusion_eg") else EXCL_NON,
            r.get("motif_exclusion_eg", ""),
            palier_teva(r),
            trimestre_de(r.get("periode")),
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

    widths = {"A": 10, "B": 15, "C": 32, "D": 20, "E": 30, "F": 22, "G": 24,
              "H": 16, "I": 9, "J": 9, "K": 12, "L": 12, "M": 10, "N": 12,
              "O": 35, "P": 26, "Q": 12, "R": 16, "S": 17, "T": 34, "U": 16, "V": 32, "W": 16, "X": 11}
    autosize(ws, widths)
    ws.freeze_panes = "A2"
    return ws, n


def build_sheet_genericqueur(wb, n_detail, labo_fills, genericqueurs):
    ws = wb.create_sheet("CA par génériqueur")
    ws["A1"] = "CA par génériqueur — vue consolidée (toutes périodes confondues)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Calculé par formules SUMIF/COUNTIF depuis l'onglet 'Détail lignes'"
    ws["A2"].font = SUBTITLE_FONT

    headers = ["Génériqueur", "Nb lignes", "Qté fact totale", "CA PPHT (€)",
               "Remise HT (€)", "% du CA génériqueurs", "CA Répertoire (€)",
               "CA Hybride (€)", "CA Hors répertoire (€)", "CA Non remboursé (€)",
               "dont exclu base remise (€)", "Assiette répertoire remisable (€)"]
    start_row = 4
    for i, h in enumerate(headers, start=1):
        ws.cell(row=start_row, column=i, value=h)
    style_header_row(ws, start_row, len(headers))

    detail_sheet = "'Détail lignes'"
    # Colonnes de l'onglet Détail lignes (cf. build_sheet_detail) :
    # D=Laboratoire, F=Type de médicament, H=Remboursement, J=Qté fact,
    # L=CA PPHT, N=Remise HT, S=Exclu base remise
    col_labo, col_type, col_remb = "D", "F", "H"
    col_qte, col_ca, col_remise = "J", "L", "N"
    col_exclu = "S"

    first_data_row = start_row + 1
    for i, gq in enumerate(genericqueurs):
        row = first_data_row + i
        rng_labo = f"{detail_sheet}!${col_labo}$2:${col_labo}${n_detail}"
        rng_type = f"{detail_sheet}!${col_type}$2:${col_type}${n_detail}"
        rng_qte = f"{detail_sheet}!${col_qte}$2:${col_qte}${n_detail}"
        rng_ca = f"{detail_sheet}!${col_ca}$2:${col_ca}${n_detail}"
        rng_remise = f"{detail_sheet}!${col_remise}$2:${col_remise}${n_detail}"
        rng_remb = f"{detail_sheet}!${col_remb}$2:${col_remb}${n_detail}"
        rng_exclu = f"{detail_sheet}!${col_exclu}$2:${col_exclu}${n_detail}"

        ws.cell(row=row, column=1, value=gq)
        ws.cell(row=row, column=2, value=f"=COUNTIF({rng_labo},A{row})")
        ws.cell(row=row, column=3, value=f"=SUMIF({rng_labo},A{row},{rng_qte})")
        ws.cell(row=row, column=4, value=f"=SUMIF({rng_labo},A{row},{rng_ca})")
        ws.cell(row=row, column=5, value=f"=SUMIF({rng_labo},A{row},{rng_remise})")
        ws.cell(row=row, column=7, value=f'=SUMIFS({rng_ca},{rng_labo},A{row},{rng_type},"Générique (répertoire)")')
        ws.cell(row=row, column=8, value=f'=SUMIFS({rng_ca},{rng_labo},A{row},{rng_type},"Hybride")')
        ws.cell(row=row, column=9, value=f'=SUMIFS({rng_ca},{rng_labo},A{row},{rng_type},"Hors répertoire")')
        ws.cell(row=row, column=10, value=f'=SUMIFS({rng_ca},{rng_labo},A{row},{rng_remb},"Non remboursé (ou non renseigné)")')
        # K : part du CA répertoire portée par des références exclues de
        # l'assiette de remontée de remise (entrées au répertoire en cours
        # d'année) — cf. EXCLUSIONS_BASE_REMISE.
        ws.cell(row=row, column=11,
                value=f'=SUMIFS({rng_ca},{rng_labo},A{row},{rng_type},"Générique (répertoire)",{rng_exclu},"{EXCL_OUI}")')
        # L : assiette réellement retenue pour le calcul de la remontée.
        ws.cell(row=row, column=12, value=f"=G{row}-K{row}")

    last_row = first_data_row + len(genericqueurs) - 1
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = BOLD_FONT
    for col in (2, 3, 4, 5, 7, 8, 9, 10, 11, 12):
        letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUM({letter}{first_data_row}:{letter}{last_row})")
        c.font = BOLD_FONT

    for row in range(first_data_row, total_row + 1):
        for col in (4, 5, 7, 8, 9, 10, 11, 12):
            ws.cell(row=row, column=col).number_format = EUR_FMT
        ws.cell(row=row, column=6, value=f"=D{row}/D${total_row}")
        ws.cell(row=row, column=6).number_format = PCT_FMT
        gq_ligne = ws.cell(row=row, column=1).value
        fill = labo_fills.get(gq_ligne)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=c)
            if cell.font != BOLD_FONT:
                cell.font = NORMAL_FONT
            if fill is not None:
                cell.fill = fill

    widths = {"A": 20, "B": 11, "C": 15, "D": 14, "E": 14, "F": 12, "G": 16,
              "H": 14, "I": 18, "J": 16, "K": 18, "L": 20}
    autosize(ws, widths)
    return ws


# Canaux détaillés pour le laboratoire Biogaran uniquement (cf. demande
# pharmacien) dans l'onglet "CA type x laboratoire" : une ligne par canal en
# plus de la ligne agrégée "Total", pour voir comment chaque catégorie de
# médicament Biogaran (générique répertoire, hors répertoire, hybride...)
# se répartit entre les 3 flux d'approvisionnement.
CANAUX_BIOGARAN_DETAIL = ["OCP", "Alliance", "Biogaran Direct"]

# Ligne affichée sous "BIOGARAN / Générique (répertoire) / Total" : part de ce
# CA portant sur les références de EXCLUSIONS_BASE_REMISE. Ligne "dont" :
# elle n'est PAS ajoutée au TOTAL GÉNÉRAL (déjà incluse dans la ligne Total).
LIGNE_EXCLUSIONS_REPERTOIRE = "dont exclusions répertoire"
EXCL_FONT_COLOR = "9C5700"
# Ligne suivante : assiette de remise = Total − exclusions (même calcul que
# la colonne « Assiette retenue » de l'onglet Assiette remise Biogaran).
# Ligne de lecture, elle non plus n'entre pas dans le TOTAL GÉNÉRAL.
LIGNE_ASSIETTE_REMISE = "Assiette remise (Total − exclusions)"
ASSIETTE_FONT_COLOR = "1F6F43"


def build_sheet_type_labo(wb, all_rows, n_detail, periodes, labo_fills):
    """
    Nouvelle vue demandée : CA par type de médicament x par laboratoire,
    avec détail mensuel (une colonne de 4 sous-mesures par période) et un
    bloc "Total annuel" à droite.
    Sous-mesures par période : CA brut (= CA PPHT), Remise HT, CA remisé
    (= CA PPHT − Remise HT, hypothèse à confirmer — voir onglet Méthodologie),
    Taux de remise pondéré (= Remise HT / CA PPHT).

    Pour BIOGARAN uniquement, chaque type de médicament est en plus détaillé
    par canal (OCP / Alliance / Biogaran Direct + une ligne Total) — voir
    CANAUX_BIOGARAN_DETAIL. Les autres laboratoires gardent une seule ligne
    agrégée, comme avant.
    """
    ws = wb.create_sheet("CA type x laboratoire")
    ws["A1"] = "CA par type de médicament et par laboratoire — détail mensuel"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("CA remisé = CA PPHT − Remise HT (hypothèse de calcul à confirmer, "
                "voir onglet Méthodologie). Ne couvre que les lignes génériqueur "
                "(hors dispositifs/nutrition/parapharmacie). Pour Biogaran, détail "
                "par canal (OCP / Alliance / Biogaran Direct) sous chaque type de "
                "médicament ; la ligne « Total » reprend l'agrégat des 3 canaux. "
                "Cet onglet présente le CA intégral : les références exclues de "
                "l'assiette de remontée de remise y sont incluses, isolées sur la "
                "ligne « dont exclusions répertoire » et détaillées sous le tableau ; "
                "la ligne « Assiette remise » en dessous = Total − exclusions "
                "— voir l'onglet « Assiette remise Biogaran » pour l'assiette corrigée.")
    ws["A2"].font = SUBTITLE_FONT

    detail_sheet = "'Détail lignes'"
    col_labo, col_type, col_canal = "D", "F", "R"
    col_periode, col_ca, col_remise = "A", "L", "N"
    rng_labo = f"{detail_sheet}!${col_labo}$2:${col_labo}${n_detail}"
    rng_type = f"{detail_sheet}!${col_type}$2:${col_type}${n_detail}"
    rng_canal = f"{detail_sheet}!${col_canal}$2:${col_canal}${n_detail}"
    rng_periode = f"{detail_sheet}!${col_periode}$2:${col_periode}${n_detail}"
    rng_ca = f"{detail_sheet}!${col_ca}$2:${col_ca}${n_detail}"
    rng_remise = f"{detail_sheet}!${col_remise}$2:${col_remise}${n_detail}"
    rng_cip = f"{detail_sheet}!$B$2:$B${n_detail}"
    rng_exclu = f"{detail_sheet}!$S$2:$S${n_detail}"

    header_row1, header_row2 = 4, 5
    ws.cell(row=header_row1, column=1, value="Laboratoire")
    ws.cell(row=header_row1, column=2, value="Type de médicament")
    ws.cell(row=header_row1, column=3, value="Canal (Biogaran uniquement)")
    ws.merge_cells(start_row=header_row1, start_column=1, end_row=header_row2, end_column=1)
    ws.merge_cells(start_row=header_row1, start_column=2, end_row=header_row2, end_column=2)
    ws.merge_cells(start_row=header_row1, start_column=3, end_row=header_row2, end_column=3)

    sub_labels = ["CA brut (€)", "Remise (€)", "CA remisé (€)", "Taux remise pondéré"]
    col = 4
    period_col_start = {}
    for periode in periodes:
        period_col_start[periode] = col
        ws.merge_cells(start_row=header_row1, start_column=col, end_row=header_row1, end_column=col + 3)
        ws.cell(row=header_row1, column=col, value=periode)
        for j, lbl in enumerate(sub_labels):
            ws.cell(row=header_row2, column=col + j, value=lbl)
        col += 4
    total_col_start = col
    ws.merge_cells(start_row=header_row1, start_column=col, end_row=header_row1, end_column=col + 3)
    ws.cell(row=header_row1, column=col, value="TOTAL ANNUEL")
    for j, lbl in enumerate(sub_labels):
        ws.cell(row=header_row2, column=col + j, value=lbl)
    n_cols_total = col + 3

    for r in range(header_row1, header_row2 + 1):
        for c in range(1, n_cols_total + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = SUBHEADER_FONT if r == header_row2 else HEADER_FONT
            cell.fill = SUBHEADER_FILL if r == header_row2 else HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

    # Combinaisons (laboratoire, type) réellement présentes dans les données,
    # restreintes aux lignes génériqueur (cohérent avec l'onglet CA par génériqueur).
    combos = sorted({
        (r["genericqueur"], r["type_medicament"])
        for r in all_rows if r["est_genericqueur"] is True
    })

    # Construit la liste finale des lignes à afficher : (labo, type_med, canal).
    # canal = None -> ligne agrégée classique (tous laboratoires hors Biogaran).
    # Pour Biogaran : une ligne par canal de CANAUX_BIOGARAN_DETAIL, puis une
    # ligne canal="Total" qui reprend l'agrégat (identique à l'ancien comportement).
    lignes_a_afficher = []
    for labo, type_med in combos:
        if labo == "BIOGARAN":
            for canal in CANAUX_BIOGARAN_DETAIL:
                lignes_a_afficher.append((labo, type_med, canal))
            lignes_a_afficher.append((labo, type_med, "Total"))
            if type_med == "Générique (répertoire)" and EXCLUSIONS_BASE_REMISE:
                lignes_a_afficher.append((labo, type_med, LIGNE_EXCLUSIONS_REPERTOIRE))
                lignes_a_afficher.append((labo, type_med, LIGNE_ASSIETTE_REMISE))
        else:
            lignes_a_afficher.append((labo, type_med, None))

    first_data_row = header_row2 + 1
    for i, (labo, type_med, canal) in enumerate(lignes_a_afficher):
        row = first_data_row + i
        ws.cell(row=row, column=1, value=labo)
        ws.cell(row=row, column=2, value=type_med)
        ws.cell(row=row, column=3, value=canal if canal else "")

        # Critère canal additionnel uniquement sur les lignes de détail par
        # canal (OCP/Alliance/Biogaran Direct) ; les lignes "Total" et les
        # lignes des autres laboratoires restent agrégées tous canaux, comme
        # avant.
        canal_crit = f',{rng_canal},"{canal}"' if canal in CANAUX_BIOGARAN_DETAIL else ""
        if canal == LIGNE_EXCLUSIONS_REPERTOIRE:
            canal_crit = f',{rng_exclu},"{EXCL_OUI}"'

        if canal == LIGNE_ASSIETTE_REMISE:
            for periode in list(periodes) + ["__TOTAL__"]:
                c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
                for off in (0, 1):
                    L = get_column_letter(c0 + off)
                    ws.cell(row=row, column=c0 + off, value=f"={L}{row - 2}-{L}{row - 1}")
                ca_cell = f"{get_column_letter(c0)}{row}"
                remise_cell = f"{get_column_letter(c0 + 1)}{row}"
                ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
                ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")
            continue

        for periode in periodes:
            c0 = period_col_start[periode]
            ca_cell = f"{get_column_letter(c0)}{row}"
            remise_cell = f"{get_column_letter(c0 + 1)}{row}"
            ws.cell(row=row, column=c0,
                    value=f'=SUMIFS({rng_ca},{rng_labo},$A{row},{rng_type},$B{row},{rng_periode},"{periode}"{canal_crit})')
            ws.cell(row=row, column=c0 + 1,
                    value=f'=SUMIFS({rng_remise},{rng_labo},$A{row},{rng_type},$B{row},{rng_periode},"{periode}"{canal_crit})')
            ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
            ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")

        c0 = total_col_start
        ca_cell = f"{get_column_letter(c0)}{row}"
        remise_cell = f"{get_column_letter(c0 + 1)}{row}"
        # Total annuel = somme sur TOUTES les lignes de la combinaison, sans filtre
        # de période (donc robuste même si une période n'est pas dans `periodes`,
        # ex. période "Inconnu" faute de PDF exploitable pour un mois donné).
        ws.cell(row=row, column=c0,
                value=f'=SUMIFS({rng_ca},{rng_labo},$A{row},{rng_type},$B{row}{canal_crit})')
        ws.cell(row=row, column=c0 + 1,
                value=f'=SUMIFS({rng_remise},{rng_labo},$A{row},{rng_type},$B{row}{canal_crit})')
        ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
        ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")

    last_row = first_data_row + len(lignes_a_afficher) - 1

    # Ligne TOTAL GÉNÉRAL — exclut explicitement les lignes de détail par canal
    # Biogaran (OCP/Alliance/Biogaran Direct) pour ne pas compter 3 fois le CA
    # déjà inclus dans la ligne "Total" de chaque type de médicament Biogaran.
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL GÉNÉRAL").font = BOLD_FONT
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    exclusion_canaux = "".join(f',{get_column_letter(3)}{first_data_row}:{get_column_letter(3)}{last_row},"<>{c}"'
                                for c in CANAUX_BIOGARAN_DETAIL + [LIGNE_EXCLUSIONS_REPERTOIRE,
                                                                   LIGNE_ASSIETTE_REMISE])
    for periode in list(periodes) + ["__TOTAL__"]:
        c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
        for offset in (0, 1):  # CA brut, Remise : sommables
            letter = get_column_letter(c0 + offset)
            ws.cell(row=total_row, column=c0 + offset,
                    value=f"=SUMIFS({letter}{first_data_row}:{letter}{last_row}{exclusion_canaux})").font = BOLD_FONT
        ca_cell = f"{get_column_letter(c0)}{total_row}"
        remise_cell = f"{get_column_letter(c0 + 1)}{total_row}"
        ws.cell(row=total_row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}").font = BOLD_FONT
        ws.cell(row=total_row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)").font = BOLD_FONT

    # Mise en forme des nombres + police + fond gris sur la ligne total,
    # fond couleur par génériqueur sur les lignes de données. Les lignes de
    # détail par canal Biogaran sont en italique (sous-détail), la ligne
    # "Total" Biogaran repasse en gras (comme l'ancienne ligne unique).
    for row in range(first_data_row, total_row + 1):
        is_total = (row == total_row)
        labo_ligne = ws.cell(row=row, column=1).value
        canal_ligne = ws.cell(row=row, column=3).value
        is_canal_detail = canal_ligne in CANAUX_BIOGARAN_DETAIL
        is_sous_total_biogaran = (canal_ligne == "Total")
        is_ligne_exclusions = (canal_ligne == LIGNE_EXCLUSIONS_REPERTOIRE)
        is_ligne_assiette = (canal_ligne == LIGNE_ASSIETTE_REMISE)
        fill = labo_fills.get(labo_ligne) if not is_total else None
        for periode in list(periodes) + ["__TOTAL__"]:
            c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
            ws.cell(row=row, column=c0).number_format = EUR_FMT
            ws.cell(row=row, column=c0 + 1).number_format = EUR_FMT
            ws.cell(row=row, column=c0 + 2).number_format = EUR_FMT
            ws.cell(row=row, column=c0 + 3).number_format = PCT_FMT
        for c in range(1, n_cols_total + 1):
            cell = ws.cell(row=row, column=c)
            if is_total:
                cell.fill = TOTAL_FILL
                if cell.font != BOLD_FONT:
                    cell.font = BOLD_FONT
            else:
                if is_ligne_assiette:
                    cell.font = Font(name=FONT_NAME, bold=True, size=9,
                                     color=ASSIETTE_FONT_COLOR)
                elif is_ligne_exclusions:
                    cell.font = Font(name=FONT_NAME, italic=True, bold=True, size=9,
                                     color=EXCL_FONT_COLOR)
                elif is_canal_detail:
                    cell.font = Font(name=FONT_NAME, italic=True, size=9, color="555555")
                elif is_sous_total_biogaran:
                    cell.font = BOLD_FONT
                elif cell.font not in (BOLD_FONT,):
                    cell.font = NORMAL_FONT
                if fill is not None:
                    cell.fill = fill

    # Liste des références exclues, sous le tableau : une ligne par CIP13 avec
    # le même détail mensuel ; leur somme = ligne « dont exclusions répertoire ».
    if EXCLUSIONS_BASE_REMISE:
        row = total_row + 2
        titre = ws.cell(row=row, column=1,
                        value="Exclusions répertoire — références Biogaran entrées au "
                              "répertoire en cours d'année, déduites de l'assiette de "
                              "remontée de remise (incluses dans le CA ci-dessus)")
        titre.font = BOLD_FONT
        row += 1
        types_par_cip = defaultdict(set)
        for r_ in all_rows:
            if r_.get("exclu_base_remise"):
                types_par_cip[cle_cip13(r_["cip13"])].add(r_.get("type_medicament") or "")
        for c, lbl in enumerate(["CIP13", "Libellé", "Classement (BDPM)"], start=1):
            cell = ws.cell(row=row, column=c, value=lbl)
            cell.font = SUBHEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.border = THIN_BORDER
        row += 1
        first_excl_row = row
        for cip13, libelle in sorted(EXCLUSIONS_BASE_REMISE.items(), key=lambda kv: kv[1]):
            ws.cell(row=row, column=1, value=cip13)
            ws.cell(row=row, column=2, value=libelle)
            ws.cell(row=row, column=3,
                    value=" / ".join(sorted(types_par_cip.get(cip13, set())))
                    or "Non acheté sur la période")
            for periode in list(periodes) + ["__TOTAL__"]:
                c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
                per_crit = f',{rng_periode},"{periode}"' if periode != "__TOTAL__" else ""
                ca_cell = f"{get_column_letter(c0)}{row}"
                remise_cell = f"{get_column_letter(c0 + 1)}{row}"
                ws.cell(row=row, column=c0,
                        value=f'=SUMIFS({rng_ca},{rng_cip},$A{row}{per_crit})')
                ws.cell(row=row, column=c0 + 1,
                        value=f'=SUMIFS({rng_remise},{rng_cip},$A{row}{per_crit})')
                ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
                ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")
            row += 1
        last_excl_row = row - 1
        # Ligne 1 : total de toutes les références listées.
        ws.cell(row=row, column=1, value="Total références exclues")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        for periode in list(periodes) + ["__TOTAL__"]:
            c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
            for offset in (0, 1):
                letter = get_column_letter(c0 + offset)
                ws.cell(row=row, column=c0 + offset,
                        value=f"=SUM({letter}{first_excl_row}:{letter}{last_excl_row})")
            ca_cell = f"{get_column_letter(c0)}{row}"
            remise_cell = f"{get_column_letter(c0 + 1)}{row}"
            ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
            ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")
        # Ligne 2 : seules les références classées Générique (répertoire) sont
        # dans l'assiette, donc réellement déduites — égale à la ligne
        # « dont exclusions répertoire » du tableau.
        row += 1
        ws.cell(row=row, column=1,
                value="dont déduit de l'assiette (classé Générique répertoire)")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        rng_cls = f"$C${first_excl_row}:$C${last_excl_row}"
        for periode in list(periodes) + ["__TOTAL__"]:
            c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
            for offset in (0, 1):
                letter = get_column_letter(c0 + offset)
                ws.cell(row=row, column=c0 + offset,
                        value=f'=SUMIFS({letter}{first_excl_row}:{letter}{last_excl_row},'
                              f'{rng_cls},"Générique (répertoire)")')
            ca_cell = f"{get_column_letter(c0)}{row}"
            remise_cell = f"{get_column_letter(c0 + 1)}{row}"
            ws.cell(row=row, column=c0 + 2, value=f"={ca_cell}-{remise_cell}")
            ws.cell(row=row, column=c0 + 3, value=f"=IFERROR({remise_cell}/{ca_cell},0)")
        total_excl_row = row
        for r in range(first_excl_row, total_excl_row + 1):
            for periode in list(periodes) + ["__TOTAL__"]:
                c0 = period_col_start[periode] if periode != "__TOTAL__" else total_col_start
                for off in (0, 1, 2):
                    ws.cell(row=r, column=c0 + off).number_format = EUR_FMT
                ws.cell(row=r, column=c0 + 3).number_format = PCT_FMT
            for c in range(1, n_cols_total + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = THIN_BORDER
                if r >= last_excl_row + 1:
                    cell.font = Font(name=FONT_NAME, bold=True, size=9, color=EXCL_FONT_COLOR)
                    cell.fill = TOTAL_FILL
                else:
                    cell.font = NORMAL_FONT
            if r <= last_excl_row:
                ws.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
                ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")

    widths = {"A": 18, "B": 22, "C": 16}
    for c in range(4, n_cols_total + 1):
        widths[get_column_letter(c)] = 13
    autosize(ws, widths)
    ws.freeze_panes = f"D{first_data_row}"
    return ws


# Plafond légal de remise cumulée sur les génériques remboursables
# (art. L138-9 CSS) : remise facturée + remontée (RDP) ne peut excéder 40 %.
PLAFOND_REMISE_GENERIQUES = 0.40


def build_sheet_assiette_remise(wb, all_rows, n_detail):
    """
    Assiette de la remontée de remise (RDP) Biogaran, mois par mois.

    Part du CA répertoire Biogaran et en déduit les références exclues
    (entrées au répertoire en cours d'année — cf. EXCLUSIONS_BASE_REMISE),
    pour obtenir l'assiette réellement retenue, le taux de remise pondéré
    sur cette assiette, et la remontée théorique restant à percevoir au
    plafond de 40 %.

    Tout est en formules pointant sur l'onglet "Détail lignes" : les
    montants restent vérifiables et recalculables dans Excel.
    """
    ws = wb.create_sheet("Assiette remise Biogaran")
    ws["A1"] = "Assiette de la remontée de remise Biogaran — détail mensuel"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Les références Biogaran entrées au répertoire des génériques en cours "
                "d'année sont déduites de l'assiette (colonne C), sur toute l'année. "
                "La remontée théorique est un calcul au plafond de 40 % "
                "(art. L138-9 CSS) : à rapprocher des avoirs RDP réellement reçus, "
                "ce n'est pas un montant dû constaté.")
    ws["A2"].font = SUBTITLE_FONT

    detail_sheet = "'Détail lignes'"
    # Colonnes de l'onglet Détail lignes (cf. build_sheet_detail) :
    # A=Période, D=Laboratoire, F=Type de médicament, L=CA PPHT, N=Remise HT,
    # S=Exclu base remise
    rng_periode = f"{detail_sheet}!$A$2:$A${n_detail}"
    rng_labo = f"{detail_sheet}!$D$2:$D${n_detail}"
    rng_type = f"{detail_sheet}!$F$2:$F${n_detail}"
    rng_ca = f"{detail_sheet}!$L$2:$L${n_detail}"
    rng_remise = f"{detail_sheet}!$N$2:$N${n_detail}"
    rng_exclu = f"{detail_sheet}!$S$2:$S${n_detail}"

    base_crit = f'{rng_labo},"BIOGARAN",{rng_type},"Générique (répertoire)"'

    headers = [
        "Période",
        "CA répertoire Biogaran (€)",
        "dont exclu — entrée répertoire en cours d'année (€)",
        "Assiette retenue (€)",
        "Remise facturée sur assiette (€)",
        "Taux de remise pondéré sur assiette",
        "Plafond 40 % de l'assiette (€)",
        "Remontée théorique restant à percevoir (€)",
    ]
    header_row = 4
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=i, value=h)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    style_header_row(ws, header_row, len(headers))
    ws.row_dimensions[header_row].height = 46

    # Périodes réellement présentes sur des lignes répertoire Biogaran : inutile
    # d'afficher un mois sans achat Biogaran au répertoire.
    periodes = sorted(
        {r["periode"] for r in all_rows
         if r.get("genericqueur") == "BIOGARAN"
         and r.get("type_medicament") == "Générique (répertoire)"},
        key=periode_sort_key,
    )

    first_data_row = header_row + 1
    for i, periode in enumerate(periodes):
        row = first_data_row + i
        per_crit = f',{rng_periode},"{periode}"'
        ws.cell(row=row, column=1, value=periode)
        ws.cell(row=row, column=2,
                value=f'=SUMIFS({rng_ca},{base_crit}{per_crit})')
        ws.cell(row=row, column=3,
                value=f'=SUMIFS({rng_ca},{base_crit}{per_crit},{rng_exclu},"{EXCL_OUI}")')
        ws.cell(row=row, column=4, value=f"=B{row}-C{row}")
        # Remise facturée sur les seules lignes conservées dans l'assiette.
        ws.cell(row=row, column=5,
                value=f'=SUMIFS({rng_remise},{base_crit}{per_crit},{rng_exclu},"{EXCL_NON}")')
        ws.cell(row=row, column=6, value=f"=IFERROR(E{row}/D{row},0)")
        ws.cell(row=row, column=7, value=f"=D{row}*{PLAFOND_REMISE_GENERIQUES}")
        ws.cell(row=row, column=8, value=f"=MAX(0,G{row}-E{row})")

    last_row = first_data_row + len(periodes) - 1
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL")
    for col in (2, 3, 4, 5):
        letter = get_column_letter(col)
        ws.cell(row=total_row, column=col,
                value=f"=SUM({letter}{first_data_row}:{letter}{last_row})")
    ws.cell(row=total_row, column=6, value=f"=IFERROR(E{total_row}/D{total_row},0)")
    ws.cell(row=total_row, column=7, value=f"=D{total_row}*{PLAFOND_REMISE_GENERIQUES}")
    ws.cell(row=total_row, column=8, value=f"=MAX(0,G{total_row}-E{total_row})")

    for row in range(first_data_row, total_row + 1):
        is_total = (row == total_row)
        for col in (2, 3, 4, 5, 7, 8):
            ws.cell(row=row, column=col).number_format = EUR_FMT
        ws.cell(row=row, column=6).number_format = PCT_FMT
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = BOLD_FONT if is_total else NORMAL_FONT
            cell.border = THIN_BORDER
            if is_total:
                cell.fill = TOTAL_FILL

    # Rappel des références exclues, directement sous le tableau, pour que la
    # ligne "dont exclu" soit lisible sans ouvrir le code.
    row = total_row + 2
    ws.cell(row=row, column=1,
            value="Références Biogaran exclues de l'assiette "
                  f"({MOTIF_EXCLUSION_REMISE.lower()})").font = BOLD_FONT
    row += 1
    for cip13, libelle in sorted(EXCLUSIONS_BASE_REMISE.items(), key=lambda kv: kv[1]):
        ws.cell(row=row, column=1, value=cip13).font = NORMAL_FONT
        ws.cell(row=row, column=2, value=libelle).font = NORMAL_FONT
        row += 1

    widths = {"A": 16, "B": 24, "C": 26, "D": 18, "E": 22, "F": 20, "G": 20, "H": 24}
    autosize(ws, widths)
    ws.freeze_panes = f"B{first_data_row}"
    return ws


def build_sheet_assiette_eg(wb, all_rows, n_detail):
    """Assiette de remise EG LABO, mois par mois, tous canaux : CA EG par
    catégorie (répertoire, hybride, autres remboursés, non remboursé), moins
    les exclusions (calcium D3 et nicotine remboursables), remise facturée et
    taux pondéré sur l'assiette. Formules sur 'Détail lignes'."""
    periodes = sorted({r["periode"] for r in all_rows if r.get("genericqueur") == GENERIQUEUR_EG},
                      key=periode_sort_key)
    if not periodes:
        return
    ws = wb.create_sheet("Assiette remise EG")
    ws["A1"] = "Assiette de remise EG LABO — détail mensuel, tous canaux"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Assiette = produits EG remboursables (génériques du répertoire, hybrides, autres "
                "remboursés) moins le calcium vitamine D3 et la nicotine remboursables. Les produits non "
                "remboursés sont hors assiette. Lignes et motifs : colonnes U/V de 'Détail lignes'.")
    ws["A2"].font = SUBTITLE_FONT
    D = "'Détail lignes'"
    rg = lambda c: f"{D}!${c}$2:${c}${n_detail}"
    PER, LAB, TYP, REM, CA, RMS, EXC, CAN = (rg(c) for c in "ADFHLNUR")
    NR = "Non remboursé (ou non renseigné)"
    base = f'{LAB},"{GENERIQUEUR_EG}"'
    headers = ["Période", "CA EG total (€)", "dont non remboursé — hors assiette (€)",
               "Générique répertoire remboursable (€)", "Hybride remboursable (€)",
               "Autres remboursables (€)", "dont exclu : calcium D3 / nicotine remb. (€)",
               "Assiette retenue (€)", "Remise facturée sur assiette (€)", "Taux de remise pondéré"]
    hr = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=hr, column=i, value=h)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    style_header_row(ws, hr, len(headers))
    ws.row_dimensions[hr].height = 58
    first = hr + 1
    for i, per in enumerate(periodes):
        row = first + i
        pc = f'{PER},"{per}"'
        ws.cell(row=row, column=1, value=per)
        ws.cell(row=row, column=2, value=f"=SUMIFS({CA},{base},{pc})")
        ws.cell(row=row, column=3, value=f'=SUMIFS({CA},{base},{pc},{REM},"{NR}")')
        ws.cell(row=row, column=4, value=f'=SUMIFS({CA},{base},{pc},{TYP},"Générique (répertoire)",{REM},"<>{NR}")')
        ws.cell(row=row, column=5, value=f'=SUMIFS({CA},{base},{pc},{TYP},"Hybride",{REM},"<>{NR}")')
        ws.cell(row=row, column=6, value=f"=B{row}-C{row}-D{row}-E{row}")
        ws.cell(row=row, column=7, value=f'=SUMIFS({CA},{base},{pc},{EXC},"{EXCL_OUI}")')
        ws.cell(row=row, column=8, value=f"=D{row}+E{row}+F{row}-G{row}")
        ws.cell(row=row, column=9, value=f'=SUMIFS({RMS},{base},{pc},{REM},"<>{NR}",{EXC},"{EXCL_NON}")')
        ws.cell(row=row, column=10, value=f"=IFERROR(I{row}/H{row},0)")
    last = first + len(periodes) - 1
    tot = last + 1
    ws.cell(row=tot, column=1, value="TOTAL")
    for col in range(2, 10):
        L = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({L}{first}:{L}{last})")
    ws.cell(row=tot, column=10, value=f"=IFERROR(I{tot}/H{tot},0)")
    for row in range(first, tot + 1):
        for col in range(1, len(headers) + 1):
            c = ws.cell(row=row, column=col)
            c.font = BOLD_FONT if row == tot else NORMAL_FONT
            c.border = THIN_BORDER
            if row == tot:
                c.fill = TOTAL_FILL
            if 2 <= col <= 9:
                c.number_format = EUR_FMT
        ws.cell(row=row, column=10).number_format = PCT_FMT

    # Assiette par canal
    canaux = sorted({r["canal"] for r in all_rows if r.get("genericqueur") == GENERIQUEUR_EG})
    row = tot + 2
    ws.cell(row=row, column=1, value="Assiette retenue par canal (€)").font = BOLD_FONT
    row += 1
    for i, h in enumerate(["Canal", "Assiette retenue (€)", "Remise facturée (€)", "Taux pondéré"], start=1):
        ws.cell(row=row, column=i, value=h)
    style_header_row(ws, row, 4)
    for canal in canaux:
        row += 1
        cc = f'{CAN},"{canal}"'
        ws.cell(row=row, column=1, value=canal)
        ws.cell(row=row, column=2, value=f'=SUMIFS({CA},{base},{cc},{REM},"<>{NR}",{EXC},"{EXCL_NON}")')
        ws.cell(row=row, column=3, value=f'=SUMIFS({RMS},{base},{cc},{REM},"<>{NR}",{EXC},"{EXCL_NON}")')
        ws.cell(row=row, column=4, value=f"=IFERROR(C{row}/B{row},0)")
        for col in range(1, 5):
            ws.cell(row=row, column=col).font = NORMAL_FONT
            ws.cell(row=row, column=col).border = THIN_BORDER
        ws.cell(row=row, column=2).number_format = EUR_FMT
        ws.cell(row=row, column=3).number_format = EUR_FMT
        ws.cell(row=row, column=4).number_format = PCT_FMT

    # Références exclues rencontrées
    exclues = {}
    for r in all_rows:
        if r.get("motif_exclusion_eg"):
            exclues.setdefault((r["motif_exclusion_eg"], r["cip13"]), r.get("libelle", ""))
    row += 2
    ws.cell(row=row, column=1, value="Références EG exclues de l'assiette (rencontrées)").font = BOLD_FONT
    if not exclues:
        row += 1
        ws.cell(row=row, column=1, value="Aucune sur la période.").font = NORMAL_FONT
    for (motif, cip), lib in sorted(exclues.items()):
        row += 1
        ws.cell(row=row, column=1, value=cip).font = NORMAL_FONT
        ws.cell(row=row, column=2, value=lib).font = NORMAL_FONT
        ws.cell(row=row, column=4, value=motif).font = NORMAL_FONT
    autosize(ws, {"A": 22, "B": 16, "C": 18, "D": 18, "E": 16, "F": 16, "G": 20, "H": 16, "I": 18, "J": 14})
    ws.freeze_panes = f"B{first}"


def build_sheet_rattrapage_teva(wb, all_rows, n_detail):
    """Rattrapage de remise TEVA par trimestre, canal (direct / indirect) et
    palier de remise facturée : CA brut, remise facturée, taux cible et
    compensation attendue = CA x cible - remise facturée."""
    eligibles = [r for r in all_rows if palier_teva(r)]
    if not eligibles:
        return
    ws = wb.create_sheet("Rattrapage Teva")
    ws["A1"] = "Rattrapage de remise TEVA — compensation attendue par trimestre"
    ws["A1"].font = TITLE_FONT
    cibles = ", ".join(f"{p:g} % → {c * 100:g} %".replace(".", ",") for p, c in TEVA_TAUX_CIBLE.items())
    ws["A2"] = ("Génériques TEVA du répertoire, remboursables, tous canaux. Taux cible selon la remise "
                f"facturée : {cibles}. Compensation attendue = CA brut × taux cible − remise facturée. "
                "« Direct » = factures Teva ; « Indirect » = achats via grossiste (OCP, Alliance). "
                "Les avoirs Teva reçus ne sont pas intégrés : à rapprocher manuellement.")
    ws["A2"].font = SUBTITLE_FONT
    D = "'Détail lignes'"
    rg = lambda c: f"{D}!${c}$2:${c}${n_detail}"
    CA, RMS, CAN, PAL, TRI = rg("L"), rg("N"), rg("R"), rg("W"), rg("X")
    trimestres = sorted({trimestre_de(r["periode"]) for r in eligibles})
    paliers = [f"{p:g} %".replace(".", ",") for p in TEVA_TAUX_CIBLE] + ["autre taux"]
    headers = ["Trimestre", "Canal", "Remise facturée (palier)", "CA brut (€)", "Remise facturée (€)",
               "Taux cible", "Remise cible (€)", "Compensation attendue (€)"]
    hr = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=hr, column=i, value=h)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    style_header_row(ws, hr, len(headers))
    ws.row_dimensions[hr].height = 40
    row = hr
    lignes_tot = []
    for t in trimestres:
        for groupe, crit in (("Direct", f'{CAN},"Teva Direct"'), ("Indirect", f'{CAN},"<>Teva Direct"')):
            for pal in paliers:
                n = sum(1 for r in eligibles if trimestre_de(r["periode"]) == t and palier_teva(r) == pal
                        and ((r["canal"] == "Teva Direct") == (groupe == "Direct")))
                if not n:
                    continue
                row += 1
                lignes_tot.append(row)
                base = f'{TRI},"{t}",{crit},{PAL},"{pal}"'
                ws.cell(row=row, column=1, value=t)
                ws.cell(row=row, column=2, value=groupe)
                ws.cell(row=row, column=3, value=pal)
                ws.cell(row=row, column=4, value=f"=SUMIFS({CA},{base})")
                ws.cell(row=row, column=5, value=f"=SUMIFS({RMS},{base})")
                cible = next((c for p, c in TEVA_TAUX_CIBLE.items()
                              if f"{p:g} %".replace(".", ",") == pal), None)
                if cible is not None:
                    ws.cell(row=row, column=6, value=cible)
                    ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
                    ws.cell(row=row, column=8, value=f"=MAX(0,G{row}-E{row})")
                else:
                    ws.cell(row=row, column=6, value="—")
                    ws.cell(row=row, column=8, value=0)
                for col in range(1, len(headers) + 1):
                    c = ws.cell(row=row, column=col)
                    c.font = NORMAL_FONT
                    c.border = THIN_BORDER
                for col in (4, 5, 7, 8):
                    ws.cell(row=row, column=col).number_format = EUR_FMT
                ws.cell(row=row, column=6).number_format = PCT_FMT
    tot = row + 1
    ws.cell(row=tot, column=1, value="TOTAL")
    for col in (4, 5, 7, 8):
        L = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({L}{hr + 1}:{L}{row})")
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=tot, column=col)
        c.font, c.fill, c.border = BOLD_FONT, TOTAL_FILL, THIN_BORDER
        if col in (4, 5, 7, 8):
            c.number_format = EUR_FMT
    autosize(ws, {"A": 12, "B": 11, "C": 16, "D": 15, "E": 17, "F": 11, "G": 15, "H": 19})
    ws.freeze_panes = f"A{hr + 1}"


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
        ws.cell(row=total_row, column=8, value=f"=SUM(H{start_row+1}:H{last_row})").font = BOLD_FONT
        ws.cell(row=total_row, column=8).number_format = EUR_FMT

        # --- Sous-totaux par laboratoire (nouveau), affichés sur la même page,
        # à droite du tableau principal pour rester visibles sans scroller ---
        labos = sorted({r["genericqueur"] for r in fuite_rows})
        sub_col = len(headers) + 2  # 2 colonnes de marge après le tableau principal
        ws.cell(row=start_row, column=sub_col, value="Laboratoire")
        ws.cell(row=start_row, column=sub_col + 1, value="CA brut (€)")
        ws.cell(row=start_row, column=sub_col + 2, value="CA remisé (€)")
        style_header_row(ws, start_row, 3, col_start=sub_col)

        rng_acheteur = f"$D${start_row+1}:$D${last_row}"
        rng_ca_brut = f"$G${start_row+1}:$G${last_row}"
        rng_remise = f"$H${start_row+1}:$H${last_row}"
        sub_row = start_row + 1
        for labo in labos:
            labo_ref = f"{get_column_letter(sub_col)}{sub_row}"
            ws.cell(row=sub_row, column=sub_col, value=labo)
            ws.cell(row=sub_row, column=sub_col + 1,
                    value=f'=SUMIF({rng_acheteur},{labo_ref},{rng_ca_brut})')
            ws.cell(row=sub_row, column=sub_col + 2,
                    value=f'=SUMIF({rng_acheteur},{labo_ref},{rng_ca_brut})-SUMIF({rng_acheteur},{labo_ref},{rng_remise})')
            ws.cell(row=sub_row, column=sub_col + 1).number_format = EUR_FMT
            ws.cell(row=sub_row, column=sub_col + 2).number_format = EUR_FMT
            for c in (sub_col, sub_col + 1, sub_col + 2):
                ws.cell(row=sub_row, column=c).font = NORMAL_FONT
            sub_row += 1
        sub_total_row = sub_row
        ws.cell(row=sub_total_row, column=sub_col, value="TOTAL").font = BOLD_FONT
        for c in (sub_col + 1, sub_col + 2):
            letter = get_column_letter(c)
            ws.cell(row=sub_total_row, column=c,
                    value=f"=SUM({letter}{start_row+1}:{letter}{sub_row-1})").font = BOLD_FONT
            ws.cell(row=sub_total_row, column=c).number_format = EUR_FMT

        autosize(ws, {get_column_letter(sub_col): 20, get_column_letter(sub_col+1): 13, get_column_letter(sub_col+2): 13})
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
    periodes = sorted({r["periode"] for r in all_rows}, key=periode_sort_key)
    ws["A2"] = f"Période(s) analysée(s) : {', '.join(periodes)}"
    ws["A2"].font = SUBTITLE_FONT

    total_ca = sum(r["ca_ppht"] for r in all_rows)
    total_remise = sum(r["remise_ht"] for r in all_rows)
    total_lignes = len(all_rows)
    total_hybrides = sum(1 for r in all_rows if r["type_medicament"] == "Hybride")

    kpis = [
        ("Nombre de lignes produit analysées", total_lignes),
        ("CA PPHT total (€)", total_ca),
        ("Remise HT totale (€)", total_remise),
        ("Nombre de lignes hybrides identifiées", total_hybrides),
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
        "→ 'Détail lignes' : chaque ligne produit enrichie (type, remboursement, hybride, fabricant)",
        "→ 'CA par génériqueur' : synthèse CA/remise par génériqueur, toutes périodes confondues",
        "→ 'CA type x laboratoire' : détail mensuel + total annuel par type de médicament et labo",
        "→ 'Assiette remise Biogaran' : assiette de la remontée de remise, hors références "
        "entrées au répertoire en cours d'année",
        "→ 'Fuite Biogaran' : achats faits ailleurs alors que Biogaran couvre le même générique, "
        "avec sous-totaux par laboratoire",
        "→ 'Méthodologie' : règles et limites de l'analyse (à lire avant interprétation)",
    ]:
        ws.cell(row=row, column=1, value=txt).font = NORMAL_FONT
        row += 1

    autosize(ws, {"A": 60, "B": 18})
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
        "CIS_CIP_bdpm.txt, CIS_GENER_bdpm.txt, ainsi que le registre ANSM des groupes "
        "hybrides (fic02grp.txt, fic03spe.txt).",
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
        "TYPE DE MÉDICAMENT",
        "- 'Générique (répertoire)' / 'Princeps' / 'Hors répertoire' : déterminé via le "
        "fichier CIS_GENER_bdpm.txt (groupes génériques officiels ANSM).",
        "- 'Hybride' : déterminé via le registre ANSM des groupes hybrides. Un hybride "
        "est traité comme une précision de 'Hors répertoire' (les hybrides ne sont pas "
        "au répertoire des génériques) — il ressort comme catégorie 'Hybride' distincte "
        "dans ce rapport plutôt que noyé dans 'Hors répertoire'.",
        "- Remboursement : taux issu de CIS_CIP_bdpm.txt (colonne officielle). "
        "Si la BDPM ne donne aucun taux pour un générique du répertoire facturé à 2,1 % de TVA, "
        "le taux commun du groupe générique est repris, avec la mention « déduit du groupe "
        "générique » (cas constaté : CLOPIDOGREL EG 75 mg B/30). "
        "'Non remboursé (ou non renseigné)' regroupe les deux cas faute de pouvoir les "
        "distinguer de façon fiable avec ces seuls fichiers.",
        "",
        "FACTURES DIRECTES DES LABORATOIRES (hors Biogaran)",
        "- EG LABO : factures émises par CSP / Movianto pour le compte d'EG (canal « EG Direct »), "
        "contrôlées au centime sur le total net HT de chaque facture. Type de médicament "
        "(répertoire, hors répertoire, hybride, princeps) et remboursement via la BDPM, comme "
        "pour les autres canaux.",
        "- SANDOZ : factures Sandoz (canal « Sandoz Direct »), contrôlées au centime. Sandoz "
        "n'imprime pas le PFHT (« NaN ») : CA PPHT = montant net facturé, remise non connue. "
        "Les avoirs Sandoz (conditions commerciales) ne sont pas pris en compte.",
        "- TEVA : factures Teva Santé (canal « Teva Direct »), deux mises en page (avant / depuis "
        "septembre 2026), contrôlées au centime.",
        "- Avoirs EG (« remise additionnelle »), Sandoz et Teva (ristournes) : ignorés pour l'instant.",
        "",
        "ASSIETTE DE REMISE EG",
        "- Produits EG remboursables, tous canaux (génériques du répertoire, hybrides, autres "
        "remboursés), moins le calcium vitamine D3 et la nicotine remboursables (règle par "
        "mots-clés : EXCLUSIONS_ASSIETTE_EG). Onglet 'Assiette remise EG'.",
        "",
        "RATTRAPAGE TEVA",
        "- Génériques TEVA du répertoire remboursables, tous canaux. Remise facturée 10/20/30 % → "
        "cible 39 % ; 2,5 % → cible 12,5 % (TEVA_TAUX_CIBLE). Compensation attendue = CA brut × "
        "cible − remise facturée, par trimestre, direct / indirect. Onglet 'Rattrapage Teva'.",
        "",
        "BIOSIMILAIRES",
        "- Dans les onglets issus des factures, les biosimilaires ne sont pas distingués : "
        "ils tombent dans 'Hors répertoire'.",
        "- Onglet 'Biosimilaires' : achats directs de biosimilaires (Sandoz...) identifiés par "
        "leur marque dans le relevé OCP, ajoutés au relevé OCP (colonne Source).",
        "- Relevé OCP (si le relevé OCP « Achats des produits "
        "BIOSIMILAIRES » est fourni) : CA net par molécule, laboratoire et trimestre, "
        "biosimilaires et bioréférents, tous achats OCP avec ou sans offre. Seul le relevé "
        "le plus récent de chaque année est retenu (cumul annuel, périmètre de molécules "
        "élargi rétroactivement par OCP). Contrôlé à chaque lecture contre les totaux "
        "imprimés du relevé. Non additionné aux autres onglets.",
        "",
        "CALCUL DU CA REMISÉ (hypothèse à confirmer)",
        "- CA remisé = CA PPHT − Remise HT. C'est l'hypothèse de calcul retenue pour "
        "l'onglet 'CA type x laboratoire' en l'absence de précision contraire — à "
        "confirmer auprès du pharmacien si un calcul différent (ex. sur la base du "
        "CA PGHT) est attendu.",
        "- Taux de remise pondéré = Remise HT / CA PPHT (donc pondéré par le volume "
        "d'achat, pas une moyenne simple des taux ligne à ligne).",
        "",
        "ASSIETTE DE LA REMONTÉE DE REMISE (RDP) BIOGARAN",
        "- La remise cumulée sur les génériques remboursables est plafonnée à 40 % "
        "(art. L138-9 CSS) : remise facturée + remontée (RDP) ≤ 40 %. L'onglet "
        "'Assiette remise Biogaran' calcule, mois par mois, l'assiette retenue et la "
        "remontée théorique restant à percevoir à ce plafond.",
        "- EXCLUSION : les molécules Biogaran entrées au répertoire des génériques EN "
        "COURS D'ANNÉE sont déduites de cette assiette, sur toute l'année. Motif : les "
        "fichiers BDPM sont toujours à jour, donc ces références ressortent 'Générique "
        "(répertoire)' sur l'ensemble des périodes analysées, y compris sur les mois où "
        "elles n'y étaient pas encore — les inclure gonflerait artificiellement "
        "l'assiette.",
        "- Les lignes concernées restent présentes partout ailleurs dans le rapport : le "
        "CA total, la vue 'CA type x laboratoire' et l'analyse de fuite sont inchangés. "
        "Seule l'assiette remise les déduit. Le marquage est tracé ligne à ligne dans "
        "'Détail lignes' (colonnes 'Exclu base remise' et 'Motif exclusion base remise'), "
        "et la liste des références est rappelée sous le tableau de l'onglet "
        "'Assiette remise Biogaran'.",
        "- La remontée théorique affichée est un CALCUL au plafond, pas un montant dû "
        "constaté : elle doit être rapprochée des avoirs RDP réellement reçus de "
        "Biogaran.",
        "- Pour ajouter ou retirer une référence de l'exclusion, modifier le "
        "dictionnaire EXCLUSIONS_BASE_REMISE (scripts/analyse_consolidee.py) : tous les "
        "onglets et formules suivent automatiquement.",
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
        "- Les sous-totaux par laboratoire de l'onglet 'Fuite Biogaran' (CA brut / CA "
        "remisé) permettent de voir en un coup d'œil chez qui la fuite se concentre.",
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


# ---------------------------------------------------------------------------
# Biosimilaires — relevé OCP « Achats des produits BIOSIMILAIRES » (BIOST)
# ---------------------------------------------------------------------------
def charger_releves_biosimilaires(pdf_paths):
    """Lit les relevés BIOST et ne garde, pour chaque année, que le PLUS
    RÉCENT : chaque relevé est un cumul depuis janvier qui reprend le détail
    trimestre par trimestre, et OCP y ajoute en cours d'année des molécules
    rétroactivement (constaté en 2026 : dénosumab, follitropine alfa,
    ranibizumab absents du relevé T1, présents sur T1 dans le relevé T3).
    Additionner plusieurs relevés compterait donc deux fois les mêmes achats."""
    from parser_biosimilaires_ocp import parse_releve_biosimilaires
    par_annee = {}
    for chemin in pdf_paths or []:
        r = parse_releve_biosimilaires(chemin)
        r["fichier"] = Path(chemin).name
        if not r["annee"]:
            print(f"  ⚠️  Relevé biosimilaires illisible (année introuvable) : {r['fichier']}",
                  file=sys.stderr)
            continue
        cle = (r["mois_fin"] or 0, r["edition"] or "")
        garde = par_annee.get(r["annee"])
        if garde is None or cle > (garde["mois_fin"] or 0, garde["edition"] or ""):
            if garde is not None:
                print(f"  Relevé biosimilaires {garde['fichier']} remplacé par le plus récent "
                      f"{r['fichier']} (cumul annuel)")
            par_annee[r["annee"]] = r
        else:
            print(f"  Relevé biosimilaires {r['fichier']} ignoré : un relevé plus récent "
                  f"de {r['annee']} est fourni ({garde['fichier']})")
    releves = [par_annee[a] for a in sorted(par_annee)]
    for r in releves:
        for c in r["controles"]:
            print(f"  ⚠️  Biosimilaires {r['fichier']} : {c}", file=sys.stderr)
    return releves


def _marque(designation):
    m = re.match(r"[A-Z]+", (designation or "").upper())
    return m.group(0) if m else ""


def biosimilaires_des_factures_directes(releves, all_rows):
    """Lignes des factures directes (canaux « ... Direct » hors Biogaran) qui
    sont des biosimilaires / bioréférents. Identification par la MARQUE (1er
    mot de la désignation : HYRIMOZ, ZIEXTENZO...) rapprochée des produits du
    relevé OCP, qui donne molécule, statut et laboratoire. Sans relevé OCP, pas
    de référentiel : rien n'est identifié (signalé)."""
    ref = {}
    for r in releves:
        for l in r["lignes"]:
            ref.setdefault(_marque(l["designation"]), l)
    directs = [r for r in all_rows
               if r.get("canal", "").endswith(" Direct") and r["canal"] != "Biogaran Direct"]
    res = []
    for r in directs:
        l = ref.get(_marque(r["libelle"]))
        if not l:
            continue
        mm, aaaa = r["periode"].split("/") if "/" in r["periode"] else ("0", "0")
        res.append({
            "molecule": l["molecule"], "statut": l["statut"], "laboratoire": l["laboratoire"],
            "designation": r["libelle"], "cip13": r["cip13"],
            "trimestre": f"{aaaa} T{(int(mm) - 1) // 3 + 1}",
            "qte": r["qte_fact"], "ca_pght": None, "ca_net": round(r["ca_ppht"] - r["remise_ht"], 2),
            "source": r["canal"],
        })
    if directs and not ref:
        print("  ⚠️  Factures directes fournies sans relevé OCP des biosimilaires : les "
              "biosimilaires achetés en direct ne peuvent pas être identifiés.", file=sys.stderr)
    return res


def build_sheet_biosimilaires_ocp(wb, releves, lignes_directes=None):
    """Onglet « Biosimilaires » : synthèse par molécule (biosimilaire /
    bioréférent, part biosimilaire), par laboratoire, puis détail produit x
    trimestre. Sources : relevé OCP (achats OCP) + factures directes des
    laboratoires. Synthèses en formules SUMIFS sur le détail du même onglet."""
    lignes_directes = lignes_directes or []
    if not releves and not lignes_directes:
        return
    lignes = []
    for r in releves:
        lignes.extend({**l, "source": "Relevé OCP"} for l in r["lignes"])
    lignes.extend(lignes_directes)
    trimestres = sorted({l["trimestre"] for l in lignes})
    molecules = sorted({l["molecule"] for l in lignes})
    labos = sorted({l["laboratoire"] for l in lignes})

    ws = wb.create_sheet("Biosimilaires")
    ws["A1"] = "Biosimilaires et bioréférents — achats OCP et achats directs"
    ws["A1"].font = TITLE_FONT
    sources = " ; ".join(
        f"{r['fichier']} (cumul janvier → mois {r['mois_fin']:02d}/{r['annee']}, édité le {r['edition']})"
        for r in releves) or "aucun relevé fourni"
    canaux_directs = sorted({l["source"] for l in lignes_directes})
    ws["A2"] = ("Sources : relevé OCP « Achats des produits BIOSIMILAIRES » — " + sources + ". "
                + (f"Factures directes ({', '.join(canaux_directs)}) : montant net facturé, "
                   "trimestre de la date de facture, produit identifié par sa marque dans le relevé OCP. "
                   if canaux_directs else "")
                + "CA net = CA PPHT − remise. Le relevé OCP couvre tous les achats OCP de ces produits, "
                "avec ou sans offre (la plupart ne figurent pas dans le BO-OFFREM) ; ses montants ne sont "
                "PAS additionnés aux autres onglets. Les achats directs, eux, figurent aussi dans "
                "'Détail lignes'. Colonne « Source » du détail pour filtrer.")
    ws["A2"].font = SUBTITLE_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 58
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4 + 2 * len(trimestres))

    # Emplacement du détail (en bas de l'onglet) : connu à l'avance pour les formules
    n_mol, n_lab = len(molecules), len(labos)
    row_mol_header = 5
    row_lab_title = row_mol_header + n_mol + 4
    row_lab_header = row_lab_title + 1
    row_det_title = row_lab_header + n_lab + 4
    row_det_header = row_det_title + 1
    first_det = row_det_header + 1
    last_det = first_det + len(lignes) - 1

    def rng(col):
        return f"${col}${first_det}:${col}${last_det}"
    # Colonnes du détail : A Molécule, B Statut, C Laboratoire, D Désignation,
    # E CIP13, F Trimestre, G Qté, H CA PGHT, I CA net
    R_MOL, R_STAT, R_LAB, R_TRIM, R_QTE, R_CA = rng("A"), rng("B"), rng("C"), rng("F"), rng("G"), rng("I")

    # --- 1. Par molécule -------------------------------------------------
    ws.cell(row=row_mol_header - 1, column=1, value="1. Par molécule — CA net (€)").font = BOLD_FONT
    headers = (["Molécule"] + [f"Biosimilaire {t}" for t in trimestres]
               + ["Biosimilaire total", "Bioréférent total", "Total", "Part biosimilaire"])
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row_mol_header, column=i, value=h)
    style_header_row(ws, row_mol_header, len(headers))
    ws.row_dimensions[row_mol_header].height = 32
    nT = len(trimestres)
    for k, mol in enumerate(molecules):
        row = row_mol_header + 1 + k
        ws.cell(row=row, column=1, value=mol)
        for j, t in enumerate(trimestres):
            ws.cell(row=row, column=2 + j,
                    value=f'=SUMIFS({R_CA},{R_MOL},$A{row},{R_STAT},"Biosimilaire",{R_TRIM},"{t}")')
        c_bs, c_br, c_tot, c_part = 2 + nT, 3 + nT, 4 + nT, 5 + nT
        ws.cell(row=row, column=c_bs, value=f'=SUMIFS({R_CA},{R_MOL},$A{row},{R_STAT},"Biosimilaire")')
        ws.cell(row=row, column=c_br, value=f'=SUMIFS({R_CA},{R_MOL},$A{row},{R_STAT},"Bioréférent")')
        L_bs, L_br, L_tot = (get_column_letter(c) for c in (c_bs, c_br, c_tot))
        ws.cell(row=row, column=c_tot, value=f"={L_bs}{row}+{L_br}{row}")
        ws.cell(row=row, column=c_part, value=f"=IFERROR({L_bs}{row}/{L_tot}{row},0)")
    tot_row = row_mol_header + 1 + n_mol
    ws.cell(row=tot_row, column=1, value="TOTAL")
    for c in range(2, 5 + nT):
        L = get_column_letter(c)
        ws.cell(row=tot_row, column=c, value=f"=SUM({L}{row_mol_header + 1}:{L}{tot_row - 1})")
    L_bs, L_tot = get_column_letter(2 + nT), get_column_letter(4 + nT)
    ws.cell(row=tot_row, column=5 + nT, value=f"=IFERROR({L_bs}{tot_row}/{L_tot}{tot_row},0)")
    for row in range(row_mol_header + 1, tot_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = BOLD_FONT if row == tot_row else NORMAL_FONT
            cell.border = THIN_BORDER
            if row == tot_row:
                cell.fill = TOTAL_FILL
            if 2 <= c <= 4 + nT:
                cell.number_format = EUR_FMT
            elif c == 5 + nT:
                cell.number_format = PCT_FMT

    # --- 2. Par laboratoire ---------------------------------------------
    ws.cell(row=row_lab_title, column=1, value="2. Par laboratoire — CA net (€)").font = BOLD_FONT
    headers_l = ["Laboratoire"] + list(trimestres) + ["Total", "dont biosimilaire", "Qté"]
    for i, h in enumerate(headers_l, start=1):
        ws.cell(row=row_lab_header, column=i, value=h)
    style_header_row(ws, row_lab_header, len(headers_l))
    for k, lab in enumerate(labos):
        row = row_lab_header + 1 + k
        ws.cell(row=row, column=1, value=lab)
        for j, t in enumerate(trimestres):
            ws.cell(row=row, column=2 + j, value=f'=SUMIFS({R_CA},{R_LAB},$A{row},{R_TRIM},"{t}")')
        ws.cell(row=row, column=2 + nT, value=f'=SUMIFS({R_CA},{R_LAB},$A{row})')
        ws.cell(row=row, column=3 + nT, value=f'=SUMIFS({R_CA},{R_LAB},$A{row},{R_STAT},"Biosimilaire")')
        ws.cell(row=row, column=4 + nT, value=f'=SUMIFS({R_QTE},{R_LAB},$A{row})')
    tot_l = row_lab_header + 1 + n_lab
    ws.cell(row=tot_l, column=1, value="TOTAL")
    for c in range(2, 5 + nT):
        L = get_column_letter(c)
        ws.cell(row=tot_l, column=c, value=f"=SUM({L}{row_lab_header + 1}:{L}{tot_l - 1})")
    for row in range(row_lab_header + 1, tot_l + 1):
        for c in range(1, len(headers_l) + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = BOLD_FONT if row == tot_l else NORMAL_FONT
            cell.border = THIN_BORDER
            if row == tot_l:
                cell.fill = TOTAL_FILL
            if 2 <= c <= 3 + nT:
                cell.number_format = EUR_FMT

    # --- 3. Détail produit x trimestre ------------------------------------
    ws.cell(row=row_det_title, column=1, value="3. Détail par produit et par trimestre").font = BOLD_FONT
    headers_d = ["Molécule", "Statut", "Laboratoire", "Désignation", "CIP13", "Trimestre",
                 "Qté", "CA PGHT (€)", "CA net (€)", "Source"]
    for i, h in enumerate(headers_d, start=1):
        ws.cell(row=row_det_header, column=i, value=h)
    style_header_row(ws, row_det_header, len(headers_d))
    for k, l in enumerate(sorted(lignes, key=lambda x: (x["molecule"], x["statut"], x["designation"], x["trimestre"]))):
        row = first_det + k
        vals = [l["molecule"], l["statut"], l["laboratoire"], l["designation"], l["cip13"],
                l["trimestre"], l["qte"], l["ca_pght"], l["ca_net"], l["source"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.font = NORMAL_FONT
            if c in (8, 9):
                cell.number_format = EUR_FMT
    tab = Table(displayName="Biosimilaires", ref=f"A{row_det_header}:J{last_det}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    autosize(ws, {"A": 24, "B": 14, "C": 26, "D": 34, "E": 15, "F": 12, "G": 12,
                  "H": 13, "I": 13, "J": 15, **{get_column_letter(c): 14 for c in range(11, 6 + nT + 1)}})
    ws.freeze_panes = "B4"


def build_report(pdf_paths, data_dir, out_path, biogaran_direct_pdfs=None, alliance_pdfs=None,
                 biosimilaires_ocp_pdfs=None, direct_labo_pdfs=None):
    bdpm = BDPM(data_dir)
    releves_biosim = charger_releves_biosimilaires(biosimilaires_ocp_pdfs)
    all_rows = enrich_rows(pdf_paths, bdpm, biogaran_direct_pdfs=biogaran_direct_pdfs, alliance_pdfs=alliance_pdfs)
    rows_direct, _ = lire_factures_direct_labo(direct_labo_pdfs, bdpm)
    all_rows.extend(rows_direct)
    catalog_biogaran = compute_biogaran_catalog(all_rows)
    fuite_rows = compute_fuite(all_rows, catalog_biogaran)
    periodes = sorted({r["periode"] for r in all_rows}, key=periode_sort_key)

    wb = Workbook()
    wb.remove(wb.active)

    genericqueurs = sorted(
        {v[0] for v in OFFRE_TO_GENERIQUEUR.values() if v[1] is True}
        | {r["genericqueur"] for r in all_rows if r["est_genericqueur"] is True and r["genericqueur"]}
    )
    labo_fills = build_labo_fill_map(genericqueurs)

    if all_rows:
        _, n_detail = build_sheet_detail(wb, all_rows)
        build_sheet_genericqueur(wb, n_detail, labo_fills, genericqueurs)
        build_sheet_type_labo(wb, all_rows, n_detail, periodes, labo_fills)
        build_sheet_assiette_remise(wb, all_rows, n_detail)
        build_sheet_assiette_eg(wb, all_rows, n_detail)
        build_sheet_rattrapage_teva(wb, all_rows, n_detail)
        _, n_fuite = build_sheet_fuite(wb, all_rows, catalog_biogaran)
        build_sheet_resume(wb, all_rows, n_fuite)
    else:
        # Seul un relevé biosimilaires a été fourni : pas de facture, donc pas
        # d'onglets de CA / fuite (ils ne contiendraient que des divisions par 0).
        print("  Aucune facture fournie : seuls les onglets biosimilaires et méthodologie sont produits.")
    build_sheet_biosimilaires_ocp(wb, releves_biosim,
                                  biosimilaires_des_factures_directes(releves_biosim, all_rows))
    build_sheet_methodologie(wb)

    wb.save(out_path)
    print(f"Rapport généré : {out_path}")
    print(f"  {len(all_rows)} lignes produit, {len(fuite_rows)} lignes de fuite détectées")
    print(f"  Périodes : {', '.join(periodes)}")
    print(f"  CA PPHT total : {sum(r['ca_ppht'] for r in all_rows):,.2f} €")

    # Contrôle de l'exclusion d'assiette : si un CIP13 listé dans
    # EXCLUSIONS_BASE_REMISE n'est jamais rencontré, autant le signaler plutôt
    # que de laisser croire que la déduction a été appliquée.
    lignes_exclues = [r for r in all_rows if r.get("exclu_base_remise")]
    ca_exclu = sum(r["ca_ppht"] for r in lignes_exclues)
    cips_vus = {cle_cip13(r["cip13"]) for r in lignes_exclues}
    ca_deduit = sum(r["ca_ppht"] for r in lignes_exclues
                    if r.get("genericqueur") == "BIOGARAN"
                    and r.get("type_medicament") == "Générique (répertoire)")
    print(f"  Assiette remise : {len(lignes_exclues)} ligne(s) exclue(s), "
          f"{ca_exclu:,.2f} € de CA PPHT au total, dont {ca_deduit:,.2f} € "
          f"classés Générique (répertoire) et donc déduits de l'assiette Biogaran")
    forces = sorted({cle_cip13(r["cip13"]) for r in lignes_exclues
                     if r.get("repertoire_force")})
    for cip in forces:
        print(f"  ℹ️  {cip} ({EXCLUSIONS_BASE_REMISE[cip]}) : absent du répertoire "
              "BDPM, classé Générique (répertoire) d'office (liste des exclusions)",
              file=sys.stderr)
    hors_rep = sorted({cle_cip13(r["cip13"]) for r in lignes_exclues
                       if r.get("genericqueur") == "BIOGARAN"
                       and r.get("type_medicament") != "Générique (répertoire)"})
    for cip in hors_rep:
        print(f"  ⚠️  Exclusion {cip} ({EXCLUSIONS_BASE_REMISE[cip]}) : classée "
              "hors répertoire — n'est pas dans l'assiette, rien à déduire",
              file=sys.stderr)
    cips_absents = sorted(set(EXCLUSIONS_BASE_REMISE) - cips_vus)
    if cips_absents:
        print(f"  ⚠️  {len(cips_absents)} référence(s) d'exclusion jamais rencontrée(s) "
              "dans les factures analysées :", file=sys.stderr)
        for cip in cips_absents:
            print(f"    {cip} — {EXCLUSIONS_BASE_REMISE[cip]}", file=sys.stderr)

    a_classifier = sorted({r["offre"] for r in all_rows if r["genericqueur"] == "À CLASSIFIER"})
    if a_classifier:
        print(f"  ⚠️  {len(a_classifier)} offre(s) non classifiée(s) dans OFFRE_TO_GENERIQUEUR :", file=sys.stderr)
        for o in a_classifier:
            print("   ", o, file=sys.stderr)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdfs", nargs="+", help="Un ou plusieurs BO-OFFREM PDF (canal OCP)")
    ap.add_argument("--data", default="data", help="Dossier des fichiers BDPM")
    ap.add_argument("--out", default="rapport_pharmacie.xlsx", help="Fichier Excel de sortie")
    ap.add_argument("--biogaran-direct", nargs="*", default=None,
                     help="Un ou plusieurs PDF de factures Biogaran Direct (canal Biogaran Direct)")
    ap.add_argument("--alliance", nargs="*", default=None,
                     help="Un ou plusieurs PDF de factures Alliance Healthcare (canal Alliance)")
    ap.add_argument("--biosimilaires-ocp", nargs="*", default=None,
                     help="Relevé(s) OCP « Achats des produits BIOSIMILAIRES » (BIOST)")
    ap.add_argument("--direct-labo", nargs="*", default=None,
                     help="Factures et avoirs directs des laboratoires (EG LABO...)")
    args = ap.parse_args()
    build_report(args.pdfs, args.data, args.out,
                  biogaran_direct_pdfs=args.biogaran_direct, alliance_pdfs=args.alliance,
                  biosimilaires_ocp_pdfs=args.biosimilaires_ocp, direct_labo_pdfs=args.direct_labo)
