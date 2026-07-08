# -*- coding: utf-8 -*-
"""
Parser pour les factures Alliance Healthcare (Cencora / "Facture" grossiste).

Format observé (texte tel qu'extrait par pdfplumber) : chaque produit est un
bloc de 4 lignes (parfois 5 si une réduction de prix ligne à ligne s'applique) :

    <DÉSIGNATION>
    Code Article : <CIP13 ou EAN>
    <N°> [Réduction de prix : X,XX %] <QTE> <UNITÉ> <PU_NET_HT> <TVA%> <MONTANT_NET_HT>
    Référence commande complémentaire : <réf>
    Prix brut unitaire : <PU_BRUT>

⚠️ Contrairement à OCP (BO-OFFREM), il n'y a PAS de champ "offre" identifiant
le génériqueur. La détection du génériqueur se fait donc par mots-clés sur la
désignation (suffixes de marque du type "BGR", "BIOG", "SDZ", "TVT"...),
répertoriés dans ALLIANCE_LABO_KEYWORDS ci-dessous. Cette liste de mots-clés
est reprise du filet de sécurité de l'ancien moteur_analyse.py (projet
analyse-factures-alliances), mais adaptée ici pour IDENTIFIER LA MARQUE
(pas seulement flaguer "est un générique" comme le faisait l'ancien moteur).

⚠️ Validé au centime uniquement sur 2 factures réelles à ce stade (21 lignes
au total, dont seulement 2 lignes génériqueur, toutes deux Biogaran). Les
mots-clés pour les autres génériqueurs (Sandoz, Teva, Viatris, etc.) sont
repris de l'ancien moteur mais n'ont pas encore été vérifiés sur des lignes
réelles de ce projet — à surveiller sur les prochaines factures.
"""
import os
import re
import pdfplumber

BLOC_PRODUIT_RE = re.compile(
    r"(?P<designation>.+)\n"
    r"Code Article\s*:\s*(?P<code_article>\d+)\n"
    r"(?:(?P<numero>\d+)\s+)?"
    r"(?:Réduction de prix\s*:\s*(?P<remise_pct>[\d.,]+)\s*%\s+)?"
    r"(?P<qte>\d+)\s+(?P<unite>\S+)\s+(?P<pu>[\d,]+)\s+(?P<tva>[\d,]+)\s*%\s+(?P<montant>[\d,]+)\n"
    r"(?:Référence commande complémentaire\s*:\s*(?P<ref_cmd>\S*)\n)?"
    r"Prix brut unitaire\s*:\s*(?P<prix_brut>[\d,]+)"
)

NUM_FACTURE_RE = re.compile(r"N°\s*de facture\s*:\s*(\S+)")
DATE_FACTURE_RE = re.compile(r"Date de facturation\s*:\s*(\d{2}/\d{2}/\d{4})")

# Mots-clés de désignation -> génériqueur. Reconstruits à partir de la liste
# LABO_GENERIQUES de l'ancien moteur (qui ne faisait que flaguer "générique",
# sans dire lequel) : ici on assume l'association marque <-> abréviation
# usuelle sur factures grossiste. Espaces volontaires pour éviter les faux
# positifs sur des syllabes internes à d'autres mots.
ALLIANCE_LABO_KEYWORDS = {
    "BIOGARAN": [" bgr", " brg", " biog", " biogar", " biogaran"],
    "SANDOZ": [" sdz", " sand ", " sandoz"],
    "TEVA": [" tvt", " tvb", " teva"],
    "VIATRIS": [" myl", " vtr", " viatris", " viatr", " mylan"],
    "ZENTIVA": [" znt", " zentiva", " zent "],
    "ACCORD HEALTHCARE": [" acrd", " accord"],
    "KRKA": [" krka", " kk "],
    "HIKMA": [" hkm", " hikma"],
    "APOTEX": [" apt", " apotex"],
    "CRISTERS": [" crs", " cris ", " cristers"],
    "ARROW": [" arr ", " arw", " arr gen", " arrow"],
    "SUBSTIPHARM": [" sfm", " substipharm"],
    "QUALIMED": [" qlm", " qualimed"],
    "PHPHARMA": [" php"],
    "ALMUS": [" alr", " almus"],
    "ZYDUS": [" zydus"],
    "STADA": [" stada"],
    "AUROBINDO": [" aur ", " aurobindo"],
    "SUN PHARMA": [" sun "],
}


def to_float(s):
    if s is None:
        return None
    return float(str(s).strip().replace(" ", "").replace(",", "."))


def identifier_genericqueur(designation):
    """Retourne (nom_genericqueur, est_genericqueur) à partir de la désignation.
    None/False si aucun mot-clé de marque générique reconnu (princeps, dispositif,
    parapharmacie, nutrition...)."""
    d = " " + designation.lower() + " "
    for labo, mots in ALLIANCE_LABO_KEYWORDS.items():
        for mot in mots:
            if mot in d:
                return labo, True
    return None, False


def extraire_facture_alliance(path_pdf):
    with pdfplumber.open(path_pdf) as pdf:
        # La dernière page est systématiquement les CGV (texte fixe, jamais de
        # tableau produit) -> on l'exclut pour ne pas gaspiller de temps dessus.
        text = "\n".join((p.extract_text() or "") for p in pdf.pages[:-1])

    num_m = NUM_FACTURE_RE.search(text)
    date_m = DATE_FACTURE_RE.search(text)
    num_facture = num_m.group(1) if num_m else ""
    date_facture = date_m.group(1) if date_m else ""

    lignes = []
    for m in BLOC_PRODUIT_RE.finditer(text):
        designation = re.sub(r"\s+", " ", m.group("designation")).strip()
        qte = int(m.group("qte"))
        pu_net = to_float(m.group("pu"))
        montant_net = to_float(m.group("montant"))
        prix_brut = to_float(m.group("prix_brut"))
        remise_pct = to_float(m.group("remise_pct")) if m.group("remise_pct") else 0.0
        tva = to_float(m.group("tva"))

        ca_ppht = round(qte * prix_brut, 2)          # CA brut (avant remise)
        remise_ht = round(ca_ppht - montant_net, 2)   # remise en valeur

        labo, est_gen = identifier_genericqueur(designation)

        lignes.append({
            "cip13": m.group("code_article"),
            "libelle": designation,
            "offre": labo,              # None si pas un génériqueur identifié
            "tva": tva,
            "qte_cdee": qte,
            "qte_fact": qte,
            "ppht_unit": prix_brut,
            "prix_net_unitaire": pu_net,
            "ca_ppht": ca_ppht,
            "ca_pght": montant_net,
            "taux_remise": remise_pct,
            "remise_ht": remise_ht,
            "canal": "Alliance",
            "est_genericqueur_detecte": est_gen,
            "num_facture": num_facture,
            "date_facture": date_facture,
        })

    return {"num_facture": num_facture, "date_facture": date_facture, "lignes": lignes}


def extraire_dossier(dossier_pdfs):
    toutes_lignes = []
    for nom in sorted(os.listdir(dossier_pdfs)):
        if not nom.lower().endswith(".pdf"):
            continue
        chemin = os.path.join(dossier_pdfs, nom)
        facture = extraire_facture_alliance(chemin)
        if not facture["lignes"]:
            print(f"⚠️  Aucune ligne produit trouvée dans {nom}")
        toutes_lignes.extend(facture["lignes"])
    return toutes_lignes


if __name__ == "__main__":
    import sys
    for path in sys.argv[1:]:
        f = extraire_facture_alliance(path)
        n_gen = sum(1 for l in f["lignes"] if l["est_genericqueur_detecte"])
        print(f"{path}: facture {f['num_facture']} du {f['date_facture']}, "
              f"{len(f['lignes'])} lignes ({n_gen} génériqueur détecté)")
