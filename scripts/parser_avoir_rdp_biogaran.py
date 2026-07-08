# -*- coding: utf-8 -*-
"""
Parser pour les avoirs de Remise de Performance (RDP) Biogaran
("RECAPITULATIF DES REMISES", fichiers avoir_*.pdf).

⚠️ Distinct des factures d'achat (facture_*.pdf, cf. parser_biogaran_direct.py) :
un avoir RDP est un agrégat MENSUEL, PAS une liste ligne par ligne par CIP.
Le détail des spécialités concernées existe chez Biogaran mais doit être
demandé séparément à l'officine (mention explicite sur le document).

Format observé (texte tel qu'extrait par pdfplumber, une seule page, colonnes
entrelacées par pdfplumber au sein d'un même paragraphe) :

    PÉRIODE DE RÉFÉRENCE
    Du 01/05/2026 au 31/05/2026
    ...
    N°Document : 9006827855
    Date document : 23/06/2026
    ...
    RDP 10.00% *** SUR
    ACHATS GROSSISTES** 2026 ou CPV, au titre des achats grossistes** réalisés sur la période de 1 10,00 -1 018,83 € 0,00
    référence (du 01/05/2026 au 31/05/2026).
    C.A brut de référence grossistes partenaires :10 187,17 €
    RDP 10.00% *** SUR
    ACHATS DIRECTS* ... réalisés sur la période de référence 1 10,00 -64,81 € 0,00
    (du 01/05/2026 au 31/05/2026).
    C.A brut de référence dépositaire :647,87 €
    ... (jusqu'à 4 blocs : 2 taux RDP (10%/20%, parfois 30% selon les mois)
        x 2 circuits d'achat)
    TOTAL -2 021,37 €

Un bloc peut être absent si le CA du circuit concerné est nul sur la période
pour ce taux (ex. avoir de février 2026 : seulement 3 blocs sur 4 possibles).

Circuits d'achat (astérisques du document) :
    - "GROSSISTES**" = achats combinés OCP + Alliance (répartiteurs pharmaceutiques)
    - "DIRECTS*"     = achats canal Biogaran Direct (dépositaires)

Règle métier confirmée par le pharmacien : la RDP ramène la remise cumulée à
40 % sur les médicaments Biogaran remboursables déjà remisés à 10/20/30 % sur
facture d'achat (plafond légal art. L138-9 CSS sur les génériques
remboursables) :
    taux_remise_facture_attendu = 40 - taux_rdp

Segmentation robuste par occurrence de "RDP X%" (plutôt que par adjacence de
mots-clés, fragile face à l'entrelacement de colonnes pdfplumber) : chaque
bloc va de "RDP X%" jusqu'à la ligne "C.A brut de référence ... : Y €" qui le
termine, capturés avec re.DOTALL sur le texte complet de la page.

⚠️ Validé au centime sur 5 avoirs réels (février à juin 2026 → périodes de
référence janvier à mai 2026) : somme des montants de remise HT de chaque
bloc = TOTAL imprimé sur le document, à chaque fois.
"""

import os
import re
import pdfplumber

BLOC_RE = re.compile(
    r"RDP\s+(?P<taux_rdp>[\d.,]+)\s*%.*?"
    r"ACHATS\s+(?P<circuit>GROSSISTES|DIRECTS)\*+.*?"
    r"(?P<qte>\d+)\s+(?P<taux_remise>[\d,]+)\s+(?P<montant_remise>-?[\d\s  ]*,\d+)\s*€?\s*(?P<tva_pct>[\d,]+).*?"
    r"C\.A brut de référence\s+(?:grossistes partenaires|dépositaire)\s*:\s*"
    r"(?P<ca_brut>[\d\s  ]*,\d+)\s*€",
    re.DOTALL,
)

PERIODE_RE = re.compile(r"Du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})")
NUM_DOCUMENT_RE = re.compile(r"N°Document\s*:\s*(\d+)")
DATE_DOCUMENT_RE = re.compile(r"Date document\s*:\s*(\d{2}/\d{2}/\d{4})")
TOTAL_RE = re.compile(r"TOTAL\s+(-?[\d\s  ]*,\d+)\s*€")

CIRCUIT_TO_CANAUX = {
    "Grossistes": ["OCP", "Alliance"],
    "Directs": ["Biogaran Direct"],
}


def to_float(s):
    if s is None:
        return None
    s = re.sub(r"[\s  ]", "", str(s)).replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def periode_depuis_debut(date_debut):
    """'JJ/MM/AAAA' -> 'MM/AAAA', cohérent avec periode_depuis_date_facture()
    d'analyse_consolidee.py (canaux Biogaran Direct / Alliance)."""
    if not date_debut or "/" not in date_debut:
        return "Inconnu"
    parts = date_debut.split("/")
    if len(parts) != 3:
        return "Inconnu"
    _, mm, aaaa = parts
    return f"{int(mm):02d}/{aaaa}"


def extraire_avoir_rdp(path_pdf):
    """
    Extrait l'entête + tous les blocs RDP d'un avoir Biogaran (1 page).
    Retourne un dict :
        {
            "num_document": str, "date_document": str,
            "periode_debut": str, "periode_fin": str, "periode": "MM/AAAA",
            "total": float,
            "lignes": [
                {
                    "taux_rdp": float, "circuit": "Grossistes"|"Directs",
                    "canaux": [...],  # canaux concernés par ce circuit
                    "taux_remise_facture_attendu": float,  # 40 - taux_rdp
                    "montant_remise_ht": float,  # négatif
                    "ca_brut_reference": float,  # positif
                    "num_document": str, "date_document": str, "periode": str,
                }, ...
            ],
        }
    """
    with pdfplumber.open(path_pdf) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    periode_m = PERIODE_RE.search(text)
    periode_debut = periode_m.group(1) if periode_m else ""
    periode_fin = periode_m.group(2) if periode_m else ""
    periode = periode_depuis_debut(periode_debut)

    num_m = NUM_DOCUMENT_RE.search(text)
    date_m = DATE_DOCUMENT_RE.search(text)
    num_document = num_m.group(1) if num_m else ""
    date_document = date_m.group(1) if date_m else ""

    total_m = TOTAL_RE.search(text)
    total = to_float(total_m.group(1)) if total_m else None

    lignes = []
    for m in BLOC_RE.finditer(text):
        taux_rdp = to_float(m.group("taux_rdp"))
        circuit = "Grossistes" if m.group("circuit") == "GROSSISTES" else "Directs"
        montant_remise_ht = to_float(m.group("montant_remise"))
        ca_brut_reference = to_float(m.group("ca_brut"))

        lignes.append({
            "taux_rdp": taux_rdp,
            "circuit": circuit,
            "canaux": CIRCUIT_TO_CANAUX[circuit],
            "taux_remise_facture_attendu": round(40 - taux_rdp, 2) if taux_rdp is not None else None,
            "montant_remise_ht": montant_remise_ht,
            "ca_brut_reference": ca_brut_reference,
            "num_document": num_document,
            "date_document": date_document,
            "periode": periode,
        })

    return {
        "num_document": num_document,
        "date_document": date_document,
        "periode_debut": periode_debut,
        "periode_fin": periode_fin,
        "periode": periode,
        "total": total,
        "lignes": lignes,
    }


def extraire_dossier(dossier_pdfs):
    """Parcourt tous les avoir_*.pdf d'un dossier et retourne la liste consolidée
    des blocs RDP (tous documents confondus)."""
    toutes_lignes = []
    for nom in sorted(os.listdir(dossier_pdfs)):
        if not nom.lower().startswith("avoir_") or not nom.lower().endswith(".pdf"):
            continue
        chemin = os.path.join(dossier_pdfs, nom)
        avoir = extraire_avoir_rdp(chemin)
        if not avoir["lignes"]:
            print(f"⚠️  Aucun bloc RDP trouvé dans {nom}")
            continue
        somme = round(sum(l["montant_remise_ht"] for l in avoir["lignes"]), 2)
        if avoir["total"] is not None and abs(somme - avoir["total"]) > 0.01:
            print(f"⚠️  {nom} : somme des blocs ({somme} €) ≠ TOTAL imprimé ({avoir['total']} €)")
        toutes_lignes.extend(avoir["lignes"])
    return toutes_lignes


if __name__ == "__main__":
    import sys
    for path in sys.argv[1:]:
        a = extraire_avoir_rdp(path)
        somme = round(sum(l["montant_remise_ht"] for l in a["lignes"]), 2)
        ok = "✅" if a["total"] is not None and abs(somme - a["total"]) <= 0.01 else "❌"
        print(f"{path}")
        print(f"  Document {a['num_document']} du {a['date_document']} — période de référence {a['periode']} "
              f"({a['periode_debut']} au {a['periode_fin']})")
        for l in a["lignes"]:
            print(f"    RDP {l['taux_rdp']}% {l['circuit']:11s} (canaux {l['canaux']}) : "
                  f"remise {l['montant_remise_ht']} € sur CA brut réf. {l['ca_brut_reference']} € "
                  f"— taux remise facture attendu = {l['taux_remise_facture_attendu']}%")
        print(f"  Somme blocs = {somme} € | TOTAL imprimé = {a['total']} € {ok}")
