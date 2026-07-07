"""
Parser BO-OFFREM (OCP) — approche par ORDRE des colonnes, pas par position X absolue.

Principe :
- Chaque ligne produit commence par un CIP13 (13 chiffres).
- La zone chiffrée démarre au taux de TVA (valeur canonique parmi 2,10 / 5,50 / 10,00 / 20,00),
  ce qui sépare proprement le libellé (longueur variable) des colonnes numériques.
- Après la TVA, les colonnes sont toujours dans le même ordre :
    Qté Cdée, Qté fact, PPHT Unit, Prix net Unitaire, CA PPHT, [CA PGHT], Taux Remise, Remise HT
  → 7 valeurs si CA PGHT absent, 8 si présent (on le déduit du nombre de tokens, pas d'une position).
- Les milliers sont parfois séparés par un espace ("6 174", "1 251,25"). On les refusionne
  UNIQUEMENT si l'écart horizontal entre les deux tokens est petit (< 6pt, séparateur de milliers)
  et jamais quand l'écart correspond à un vrai espacement de colonne (~15-30pt).
"""
import re
import pdfplumber
from collections import defaultdict

TVA_CANONIQUES = {"2,10", "5,50", "10,00", "20,00"}
GAP_MILLIERS_MAX = 6.0  # pt : en dessous = séparateur de milliers, au dessus = vraie colonne

CIP_RE = re.compile(r"^\d{13}$")
BARE_INT_RE = re.compile(r"^\d{1,3}$")
CONT_3DIGIT_RE = re.compile(r"^\d{3}$")
CONT_3DIGIT_DEC_RE = re.compile(r"^\d{3},\d{2}$")


def fr_to_float(s):
    return float(s.replace(" ", "").replace(",", "."))


def fr_to_int(s):
    return int(s.replace(" ", ""))


def get_rows(page, ytol=2.5):
    """Groupe les mots en lignes (par proximité verticale), triées par x0 au sein de chaque ligne."""
    words = page.extract_words()
    words.sort(key=lambda w: (w["top"], w["x0"]))
    rows, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= ytol:
            cur.append(w)
            if cur_top is None:
                cur_top = w["top"]
        else:
            rows.append(sorted(cur, key=lambda x: x["x0"]))
            cur = [w]
            cur_top = w["top"]
    if cur:
        rows.append(sorted(cur, key=lambda x: x["x0"]))
    return rows


def merge_thousands(tokens):
    """Refusionne les tokens séparés par un espace de milliers (gap < GAP_MILLIERS_MAX pt)."""
    merged = []
    i = 0
    while i < len(tokens):
        cur = tokens[i]
        if i + 1 < len(tokens):
            nxt = tokens[i + 1]
            gap = nxt["x0"] - cur["x1"]
            cur_is_bare_int = BARE_INT_RE.fullmatch(cur["text"]) is not None
            nxt_is_cont = (
                CONT_3DIGIT_RE.fullmatch(nxt["text"]) is not None
                or CONT_3DIGIT_DEC_RE.fullmatch(nxt["text"]) is not None
            )
            if cur_is_bare_int and nxt_is_cont and gap < GAP_MILLIERS_MAX:
                merged.append(
                    {
                        "text": cur["text"] + nxt["text"],
                        "x0": cur["x0"],
                        "x1": nxt["x1"],
                    }
                )
                i += 2
                continue
        merged.append(cur)
        i += 1
    return merged


def parse_data_row(tokens, offre, page_num, row_index):
    """Parse une ligne produit. Retourne un dict ou None + liste d'anomalies."""
    cip = tokens[0]["text"]

    # Cherche l'ancre TVA (première valeur canonique après le CIP)
    idx_tva = None
    for i in range(1, len(tokens)):
        if tokens[i]["text"] in TVA_CANONIQUES:
            idx_tva = i
            break
    if idx_tva is None:
        return None, [f"[p{page_num} l{row_index}] CIP {cip}: pas d'ancre TVA trouvée"]

    label = " ".join(t["text"] for t in tokens[1:idx_tva])
    tva = tokens[idx_tva]["text"]
    rest = merge_thousands(tokens[idx_tva + 1 :])

    anomalies = []
    ca_pght = None
    if len(rest) == 7:
        qte_cdee, qte_fact, ppht_unit, prix_net, ca_ppht, taux_remise, remise_ht = [
            t["text"] for t in rest
        ]
    elif len(rest) == 8:
        (
            qte_cdee,
            qte_fact,
            ppht_unit,
            prix_net,
            ca_ppht,
            ca_pght,
            taux_remise,
            remise_ht,
        ) = [t["text"] for t in rest]
    else:
        anomalies.append(
            f"[p{page_num} l{row_index}] CIP {cip} '{label}': {len(rest)} tokens "
            f"après TVA (attendu 7 ou 8) -> {[t['text'] for t in rest]}"
        )
        return None, anomalies

    try:
        row = {
            "cip13": cip,
            "libelle": label,
            "offre": offre,
            "tva": fr_to_float(tva),
            "qte_cdee": fr_to_int(qte_cdee),
            "qte_fact": fr_to_int(qte_fact),
            "ppht_unit": fr_to_float(ppht_unit),
            "prix_net_unitaire": fr_to_float(prix_net),
            "ca_ppht": fr_to_float(ca_ppht),
            "ca_pght": fr_to_float(ca_pght) if ca_pght is not None else None,
            "taux_remise": fr_to_float(taux_remise),
            "remise_ht": fr_to_float(remise_ht),
        }
    except ValueError as e:
        anomalies.append(f"[p{page_num} l{row_index}] CIP {cip}: erreur conversion ({e})")
        return None, anomalies

    return row, anomalies


def parse_bo_offrem(pdf_path):
    rows_out = []
    anomalies = []
    table_totals = []  # lignes "Total" imprimées, pour validation croisée

    current_offre = None
    current_table_rows = []

    pdf = pdfplumber.open(pdf_path)
    for page_num, page in enumerate(pdf.pages, start=1):
        rows = get_rows(page)
        n_rows = len(rows)
        for row_index, tokens in enumerate(rows):
            if not tokens:
                continue
            first = tokens[0]["text"]

            if first == "Offre":
                nom_offre = " ".join(t["text"] for t in tokens[1:]).strip()
                if not nom_offre:
                    # Cas observé sur certains mois (ex. mai 2026) : le mot "Offre"
                    # est seul sur sa ligne, le nom réel de l'offre est sur la
                    # ligne suivante (retour à la ligne dans le PDF). On va le
                    # chercher sur la prochaine ligne non vide, à condition que
                    # ce ne soit ni une ligne CIP13, ni "Total", ni "Offre".
                    for lookahead in range(row_index + 1, min(row_index + 3, n_rows)):
                        next_tokens = rows[lookahead]
                        if not next_tokens:
                            continue
                        next_first = next_tokens[0]["text"]
                        if CIP_RE.match(next_first) or next_first in ("Total", "Offre"):
                            break
                        nom_offre = " ".join(t["text"] for t in next_tokens).strip()
                        break
                current_offre = nom_offre
                current_table_rows = []
                continue

            if first == "Total":
                # ligne de total imprimée par le PDF -> on la garde pour validation
                nums = [t["text"] for t in tokens[1:]]
                table_totals.append(
                    {
                        "offre": current_offre,
                        "page": page_num,
                        "raw": nums,
                        "rows_captured": len(current_table_rows),
                    }
                )
                continue

            if CIP_RE.match(first):
                row, row_anomalies = parse_data_row(
                    tokens, current_offre, page_num, row_index
                )
                anomalies.extend(row_anomalies)
                if row:
                    rows_out.append(row)
                    current_table_rows.append(row)

    return rows_out, anomalies, table_totals


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/recap_offres_avril.pdf"
    rows, anomalies, totals = parse_bo_offrem(path)

    print(f"Lignes produit extraites : {len(rows)}")
    print(f"Anomalies : {len(anomalies)}")
    for a in anomalies[:30]:
        print("  ", a)

    grand_ca_ppht = sum(r["ca_ppht"] for r in rows)
    grand_qte_cdee = sum(r["qte_cdee"] for r in rows)
    grand_qte_fact = sum(r["qte_fact"] for r in rows)
    grand_remise = sum(r["remise_ht"] for r in rows)

    print(f"\nTotal CA PPHT calculé : {grand_ca_ppht:.2f}")
    print(f"Total Qté Cdée calculé : {grand_qte_cdee}")
    print(f"Total Qté fact calculé : {grand_qte_fact}")
    print(f"Total Remise HT calculée : {grand_remise:.2f}")
