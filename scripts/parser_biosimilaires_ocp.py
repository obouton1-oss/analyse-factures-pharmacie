# -*- coding: utf-8 -*-
"""
Parser du relevé OCP « Achats en cumul annuel des Produits BIOSIMILAIRES »
(code BIOST, édité chaque trimestre, cumul depuis janvier).

Structure du document :
- Section « Molécule X » : une ligne par produit (désignation, CIP13,
  BIOSIMILAIRE / BIOREFERENT), puis pour CHAQUE trimestre présent un triplet
  CA Net / CA PGHT / Qté, puis le CA Net total et le poids du produit.
  Un produit sans achat sur un trimestre a ses cases vides : l'affectation des
  montants aux trimestres se fait donc par la POSITION HORIZONTALE des
  colonnes (en-têtes), pas par l'ordre des nombres.
- Section « Par LABORATOIRE » : mêmes produits, regroupés par laboratoire
  (fabricant). Sert à rattacher chaque CIP à son laboratoire.
- Page « Récapitulatif de la période » : totaux 2026 à date par type
  (BIOSIMILAIRE / BIOREFERENT) — Qté commandée, Qté facturée, CA PPHT,
  CA PGHT, Mt Remise, CA NET. Sert de contrôle.

« CA Net » = CA PPHT − remise (la remise n'existe que pour les achats passés
dans une offre, ex. FRESENIUS BIOSIMILAIRE RSF).

Le relevé couvre TOUS les achats de biosimilaires / bioréférents chez OCP,
avec ou sans offre : la plupart n'apparaissent donc PAS dans le BO-OFFREM.

Usage autonome :
    python parser_biosimilaires_ocp.py "Biosim 2026 T3 OCP.pdf"
"""
import re
import sys
from collections import defaultdict

import pdfplumber

STATUTS = ("BIOSIMILAIRE", "BIOREFERENT")
RE_CIP = re.compile(r"^\d{13}$")
RE_MONTANT_TOKEN = re.compile(r"^-?[\d]+(,\d+)?$")
MOIS = {"JANVIER": 1, "FEVRIER": 2, "FÉVRIER": 2, "MARS": 3, "AVRIL": 4, "MAI": 5,
        "JUIN": 6, "JUILLET": 7, "AOUT": 8, "AOÛT": 8, "SEPTEMBRE": 9,
        "OCTOBRE": 10, "NOVEMBRE": 11, "DECEMBRE": 12, "DÉCEMBRE": 12}


def _num(s):
    return float(s.replace(" ", "").replace(",", "."))


def est_releve_biosimilaires(texte_page1):
    """Reconnaissance du document (1re page)."""
    t = (texte_page1 or "").upper()
    return "PRODUITS BIOSIMILAIRES" in t and "BIOST" in t


def _lignes(page):
    """Mots de la page regroupés par ligne (tolérance verticale 2 pt)."""
    mots = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
    lignes = []
    for m in sorted(mots, key=lambda w: (round(w["top"]), w["x0"])):
        if lignes and abs(lignes[-1][0] - m["top"]) <= 2:
            lignes[-1][1].append(m)
        else:
            lignes.append([m["top"], [m]])
    return [(top, sorted(ws, key=lambda w: w["x0"])) for top, ws in lignes]


def _fusionner_nombres(mots):
    """Recolle les nombres coupés par le séparateur de milliers
    ("2" + "364,48" -> 2364,48) : un groupe de 1 à 3 chiffres suivi à moins
    de 3 pt d'un autre groupe numérique. Renvoie [(valeur_texte, x1)]."""
    res = []
    i = 0
    while i < len(mots):
        txt, x0, x1 = mots[i]["text"], mots[i]["x0"], mots[i]["x1"]
        while (re.fullmatch(r"-?\d{1,3}", txt) and i + 1 < len(mots)
               and re.fullmatch(r"\d{3}(,\d+)?", mots[i + 1]["text"])
               and mots[i + 1]["x0"] - x1 < 3):
            i += 1
            txt += " " + mots[i]["text"]
            x1 = mots[i]["x1"]
        res.append((txt, x0, x1))
        i += 1
    return res


def parse_releve_biosimilaires(pdf_path):
    """Renvoie un dict :
      annee, mois_fin, edition,
      lignes : [{molecule, designation, cip13, statut, laboratoire,
                 trimestre ("2026 T1"), ca_net, ca_pght, qte}]
      totaux_lignes : {cip13: ca_net_total imprimé}
      recap : {statut: {qte_cdee, qte_fact, ca_ppht, ca_pght, remise, ca_net}}
      controles : [messages]  (vide = tout concorde)
    """
    lignes_mol = []      # lignes de la section molécule
    labo_par_cip = {}
    totaux = {}
    recap = {}
    annee = mois_fin = edition = None

    section = None       # "molecule" | "labo" | "recap" | None
    courant = None       # nom de molécule ou de labo
    colonnes = []        # [(x0, nom_colonne, trimestre)] triées par x0
    trimestres_x = []    # [(x0, "2026 T1")]

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texte = page.extract_text() or ""
            if edition is None:
                m = re.search(r"Edit[ée] le : (\d\d/\d\d/\d\d)", texte)
                edition = m.group(1) if m else None
            if mois_fin is None:
                m = re.search(r"Ann[ée]e en cours : \w+ [àa] (\w+) (\d{4})", texte)
                if m:
                    mois_fin = MOIS.get(m.group(1).upper())
                    annee = int(m.group(2))
            if "Récapitulatif de la période" in texte or "capitulatif de la p" in texte:
                section = "recap"
            for top, mots in _lignes(page):
                txts = [w["text"] for w in mots]
                ligne = " ".join(txts)

                if txts[:1] == ["Molécule"] and len(txts) >= 2:
                    section, courant = "molecule", " ".join(txts[1:])
                    continue
                if txts[:1] == ["Laboratoire"] and len(txts) >= 2 and "bioréférent" not in ligne:
                    section, courant = "labo", " ".join(txts[1:])
                    continue
                if section in ("molecule", "labo") and re.fullmatch(r"(\d{4} T\d ?)+", ligne):
                    trimestres_x = [(mots[i]["x0"], f"{mots[i]['text']} {mots[i + 1]['text']}")
                                    for i in range(0, len(mots) - 1, 2)]
                    colonnes = []
                    continue
                if section in ("molecule", "labo") and txts and txts[0] in ("Total", "Poids"):
                    colonnes.append((mots[0]["x0"], "total" if txts[0] == "Total" else "poids", None))
                    continue
                if section in ("molecule", "labo") and txts[:2] == ["Désignation", "produit"]:
                    # Triplets CA Net / CA PGHT / Qté, dans l'ordre des trimestres
                    sous = []
                    for i, w in enumerate(mots):
                        if w["text"] == "CA" and i + 1 < len(mots) and mots[i + 1]["text"] in ("Net", "PGHT"):
                            sous.append((w["x0"], "ca_net" if mots[i + 1]["text"] == "Net" else "ca_pght"))
                        elif w["text"] == "Qté":
                            sous.append((w["x0"], "qte"))
                    trims = [t for _, t in sorted(trimestres_x)]
                    cols = []
                    for k, (x0, nom) in enumerate(sous):
                        t = trims[k // 3] if k // 3 < len(trims) else None
                        cols.append((x0, nom, t))
                    colonnes = sorted(cols + [c for c in colonnes if c[1] in ("total", "poids")])
                    continue

                if section in ("molecule", "labo"):
                    i_cip = next((i for i, t in enumerate(txts) if RE_CIP.match(t)), None)
                    if i_cip is None or i_cip + 1 >= len(txts) or txts[i_cip + 1] not in STATUTS:
                        continue
                    cip = txts[i_cip]
                    statut = txts[i_cip + 1]
                    designation = " ".join(txts[:i_cip])
                    if section == "labo":
                        labo_par_cip[cip] = courant
                        continue
                    valeurs = defaultdict(dict)
                    for txt, x0, x1 in _fusionner_nombres(mots[i_cip + 2:]):
                        if txt.endswith("%"):
                            continue
                        col = None
                        for cx0, nom, trim in colonnes:
                            if cx0 <= x1:
                                col = (nom, trim)
                        if col is None:
                            continue
                        nom, trim = col
                        if nom == "total":
                            totaux[cip] = totaux.get(cip, 0) + _num(txt)
                        elif nom != "poids" and trim:
                            valeurs[trim][nom] = _num(txt)
                    for trim, v in valeurs.items():
                        lignes_mol.append({
                            "molecule": courant.replace("_", " "),
                            "designation": designation,
                            "cip13": cip,
                            "statut": "Biosimilaire" if statut == "BIOSIMILAIRE" else "Bioréférent",
                            "trimestre": trim,
                            "ca_net": v.get("ca_net", 0.0),
                            "ca_pght": v.get("ca_pght", 0.0),
                            "qte": int(v.get("qte", 0)),
                        })
                    continue

                if section == "recap":
                    m = re.match(r"(\d{4}) à date (BIOSIMILAIRE|BIOREFERENT) (\d+) (\d+) (.*)$", ligne)
                    if m and int(m.group(1)) == annee:
                        montants = re.findall(r"\d{1,3}(?: \d{3})*,\d{1,2}", m.group(5))
                        if len(montants) >= 4:
                            ca_ppht, ca_pght, remise, ca_net = (_num(x) for x in montants[:4])
                            recap["Biosimilaire" if m.group(2) == "BIOSIMILAIRE" else "Bioréférent"] = {
                                "qte_cdee": int(m.group(3)), "qte_fact": int(m.group(4)),
                                "ca_ppht": ca_ppht, "ca_pght": ca_pght,
                                "remise": remise, "ca_net": ca_net,
                            }

    for l in lignes_mol:
        l["laboratoire"] = labo_par_cip.get(l["cip13"], "")

    # ---- Contrôles -------------------------------------------------------
    controles = []
    par_cip = defaultdict(float)
    for l in lignes_mol:
        par_cip[l["cip13"]] += l["ca_net"]
    for cip, tot in totaux.items():
        if abs(par_cip[cip] - tot) > 0.02:
            controles.append(f"CIP {cip} : somme des trimestres {par_cip[cip]:.2f} ≠ total imprimé {tot:.2f}")
    for cip in par_cip:
        if cip not in labo_par_cip:
            controles.append(f"CIP {cip} : laboratoire non trouvé dans la section « Par laboratoire »")
    for statut, r in recap.items():
        ca = sum(l["ca_net"] for l in lignes_mol if l["statut"] == statut)
        q = sum(l["qte"] for l in lignes_mol if l["statut"] == statut)
        if abs(ca - r["ca_net"]) > 0.05:
            controles.append(f"{statut} : CA net des lignes {ca:.2f} ≠ récapitulatif {r['ca_net']:.2f}")
        if q != r["qte_fact"]:
            controles.append(f"{statut} : quantité des lignes {q} ≠ quantité facturée récapitulatif {r['qte_fact']}")
    if not recap:
        controles.append("Récapitulatif de la période introuvable : contrôle des totaux impossible")

    return {
        "annee": annee, "mois_fin": mois_fin, "edition": edition,
        "lignes": lignes_mol, "totaux_lignes": totaux, "recap": recap,
        "controles": controles,
    }


if __name__ == "__main__":
    for chemin in sys.argv[1:]:
        r = parse_releve_biosimilaires(chemin)
        print(f"== {chemin} : {r['annee']} jusqu'au mois {r['mois_fin']}, édité le {r['edition']}")
        print(f"   {len(r['lignes'])} lignes (produit x trimestre)")
        for s, v in r["recap"].items():
            print(f"   récap {s} : {v}")
        print("   contrôles :", "OK" if not r["controles"] else "")
        for c in r["controles"]:
            print("    ⚠️", c)
