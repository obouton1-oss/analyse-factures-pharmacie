# -*- coding: utf-8 -*-
"""
Parser des factures « directes » des laboratoires génériqueurs autres que
Biogaran : EG LABO et SANDOZ (Teva... à ajouter quand leurs factures seront
fournies).

Les factures EG sont émises par le même prestataire que celles de Biogaran
(Centre Spécialités Pharmaceutiques, puis Movianto) « pour ordre et pour
compte » du laboratoire : on retrouve exactement les deux mêmes mises en
page. Les motifs de lignes produit sont donc repris de parser_biogaran_direct.py
(validé au centime), sans modifier ce dernier. Ce module ajoute :

- l'identification du laboratoire (LABORATOIRES ci-dessous), qui donne
  l'offre / le génériqueur et le canal (« EG Direct ») ;
- le numéro et la date des factures au format CSP d'EG (« 18/09/2026
  T460911763 », date AVANT le numéro, contrairement à Biogaran) ;
- les avoirs de « REMISE ADDITIONNELLE » (lignes AVPPGRO / AVPPDIR, sans
  CIP) : ce ne sont pas des achats, ils sont renvoyés à part avec leur
  montant HT et le mois auquel ils se rapportent (« 08.26REMISEADD ») ;
- un contrôle au centime pour le format CSP : somme des lignes = somme des
  bases HT du pied de facture, et quantité = « TOTAL QTES FACTUREES ».
"""
import re

import pdfplumber

from parser_biogaran_direct import LIGNE_PRODUIT, to_float, _extraire_format_movianto

# Laboratoire : (motifs reconnus dans le texte, génériqueur, nom de canal)
LABORATOIRES = [
    ((r"POUR COMPTE DES LABORATOIRES EG LABO", r"D[ée]nomination sociale\s*:\s*EG LABO"),
     "EG LABO", "EG Direct"),
    ((r"D[ée]nomination sociale\s*:\s*SANDOZ\b",), "SANDOZ", "Sandoz Direct"),
]

# Format Sandoz (plateforme Cegedim, différent de Movianto) :
#   <désignation, parfois sur 2 lignes>
#   <N°> <Qté> PCE <PFHT|NaN> EUR [<remise> EUR] <prix net> EUR <montant net> EUR <TVA>%
#   Code Article : <CIP13>
#   N° de lot : <lot>
# Sandoz n'imprime pas le PFHT (« NaN ») : seul le net est connu, et le
# « Montant Total PFHT » imprimé est égal au net. Les avoirs (« Avoir n° »,
# conditions commerciales sur CIP fictif 3400999999999) sont ignorés.
LIGNE_SANDOZ_RE = re.compile(
    r"^(?P<num>\d+)\s+(?P<qte>-?\d+)\s+(?P<unite>[A-Z]{2,4})\s+(?P<montants>(?:(?:NaN|-?[\d ]+,\d+)\s+EUR\s+)+)"
    r"(?P<tva>[\d,]+)\s*%$")
MONTANT_EUR_RE = re.compile(r"(NaN|-?[\d ]+,\d+)\s+EUR")
CIP_SANDOZ_RE = re.compile(r"^Code Article\s*:\s*(\d{13})")
TYPE_SANDOZ_RE = re.compile(r"^(Facture|Avoir)\s+(\d+)\s*$", re.M)
DATE_SANDOZ_RE = re.compile(r"Date de facturation\s*:\s*(\d{2}/\d{2}/\d{4})")
TOTAL_SANDOZ_RE = re.compile(r"Montant Total net H\.?T\.?\s+(-?[\d ]+,\d+)\s+EUR")


def _num_fr(x):
    return float(x.replace(" ", "").replace(",", "."))


def _extraire_format_sandoz(texte, res):
    mt = TYPE_SANDOZ_RE.search(texte)
    md = DATE_SANDOZ_RE.search(texte)
    res["num_facture"] = mt.group(2) if mt else ""
    res["date_facture"] = md.group(1) if md else ""
    mtot = TOTAL_SANDOZ_RE.search(texte)
    res["total_net_ht"] = _num_fr(mtot.group(1)) if mtot else None
    if mt and mt.group(1) == "Avoir":
        res["type_document"] = "avoir_ignore"
        return []
    res["type_document"] = "facture"
    lignes_txt = texte.split("\n")
    debut = next((i for i, l in enumerate(lignes_txt) if l.startswith("CIP 13")), 0) + 1
    lignes, tampon = [], []
    i = debut
    while i < len(lignes_txt):
        l = lignes_txt[i].strip()
        m = LIGNE_SANDOZ_RE.match(l)
        if m:
            montants = MONTANT_EUR_RE.findall(m.group("montants"))
            qte = int(m.group("qte"))
            pfht = None if montants[0] == "NaN" else _num_fr(montants[0])
            prix_net = _num_fr(montants[-2])
            montant_net = _num_fr(montants[-1])
            cip = ""
            for j in range(i + 1, min(i + 4, len(lignes_txt))):
                mc = CIP_SANDOZ_RE.match(lignes_txt[j].strip())
                if mc:
                    cip = mc.group(1)
                    break
            ca_ppht = round(qte * pfht, 2) if pfht is not None else montant_net
            lignes.append({
                "cip13": cip,
                "libelle": re.sub(r"\s+", " ", " ".join(tampon)).strip(),
                "tva": _num_fr(m.group("tva")),
                "qte_cdee": qte, "qte_fact": qte,
                "ppht_unit": pfht if pfht is not None else prix_net,
                "prix_net_unitaire": prix_net,
                "ca_ppht": ca_ppht,
                "ca_pght": montant_net,
                "taux_remise": round((1 - montant_net / ca_ppht) * 100, 2) if ca_ppht else 0.0,
                "remise_ht": round(ca_ppht - montant_net, 2),
                "pfht_non_communique": pfht is None,
            })
            tampon = []
        elif l.startswith(("Code Article", "N° de lot", "Montant Total")):
            tampon = []
            if l.startswith("Montant Total"):
                break
        elif l:
            tampon.append(l)
        i += 1
    return lignes

NUM_DATE_CSP_RE = re.compile(r"(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<num>[A-Z]\d{9})\b")
TOTAL_QTES_RE = re.compile(r"TOTAL QTES FACTUREES\s+(-?\d+)")
LIGNE_AVOIR_RE = re.compile(
    r"^(?P<code>AV[A-Z]+)\s+(?P<libelle>.+?)\s+(?P<ref>\S+)\s+(?P<qte>\d+)\s+"
    r"(?P<pu>[\d.]+)\s+(?P<pu_net>[\d.]+)\s+(?P<montant>[\d.]+)\s+(?P<tva>[\d.]+)$")
MOIS_REMISE_RE = re.compile(r"(\d{2})\.(\d{2})REMISEADD")
# Pied de facture CSP : "base taux montant_tva [total_tva]"
LIGNE_TVA_RE = re.compile(r"^(-?\d+\.\d{2})\s+(\d+\.\d{2})\s+(-?\d+\.\d{2})(?:\s+-?\d+\.\d{2})?$")


def identifier_laboratoire(texte):
    """(génériqueur, canal) du laboratoire émetteur, ou (None, None)."""
    for motifs, generiqueur, canal in LABORATOIRES:
        if any(re.search(m, texte, re.I) for m in motifs):
            return generiqueur, canal
    return None, None


def extraire_facture_direct_labo(path_pdf):
    """Renvoie un dict :
      laboratoire, canal, num_facture, date_facture, type_document
      ("facture" | "avoir_remise" | "avoir_ignore" | "inconnu"), lignes (format commun aux autres
      parsers), avoirs [{code, libelle, montant_ht, mois_concerne}],
      total_net_ht (imprimé ou reconstitué du pied de facture), controles."""
    with pdfplumber.open(path_pdf) as pdf:
        texte = "\n".join(p.extract_text() or "" for p in pdf.pages)

    labo, canal = identifier_laboratoire(texte)
    res = {"laboratoire": labo, "canal": canal, "num_facture": "", "date_facture": "",
           "type_document": "inconnu", "lignes": [], "avoirs": [], "total_net_ht": None,
           "controles": []}
    if not labo:
        res["controles"].append("laboratoire émetteur non reconnu")
        return res

    if labo == "SANDOZ":
        lignes = _extraire_format_sandoz(texte, res)
        if res["type_document"] == "avoir_ignore":
            return res
        if any(l.get("pfht_non_communique") for l in lignes):
            res["pfht_non_communique"] = True
    elif "Code Article" in texte and "Prix brut" in texte:
        # Format Movianto : même mise en page que Biogaran
        f = _extraire_format_movianto(texte)
        res.update(num_facture=f["num_facture"], date_facture=f["date_facture"],
                   total_net_ht=f["total_net_ht"], type_document="facture")
        lignes = f["lignes"]
    else:
        m = NUM_DATE_CSP_RE.search(texte)
        if m:
            res["num_facture"], res["date_facture"] = m.group("num"), m.group("date")
        lignes = []
        for brute in texte.split("\n"):
            brute = brute.strip()
            mp = LIGNE_PRODUIT.match(brute)
            if mp:
                qte = int(mp.group("qte"))
                pu_ht = to_float(mp.group("pu_ht"))
                montant_net = to_float(mp.group("montant"))
                ca_ppht = round(qte * pu_ht, 2)
                lignes.append({
                    "cip13": mp.group("cip"),
                    "libelle": re.sub(r"\s+", " ", mp.group("designation")).strip(),
                    "tva": to_float(mp.group("tva")),
                    "qte_cdee": qte, "qte_fact": qte,
                    "ppht_unit": pu_ht,
                    "prix_net_unitaire": to_float(mp.group("pu_apres")),
                    "ca_ppht": ca_ppht,
                    "ca_pght": montant_net,
                    "taux_remise": to_float(mp.group("remise")),
                    "remise_ht": round(ca_ppht - montant_net, 2),
                })
                continue
            ma = LIGNE_AVOIR_RE.match(brute)
            if ma:
                mm = MOIS_REMISE_RE.search(texte)
                res["avoirs"].append({
                    "code": ma.group("code"),
                    "libelle": ma.group("libelle"),
                    "montant_ht": -to_float(ma.group("montant")),
                    "mois_concerne": f"{mm.group(1)}/20{mm.group(2)}" if mm else "",
                })
        # Pied de facture : somme des bases HT
        bloc = texte.split("TOTAL QTES FACTUREES", 1)[-1].split("Pas d", 1)[0]
        bases = [to_float(x.group(1)) for x in
                 (LIGNE_TVA_RE.match(l.strip()) for l in bloc.split("\n")) if x]
        if bases:
            res["total_net_ht"] = round(sum(bases), 2)
        qm = TOTAL_QTES_RE.search(texte)
        if res["avoirs"] and not lignes:
            res["type_document"] = "avoir_remise"
            somme = round(sum(a["montant_ht"] for a in res["avoirs"]), 2)
            if res["total_net_ht"] is not None and abs(somme - res["total_net_ht"]) > 0.005:
                res["controles"].append(f"avoir : lignes {somme:.2f} ≠ pied de facture {res['total_net_ht']:.2f}")
        elif lignes:
            res["type_document"] = "facture"
            q = sum(l["qte_fact"] for l in lignes)
            if qm and int(qm.group(1)) != q:
                res["controles"].append(f"quantité lignes {q} ≠ TOTAL QTES FACTUREES {qm.group(1)}")

    for l in lignes:
        l.update({"offre": labo, "canal": canal,
                  "num_facture": res["num_facture"], "date_facture": res["date_facture"]})
    res["lignes"] = lignes
    if res["type_document"] == "facture" and res["total_net_ht"] is not None:
        somme = round(sum(l["ca_pght"] for l in lignes), 2)
        if abs(somme - res["total_net_ht"]) > 0.005:
            res["controles"].append(f"lignes {somme:.2f} € ≠ total net HT facture {res['total_net_ht']:.2f} €")
    if res["type_document"] == "facture" and res["total_net_ht"] is None:
        res["controles"].append("total net HT de la facture introuvable : contrôle impossible")
    if not res["date_facture"]:
        res["controles"].append("date de facture introuvable")
    return res


if __name__ == "__main__":
    import sys
    for chemin in sys.argv[1:]:
        r = extraire_facture_direct_labo(chemin)
        print(f"== {chemin} : {r['laboratoire']} {r['type_document']} {r['num_facture']} du "
              f"{r['date_facture']} — {len(r['lignes'])} ligne(s), total net HT {r['total_net_ht']}")
        for a in r["avoirs"]:
            print(f"   avoir {a['code']} {a['libelle']} {a['montant_ht']:.2f} € (mois {a['mois_concerne']})")
        print("   contrôles :", "OK" if not r["controles"] else " | ".join(r["controles"]))
