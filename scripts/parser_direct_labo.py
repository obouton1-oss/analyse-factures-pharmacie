# -*- coding: utf-8 -*-
"""
Parser des factures « directes » des laboratoires génériqueurs autres que
Biogaran (EG LABO pour l'instant ; Teva, Sandoz... à ajouter quand leurs
factures seront fournies).

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
]

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
      ("facture" | "avoir_remise" | "inconnu"), lignes (format commun aux autres
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

    if "Code Article" in texte and "Prix brut" in texte:
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
