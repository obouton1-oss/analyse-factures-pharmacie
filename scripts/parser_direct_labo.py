# -*- coding: utf-8 -*-
"""
Parser des factures « directes » des laboratoires génériqueurs autres que
Biogaran : EG LABO, SANDOZ et TEVA.

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
    ((r"\bTeva Sant[ée]\b", r"\bTEVA SANTE SAS\b"), "TEVA", "Teva Direct"),
]

# ---------------------------------------------------------------------------
# Format Teva : deux mises en page.
#  - « SAP » (jusqu'à août 2026) : bloc produit sur plusieurs lignes
#      <désignation>            / <CIP13> / <suite désignation> <code douane 8 ch.>
#      Lot n° Exp. Date Quantité HTS# (TARIFF)
#      <qté> <PU brut> <remise>% <PU net> <TVA> % <montant HT>
#    Total : « Total <qté> <brut HT> <net HT> <TVA> <TTC> ».
#  - « DigiPA » (depuis septembre 2026, « Rendu lisible de la facture
#    électronique ») : une ligne par produit
#      <désignation courte> <qté> <PU brut> € <remise> % <PU net> € <TVA> % <montant> €
#      puis « EAN 0<CIP13> ». Total : « Total HT <montant> € ».
#  Dans les deux cas la quantité et le PU brut se touchent (« 1 270,00 » =
#  1 x 270,00, pas 1 270,00) : on retient le découpage qui vérifie
#  PU brut x (1 - remise) = PU net.
#  Avoirs Teva (« Avoir: n° », ristournes additionnelles) : ignorés.
# ---------------------------------------------------------------------------
TEVA_SAP_VAL_RE = re.compile(
    r"^(?P<avant>[\d ]+,\d+)\s+(?P<rem>[\d,]+)%\s+(?P<net>[\d ]*\d,\d+)\s+(?P<tva>[\d.,]+)\s*%\s+"
    r"(?P<mt>[\d ]*\d,\d+)$")
TEVA_DIGIPA_VAL_RE = re.compile(
    r"^(?P<des>.*?)\s+(?P<avant>\d[\d ]*,\d+)\s*€\s+(?P<rem>[\d,]+)\s*%\s+(?P<net>[\d ]*\d,\d+)\s*€\s+"
    r"(?P<tva>[\d.,]+)\s*%\s+(?P<mt>[\d ]*\d,\d+)\s*€$")


def _qte_et_brut(avant, rem, net, montant):
    """« avant » = [fin de désignation] + quantité + PU brut accolés
    (« 1 270,00 », « 200 1,96 », « 1 5 6,60 » où le 1er « 1 » termine la
    désignation). Renvoie (nb_jetons_designation, qte, brut) tel que
    brut x (1 - rem) = net et qte x net = montant, ou (0, None, None)."""
    toks = avant.split()
    for debut in range(0, len(toks) - 1):
        for k in range(debut + 1, len(toks)):
            try:
                qte = int("".join(toks[debut:k]))
                brut = _num_fr("".join(toks[k:]))
            except ValueError:
                continue
            if (qte > 0 and abs(brut * (1 - rem / 100) - net) <= 0.011
                    and abs(qte * net - montant) <= max(0.02, qte * 0.0051)):
                return debut, qte, brut
    return 0, None, None


def _ligne_teva(cip, libelle, qte, brut, rem, net, tva, montant):
    ca_ppht = round(qte * brut, 2)
    return {"cip13": cip, "libelle": libelle, "tva": tva, "qte_cdee": qte, "qte_fact": qte,
            "ppht_unit": brut, "prix_net_unitaire": net, "ca_ppht": ca_ppht, "ca_pght": montant,
            "taux_remise": rem, "remise_ht": round(ca_ppht - montant, 2)}


def _extraire_format_teva(texte, res):
    lignes_txt = [l.strip() for l in texte.split("\n")]
    if re.search(r"^Avoir\s*:", texte, re.M):
        m = re.search(r"Avoir\s*:\s*(\d+)\s+Date Document\s*:\s*(\d{2}[/.]\d{2}[/.]\d{4})", texte)
        if m:
            res["num_facture"], res["date_facture"] = m.group(1), m.group(2).replace(".", "/")
        res["type_document"] = "avoir_ignore"
        return []
    res["type_document"] = "facture"
    lignes = []
    if "DigiPA" in texte or "N° de facture" in texte:
        m = re.search(r"N° de facture\s+(\d+)", texte)
        d = re.search(r"Date d.émission\s+(\d{2}/\d{2}/\d{4})", texte)
        res["num_facture"] = m.group(1) if m else ""
        res["date_facture"] = d.group(1) if d else ""
        t = re.search(r"Total HT\s+([\d ]*\d,\d+)\s*€", texte)
        res["total_net_ht"] = _num_fr(t.group(1)) if t else None
        for i, l in enumerate(lignes_txt):
            m = TEVA_DIGIPA_VAL_RE.match(l)
            if not m:
                continue
            rem, net, mt = _num_fr(m.group("rem")), _num_fr(m.group("net")), _num_fr(m.group("mt"))
            # La quantité est le dernier nombre entier de la désignation courte
            # (« ... 120P1 20 3,13 € ») : on la recolle à « avant » pour le découpage.
            avant_toks = m.group("avant").split()
            nb_des, qte, brut = _qte_et_brut(m.group("avant"), rem, net, mt)
            libelle = " ".join([m.group("des")] + avant_toks[:nb_des])
            cip = ""
            for j in range(i + 1, min(i + 6, len(lignes_txt))):
                me = re.search(r"EAN\s+0?(\d{13})", lignes_txt[j])
                if me:
                    cip = me.group(1)
                    break
            if qte is None:
                res["controles"].append(f"ligne non décodée : {l}")
                continue
            lignes.append(_ligne_teva(cip, libelle, qte, brut, rem, net, _num_fr(m.group("tva")), mt))
    else:
        m = re.search(r"Facture\s*:\s*(\d+)\s+Date Document\s*:\s*(\d{2}[/.]\d{2}[/.]\d{4})", texte)
        if m:
            res["num_facture"], res["date_facture"] = m.group(1), m.group(2).replace(".", "/")
        mt_ = re.search(r"\bTotal\s+([\d ,]+)$", texte, re.M)
        if mt_:
            # Groupes de montants : un groupe se termine par un jeton à virgule ;
            # le net HT est le 3e groupe en partant de la droite.
            groupes, courant = [], []
            for tok in mt_.group(1).split():
                courant.append(tok)
                if "," in tok:
                    groupes.append(courant)
                    courant = []
            if len(groupes) >= 3:
                net_toks = groupes[-3]
                # le 1er groupe colle quantité + brut ; les suivants sont propres
                res["total_net_ht"] = _num_fr("".join(net_toks))
        # Zone produits : entre l'en-tête « Désignation CIP Code ... » (+ ligne
        # « HT ») et « Détails sur le versement ».
        try:
            debut = next(i for i, l in enumerate(lignes_txt) if l.startswith("Désignation CIP Code")) + 1
        except StopIteration:
            debut = 0
        fin = next((i for i, l in enumerate(lignes_txt) if l.startswith("Détails sur le versement")),
                   len(lignes_txt))
        LOT_RE = re.compile(r"^\S+ \d{1,2} \S+ \d{2} \d+ \d{8}$")   # « TBYC22D 31 mars 28 1 30049000 »
        designation, cip = [], ""
        for l in lignes_txt[debut:fin]:
            if l == "HT" or l.startswith("Lot n°") or LOT_RE.match(l):
                continue
            if re.fullmatch(r"\d{13}", l):
                cip = l
                continue
            m = TEVA_SAP_VAL_RE.match(l)
            if m:
                rem, net, mt = _num_fr(m.group("rem")), _num_fr(m.group("net")), _num_fr(m.group("mt"))
                _, qte, brut = _qte_et_brut(m.group("avant"), rem, net, mt)
                libelle = re.sub(r"\s+\d{8}$", "", " ".join(designation)).strip()
                if qte is None:
                    res["controles"].append(f"ligne non décodée : {l}")
                else:
                    lignes.append(_ligne_teva(cip, libelle, qte, brut, rem, net,
                                              _num_fr(m.group("tva")), mt))
                designation, cip = [], ""
                continue
            designation.append(l)
    return lignes

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
      ("facture" | "facture_sans_produit" | "avoir_remise" | "avoir_ignore" |
      "inconnu"), lignes (format commun aux autres
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

    if labo == "TEVA":
        lignes = _extraire_format_teva(texte, res)
        if res["type_document"] == "avoir_ignore":
            return res
    elif labo == "SANDOZ":
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
        if not lignes and not res["avoirs"] and res["total_net_ht"] in (0.0, None) and qm:
            res["type_document"] = "facture_sans_produit"   # ex. brochures gratuites (0,00 €)
        elif res["avoirs"] and not lignes:
            res["type_document"] = "avoir_remise"
            somme = round(sum(a["montant_ht"] for a in res["avoirs"]), 2)
            if res["total_net_ht"] is not None and abs(somme - res["total_net_ht"]) > 0.005:
                res["controles"].append(f"avoir : lignes {somme:.2f} ≠ pied de facture {res['total_net_ht']:.2f}")
        elif lignes:
            res["type_document"] = "facture"
            q = sum(l["qte_fact"] for l in lignes)
            if qm and int(qm.group(1)) != q:
                res["controles"].append(f"quantité lignes {q} ≠ TOTAL QTES FACTUREES {qm.group(1)}")

    if res["type_document"] == "inconnu":
        res["controles"].append("aucune ligne produit ni avoir reconnu : présentation de facture inconnue ?")
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
