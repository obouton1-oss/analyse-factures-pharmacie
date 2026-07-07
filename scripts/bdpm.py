"""
Module bdpm.py — Mapping CIP13 -> fabricant / nature du médicament.

Basé sur 3 fichiers officiels ANSM (téléchargeables gratuitement sur
https://base-donnees-publique.medicaments.gouv.fr/telechargement.php) :

  - CIS_bdpm.txt      : encodage latin-1, colonne 11 (index 10) = Titulaire de l'AMM
  - CIS_CIP_bdpm.txt  : encodage utf-8,   colonne 7 (index 6)  = CIP13
                                            colonne 9 (index 8)  = Taux de remboursement
  - CIS_GENER_bdpm.txt: encodage latin-1, colonne 4 (index 3)  = type générique
                                            (0=princeps, 1/2/4=générique/substituable)

⚠️ LIMITE CONNUE — à lire avant utilisation :
Ces 3 fichiers permettent de déterminer de façon fiable et 100% officielle :
  - le TITULAIRE de l'AMM (fabricant() ) — mais PAS le fournisseur commercial réel
    (ex: un générique vendu sous marque Biogaran peut avoir un titulaire GSK,
    cf. cas de co-exploitation identifiés précédemment). Pour le CA par génériqueur,
    utiliser le nom de l'OFFRE du BO-OFFREM (voir parser_bo_offrem.py), pas ce module.
  - le statut RÉPERTOIRE / PRINCEPS / HORS RÉPERTOIRE (via CIS_GENER)
  - le taux de REMBOURSEMENT (via CIS_CIP), donc les produits NON REMBOURSÉS

Ces 3 fichiers NE PERMETTENT PAS, en l'état, de distinguer de façon fiable les
"HYBRIDES" ni les "BIOSIMILAIRES" — l'ANSM ne les code pas dans ces tables standard.
Ces deux catégories nécessitent une source complémentaire (ex. liste ANSM des
médicaments biosimilaires, ou le champ "type de procédure" combiné à une liste
externe). Le champ `nature()` renvoie donc explicitement "Répertoire (générique)",
"Princeps / hors répertoire" ou "Non identifié" — PAS "hybride" ou "biosimilaire"
tant que cette source complémentaire n'est pas intégrée.
"""
import csv
from pathlib import Path

# ---------------------------------------------------------------------------
# Normalisation des variantes de raison sociale (rachats, orthographes multiples)
# ---------------------------------------------------------------------------
NORMALISATION_TITULAIRE = {
    "MYLAN SAS": "VIATRIS",
    "VIATRIS SANTE": "VIATRIS",
    "VIATRIS SANTÉ": "VIATRIS",
    "EG LABO": "EG LABO - LABORATOIRES EUROGENERICS",
    "EG LABO - LABORATOIRES EUROGENERICS": "EG LABO - LABORATOIRES EUROGENERICS",
    "TEVA SANTE": "TEVA",
    "TEVA SANTÉ": "TEVA",
    "SANDOZ": "SANDOZ",
}


def _normalise(titulaire: str) -> str:
    t = titulaire.strip()
    return NORMALISATION_TITULAIRE.get(t.upper(), t)


class BDPM:
    def __init__(self, data_dir):
        data_dir = Path(data_dir)
        self.cip13_to_cis = {}
        self.cip13_to_taux_remb = {}
        self.cis_to_titulaire = {}
        self.cis_to_type_generique = {}  # 0=princeps, 1/2/4=générique

        self._load_cis_cip(data_dir / "CIS_CIP_bdpm.txt")
        self._load_cis(data_dir / "CIS_bdpm.txt")
        self._load_cis_gener(data_dir / "CIS_GENER_bdpm.txt")

    def _load_cis_cip(self, path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                cols = line.rstrip("\n").split("\t")
                if len(cols) < 9:
                    continue
                cis, cip13, taux_remb = cols[0], cols[6], cols[8]
                self.cip13_to_cis[cip13] = cis
                self.cip13_to_taux_remb[cip13] = taux_remb.strip() or None

    def _load_cis(self, path):
        with open(path, encoding="latin-1") as f:
            for line in f:
                cols = line.rstrip("\n").split("\t")
                if len(cols) < 11:
                    continue
                cis, titulaire = cols[0], cols[10]
                if titulaire.strip():
                    self.cis_to_titulaire[cis] = _normalise(titulaire)

    def _load_cis_gener(self, path):
        with open(path, encoding="latin-1") as f:
            for line in f:
                cols = line.rstrip("\n").split("\t")
                if len(cols) < 4:
                    continue
                cis, type_gen = cols[2], cols[3]
                self.cis_to_type_generique[cis] = type_gen.strip()

    # -----------------------------------------------------------------
    def fabricant(self, cip13: str):
        """Titulaire de l'AMM (officiel BDPM). ⚠️ pas le fournisseur commercial."""
        cis = self.cip13_to_cis.get(cip13)
        if not cis:
            return None
        return self.cis_to_titulaire.get(cis)

    def nature(self, cip13: str):
        """
        Retourne un dict avec :
          - repertoire: "Générique (répertoire)" / "Princeps" / "Hors répertoire"
          - remboursement: taux (ex: "65%") ou "Non remboursé (ou non renseigné)"
        Ne distingue PAS hybride / biosimilaire (cf. limite en tête de module).
        """
        cis = self.cip13_to_cis.get(cip13)
        taux = self.cip13_to_taux_remb.get(cip13)

        if cis and cis in self.cis_to_type_generique:
            type_gen = self.cis_to_type_generique[cis]
            repertoire = "Princeps" if type_gen == "0" else "Générique (répertoire)"
        else:
            repertoire = "Hors répertoire"

        remboursement = taux if taux else "Non remboursé (ou non renseigné)"

        return {"repertoire": repertoire, "remboursement": remboursement}


if __name__ == "__main__":
    import sys

    data_dir = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads"
    bdpm = BDPM(data_dir)
    print(f"CIP13 -> CIS chargés : {len(bdpm.cip13_to_cis)}")
    print(f"CIS -> Titulaire chargés : {len(bdpm.cis_to_titulaire)}")
    print(f"CIS -> Type générique chargés : {len(bdpm.cis_to_type_generique)}")
