"""
Module bdpm.py — Mapping CIP13 -> fabricant / nature du médicament.

Basé sur 3 fichiers officiels ANSM (téléchargeables gratuitement sur
https://base-donnees-publique.medicaments.gouv.fr/telechargement.php) :

  - CIS_bdpm.txt      : encodage latin-1, colonne 11 (index 10) = Titulaire de l'AMM
  - CIS_CIP_bdpm.txt  : encodage utf-8,   colonne 7 (index 6)  = CIP13
                                            colonne 9 (index 8)  = Taux de remboursement
  - CIS_GENER_bdpm.txt: encodage latin-1, colonne 4 (index 3)  = type générique
                                            (0=princeps, 1/2/4=générique/substituable)

Complété par le registre ANSM des groupes hybrides (téléchargeable sur
https://ansm.sante.fr/documents/reference/registre-des-groupes-hybrides), 2 fichiers :

  - fic02grp.txt : encodage latin-1, colonne 2 (index 1) = code groupe,
                                       colonne 4 (index 3) = libellé du groupe
  - fic03spe.txt : encodage latin-1, colonne 1 (index 0) = code groupe,
                                       colonne 2 (index 1) = code CIS,
                                       colonne 3 (index 2) = rôle "R" (référence)
                                       ou "H" (hybride)

⚠️ LIMITE CONNUE — à lire avant utilisation :
Ces 3 fichiers BDPM permettent de déterminer de façon fiable et 100% officielle :
  - le TITULAIRE de l'AMM (fabricant() ) — mais PAS le fournisseur commercial réel
    (ex: un générique vendu sous marque Biogaran peut avoir un titulaire GSK,
    cf. cas de co-exploitation identifiés précédemment). Pour le CA par génériqueur,
    utiliser le nom de l'OFFRE du BO-OFFREM (voir parser_bo_offrem.py), pas ce module.
  - le statut RÉPERTOIRE / PRINCEPS / HORS RÉPERTOIRE (via CIS_GENER)
  - le taux de REMBOURSEMENT (via CIS_CIP), donc les produits NON REMBOURSÉS
  - le statut HYBRIDE (via le registre ANSM ci-dessus, ajouté en complément)

Le statut BIOSIMILAIRE n'est en revanche PAS distingué à ce stade — l'ANSM ne le
code pas dans ces tables standard et le registre équivalent n'a pas encore été
intégré à ce module (piste à investiguer si besoin).

Un médicament hybride est traité comme une PRÉCISION de "Hors répertoire" (les
hybrides ne sont pas au répertoire des génériques) : le champ `repertoire` vaut
"Hors répertoire (Hybride)" au lieu de "Hors répertoire" dans ce cas, et le détail
complet (rôle R/H, libellé du groupe hybride) reste disponible via le champ `hybride`.
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
        self.cis_to_groupe = {}          # CIS -> code groupe générique
        self.groupe_to_cis = {}          # code groupe -> set(CIS)
        self.groupe_libelle = {}         # code groupe -> libellé du groupe

        # --- Registre ANSM des groupes hybrides ---
        self.cis_to_hybride_groupe = {}   # CIS -> code_groupe_hybride
        self.cis_to_hybride_role = {}     # CIS -> "R" ou "H"
        self.hybride_groupe_libelle = {}  # code_groupe_hybride -> libellé du groupe

        self._load_cis_cip(data_dir / "CIS_CIP_bdpm.txt")
        self._load_cis(data_dir / "CIS_bdpm.txt")
        self._load_cis_gener(data_dir / "CIS_GENER_bdpm.txt")
        self._load_hybrides(data_dir / "fic02grp.txt", data_dir / "fic03spe.txt")
        self._biogaran_groupes = self._compute_biogaran_groupes()

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
                groupe_id, libelle, cis, type_gen = cols[0], cols[1], cols[2], cols[3]
                self.cis_to_type_generique[cis] = type_gen.strip()
                self.cis_to_groupe[cis] = groupe_id
                self.groupe_to_cis.setdefault(groupe_id, set()).add(cis)
                self.groupe_libelle.setdefault(groupe_id, libelle.strip())

    def _load_hybrides(self, path_grp, path_spe):
        """Charge le registre ANSM des groupes hybrides, si les fichiers sont
        présents dans data_dir. Ne lève pas d'erreur si absents (fonctionnalité
        optionnelle tant que le workflow de téléchargement n'a pas tourné) —
        dans ce cas hybride() renvoie systématiquement est_hybride=False."""
        if path_grp.exists():
            with open(path_grp, encoding="latin-1") as f:
                for line in f:
                    cols = line.rstrip("\n").split("\t")
                    if len(cols) < 4:
                        continue
                    code_groupe, libelle = cols[1].strip(), cols[3].strip()
                    self.hybride_groupe_libelle[code_groupe] = libelle

        if path_spe.exists():
            with open(path_spe, encoding="latin-1") as f:
                for line in f:
                    cols = line.rstrip("\n").split("\t")
                    if len(cols) < 3:
                        continue
                    code_groupe, cis, role = cols[0].strip(), cols[1].strip(), cols[2].strip()
                    self.cis_to_hybride_groupe[cis] = code_groupe
                    self.cis_to_hybride_role[cis] = role

    def _compute_biogaran_groupes(self):
        """Codes groupe générique pour lesquels BIOGARAN est titulaire d'au moins un CIS.
        Sert de proxy pour 'ce générique est disponible chez Biogaran' (cf. limite
        dans le docstring du module : le titulaire n'est pas toujours strictement
        identique à la marque commerciale en cas de co-exploitation)."""
        groupes = set()
        for groupe_id, cis_set in self.groupe_to_cis.items():
            for cis in cis_set:
                if self.cis_to_titulaire.get(cis) == "BIOGARAN":
                    groupes.add(groupe_id)
                    break
        return groupes

    # -----------------------------------------------------------------
    def fabricant(self, cip13: str):
        """Titulaire de l'AMM (officiel BDPM). ⚠️ pas le fournisseur commercial."""
        cis = self.cip13_to_cis.get(cip13)
        if not cis:
            return None
        return self.cis_to_titulaire.get(cis)

    def hybride(self, cip13: str):
        """
        Retourne le statut hybride d'un CIP13, basé sur le registre ANSM des
        groupes hybrides (fic02grp.txt / fic03spe.txt dans data_dir).
        -> {"est_hybride": bool, "role": "R"/"H"/None, "groupe_hybride": libellé ou None}
        Si les fichiers du registre ne sont pas présents dans data_dir, renvoie
        toujours est_hybride=False (dégradation silencieuse, pas d'erreur).
        """
        cis = self.cip13_to_cis.get(cip13)
        if not cis or cis not in self.cis_to_hybride_groupe:
            return {"est_hybride": False, "role": None, "groupe_hybride": None}
        code_groupe = self.cis_to_hybride_groupe[cis]
        return {
            "est_hybride": True,
            "role": self.cis_to_hybride_role.get(cis),
            "groupe_hybride": self.hybride_groupe_libelle.get(code_groupe),
        }

    def nature(self, cip13: str):
        """
        Retourne un dict avec :
          - repertoire: "Générique (répertoire)" / "Princeps" / "Hors répertoire"
            / "Hors répertoire (Hybride)"
          - remboursement: taux (ex: "65%") ou "Non remboursé (ou non renseigné)"
          - hybride: détail complet, cf. méthode hybride() ci-dessus
        Ne distingue PAS biosimilaire (cf. limite en tête de module).
        """
        cis = self.cip13_to_cis.get(cip13)
        taux = self.cip13_to_taux_remb.get(cip13)

        if cis and cis in self.cis_to_type_generique:
            type_gen = self.cis_to_type_generique[cis]
            repertoire = "Princeps" if type_gen == "0" else "Générique (répertoire)"
        else:
            repertoire = "Hors répertoire"

        hyb = self.hybride(cip13)
        if repertoire == "Hors répertoire" and hyb["est_hybride"]:
            repertoire = "Hors répertoire (Hybride)"

        remboursement = taux if taux else "Non remboursé (ou non renseigné)"

        return {"repertoire": repertoire, "remboursement": remboursement, "hybride": hyb}

    def groupe_generique(self, cip13: str):
        """Retourne (code_groupe, libellé_groupe) ou (None, None) si hors répertoire."""
        cis = self.cip13_to_cis.get(cip13)
        if not cis:
            return None, None
        groupe_id = self.cis_to_groupe.get(cis)
        if not groupe_id:
            return None, None
        return groupe_id, self.groupe_libelle.get(groupe_id)

    def biogaran_disponible(self, cip13: str) -> bool:
        """True si le groupe générique de ce CIP13 compte au moins un CIS dont
        BIOGARAN est titulaire (donc un équivalent générique existe chez Biogaran)."""
        groupe_id, _ = self.groupe_generique(cip13)
        if not groupe_id:
            return False
        return groupe_id in self._biogaran_groupes


if __name__ == "__main__":
    import sys

    data_dir = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads"
    bdpm = BDPM(data_dir)
    print(f"CIP13 -> CIS chargés : {len(bdpm.cip13_to_cis)}")
    print(f"CIS -> Titulaire chargés : {len(bdpm.cis_to_titulaire)}")
    print(f"CIS -> Type générique chargés : {len(bdpm.cis_to_type_generique)}")
    print(f"CIS -> Groupe hybride chargés : {len(bdpm.cis_to_hybride_groupe)}")
