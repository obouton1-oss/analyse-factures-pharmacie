# -*- coding: utf-8 -*-
"""
Interface Streamlit pour l'analyse des factures pharmacie (canaux OCP,
Alliance Healthcare, Biogaran Direct).

Lancement :
    cd ~/projets_pharmacie/analyse-factures-pharmacie
    source ~/projets_pharmacie/venv/bin/activate
    streamlit run app.py

L'app ne fait qu'appeler build_report() de scripts/analyse_consolidee.py — la
logique d'analyse elle-même (parsers, règles métier) ne change pas.
"""
import gc
import sys
import tempfile
import threading
import shutil
import traceback
import unicodedata
import io
import re
import zipfile
import hashlib
from contextlib import contextmanager, redirect_stdout, redirect_stderr
from datetime import datetime
from pathlib import Path

import pdfplumber
import streamlit as st

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "scripts"))

from analyse_consolidee import build_report  # noqa: E402


# ---------------------------------------------------------------------------
# Config page + style
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Analyses factures génériques",
    page_icon="💊",
    layout="wide",
)

NAVY = "#1F3864"
NAVY_LIGHT = "#2E5395"

st.markdown(f"""
<style>
    .main {{ background-color: #FAFBFC; }}
    .app-header {{
        background: linear-gradient(135deg, {NAVY} 0%, {NAVY_LIGHT} 100%);
        padding: 1.6rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.6rem;
        color: white;
    }}
    .app-header h1 {{ margin: 0; font-size: 1.6rem; }}
    .app-header p {{ margin: 0.3rem 0 0 0; opacity: 0.85; font-size: 0.95rem; }}

    .canal-card {{
        background: white;
        border: 1px solid #E5E7EB;
        border-top: 5px solid var(--accent, {NAVY});
        border-radius: 10px;
        padding: 1rem 1.1rem 0.6rem 1.1rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .canal-card h3 {{ margin-top: 0; font-size: 1.05rem; }}
    .canal-card .sub {{ color: #6B7280; font-size: 0.85rem; margin-bottom: 0.6rem; }}

    .kpi-box {{
        background: white;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 0.9rem 1rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }}
    .kpi-box .val {{ font-size: 1.5rem; font-weight: 700; color: {NAVY}; }}
    .kpi-box .lbl {{ font-size: 0.8rem; color: #6B7280; margin-top: 0.2rem; }}

    div[data-testid="stFileUploader"] section {{
        border-radius: 8px;
    }}
    /* On garde la liste des fichiers déposés visible (chaque fichier a sa
       propre croix pour le retirer en cas d'erreur), mais on la limite en
       hauteur avec défilement interne, et surtout on fait remonter le
       bouton "+" (ajouter des fichiers) tout en haut de cette zone au lieu
       de le laisser tout en bas après le dernier fichier : avec beaucoup de
       PDF (factures Alliance en particulier), il finissait repoussé hors de
       vue, obligeant à dérouler toute la liste pour le retrouver.
       Le sélecteur "stFileUploaderFile" est gardé en compatibilité avec les
       anciennes versions de Streamlit (avant le passage au système de
       "chips" "stFileChips"). */
    div[data-testid="stFileChips"] {{
        max-height: 170px;
        overflow-y: auto;
        align-content: flex-start;
        padding-right: 4px;
    }}
    div[data-testid="stFileChips"] > button[aria-label="Add files"] {{
        order: -1;
        position: sticky;
        top: 0;
        z-index: 1;
        background: {NAVY};
        color: white;
        width: 2.2rem;
        height: 2.2rem;
        min-height: 2.2rem;
        border-radius: 50%;
        box-shadow: 0 0 0 3px #FAFBFC;
    }}
    div[data-testid="stFileChips"] > button[aria-label="Add files"]:hover {{
        background: {NAVY_LIGHT};
        color: white;
    }}
    div[data-testid="stFileChips"] > button[aria-label="Add files"] span {{
        font-size: 1.4rem;
    }}
    div[data-testid="stFileUploaderFile"] {{
        display: none !important;
    }}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Protection par mot de passe
# ---------------------------------------------------------------------------
def verifier_mot_de_passe():
    """Affiche un écran de connexion tant que le bon mot de passe n'a pas été
    saisi. Le mot de passe attendu est lu dans st.secrets (fichier
    .streamlit/secrets.toml en local, section "Secrets" des réglages de
    l'app sur Streamlit Cloud) — jamais écrit en clair dans le code."""
    if st.session_state.get("authenticated"):
        return True

    mot_de_passe_attendu = st.secrets.get("password")

    st.markdown("""
    <div class="app-header">
        <h1>💊 Analyses factures génériques</h1>
        <p>Accès protégé par mot de passe.</p>
    </div>
    """, unsafe_allow_html=True)

    if not mot_de_passe_attendu:
        st.error(
            "Aucun mot de passe configuré. Ajoute-le dans .streamlit/secrets.toml "
            "(en local) ou dans les réglages « Secrets » de l'app sur Streamlit Cloud, "
            'sous la forme : password = "..."'
        )
        st.stop()

    _, col_centre, _ = st.columns([1, 1, 1])
    with col_centre:
        saisie = st.text_input("Mot de passe", type="password", key="pwd_input")
        if st.button("Se connecter", type="primary", use_container_width=True):
            if saisie == mot_de_passe_attendu:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Mot de passe incorrect.")
    st.stop()


verifier_mot_de_passe()


st.markdown(f"""
<div class="app-header">
    <h1>💊 Analyses factures génériques</h1>
    <p>Dépose le dossier de tes factures (tous mois confondus), puis génère le rapport Excel consolidé.</p>
</div>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# État de session
# ---------------------------------------------------------------------------
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0
if "report_bytes" not in st.session_state:
    st.session_state.report_bytes = None
if "report_name" not in st.session_state:
    st.session_state.report_name = None
if "log_text" not in st.session_state:
    st.session_state.log_text = ""
if "kpis" not in st.session_state:
    st.session_state.kpis = None
if "error_text" not in st.session_state:
    st.session_state.error_text = None

k = st.session_state.reset_counter


def nouvelle_analyse():
    st.session_state.reset_counter += 1
    st.session_state.fichiers_liberes = False
    st.session_state.report_bytes = None
    st.session_state.report_name = None
    st.session_state.log_text = ""
    st.session_state.kpis = None
    st.session_state.error_text = None


def deplier_fichiers(files):
    """Retourne une liste de (nom_fichier, contenu, clé) à partir des
    fichiers déposés : les PDF sont pris tels quels, les .zip sont dépliés
    (chaque PDF trouvé à l'intérieur est extrait). Les fichiers système macOS
    (__MACOSX, ._xxx) et tout ce qui n'est pas un PDF (.DS_Store, Excel...)
    sont ignorés — utile en dépôt de dossier, où tout le contenu du dossier
    est envoyé. Le nom est réduit au nom de fichier seul (sans le chemin des
    sous-dossiers)."""
    # Mémoire : un PDF déposé est passé tel quel (le fichier déjà chargé par
    # Streamlit), lu sans copie au moment voulu (voir octets()). Seuls les PDF extraits
    # d'un .zip occupent de la mémoire en plus (d'où l'intérêt de déposer des
    # dossiers plutôt que des .zip). Chaque élément porte aussi une clé stable
    # (identifiant Streamlit du fichier) pour ne calculer l'empreinte qu'une fois.
    resultat = []
    for f in files or []:
        nom = Path(f.name).name
        if nom.startswith("._"):
            continue
        cle = getattr(f, "file_id", None) or f"{f.name}:{id(f)}"
        if nom.lower().endswith(".zip"):
            try:
                f.seek(0)
                with zipfile.ZipFile(f) as zf:
                    for entry in zf.namelist():
                        if not entry.lower().endswith(".pdf"):
                            continue
                        if entry.startswith("__MACOSX") or "/._" in entry or entry.startswith("._"):
                            continue
                        resultat.append((Path(entry).name, zf.read(entry), f"{cle}/{entry}"))
            except zipfile.BadZipFile:
                st.warning(f"⚠️ {nom} n'est pas une archive .zip valide, ignoré.")
        elif nom.lower().endswith(".pdf"):
            resultat.append((nom, f, cle))   # le fichier lui-même, lu à la demande
    return resultat


@contextmanager
def octets(contenu):
    """Accès aux octets d'un PDF : bytes (extrait d'un .zip) ou fichier déposé.
    Pour un fichier déposé, vue mémoire SANS copie, libérée dès la sortie du
    bloc (une vue laissée ouverte empêche Streamlit de libérer le fichier)."""
    if isinstance(contenu, (bytes, bytearray)):
        yield contenu
    else:
        with contenu.getbuffer() as vue:
            yield vue


# Reconnaissance du canal d'un PDF d'après le texte de sa 1re page (accents
# retirés, majuscules). Règles vérifiées le 20/11/2026 sur les 2 417 PDF de
# factures_a_traiter/archive : 100 % concordants avec le tri existant
# (récaps OCP, Alliance, Biogaran Direct CSP et Biogaran en direct, avoirs
# RDP Biogaran ; factures OCP individuelles et EMAC non reconnues, donc
# ignorées comme attendu).
def _normaliser(texte):
    return "".join(
        c for c in unicodedata.normalize("NFKD", texte) if not unicodedata.combining(c)
    ).upper()


def detecter_canal(contenu):
    """Canal du PDF : "ocp" (récap BO-OFFREM), "biosim_ocp" (relevé OCP des
    achats de biosimilaires), "direct_labo" (factures directes EG LABO, SANDOZ, TEVA),
    "alliance", "biogaran",
    "avoir_biogaran" (avoir RDP, non pris en compte), ou None s'il n'est pas
    reconnu (facture OCP individuelle, EMAC, autre fournisseur, PDF
    illisible...)."""
    try:
        with octets(contenu) as b, pdfplumber.open(io.BytesIO(b)) as pdf:
            if not pdf.pages:
                return None
            t = _normaliser(pdf.pages[0].extract_text() or "")
    except Exception:
        return None
    if "PRODUITS BIOSIMILAIRES" in t and "BIOST" in t:
        return "biosim_ocp"   # relevé OCP « Achats des produits BIOSIMILAIRES »
    if "BO-OFFREM" in t:
        return "ocp"
    if "ALLIANCE HEALTHCARE" in t:
        return "alliance"
    if ("POUR COMPTE DES LABORATOIRES EG LABO" in t or "DENOMINATION SOCIALE : EG LABO" in t
            or "DENOMINATION SOCIALE : SANDOZ" in t or "TEVA SANTE" in t):
        return "direct_labo"  # factures et avoirs EG LABO / SANDOZ / TEVA en direct
    if "BIOGARAN" in t and "RECAPITULATIF DES REMISES" in t:
        return "avoir_biogaran"
    if ("BIOGARAN" in t and "AU NOM ET POUR LE COMPTE DE" in t) \
            or "DENOMINATION SOCIALE : BIOGARAN" in t:
        return "biogaran"
    return None


def classer_automatiquement(fichiers):
    """Ajoute le canal détecté à chaque (nom, contenu, empreinte). Le résultat
    est mémorisé par empreinte pour la session : Streamlit relance le script à
    chaque clic, on ne relit donc chaque PDF qu'une seule fois."""
    cache = st.session_state.setdefault("canal_cache", {})
    a_lire = [x for x in fichiers if x[2] not in cache]
    if a_lire:
        barre = st.progress(0.0, text=f"Reconnaissance des factures (0/{len(a_lire)})…")
        for i, (nom, contenu, empreinte) in enumerate(a_lire, 1):
            cache[empreinte] = detecter_canal(contenu)
            if i % 10 == 0 or i == len(a_lire):
                barre.progress(i / len(a_lire), text=f"Reconnaissance des factures ({i}/{len(a_lire)})…")
        barre.empty()
    return [(nom, contenu, empreinte, cache[empreinte]) for nom, contenu, empreinte in fichiers]


def avec_empreinte(fichiers):
    """(nom, contenu, clé) -> (nom, contenu, empreinte sha256). L'empreinte
    est mémorisée par clé : chaque fichier n'est haché qu'une fois par session,
    pas à chaque clic."""
    cache = st.session_state.setdefault("empreinte_cache", {})
    res = []
    for nom, contenu, cle in fichiers:
        if cle not in cache:
            with octets(contenu) as b:
                cache[cle] = hashlib.sha256(b).hexdigest()
        res.append((nom, contenu, cache[cle]))
    return res


@st.cache_resource
def file_attente():
    """File d'attente commune à tous les utilisateurs de l'appli (même
    processus) : une seule analyse à la fois, pour ne pas dépasser la mémoire
    allouée par Streamlit Cloud ; les suivantes attendent leur tour."""
    return {"verrou": threading.Lock(), "mutex": threading.Lock(), "inscrits": 0}


def liberer_memoire():
    """Après une analyse réussie : nouvelles clés de zones de dépôt (Streamlit
    libère les fichiers déposés au rechargement), vidage des caches de
    session, ramasse-miettes Python."""
    st.session_state.reset_counter += 1
    st.session_state.pop("canal_cache", None)
    st.session_state.pop("empreinte_cache", None)
    st.session_state.fichiers_liberes = True
    gc.collect()


def uploader_dossiers(label, key):
    """Zone de dépôt acceptant des dossiers entiers (sous-dossiers compris).
    Repli sur le dépôt multi-fichiers classique si la version de Streamlit
    installée ne connaît pas encore le mode "directory"."""
    # Pas de filtre de type côté navigateur en mode dossier : sinon chaque
    # fichier non-PDF des dossiers (.DS_Store créé par macOS...) s'affiche
    # en erreur rouge. Le tri PDF/.zip est fait par deplier_fichiers().
    try:
        return st.file_uploader(
            label, accept_multiple_files="directory",
            key=key, label_visibility="collapsed",
        )
    except Exception:
        return st.file_uploader(
            label, type=["pdf", "zip"], accept_multiple_files=True,
            key=key, label_visibility="collapsed",
        )


# ---------------------------------------------------------------------------
# Liste des documents à fournir
# ---------------------------------------------------------------------------
with st.expander("📄 Documents à fournir pour un rapport complet", expanded=False):
    st.markdown("""
| Document | Ce qu'il alimente dans le rapport | Remarque |
|---|---|---|
| **OCP — récapitulatif mensuel des offres** (« BO-OFFREM ») | CA OCP par offre et par génériqueur, remises, fuite Biogaran, assiettes Biogaran et EG, rattrapage Teva (grossistes), remises des biosimilaires achetés chez OCP | Un par mois |
| **OCP — relevé des achats de produits biosimilaires** (« BIOST ») | Onglet *Biosimilaires* (achats OCP avec ou sans offre, part de marché Sandoz) | Cumul depuis janvier : seul le plus récent de l'année est utilisé |
| **Alliance Healthcare — factures** | CA Alliance par génériqueur, fuite, assiettes Biogaran et EG, rattrapage Teva (grossistes) | Les relevés mensuels Alliance sont ignorés automatiquement |
| **Biogaran — factures directes** (CSP / Movianto) | Canal *Biogaran Direct*, assiette de remise Biogaran | |
| **EG LABO — factures directes** (CSP / Movianto) | Canal *EG Direct*, assiette de remise EG | |
| **Teva Santé — factures directes** | Canal *Teva Direct*, rattrapage de remise Teva (direct) | |
| **Sandoz — factures directes** | Canal *Sandoz Direct*, onglet *Biosimilaires* | |

**Inutiles (reconnus et ignorés)** : factures OCP individuelles, relevés EMAC, relevés de compte OCP, avoirs Biogaran (RDP), EG, Teva et Sandoz.

Pour un rapport complet sur une période, fournir **tous** les documents de la période pour chaque canal utilisé : un mois manquant chez un fournisseur fait baisser son CA sans message d'erreur.
""")


# ---------------------------------------------------------------------------
# Dépôt global : un ou plusieurs dossiers, tri automatique par canal
# ---------------------------------------------------------------------------
st.markdown('<div class="canal-card" style="--accent:#1F3864">', unsafe_allow_html=True)
st.markdown("### 📂 Dépôt des factures — tous canaux, tous mois")
st.markdown(
    '<div class="sub">Clique sur <b>Upload directories</b> et choisis le dossier qui contient tous tes '
    'dossiers mensuels (par ex. « Factures 2026 ») : tous les sous-dossiers sont lus d\'un coup. '
    'Chaque PDF est reconnu automatiquement (récap OCP, relevé OCP des biosimilaires, Alliance, Biogaran Direct, factures directes EG, Sandoz ou Teva) d\'après son contenu, '
    'peu importe le dossier où il est rangé. Le bouton <b>+</b> permet ensuite d\'ajouter un autre dossier ; '
    'les <b>+</b> sous les compteurs ajoutent des PDF directement dans un canal.</div>',
    unsafe_allow_html=True,
)
st.info(
    "⚠️ Les factures OCP individuelles (plus d'un millier par an, ~200 Mo) et les EMAC "
    "ne servent pas au rapport : si elles se trouvent dans le dossier choisi, elles sont "
    "envoyées puis ignorées, ce qui rallonge nettement l'envoi. Mieux vaut les ranger "
    "dans un dossier à part avant le dépôt."
)
global_raw = uploader_dossiers("Dossiers de factures", key=f"global_{k}")
st.markdown('</div>', unsafe_allow_html=True)

# Compteurs par canal, chacun avec son « ➕ » pour ajouter des PDF
# directement dans ce canal (sans reconnaissance automatique) — utile pour
# un PDF que la zone ci-dessus n'a pas reconnu. Les bulles sont remplies plus
# bas, une fois tous les dépôts regroupés.
m1, m2, m3, m4, m5 = st.columns(5)
bulles = {}
manuels = {}
for col, canal, titre in (
    (m1, "ocp", "Récaps OCP"),
    (m2, "alliance", "PDF Alliance"),
    (m3, "biogaran", "Factures Biogaran"),
    (m4, "direct_labo", "Directs labos"),
):
    with col:
        bulles[canal] = st.empty()
        with st.popover("➕", help=f"Ajouter des PDF dans « {titre} »", width="stretch"):
            st.caption("PDF ou .zip ajoutés directement dans ce canal."
                       + (" Le relevé OCP des biosimilaires peut aussi être ajouté ici."
                          if canal == "ocp" else ""))
            manuels[canal] = st.file_uploader(
                f"Ajouter {titre}", type=["pdf", "zip"], accept_multiple_files=True,
                key=f"{canal}_{k}", label_visibility="collapsed",
            )
with m5:
    bulles["ignores"] = st.empty()
ocp_raw, alliance_raw, biogaran_raw = manuels["ocp"], manuels["alliance"], manuels["biogaran"]
direct_raw = manuels["direct_labo"]

# Regroupement : dépôt global (canal détecté) + dépôts manuels (canal imposé),
# dédoublonnage par contenu sur l'ensemble.
entrees = classer_automatiquement(avec_empreinte(deplier_fichiers(global_raw)))
# Ajout manuel OCP : le relevé biosimilaires y est reconnu et mis à part,
# tout le reste est traité comme récap BO-OFFREM.
entrees += [(nom, c, e, "biosim_ocp" if canal == "biosim_ocp" else "ocp")
            for nom, c, e, canal in classer_automatiquement(avec_empreinte(deplier_fichiers(ocp_raw)))]
for raw, canal in ((alliance_raw, "alliance"), (biogaran_raw, "biogaran"), (direct_raw, "direct_labo")):
    entrees += [(nom, c, e, canal) for nom, c, e in avec_empreinte(deplier_fichiers(raw))]

par_canal = {"ocp": [], "biosim_ocp": [], "alliance": [], "biogaran": [], "direct_labo": []}
ignores = []   # (nom, raison)
vus = set()
doublons = 0
for nom, contenu, empreinte, canal in entrees:
    if empreinte in vus:
        doublons += 1
        continue
    vus.add(empreinte)
    if canal is None:
        ignores.append((nom, "non reconnu (facture OCP individuelle, EMAC, autre fournisseur…)"))
    elif canal == "avoir_biogaran" or (canal == "biogaran" and nom.lower().startswith("avoir_")):
        ignores.append((nom, "avoir Biogaran (non pris en compte)"))
    else:
        par_canal[canal].append((nom, contenu))

ocp_files = par_canal["ocp"]
alliance_files = par_canal["alliance"]
biogaran_files = par_canal["biogaran"]
biosim_files = par_canal["biosim_ocp"]
direct_files = par_canal["direct_labo"]

for cle, val, lbl in (
    ("ocp", len(ocp_files), "📋 Récaps OCP"),
    ("alliance", len(alliance_files), "🏥 PDF Alliance"),
    ("biogaran", len(biogaran_files), "💊 Factures Biogaran"),
    ("direct_labo", len(direct_files), "🏭 Directs labos (EG, Sandoz, Teva)"),
    ("ignores", len(ignores), "Ignorés"),
):
    bulles[cle].markdown(f"""
    <div class="kpi-box">
        <div class="val">{val}</div>
        <div class="lbl">{lbl}</div>
    </div>
    """, unsafe_allow_html=True)

if entrees:
    if biosim_files:
        st.caption(f"📈 {len(biosim_files)} relevé(s) OCP des biosimilaires reconnu(s) — onglet "
                   "« Biosimilaires » ; seul le plus récent de chaque année est utilisé "
                   "(chaque relevé reprend le cumul depuis janvier).")
    if doublons:
        st.caption(f"🔁 {doublons} doublon(s) détecté(s) et ignoré(s) (fichier identique déjà déposé).")
    if ignores:
        with st.expander(f"Voir les {len(ignores)} fichier(s) ignoré(s)"):
            st.dataframe(
                [{"Fichier": n, "Raison": r} for n, r in sorted(ignores)],
                width="stretch", hide_index=True,
            )

st.write("")
bcol1, bcol2, _ = st.columns([1, 1, 3])
generer = bcol1.button("🚀 Générer le rapport", type="primary", use_container_width=True)
if bcol2.button("🔄 Nouvelle analyse", use_container_width=True):
    nouvelle_analyse()
    st.rerun()

total_fichiers = (len(ocp_files) + len(alliance_files) + len(biogaran_files) + len(biosim_files)
                  + len(direct_files))


# ---------------------------------------------------------------------------
# Génération du rapport
# ---------------------------------------------------------------------------
def parse_kpis(log_text):
    kpis = {}
    m = re.search(r"(\d+) lignes produit, (\d+) lignes de fuite", log_text)
    if m:
        kpis["lignes"] = m.group(1)
        kpis["fuite"] = m.group(2)
    m = re.search(r"CA PPHT total\s*:\s*([\d\s.,]+)\s*€", log_text)
    if m:
        kpis["ca_total"] = m.group(1).strip() + " €"
    m = re.search(r"Périodes\s*:\s*(.+)", log_text)
    if m:
        kpis["periodes"] = m.group(1).strip()
    kpis["n_warnings"] = log_text.count("⚠️")
    return kpis


tour = False
if generer:
    if total_fichiers == 0:
        st.warning("Aucune facture reconnue : dépose au moins un dossier ou un PDF avant de générer le rapport.")
    else:
        fa = file_attente()
        with fa["mutex"]:
            fa["inscrits"] += 1
        message = st.empty()
        try:
            # Attente de son tour, message mis à jour toutes les 2 s. Si
            # l'utilisateur ferme la page, l'appel Streamlit suivant
            # interrompt l'attente et la place est rendue (bloc finally).
            while not fa["verrou"].acquire(timeout=2):
                avant = max(1, fa["inscrits"] - 1)
                message.info(f"⏳ {avant} analyse(s) en cours ou en attente avant la vôtre. "
                             "Elle démarrera automatiquement : gardez la page ouverte.")
            tour = True
            message.empty()
        finally:
            if not tour:
                with fa["mutex"]:
                    fa["inscrits"] -= 1
    if generer and total_fichiers and tour:
        with st.spinner("Analyse en cours… (peut prendre plusieurs minutes si beaucoup de factures Alliance)"):
            tmp_dir = Path(tempfile.mkdtemp(prefix="analyse_pharma_"))
            try:
                def sauver(files, sous_dossier):
                    dossier = tmp_dir / sous_dossier
                    dossier.mkdir(parents=True, exist_ok=True)
                    chemins = []
                    for nom, contenu in files or []:
                        # Deux PDF différents peuvent porter le même nom
                        # (ex. un récap OCP par mois rangé dans des dossiers
                        # mensuels différents) : on numérote au lieu
                        # d'écraser.
                        chemin = dossier / nom
                        n = 1
                        while chemin.exists():
                            chemin = dossier / f"{Path(nom).stem}_{n}{Path(nom).suffix}"
                            n += 1
                        with octets(contenu) as b:
                            chemin.write_bytes(b)
                        chemins.append(str(chemin))
                    return chemins

                ocp_paths = sauver(ocp_files, "ocp")
                alliance_paths = sauver(alliance_files, "alliance")
                biogaran_paths = sauver(biogaran_files, "biogaran_direct")
                biosim_paths = sauver(biosim_files, "biosimilaires_ocp")
                direct_paths = sauver(direct_files, "direct_labo")

                out_path = tmp_dir / "rapport_pharmacie.xlsx"
                data_dir = str(ROOT / "data")

                log_buffer = io.StringIO()
                with redirect_stdout(log_buffer), redirect_stderr(log_buffer):
                    build_report(
                        ocp_paths, data_dir, str(out_path),
                        biogaran_direct_pdfs=biogaran_paths or None,
                        alliance_pdfs=alliance_paths or None,
                        biosimilaires_ocp_pdfs=biosim_paths or None,
                        direct_labo_pdfs=direct_paths or None,
                    )

                log_text = log_buffer.getvalue()
                st.session_state.log_text = log_text
                st.session_state.kpis = parse_kpis(log_text)
                st.session_state.report_bytes = out_path.read_bytes()
                st.session_state.report_name = f"rapport_pharmacie_{datetime.now():%Y%m%d_%H%M}.xlsx"
                st.session_state.error_text = None
                liberer_memoire()
            except Exception:
                st.session_state.error_text = traceback.format_exc()
                st.session_state.report_bytes = None
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
                fa["verrou"].release()
                with fa["mutex"]:
                    fa["inscrits"] -= 1
                gc.collect()
        st.rerun()


# ---------------------------------------------------------------------------
# Résultats
# ---------------------------------------------------------------------------
if st.session_state.error_text:
    st.error("Une erreur est survenue pendant la génération du rapport.")
    with st.expander("Détails de l'erreur"):
        st.code(st.session_state.error_text)

if st.session_state.report_bytes:
    st.success("Rapport généré avec succès ✅")
    if st.session_state.get("fichiers_liberes"):
        st.caption("🧹 Les fichiers déposés ont été retirés de la mémoire après l'analyse. "
                   "Pour une nouvelle analyse, redéposez vos dossiers.")

    kpis = st.session_state.kpis or {}
    kcols = st.columns(4)
    kpi_defs = [
        ("lignes", "Lignes produit"),
        ("ca_total", "CA PPHT total"),
        ("fuite", "Lignes de fuite"),
        ("n_warnings", "Avertissements"),
    ]
    for col, (key, label) in zip(kcols, kpi_defs):
        val = kpis.get(key, "—")
        col.markdown(f"""
        <div class="kpi-box">
            <div class="val">{val}</div>
            <div class="lbl">{label}</div>
        </div>
        """, unsafe_allow_html=True)

    if kpis.get("periodes"):
        st.caption(f"Périodes analysées : {kpis['periodes']}")

    st.write("")
    st.download_button(
        "⬇️ Télécharger le rapport Excel",
        data=st.session_state.report_bytes,
        file_name=st.session_state.report_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    with st.expander("Voir le détail (avertissements, anomalies…)"):
        st.code(st.session_state.log_text or "(aucune sortie)")
