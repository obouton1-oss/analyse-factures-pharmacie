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
import sys
import tempfile
import shutil
import traceback
import io
import re
import zipfile
import hashlib
from contextlib import redirect_stdout, redirect_stderr
from datetime import datetime
from pathlib import Path

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
    /* On masque la liste des fichiers déjà déposés : avec beaucoup de PDF
       (factures Alliance en particulier), cette liste peut devenir très
       longue et repousser le bouton "Browse files" hors de vue, obligeant à
       dérouler toute la zone pour en ajouter d'autres. Le nombre de fichiers
       détectés est de toute façon déjà affiché juste en dessous (st.caption),
       donc cette liste native est redondante. */
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
    <p>Dépose les PDF de chaque canal ci-dessous, puis génère le rapport Excel consolidé.</p>
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
    st.session_state.report_bytes = None
    st.session_state.report_name = None
    st.session_state.log_text = ""
    st.session_state.kpis = None
    st.session_state.error_text = None


def expand_uploads(files):
    """Retourne une liste de (nom_fichier, contenu_bytes) à partir des
    fichiers déposés : les PDF sont pris tels quels, les .zip sont dépliés
    (chaque PDF trouvé à l'intérieur est extrait, les dossiers/fichiers
    système macOS type __MACOSX ou ._xxx sont ignorés). Permet de déposer un
    dossier entier compressé plutôt que de sélectionner chaque PDF un par
    un — le navigateur ne permet pas de glisser un dossier non compressé
    directement sur la zone de dépôt.

    Déduplique aussi au passage par empreinte du contenu (sha256) : si le
    même PDF est déposé deux fois (dépôt en plusieurs fois, PDF présent à la
    fois seul et dans un .zip...), il n'est gardé qu'une fois."""
    resultat = []
    vus = set()
    doublons = 0

    def ajouter(nom, contenu):
        nonlocal doublons
        empreinte = hashlib.sha256(contenu).hexdigest()
        if empreinte in vus:
            doublons += 1
            return
        vus.add(empreinte)
        resultat.append((nom, contenu))

    for f in files or []:
        nom = f.name
        if nom.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(f.getvalue())) as zf:
                    for entry in zf.namelist():
                        if not entry.lower().endswith(".pdf"):
                            continue
                        if entry.startswith("__MACOSX") or "/._" in entry or entry.startswith("._"):
                            continue
                        ajouter(Path(entry).name, zf.read(entry))
            except zipfile.BadZipFile:
                st.warning(f"⚠️ {nom} n'est pas une archive .zip valide, ignoré.")
        else:
            ajouter(nom, f.getvalue())

    if doublons:
        st.caption(f"🔁 {doublons} doublon(s) détecté(s) et ignoré(s) (fichier identique déjà déposé).")
    return resultat


# ---------------------------------------------------------------------------
# Zones de dépôt (3 canaux)
# ---------------------------------------------------------------------------
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="canal-card" style="--accent:#2E86C1">', unsafe_allow_html=True)
    st.markdown("### 📋 OCP")
    st.markdown('<div class="sub">Récapitulatifs mensuels BO-OFFREM — PDF individuels ou dossier zippé</div>', unsafe_allow_html=True)
    ocp_raw = st.file_uploader(
        "PDF OCP", type=["pdf", "zip"], accept_multiple_files=True,
        key=f"ocp_{k}", label_visibility="collapsed",
    )
    ocp_files = expand_uploads(ocp_raw)
    st.caption(f"{len(ocp_files)} PDF détecté(s)" if ocp_files else "Aucun fichier pour l'instant")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="canal-card" style="--accent:#28B463">', unsafe_allow_html=True)
    st.markdown("### 🏥 Alliance Healthcare")
    st.markdown('<div class="sub">Factures grossiste (relevés mensuels ignorés automatiquement) — PDF individuels ou dossier zippé</div>', unsafe_allow_html=True)
    alliance_raw = st.file_uploader(
        "PDF Alliance", type=["pdf", "zip"], accept_multiple_files=True,
        key=f"alliance_{k}", label_visibility="collapsed",
    )
    alliance_files = expand_uploads(alliance_raw)
    st.caption(f"{len(alliance_files)} PDF détecté(s)" if alliance_files else "Aucun fichier pour l'instant")
    st.markdown('</div>', unsafe_allow_html=True)

with col3:
    st.markdown('<div class="canal-card" style="--accent:#E67E22">', unsafe_allow_html=True)
    st.markdown("### 💊 Biogaran Direct")
    st.markdown('<div class="sub">Factures (facture_*.pdf) — PDF individuels ou dossier zippé. Les avoirs (avoir_*.pdf) sont ignorés s\'ils sont déposés.</div>', unsafe_allow_html=True)
    biogaran_raw = st.file_uploader(
        "PDF Biogaran", type=["pdf", "zip"], accept_multiple_files=True,
        key=f"biogaran_{k}", label_visibility="collapsed",
    )
    biogaran_files_brutes = expand_uploads(biogaran_raw)
    biogaran_files = [(nom, c) for nom, c in biogaran_files_brutes if not nom.lower().startswith("avoir_")]
    n_avoir_ignores = len(biogaran_files_brutes) - len(biogaran_files)
    if biogaran_files:
        st.caption(f"{len(biogaran_files)} facture(s) détectée(s)" + (f" — {n_avoir_ignores} avoir(s) ignoré(s)" if n_avoir_ignores else ""))
    elif n_avoir_ignores:
        st.caption(f"{n_avoir_ignores} avoir(s) déposé(s), ignoré(s) (non pris en compte)")
    else:
        st.caption("Aucun fichier pour l'instant")
    st.markdown('</div>', unsafe_allow_html=True)

st.caption("💡 Astuce : pour déposer un dossier entier d'un coup, compresse-le en .zip sur ton Mac (clic droit → Compresser) puis glisse le .zip — les PDF à l'intérieur seront détectés automatiquement.")

st.write("")
bcol1, bcol2, _ = st.columns([1, 1, 3])
generer = bcol1.button("🚀 Générer le rapport", type="primary", use_container_width=True)
if bcol2.button("🔄 Nouvelle analyse", use_container_width=True):
    nouvelle_analyse()
    st.rerun()

total_fichiers = len(ocp_files) + len(alliance_files) + len(biogaran_files)


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


if generer:
    if total_fichiers == 0:
        st.warning("Dépose au moins un PDF dans une des 3 zones avant de générer le rapport.")
    else:
        with st.spinner("Analyse en cours… (peut prendre plusieurs minutes si beaucoup de factures Alliance)"):
            tmp_dir = Path(tempfile.mkdtemp(prefix="analyse_pharma_"))
            try:
                def sauver(files, sous_dossier):
                    dossier = tmp_dir / sous_dossier
                    dossier.mkdir(parents=True, exist_ok=True)
                    chemins = []
                    for nom, contenu in files or []:
                        chemin = dossier / nom
                        chemin.write_bytes(contenu)
                        chemins.append(str(chemin))
                    return chemins

                ocp_paths = sauver(ocp_files, "ocp")
                alliance_paths = sauver(alliance_files, "alliance")
                biogaran_paths = sauver(biogaran_files, "biogaran_direct")

                out_path = tmp_dir / "rapport_pharmacie.xlsx"
                data_dir = str(ROOT / "data")

                log_buffer = io.StringIO()
                with redirect_stdout(log_buffer), redirect_stderr(log_buffer):
                    build_report(
                        ocp_paths, data_dir, str(out_path),
                        biogaran_direct_pdfs=biogaran_paths or None,
                        alliance_pdfs=alliance_paths or None,
                    )

                log_text = log_buffer.getvalue()
                st.session_state.log_text = log_text
                st.session_state.kpis = parse_kpis(log_text)
                st.session_state.report_bytes = out_path.read_bytes()
                st.session_state.report_name = f"rapport_pharmacie_{datetime.now():%Y%m%d_%H%M}.xlsx"
                st.session_state.error_text = None
            except Exception:
                st.session_state.error_text = traceback.format_exc()
                st.session_state.report_bytes = None
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
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
