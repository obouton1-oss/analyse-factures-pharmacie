# -*- coding: utf-8 -*-
"""
Interface Streamlit pour l'analyse des factures pharmacie (canaux OCP,
Alliance Healthcare, Biogaran Direct + avoirs RDP).

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
    directement sur la zone de dépôt."""
    resultat = []
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
                        resultat.append((Path(entry).name, zf.read(entry)))
            except zipfile.BadZipFile:
                st.warning(f"⚠️ {nom} n'est pas une archive .zip valide, ignoré.")
        else:
            resultat.append((nom, f.getvalue()))
    return resultat


# ---------------------------------------------------------------------------
# Zones de dépôt (3 canaux)
# ---------------------------------------------------------------------------
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="canal-card" style="--accent:#2E86C1">', unsafe_allow_html=True)
    st.markdown("### 📋 OCP")
    st.markdown('<div class="sub">Récapitulatifs mensuels BO-OFFREM — PDF individuels ou dossier zippé</div>', unsafe_allow_html=True)
    with st.container(height=180, border=False):
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
    with st.container(height=180, border=False):
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
    st.markdown('<div class="sub">Factures (facture_*.pdf) + avoirs RDP (avoir_*.pdf) — triés automatiquement, PDF individuels ou dossier zippé</div>', unsafe_allow_html=True)
    with st.container(height=180, border=False):
        biogaran_raw = st.file_uploader(
            "PDF Biogaran", type=["pdf", "zip"], accept_multiple_files=True,
            key=f"biogaran_{k}", label_visibility="collapsed",
        )
    biogaran_files = expand_uploads(biogaran_raw)
    n_fact = sum(1 for nom, _ in biogaran_files if nom.lower().startswith("facture_"))
    n_avoir = sum(1 for nom, _ in biogaran_files if nom.lower().startswith("avoir_"))
    n_autre = len(biogaran_files) - n_fact - n_avoir
    if biogaran_files:
        st.caption(f"{n_fact} facture(s), {n_avoir} avoir(s)" + (f", {n_autre} fichier(s) non reconnu(s) ⚠️" if n_autre else ""))
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
    m = re.search(r"(\d+) bloc\(s\) RDP Biogaran", log_text)
    if m:
        kpis["rdp"] = m.group(1)
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

                facture_files = [(nom, c) for nom, c in biogaran_files if nom.lower().startswith("facture_")]
                avoir_files = [(nom, c) for nom, c in biogaran_files if nom.lower().startswith("avoir_")]
                biogaran_paths = sauver(facture_files, "biogaran_direct")
                avoir_paths = sauver(avoir_files, "biogaran_avoirs")

                out_path = tmp_dir / "rapport_pharmacie.xlsx"
                data_dir = str(ROOT / "data")

                log_buffer = io.StringIO()
                with redirect_stdout(log_buffer), redirect_stderr(log_buffer):
                    build_report(
                        ocp_paths, data_dir, str(out_path),
                        biogaran_direct_pdfs=biogaran_paths or None,
                        alliance_pdfs=alliance_paths or None,
                        avoirs_rdp_pdfs=avoir_paths or None,
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
    kcols = st.columns(5)
    kpi_defs = [
        ("lignes", "Lignes produit"),
        ("ca_total", "CA PPHT total"),
        ("fuite", "Lignes de fuite"),
        ("rdp", "Blocs RDP contrôlés"),
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
