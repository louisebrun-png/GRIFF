# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle, historique M-1 et suivi des chantiers
# AUTEUR      : Louise BRUN - Koesio Groupe
# ==============================================================================

import streamlit as st
import pandas as pd
import re
import os

try:
    import pypdf
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

# ------------------------------------------------------------------------------
# 1. CONFIGURATION DE LA PAGE STREAMLIT
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="GRIFF - Koesio",
    page_icon="🦅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------------------------------------------------------------
# 2. GESTION DE LA REMISE À ZÉRO (RAZ)
# ------------------------------------------------------------------------------
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

def reset_application():
    st.cache_data.clear()
    st.session_state["uploader_key"] += 1
    st.rerun()

# ------------------------------------------------------------------------------
# 3. BARRE LATÉRALE (SIDEBAR) : LOGO, TITRE CLAIR & RAZ
# ------------------------------------------------------------------------------
logo_path = "LOGO-SD-Long-Image+Nom.png"
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, use_container_width=True)

st.sidebar.header("📂 Chargement des exports mensuels")
st.sidebar.markdown("Téléversez vos fichiers d'exports du mois pour générer l'analyse.")

if st.sidebar.button("🔄 Remise à zéro (RAZ)", use_container_width=True, type="secondary"):
    reset_application()

st.sidebar.markdown("---")

u_key = st.session_state["uploader_key"]

files_isiwork_input = st.sidebar.file_uploader(
    "Exports Isiwork / Zeendoc (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    key=f"isi_{u_key}",
    help="Exports Zeendoc de vos différentes filiales (multi-fichiers acceptés)."
)

files_artis_input = st.sidebar.file_uploader(
    "Exports Artis ERP (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    key=f"art_{u_key}",
    help="Exports Artis ERP de vos différentes filiales."
)

st.sidebar.markdown("---")
st.sidebar.header("📜 Historique M-1 (PDF Mail Passé)")
file_history_input = st.sidebar.file_uploader(
    "Rapport / Mail PDF M-1 (.pdf)",
    type=["pdf"],
    accept_multiple_files=False,
    key=f"hist_{u_key}",
    help="Téléversez le rapport PDF du mois précédent pour extraire l'historique M-1."
)

all_uploaded_files = (files_isiwork_input or []) + (files_artis_input or [])

# ------------------------------------------------------------------------------
# 4. EN-TÊTE & ACRONYME GRIFF
# ------------------------------------------------------------------------------
st.title("🦅 GRIFF — Tour de Contrôle Koesio")

st.info(
    "**GRIFF** : **G**estion & **R**approchement **I**ntelligent des **F**actures et **F**lux\n\n"
    "Application d'analyse mensuelle et de contrôle automatisé des rapprochements Bon de Livraison (BL) / Factures d'Achat "
    "entre les systèmes **Isiwork/Zeendoc** et **Artis (ERP)** pour les filiales Koesio."
)

st.markdown("---")

# ------------------------------------------------------------------------------
# 5. HISTORIQUE M-1 (JUILLET 2026) PAR DÉFAUT ET DEPUIS LE MAIL PDF
# ------------------------------------------------------------------------------
HISTORIQUE_DEFAULT_M1 = {
    "KOESIO CENTRE EST": 48.5,
    "KOESIO EST": 40.3,
    "KOESIO NORD OUEST": 34.1,
    "KOESIO NORD-OUEST": 34.1,
    "KOESIO ILE DE FRANCE": 29.6,
    "KOESIO OCCITANIE": 36.9,
    "EOLE": 34.2,
    "2IT SOLUTIONS": 8.8,
    "KOESIO PACA": 39.2,
    "KOESIO MEDITERRANEE": 25.4,
    "TOTAL GROUPE CONSOLIDÉ": 36.6
}

@st.cache_data
def parse_pdf_history_full(pdf_file):
    hist_data = HISTORIQUE_DEFAULT_M1.copy()
    chantiers = [
        "Alerte Référentiels Prix : SHARP, ACTUAL SYSTEMES, KATUN FRANCE.",
        "Correctif API TVA : Suivi de la livraison globale sur Artis v2."
    ]
    if not pdf_file or not HAS_PYPDF:
        return hist_data, chantiers
    try:
        reader = pypdf.PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() + "\n"
            
        pattern_ged = re.findall(r'KOESIO\s*([A-Za-zÀ-ÿ\s\-\diT]+?)\s*(\d[\d\s]*)\s+(\d+)\s+(\d+)\s+([\d,]+)\s*%\s*\(([\d\s]+)\)', full_text)
        for match in pattern_ged:
            name = match[0].strip().upper()
            auto_pct = float(match[4].replace(',', '.'))
            hist_data[f"KOESIO {name}"] = auto_pct
            
        return hist_data, chantiers
    except Exception:
        return hist_data, chantiers

dict_hist_m1, list_chantiers_m1 = parse_pdf_history_full(file_history_input)

# ------------------------------------------------------------------------------
# 6. CHARGEMENT ET DÉTECTION PAR SIGNATURE ZEENDOC
# ------------------------------------------------------------------------------
@st.cache_data(show_spinner="Analyse rapide des fichiers Excel...")
def load_all_data(uploaded_files):
    isi_list = []
    art_list = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.xlsx'):
                df_temp = pd.read_excel(file, header=None, nrows=10)
            else:
                df_temp = pd.read_csv(file, header=None, nrows=10)
                
            is_isiwork = False
            header_row = 0
            
            row_0_str = " ".join([str(v) for v in df_temp.iloc[0].values if pd.notna(v)]).strip() if len(df_temp) > 0 else ""
            entite_detectee = row_0_str if row_0_str else file.name
            
            for idx, row in df_temp.iterrows():
                row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
                if 'Statut de rapprochement' in row_str or 'Identifiant' in row_str:
                    is_isiwork = True
                    header_row = idx
                    break
            
            if is_isiwork:
                if file.name.endswith('.xlsx'):
                    df_isi = pd.read_excel(file, header=header_row)
                else:
                    df_isi = pd.read_csv(file, header=header_row)
                
                df_isi['Organisation_Zeendoc'] = entite_detectee
                isi_list.append(df_isi)
            else:
                if file.name.endswith('.xlsx'):
                    df_art = pd.read_excel(file)
                else:
                    df_art = pd.read_csv(file)
                
                if 'Organisation interne' in df_art.columns:
                    df_art['Organisation interne'] = df_art['Organisation interne'].replace('KOESIO AQUITAINE (EX GTO)', 'KOESIO AQUITAINE')
                    
                art_list.append(df_art)
                
        except Exception as e:
            st.error(f"Erreur lors de la lecture de {file.name}: {e}")
            
    df_isi_concat = pd.concat(isi_list, ignore_index=True) if isi_list else pd.DataFrame()
    df_art_concat = pd.concat(art_list, ignore_index=True) if art_list else pd.DataFrame()
    
    return df_isi_concat, df_art_concat

df_isiwork, df_artis = load_all_data(all_uploaded_files)

# ------------------------------------------------------------------------------
# 7. DÉTECTION DYNAMIQUE DE LA FILIALE EN VISION INDIVIDUELLE
# ------------------------------------------------------------------------------
def detect_filiale(df_art):
    if not df_art.empty and 'Organisation interne' in df_art.columns:
        filiales = df_art['Organisation interne'].dropna().unique().tolist()
        if filiales:
            return ", ".join(map(str, filiales))
    return "En attente de fichier..."

filiale_detectee = detect_filiale(df_artis)

# ------------------------------------------------------------------------------
# 8. DÉTERMINATION DE LA PÉRIODE GLOBALE (AVEC CONSERVATION D'AOÛT & SEPTEMBRE)
# ------------------------------------------------------------------------------
mois_fr = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

def get_all_periods(df_art, df_isi):
    p_art, p_isi = [], []
    if not df_art.empty and 'Date' in df_art.columns:
        p_art = pd.to_datetime(df_art['Date'], errors='coerce').dt.to_period('M').dropna().unique().tolist()
    if not df_isi.empty and 'Date facture' in df_isi.columns:
        p_isi = pd.to_datetime(df_isi['Date facture'], errors='coerce').dt.to_period('M').dropna().unique().tolist()
    return sorted(list(set(p_art + p_isi)))

all_periods_p = get_all_periods(df_artis, df_isiwork)

if all_periods_p:
    if len(all_periods_p) == 1:
        periode_globale = f"{mois_fr.get(all_periods_p[0].month, str(all_periods_p[0].month))} {all_periods_p[0].year}"
    else:
        m_start = f"{mois_fr.get(all_periods_p[0].month, str(all_periods_p[0].month))} {all_periods_p[0].year}"
        m_end = f"{mois_fr.get(all_periods_p[-1].month, str(all_periods_p[-1].month))} {all_periods_p[-1].year}"
        periode_globale = f"{m_start} et {m_end}"
else:
    periode_globale = "Mois non détecté"

df_artis_filtered = df_artis
df_isiwork_filtered = df_isiwork

# Ventilation par mois pour la Vision Filiale
@st.cache_data
def compute_monthly_breakdown(df_art_in, df_isi_in):
    df_art = df_art_in.copy()
    df_isi = df_isi_in.copy()
    
    if not df_art.empty and 'Date' in df_art.columns:
        df_art['Date_dt'] = pd.to_datetime(df_art['Date'], errors='coerce')
        df_art['Mois_Period'] = df_art['Date_dt'].dt.to_period('M')
    else:
        df_art['Mois_Period'] = None

    if not df_isi.empty and 'Date facture' in df_isi.columns:
        df_isi['Date_dt'] = pd.to_datetime(df_isi['Date facture'], errors='coerce')
        df_isi['Mois_Period'] = df_isi['Date_dt'].dt.to_period('M')
    else:
        df_isi['Mois_Period'] = None

    periods_isi = df_isi['Mois_Period'].dropna().unique().tolist() if 'Mois_Period' in df_isi else []
    all_periods = sorted([p for p in periods_isi if p is not None])

    rows_ged = []
    rows_erp = []

    for p in all_periods:
        label_mois = f"{mois_fr.get(p.month, str(p.month))} {p.year}"
        
        sub_art = df_art[df_art['Mois_Period'] == p] if 'Mois_Period' in df_art else pd.DataFrame()
        sub_isi = df_isi[df_isi['Mois_Period'] == p] if 'Mois_Period' in df_isi else pd.DataFrame()
        
        tot_art = len(sub_art)
        if tot_art > 0 and 'Statut' in sub_art.columns:
            art_batch = len(sub_art[sub_art['Statut'].notna() & (sub_art['Statut'] != '')])
            art_man = tot_art - art_batch
            pct_batch = (art_batch / tot_art) * 100
            pct_man = (art_man / tot_art) * 100
        else:
            art_batch, art_man, pct_batch, pct_man = 0, 0, 0.0, 0.0

        tot_isi_brut = len(sub_isi)
        if tot_isi_brut > 0 and 'Statut de rapprochement' in sub_isi.columns:
            auto_c = len(sub_isi[sub_isi['Statut de rapprochement'] == 'Rapproché automatiquement'])
            
            not_auto_mask = sub_isi['Statut de rapprochement'] != 'Rapproché automatiquement'
            
            notes_all = sub_isi['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in sub_isi.columns else pd.Series('', index=sub_isi.index)
            export_erp = sub_isi['Date export ERP'].fillna('').astype(str) if 'Date export ERP' in sub_isi.columns else pd.Series('', index=sub_isi.index)
            
            pattern_ged_pure = 'FAF|TRAIT|ARTIS|ARCHIV|DÉJÀ TRAITÉ'
            
            is_faf = notes_all.str.contains(pattern_ged_pure, case=False, regex=True) | export_erp.str.contains(pattern_ged_pure, case=False, regex=True) | sub_isi['Statut de rapprochement'].isna()
            
            ged_pure_mask = not_auto_mask & is_faf
            ged_pure_c = ged_pure_mask.sum()
            
            vol_epure = max(0, tot_isi_brut - ged_pure_c)
            
            sub_isi_epure = sub_isi[~ged_pure_mask].copy()
            
            df_man_epure = sub_isi_epure[sub_isi_epure['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)].copy()
            notes_man_epure = df_man_epure['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in df_man_epure.columns else pd.Series('', index=df_man_epure.index)
            
            faux_c = len(df_man_epure[notes_man_epure.str.contains('FRAIS|PORT|D3EE|DEEE', case=False, regex=True)])
            vrai_c = max(0, len(df_man_epure) - faux_c)
            
            attente_c = max(0, vol_epure - (auto_c + faux_c + vrai_c))
            
            denom = vol_epure if vol_epure > 0 else 1
            
            pct_a = (auto_c / denom) * 100
            pct_f = (faux_c / denom) * 100
            pct_v = (vrai_c / denom) * 100
            pct_att = (attente_c / denom) * 100
        else:
            auto_c, faux_c, vrai_c, ged_pure_c, attente_c, vol_epure = 0, 0, 0, 0, 0, 0
            pct_a, pct_f, pct_v, pct_att = 0.0, 0.0, 0.0, 0.0

        rows_ged.append({
            "Mois Comptable": label_mois,
            "📋 Vol. Brut GED": tot_isi_brut,
            "📁 Flux GED Pure (Archivage / FAF)": f"{ged_pure_c}",
            "📊 Vol. Épuré": vol_epure,
            "🤖 Rapprochement Auto": f"{pct_a:.1f}% ({auto_c})",
            "🚚 Faux Manuel (Frais/Port)": f"{pct_f:.1f}% ({faux_c})",
            "🛠️ Vrai Manuel (Écarts)": f"{pct_v:.1f}% ({vrai_c})",
            "⏳ En attente": f"{pct_att:.1f}% ({attente_c})"
        })

        rows_erp.append({
            "Mois Comptable": label_mois,
            "💻 Volume Intégré Artis": tot_art,
            "🤖 Via Import Automatique": f"{pct_batch:.1f}% ({art_batch})",
            "👤 Saisie Directe / Forçage Manuel ERP": f"{pct_man:.1f}% ({art_man})"
        })

    return pd.DataFrame(rows_ged), pd.DataFrame(rows_erp), all_periods

df_ged_monthly, df_erp_monthly, valid_periods = compute_monthly_breakdown(df_artis, df_isiwork)

# ------------------------------------------------------------------------------
# 9. EXTRACTION DYNAMIQUE TOPS FOURNISSEURS
# ------------------------------------------------------------------------------
txt_top_vrai_html = ""
txt_top_faux_html = ""

if not df_isiwork_filtered.empty and 'Statut de rapprochement' in df_isiwork_filtered.columns:
    sub_isi_f = df_isiwork_filtered.copy()
    notes_all = sub_isi_f['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in sub_isi_f.columns else pd.Series('', index=sub_isi_f.index)
    export_erp = sub_isi_f['Date export ERP'].fillna('').astype(str) if 'Date export ERP' in sub_isi_f.columns else pd.Series('', index=sub_isi_f.index)
    pattern_ged_pure = 'FAF|TRAIT|ARTIS|ARCHIV|DÉJÀ TRAITÉ'
    
    not_auto_mask = sub_isi_f['Statut de rapprochement'] != 'Rapproché automatiquement'
    is_faf = notes_all.str.contains(pattern_ged_pure, case=False, regex=True) | export_erp.str.contains(pattern_ged_pure, case=False, regex=True) | sub_isi_f['Statut de rapprochement'].isna()
    ged_pure_mask = not_auto_mask & is_faf
    
    sub_isi_epure = sub_isi_f[~ged_pure_mask].copy()
    df_man_epure = sub_isi_epure[sub_isi_epure['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)].copy()
    notes_man_epure = df_man_epure['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in df_man_epure.columns else pd.Series('', index=df_man_epure.index)
    mask_faux = notes_man_epure.str.contains('FRAIS|PORT|D3EE|DEEE', case=False, regex=True)
    
    df_vrai_fourn = df_man_epure[~mask_faux]
    df_faux_fourn = df_man_epure[mask_faux]
    
    if not df_vrai_fourn.empty and 'Fournisseur' in df_vrai_fourn.columns:
        s_vrai = df_vrai_fourn['Fournisseur'].value_counts().head(5)
        txt_top_vrai_html = "".join([f"<li><b>{fourn}</b> : {cnt} factures</li>" for fourn, cnt in s_vrai.items()])
    
    if not df_faux_fourn.empty and 'Fournisseur' in df_faux_fourn.columns:
        s_faux = df_faux_fourn['Fournisseur'].value_counts().head(5)
        txt_top_faux_html = "".join([f"<li><b>{fourn}</b> : {cnt} factures</li>" for fourn, cnt in s_faux.items()])

if not txt_top_vrai_html:
    txt_top_vrai_html = "<li><b>TD SYNNEX</b></li><li><b>ACTUAL SYSTEMES</b></li><li><b>CANON FRANCE</b></li>"

if not txt_top_faux_html:
    txt_top_faux_html = "<li>Aucun Faux Manuel détecté</li>"

# ------------------------------------------------------------------------------
# 10. ONGLETS DE NAVIGATION
# ------------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🏢 Vision Filiale (Rapport & Mail)", "🌍 Consolidation Groupe"])

# ==============================================================================
# ONGLET 1 : VISION FILIALE
# ==============================================================================
with tab1:
    st.subheader(f"📊 Bilan de Production — {filiale_detectee}")
    
    if not df_isiwork.empty or not df_artis.empty:
        st.success(f"📅 Périodes analysées (mois clôturés) : **{periode_globale}**")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Factures Artis", f"{len(df_artis_filtered):,}".replace(',', ' '))
        with c2:
            st.metric("Total Factures Isiwork", f"{len(df_isiwork_filtered):,}".replace(',', ' '))
        with c3:
            st.metric("Mois Traités", len(df_ged_monthly))
        with c4:
            st.metric("Filiale Détectée", filiale_detectee)
            
        st.markdown("---")
        
        st.markdown("### 📈 1. VISION GED / ISIWORK — Typologie des flux dans la GED")
        st.dataframe(df_ged_monthly, hide_index=True, use_container_width=True)
        
        st.markdown("### 💻 2. VISION ERP ARTIS — Intégration et traitement réel dans l'ERP")
        st.dataframe(df_erp_monthly, hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### ✉️ Aperçu du Mail (Prêt à copier pour Outlook / Teams)")
        
        table_ged_html = df_ged_monthly.to_html(index=False, border=1, classes='content-table')
        table_erp_html = df_erp_monthly.to_html(index=False, border=1, classes='content-table')
        
        mail_html = f"""
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.5; padding: 10px; background-color: #ffffff; border-radius: 5px;">
            <p>Bonjour à tous,</p>
            
            <p>Vous trouverez ci-dessous l'analyse détaillée des performances du flux de rapprochement automatique des factures fournisseurs pour <b>{filiale_detectee}</b>, consolidée sur la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément aux directives du Groupe, nous appliquons la méthodologie d'<b>isolation des flux « GED Pure / FAF Artis »</b> afin de refléter l'efficacité machine réelle d'Isiwork.</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — Typologie des flux dans la GED</h3>
            {table_ged_html}
            
            <h3 style="color: #0056b3; margin-top: 20px;">💻 2. VISION ERP ARTIS — Intégration et traitement réel dans l'ERP</h3>
            {table_erp_html}
            
            <h3 style="color: #0056b3; margin-top: 20px;">🔍 3. FOCUS ACTIONS — Analyse des "Vrais Manuels" & Frictions</h3>
            <p><b>• Top 5 Fournisseurs bloquants :</b></p>
            <ul>
                {txt_top_vrai_html}
            </ul>
            
            <p>Cordialement,</p>
            <p><b>Louise BRUN</b><br>
            Chef de projet pour KOESIO GROUPE & KAM<br>
            Potentiale</p>
        </div>
        """
        st.components.v1.html(mail_html, height=800, scrolling=True)
        
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork/Zeendoc** et **Artis** dans le panneau latéral.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION GROUPE (AVEC HISTORIQUE M-1 ET MODÈLE MAIL HISTORIQUE)
# ==============================================================================
with tab2:
    st.subheader(f"🌍 Consolidation Inter-Filiales — Groupe Koesio ({periode_globale})")
    
    if not df_artis_filtered.empty and 'Organisation interne' in df_artis_filtered.columns:
        filiales_groupe = df_artis_filtered['Organisation interne'].dropna().unique().tolist()
        
        st.success(f"🏢 **{len(filiales_groupe)} filiales actives chargées pour {periode_globale}** : {', '.join(map(str, filiales_groupe))}")
        
        rows_ged_raw = []
        tot_vol_brut, tot_ged_pure, tot_vol_epure, tot_auto, tot_faux, tot_vrai, tot_att = 0, 0, 0, 0, 0, 0, 0
        
        CORRESPONDANCE_ZEENDOC = {
            "KOESIO NORD OUEST": ["KOESIO NORD-OUEST", "KOESIO NORD OUEST", "KNO"],
            "KOESIO NORD-OUEST": ["KOESIO NORD-OUEST", "KOESIO NORD OUEST", "KNO"],
            "KOESIO OUEST": ["KOESIO OUEST"],
            "KOESIO ILE DE FRANCE": ["KOESIO IDF", "KOESIO|Rappro Artis IDF", "IDF", "ILE DE FRANCE"],
            "KOESIO AQUITAINE": ["JEAPI Aquitaine", "AQUITAINE"],
            "KOESIO OCCITANIE": ["BCMP", "OCCITANIE", "OCC"],
            "KOESIO CENTRE EST": ["Esus", "CENTRE EST", "CEE"],
            "2iT solutions": ["RHONE ALPES REPRO INFORMATIQUE", "2IT"],
            "KOESIO MEDITERRANEE": ["KOESIO MED", "MEDITERRANEE"],
            "KOESIO PACA": ["KOESIO PACA", "PACA"],
            "KOESIO EST": ["KOESIO EST"],
            "EOLE": ["EOLE"]
        }
        
        for f in filiales_groupe:
            sub_art_f = df_artis_filtered[df_artis_filtered['Organisation interne'] == f]
            
            if 'Organisation_Zeendoc' in df_isiwork_filtered.columns:
                termes_zeendoc = CORRESPONDANCE_ZEENDOC.get(f, [f])
                mask_isi = df_isiwork_filtered['Organisation_Zeendoc'].astype(str).str.upper().apply(
                    lambda x: any(term.upper() in x for term in termes_zeendoc)
                )
                sub_isi_f = df_isiwork_filtered[mask_isi].copy()
            else:
                sub_isi_f = pd.DataFrame()
            
            vol_brut = len(sub_isi_f)
            tot_art_f = len(sub_art_f)
            
            if vol_brut > 0 and 'Statut de rapprochement' in sub_isi_f.columns:
                auto_c = len(sub_isi_f[sub_isi_f['Statut de rapprochement'] == 'Rapproché automatiquement'])
                not_auto_mask = sub_isi_f['Statut de rapprochement'] != 'Rapproché automatiquement'
                
                notes_f = sub_isi_f['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in sub_isi_f.columns else pd.Series('', index=sub_isi_f.index)
                export_erp_f = sub_isi_f['Date export ERP'].fillna('').astype(str) if 'Date export ERP' in sub_isi_f.columns else pd.Series('', index=sub_isi_f.index)
                pattern_ged_pure = 'FAF|TRAIT|ARTIS|ARCHIV|DÉJÀ TRAITÉ'
                is_faf = notes_f.str.contains(pattern_ged_pure, case=False, regex=True) | export_erp_f.str.contains(pattern_ged_pure, case=False, regex=True) | sub_isi_f['Statut de rapprochement'].isna()
                
                ged_pure_mask = not_auto_mask & is_faf
                ged_pure = ged_pure_mask.sum()
                vol_epure = max(0, vol_brut - ged_pure)
                
                sub_isi_f_epure = sub_isi_f[~ged_pure_mask].copy()
                df_man_f_epure = sub_isi_f_epure[sub_isi_f_epure['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)].copy() if 'Statut de rapprochement' in sub_isi_f_epure.columns else pd.DataFrame()
                notes_f_epure = df_man_f_epure['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in df_man_f_epure.columns else pd.Series('', index=df_man_f_epure.index)
                
                faux_c = len(df_man_f_epure[notes_f_epure.str.contains('FRAIS|PORT|D3EE|DEEE', case=False, regex=True)]) if not df_man_f_epure.empty else 0
                vrai_c = max(0, len(df_man_f_epure) - faux_c)
                att_c = max(0, vol_epure - (auto_c + faux_c + vrai_c))
            else:
                auto_c, faux_c, vrai_c, ged_pure, att_c, vol_epure = 0, 0, 0, 0, 0, 0

            denom = vol_epure if vol_epure > 0 else 1
            pct_a = (auto_c / denom * 100) if vol_epure > 0 else 0
            pct_f = (faux_c / denom * 100) if vol_epure > 0 else 0
            pct_v = (vrai_c / denom * 100) if vol_epure > 0 else 0
            pct_att = (att_c / denom * 100) if vol_epure > 0 else 0
            pct_pot = pct_a + pct_f
            
            tot_vol_brut += vol_brut
            tot_ged_pure += ged_pure
            tot_vol_epure += vol_epure
            tot_auto += auto_c
            tot_faux += faux_c
            tot_vrai += vrai_c
            tot_att += att_c
            
            rows_ged_raw.append({
                "Filiale Groupe": f,
                "Vol. Brut GED": vol_brut,
                "Flux GED Pure": ged_pure,
                "Vol. Épuré": vol_epure,
                "pct_a": pct_a,
                "auto_c": auto_c,
                "pct_f": pct_f,
                "faux_c": faux_c,
                "pct_v": pct_v,
                "vrai_c": vrai_c,
                "pct_att": pct_att,
                "att_c": att_c,
                "pct_pot": pct_pot
            })

        # TRI TABLEAU 1 : % AUTO DESC, PUIS VOL. ÉPURÉ DESC
        df_ged_temp = pd.DataFrame(rows_ged_raw)
        df_ged_temp = df_ged_temp.sort_values(by=["pct_a", "Vol. Épuré"], ascending=[False, False]).reset_index(drop=True)
        
        rows_ged_groupe = []
        for idx, row in df_ged_temp.iterrows():
            f_name = row["Filiale Groupe"]
            pct_a_val = row["pct_a"]
            auto_c_val = row["auto_c"]
            
            str_eval_auto = f"{pct_a_val:.1f}% ({auto_c_val})"
            m1_val = None
            for key_h, val_h in dict_hist_m1.items():
                if f_name.upper() in key_h.upper() or key_h.upper() in f_name.upper():
                    m1_val = val_h
                    break
                    
            if m1_val is not None:
                delta_a = pct_a_val - m1_val
                if delta_a > 0.05:
                    symb = "↗️"
                elif delta_a < -0.05:
                    symb = "↘️"
                else:
                    symb = "➡️"
                str_eval_auto += f" [{symb} {delta_a:+.1f}% vs M-1]"

            rows_ged_groupe.append({
                "Rang": idx + 1,
                "Filiale Groupe": f_name,
                "Vol. Brut GED": row["Vol. Brut GED"],
                "Flux GED Pure": f"{row['Flux GED Pure']}",
                "Vol. Épuré": row["Vol. Épuré"],
                "🤖 Rapprochement Auto (Évolution M vs M-1)": str_eval_auto,
                "🚚 Faux Manuel (Frais)": f"{row['pct_f']:.1f}% ({row['faux_c']})",
                "🛠️ Vrai Manuel (Litiges)": f"{row['pct_v']:.1f}% ({row['vrai_c']})",
                "⏳ En attente": f"{row['pct_att']:.1f}% ({row['att_c']})",
                "🎯 Potentiel Cible": f"{row['pct_pot']:.1f}%"
            })
            
        pct_tot_a = (tot_auto / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_f = (tot_faux / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_v = (tot_vrai / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_att = (tot_att / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        
        m1_tot_val = dict_hist_m1.get("TOTAL GROUPE CONSOLIDÉ", 36.6)
        delta_tot = pct_tot_a - m1_tot_val
        symb_tot = "↗️" if delta_tot > 0.05 else "↘️" if delta_tot < -0.05 else "➡️"
        
        rows_ged_groupe.append({
            "Rang": "—",
            "Filiale Groupe": "TOTAL GROUPE CONSOLIDÉ",
            "Vol. Brut GED": tot_vol_brut,
            "Flux GED Pure": f"{tot_ged_pure}",
            "Vol. Épuré": tot_vol_epure,
            "🤖 Rapprochement Auto (Évolution M vs M-1)": f"{pct_tot_a:.1f}% ({tot_auto}) [{symb_tot} {delta_tot:+.1f}% vs M-1]",
            "🚚 Faux Manuel (Frais)": f"{pct_tot_f:.1f}% ({tot_faux})",
            "🛠️ Vrai Manuel (Litiges)": f"{pct_tot_v:.1f}% ({tot_vrai})",
            "⏳ En attente": f"{pct_tot_att:.1f}% ({tot_att})",
            "🎯 Potentiel Cible": f"{(pct_tot_a + pct_tot_f):.1f}%"
        })
        
        df_ged_groupe = pd.DataFrame(rows_ged_groupe)
        
        # TABLEAU 2 (ARTIS) : STATUT NON VIDE = IMPORT AUTO / STATUT VIDE = SAISIE DIRECTE
        rows_erp_raw = []
        tot_art_all, tot_batch_all, tot_man_all = 0, 0, 0
        
        for f in filiales_groupe:
            sub_art_f = df_artis_filtered[df_artis_filtered['Organisation interne'] == f]
            tot_art_f = len(sub_art_f)
            
            if tot_art_f > 0 and 'Statut' in sub_art_f.columns:
                batch_c = len(sub_art_f[sub_art_f['Statut'].notna() & (sub_art_f['Statut'] != '')])
                man_c = tot_art_f - batch_c
                pct_b = (batch_c / tot_art_f) * 100
                pct_m = (man_c / tot_art_f) * 100
            else:
                batch_c, man_c, pct_b, pct_m = 0, 0, 0.0, 0.0
                
            tot_art_all += tot_art_f
            tot_batch_all += batch_c
            tot_man_all += man_c
            
            rows_erp_raw.append({
                "Filiale Groupe": f,
                "Total Pièces Validées Artis": tot_art_f,
                "batch_c": batch_c,
                "pct_b": pct_b,
                "man_c": man_c,
                "pct_m": pct_m
            })
            
        df_erp_temp = pd.DataFrame(rows_erp_raw)
        df_erp_temp = df_erp_temp.sort_values(by=["pct_b", "Total Pièces Validées Artis"], ascending=[False, False]).reset_index(drop=True)
        
        rows_erp_groupe = []
        for idx, row in df_erp_temp.iterrows():
            rows_erp_groupe.append({
                "Rang": idx + 1,
                "Filiale Groupe": row["Filiale Groupe"],
                "Total Pièces Validées Artis": row["Total Pièces Validées Artis"],
                "🤖 Via Import Automatique": f"{row['pct_b']:.1f}% ({row['batch_c']})",
                "👤 Saisie Directe / Forçage Manuel ERP": f"{row['pct_m']:.1f}% ({row['man_c']})"
            })
            
        pct_tot_batch = (tot_batch_all / tot_art_all * 100) if tot_art_all > 0 else 0
        pct_tot_man_erp = (tot_man_all / tot_art_all * 100) if tot_art_all > 0 else 0
        
        rows_erp_groupe.append({
            "Rang": "—",
            "Filiale Groupe": "TOTAL ERP CONSOLIDÉ",
            "Total Pièces Validées Artis": tot_art_all,
            "🤖 Via Import Automatique": f"{pct_tot_batch:.1f}% ({tot_batch_all})",
            "👤 Saisie Directe / Forçage Manuel ERP": f"{pct_tot_man_erp:.1f}% ({tot_man_all})"
        })
        
        df_erp_groupe = pd.DataFrame(rows_erp_groupe)
        
        g1, g2, g3, g4 = st.columns(4)
        with g1:
            st.metric("Total Factures Artis Groupe", f"{tot_art_all:,}".replace(',', ' '))
        with g2:
            st.metric("Total Factures Isiwork Groupe", f"{tot_vol_brut:,}".replace(',', ' '))
        with g3:
            st.metric("Taux Injection Batch ERP Groupe", f"{pct_tot_batch:.1f}%")
        with g4:
            st.metric("Rapprochement Auto GED Groupe", f"{pct_tot_a:.1f}%")
            
        st.markdown("---")
        
        st.markdown(f"### 📈 1. VISION GED / ISIWORK — Rapprochement National Consolidé ({periode_globale})")
        st.dataframe(df_ged_groupe, hide_index=True, use_container_width=True)
        
        st.markdown(f"### 💻 2. VISION ERP ARTIS — Journal d'Intégration National Consolidé ({periode_globale})")
        st.dataframe(df_erp_groupe, hide_index=True, use_container_width=True)
        
        st.caption("ℹ️ **Note sur l'intégration ERP** : Certaines factures d'achat (abonnements, sous-traitance) peuvent se retrouver en saisie directe dans ARTIS, mais l'essentiel du flux doit transiter par Isiwork. Si ce n'est pas le cas, une analyse des pièces saisies en direct s'impose (ex. cas connu chez EOLE : TVA à 0%).")
        
        st.markdown("---")
        st.markdown("### ✉️ Modèle de Mail National Consolidé avec Historique (Prêt à copier)")
        
        table_ged_groupe_html = df_ged_groupe.to_html(index=False, border=1, classes='content-table')
        table_erp_groupe_html = df_erp_groupe.to_html(index=False, border=1, classes='content-table')
        
        txt_chantiers_html = "".join([f"<li>{c}</li>" for c in list_chantiers_m1])
        
        mail_groupe_html = f"""
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.5; padding: 10px; background-color: #ffffff; border-radius: 5px;">
            <p>Bonjour à tous,</p>
            
            <p>Dans le cadre du suivi technique et opérationnel de notre projet de transition de la saisie manuelle vers le pilotage financier automatique, vous trouverez ci-dessous le rapport national récapitulatif consolidé pour l'ensemble de nos filiales actives du Groupe Koesio pour la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément à nos spécifications de pilotage Groupe validées en comité, ce rapport isole de manière stricte les flux qualifiés de <b>« GED Pure / FAF Artis »</b> (factures d'archivage directement traitées manuellement de bout en bout dans l'ERP) afin de ne pas fausser l'efficacité machine réelle de rapprochement de notre plateforme Isiwork.</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — Rapprochement National & Évolutions M-1</h3>
            {table_ged_groupe_html}
            
            <h3 style="color: #0056b3; margin-top: 20px;">💻 2. VISION ERP ARTIS — Journal d'Intégration National</h3>
            {table_erp_groupe_html}
            <p style="font-size: 12px; color: #555; margin-top: 5px;">
                <i>* Note : Certaines factures d'achat (abonnements, sous-traitance) peuvent se retrouver en saisie directe dans ARTIS, mais le flux doit être majoritairement issu d'Isiwork. Si ce n'est pas le cas, une analyse des pièces saisies en direct s'impose (ex. cas connu chez EOLE : TVA à 0%).</i>
            </p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">🔍 3. FOCUS ACTIONS — Enseignements Clés du Groupe</h3>
            <p><b>• Alerte Référentiels Prix :</b> L'analyse des blocages « Vrais Manuels » montre que la charge de traitement réside principalement sur un problème de mise à jour tarifaire en amont sur nos partenaires stratégiques (TD SYNNEX, ACTUAL SYSTEMES, CANON FRANCE, SHARP).</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">⚙️ 4. SUIVI DES CHANTIERS TECHNIQUES & PERSPECTIVES (Rappel M-1 & Avancées)</h3>
            <ul>
                {txt_chantiers_html}
            </ul>
            
            <p>Cordialement,</p>
            <p><b>Louise BRUN</b><br>
            Chef de projet pour KOESIO GROUPE & KAM<br>
            Potentiale</p>
        </div>
        """
        
        st.components.v1.html(mail_groupe_html, height=850, scrolling=True)
        
    else:
        st.warning("👈 Téléversez les fichiers de vos différentes filiales dans la barre latérale pour afficher la consolidation groupe.")
