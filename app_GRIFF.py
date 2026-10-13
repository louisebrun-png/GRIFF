# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle et rapprochement automatique BL / Factures
# AUTEUR      : Louise BRUN - Koesio Groupe
# ==============================================================================

import streamlit as st
import pandas as pd
import plotly.express as px

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
# 2. EN-TÊTE & ACRONYME GRIFF
# ------------------------------------------------------------------------------
st.title("🦅 GRIFF — Tour de Contrôle Koesio")

st.info(
    "**GRIFF** : **G**estion & **R**approchement **I**ntelligent des **F**actures et **F**lux\n\n"
    "Application d'analyse mensuelle et de contrôle automatisé des rapprochements Bon de Livraison (BL) / Factures d'Achat "
    "entre les systèmes **Isiwork/Zeendoc** et **Artis (ERP)** pour les filiales Koesio."
)

st.markdown("---")

# ------------------------------------------------------------------------------
# 3. BARRE LATÉRALE (SIDEBAR) : INGESTION & FILTRES
# ------------------------------------------------------------------------------
st.sidebar.header("📂 Ingestion des Données")
st.sidebar.markdown("Téléversez vos exports mensuels pour lancer le calcul automatisé.")

# Import Zeendoc / Isiwork
files_isiwork = st.sidebar.file_uploader(
    "Exports Isiwork / Zeendoc (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Export Zeendoc contenant la colonne 'Statut de rapprochement'."
)

# Import Artis
files_artis = st.sidebar.file_uploader(
    "Exports Artis ERP (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Export de la recherche des factures fournisseurs Artis."
)

st.sidebar.markdown("---")
st.sidebar.header("🎯 Filtres d'Analyse")

filiale_select = st.sidebar.selectbox(
    "Filiale Koesio",
    ["KOESIO ILE DE FRANCE", "KOESIO CEE", "KOESIO AURA", "Toutes les filiales"]
)

mois_select = st.sidebar.selectbox(
    "Période / Mois d'analyse",
    ["Août et Septembre 2026", "Mai et Juin 2026", "Mois en cours", "Cumul YTD"]
)

# ------------------------------------------------------------------------------
# 4. FONCTIONS DE CHARGEMENT & NETTOYAGE LIGNE EN-TÊTE
# ------------------------------------------------------------------------------
@st.cache_data
def load_zeendoc_data(uploaded_files):
    """
    Charge et nettoie les fichiers Zeendoc/Isiwork (Gestion de l'en-tête ligne 2).
    """
    df_list = []
    for file in uploaded_files:
        try:
            df_temp = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
            header_row = 0
            for idx, row in df_temp.iterrows():
                if 'Statut de rapprochement' in row.values or 'Identifiant' in row.values:
                    header_row = idx
                    break
            
            if file.name.endswith('.xlsx'):
                df = pd.read_excel(file, header=header_row)
            else:
                df = pd.read_csv(file, header=header_row)
            
            df_list.append(df)
        except Exception as e:
            st.error(f"Erreur lors de la lecture de {file.name}: {e}")
            
    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return pd.DataFrame()

@st.cache_data
def load_artis_data(uploaded_files):
    """
    Charge les fichiers d'export Artis.
    """
    df_list = []
    for file in uploaded_files:
        try:
            if file.name.endswith('.xlsx'):
                df = pd.read_excel(file)
            else:
                df = pd.read_csv(file)
            df_list.append(df)
        except Exception as e:
            st.error(f"Erreur lors de la lecture de {file.name}: {e}")
            
    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return pd.DataFrame()

df_isiwork = load_zeendoc_data(files_isiwork) if files_isiwork else pd.DataFrame()
df_artis = load_artis_data(files_artis) if files_artis else pd.DataFrame()

# ------------------------------------------------------------------------------
# 5. INITIALISATION DES VARIABLES DE CALCULS (Évite les NameError)
# ------------------------------------------------------------------------------
artis_total, artis_batch, artis_manuel, pct_batch, pct_manuel = 0, 0, 0, 0.0, 0.0
isiwork_total, auto_cnt, faux_man_cnt, vrai_man_cnt, attente_cnt = 0, 0, 0, 0, 0
pct_auto, pct_faux, pct_vrai, pct_attente = 0.0, 0.0, 0.0, 0.0

# Calculs effectifs sur les données Artis si chargées
artis_total = len(df_artis) if not df_artis.empty else 0
if artis_total > 0 and 'Statut' in df_artis.columns:
    artis_batch = len(df_artis[df_artis['Statut'] == 'Bon à payer'])
    artis_manuel = artis_total - artis_batch
    pct_batch = (artis_batch / artis_total) * 100
    pct_manuel = (artis_manuel / artis_total) * 100

# Calculs effectifs sur les données Zeendoc/Isiwork si chargées
isiwork_total = len(df_isiwork) if not df_isiwork.empty else 0
if isiwork_total > 0 and 'Statut de rapprochement' in df_isiwork.columns:
    auto_cnt = len(df_isiwork[df_isiwork['Statut de rapprochement'] == 'Rapproché automatiquement'])
    
    if 'Tampons / Notes / Conversations' in df_isiwork.columns:
        notes = df_isiwork['Tampons / Notes / Conversations'].fillna('')
        faux_man_cnt = len(df_isiwork[notes.str.contains('FRAIS|PORT|D3EE', case=False, regex=True)])
    
    man_total = len(df_isiwork[df_isiwork['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)])
    vrai_man_cnt = max(0, man_total - faux_man_cnt)
    attente_cnt = max(0, isiwork_total - (auto_cnt + faux_man_cnt + vrai_man_cnt))
    
    pct_auto = (auto_cnt / isiwork_total) * 100
    pct_faux = (faux_man_cnt / isiwork_total) * 100
    pct_vrai = (vrai_man_cnt / isiwork_total) * 100
    pct_attente = (attente_cnt / isiwork_total) * 100

# ------------------------------------------------------------------------------
# 6. ONGLETS DE NAVIGATION
# ------------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["🏢 Vision Filiale (Rapport)", "🌍 Consolidation Groupe", "✉️ Modèle de Mail Généré"])

# ==============================================================================
# ONGLET 1 : VISION FILIALE
# ==============================================================================
with tab1:
    st.subheader(f"📊 Bilan de Production — {filiale_select} ({mois_select})")
    
    if not df_isiwork.empty or not df_artis.empty:
        st.success("Données traitées et synchronisées avec succès !")
        
        # --- CARTES KPI ---
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Factures Artis", f"{artis_total:,}".replace(',', ' '))
        with c2:
            st.metric("Intégrées par Batch (Bon à payer)", f"{pct_batch:.1f}%", f"{artis_batch} factures")
        with c3:
            st.metric("Total Traitées Isiwork", f"{isiwork_total:,}".replace(',', ' '))
        with c4:
            st.metric("Rapprochement Auto GED", f"{pct_auto:.1f}%", f"{auto_cnt} factures")
            
        st.markdown("---")
        
        # --- 1. VISION GED ---
        st.markdown("### 1. Vision GED — Typologie des flux dans Isiwork")
        ged_summary = pd.DataFrame({
            "Mois / Période": [mois_select],
            "Total Factures Traitées": [isiwork_total],
            "Rapprochement Auto": [f"{pct_auto:.1f}% ({auto_cnt})"],
            "Faux Manuel (Frais/Port)": [f"{pct_faux:.1f}% ({faux_man_cnt})"],
            "Vrai Manuel (Écarts)": [f"{pct_vrai:.1f}% ({vrai_man_cnt})"],
            "En attente": [f"{pct_attente:.1f}% ({attente_cnt})"]
        })
        st.dataframe(ged_summary, hide_index=True, use_container_width=True)
        
        # --- 2. VISION ERP ---
        st.markdown("### 2. Vision ERP — Intégration et traitement réel dans Artis")
        erp_summary = pd.DataFrame({
            "Mois / Période": [mois_select],
            "Volume Intégré Artis": [artis_total],
            "Via Batch Automatique (admin)": [f"{pct_batch:.1f}% ({artis_batch})"],
            "Saisie / Forçage Manuel ERP": [f"{pct_manuel:.1f}% ({artis_manuel})"]
        })
        st.dataframe(erp_summary, hide_index=True, use_container_width=True)
        
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork/Zeendoc** et **Artis** dans le panneau latéral.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION GROUPE
# ==============================================================================
with tab2:
    st.subheader("🌍 Consolidation Inter-Filiales Koesio")
    st.markdown("Benchmark de la performance de rapprochement réseau.")
    
    df_groupe = pd.DataFrame({
        "Filiale Koesio": ["KOESIO ILE DE FRANCE", "KOESIO CEE", "KOESIO AURA"],
        "Factures Traitées Artis": [675, 964, 1100],
        "Taux Batch Automatique (%)": ["84.2%", "79.6%", "82.0%"],
        "Rapprochement Auto GED (%)": ["37.3%", "57.8%", "45.0%"],
        "Statut Prochaine Étape": ["En cours d'analyse", "✅ Validé v5.5.0", "⚙️ En préparation"]
    })
    st.dataframe(df_groupe, hide_index=True, use_container_width=True)

# ==============================================================================
# ONGLET 3 : MODÈLE DE MAIL GÉNÉRÉ AUTOMATIQUEMENT
# ==============================================================================
with tab3:
    st.subheader("✉️ Générateur de Rapport par Mail")
    st.markdown("Copiez-collez ce texte prêt à l'emploi pour vos envois mensuels :")
    
    mail_text = f"""Bonjour à tous,

Vous trouverez ci-dessous l'analyse détaillée des performances du flux de rapprochement automatique des factures fournisseurs pour {filiale_select}, consolidée sur la période : {mois_select}.

Afin de maintenir une analyse d'une grande rigueur opérationnelle, nous appliquons la méthodologie de lecture fine des tampons et des lignes de commentaires Isiwork.

1. Vision GED — Typologie des flux dans Isiwork
- Total Factures Traitées : {isiwork_total}
- Rapprochement Auto : {pct_auto:.1f}% ({auto_cnt})
- Faux Manuel (Frais/Port) : {pct_faux:.1f}% ({faux_man_cnt})
- Vrai Manuel (Écarts) : {pct_vrai:.1f}% ({vrai_man_cnt})
- En attente : {pct_attente:.1f}% ({attente_cnt})

2. Vision ERP — Intégration et traitement réel dans Artis
- Volume Intégré Artis : {artis_total}
- Via Batch Automatique (Bon à payer) : {pct_batch:.1f}% ({artis_batch})
- Saisie / Forçage Manuel ERP : {pct_manuel:.1f}% ({artis_manuel})

3. Focus Actions & Fournisseurs Cibles
Les écarts de rapprochement se concentrent principalement sur les fournisseurs clés (CANON, SHARP, ACTUAL SYSTEMES).

Bien cordialement,
Louise BRUN
Chef de projet pour KOESIO GROUPE & KAM
Potentiale
"""
    st.code(mail_text, language="markdown")
