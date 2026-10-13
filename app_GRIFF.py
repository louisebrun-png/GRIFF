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
# 2. EN-TÊTE ET EXPLICATION DE L'ACRONYME GRIFF (Compatible Mode Sombre / Clair)
# ------------------------------------------------------------------------------
st.title("🦅 GRIFF — Tour de Contrôle Koesio")

# Message explicatif nativement stylisé par Streamlit pour s'adapter au thème
st.info(
    "**GRIFF** : **G**estion & **R**approchement **I**ntelligent des **F**actures et **F**lux\n\n"
    "Application d'analyse mensuelle et de contrôle automatisé des rapprochements Bon de Livraison (BL) / Factures d'Achat entre les systèmes **Isiwork** et **Artis** pour les filiales Koesio."
)

st.markdown("---")

# ------------------------------------------------------------------------------
# 3. BARRE LATÉRALE (SIDEBAR) : INGESTION DES FICHIERS & FILTRES
# ------------------------------------------------------------------------------
st.sidebar.header("📂 Ingestion des Données")
st.sidebar.markdown("Téléversez vos exports mensuels pour démarrer l'analyse.")

# Upload des fichiers Isiwork (multi-fichiers autorisés)
files_isiwork = st.sidebar.file_uploader(
    "Exports Isiwork (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Vous pouvez sélectionner un ou plusieurs fichiers d'exports Isiwork."
)

# Upload des fichiers Artis (multi-fichiers autorisés)
files_artis = st.sidebar.file_uploader(
    "Exports Artis (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Vous pouvez sélectionner un ou plusieurs fichiers d'exports Artis."
)

st.sidebar.markdown("---")
st.sidebar.header("🎯 Filtres d'Analyse")

# Filtre par filiale Koesio
filiale_select = st.sidebar.selectbox(
    "Filiale Koesio",
    ["Toutes les filiales", "Koesio CEE", "Koesio AURA", "Koesio Aquitaine", "Autres filiales"]
)

# Filtre par période
mois_select = st.sidebar.selectbox(
    "Période d'analyse",
    ["Mois en cours", "Dernier mois clos", "Cumul annuel (YTD)"]
)

# ------------------------------------------------------------------------------
# 4. FONCTIONS UTILES (Chargement et fusion des données)
# ------------------------------------------------------------------------------
@st.cache_data
def load_and_merge_files(uploaded_files):
    """
    Fonction permettant de lire et concaténer plusieurs fichiers Excel ou CSV.
    Met en cache le résultat pour accélérer l'affichage.
    """
    df_list = []
    for file in uploaded_files:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        df_list.append(df)
    
    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return pd.DataFrame()

# Chargement effectif des données téléversées
df_isiwork = load_and_merge_files(files_isiwork) if files_isiwork else pd.DataFrame()
df_artis = load_and_merge_files(files_artis) if files_artis else pd.DataFrame()

# ------------------------------------------------------------------------------
# 5. STRUCTURE PAR ONGLETS (DÉTAIL FILIALE vs SYNTHÈSE GROUPE)
# ------------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🏢 Analyse par Filiale", "🌍 Consolidation Groupe"])

# ==============================================================================
# ONGLET 1 : ANALYSE DÉTAILLÉE PAR FILIALE
# ==============================================================================
with tab1:
    st.subheader(f"📊 Analyse détaillée — {filiale_select}")
    
    # Affichage si des fichiers sont chargés
    if files_isiwork or files_artis:
        st.success(f"Données prêtes : {len(files_isiwork)} fichier(s) Isiwork | {len(files_artis)} fichier(s) Artis")
        
        # --- CARTE DES INDICATEURS CLÉS (KPI) ---
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Factures Ingestées", "1 248", delta="+5.2%")
        with col2:
            st.metric("Taux Rapprochement Auto", "84.5%", delta="+2.1%")
        with col3:
            st.metric("Écarts de Prix / Montants", "42", delta="-8", delta_color="inverse")
        with col4:
            st.metric("Dossiers à Dérapprocher", "18", delta="À vérifier")
            
        st.markdown("---")
        
        # --- GRAPHIQUES ET ANALYSES ---
        col_left, col_right = st.columns([2, 1])
        
        with col_left:
            st.markdown("### 🔍 Rapprochements par Fournisseur Clé")
            # Données de démonstration
            df_demo = pd.DataFrame({
                "Fournisseur": ["SHARP", "EDOX", "TECH DATA / TD SYNNEX", "HP", "Canon"],
                "Rapprochés Auto": [450, 210, 180, 120, 80],
                "Écarts / Erreurs": [15, 8, 22, 5, 2]
            })
            fig = px.bar(
                df_demo, 
                x="Fournisseur", 
                y=["Rapprochés Auto", "Écarts / Erreurs"],
                title="Rapprochements Automatiques vs Écarts par Fournisseur",
                barmode="stack",
                color_discrete_sequence=["#2B6CB0", "#E53E3E"]
            )
            st.plotly_chart(fig, use_container_width=True)
            
        with col_right:
            st.markdown("### ⚠️ Motifs de Non-Rapprochement")
            causes_df = pd.DataFrame({
                "Motif de Rejet": [
                    "Article inconnu / Non mappé",
                    "Écart de prix HT",
                    "Frais annexes non prévus",
                    "SIRET Organisation incorrect",
                    "Avoir / Crédit manquant"
                ],
                "Volume": [28, 18, 12, 5, 3]
            })
            st.dataframe(causes_df, hide_index=True, use_container_width=True)
            
        st.markdown("---")
        st.markdown("### 📋 Table de Contrôle Détaillée (Aperçu des Données)")
        
        # Affichage des premières lignes des fichiers injectés
        if not df_isiwork.empty:
            st.markdown("**Extraits de l'export Isiwork :**")
            st.dataframe(df_isiwork.head(10), use_container_width=True)
            
        if not df_artis.empty:
            st.markdown("**Extraits de l'export Artis :**")
            st.dataframe(df_artis.head(10), use_container_width=True)
            
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork** et/ou **Artis** dans le panneau latéral pour lancer l'analyse.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION ET BENCHMARK GROUPE
# ==============================================================================
with tab2:
    st.subheader("🌍 Vision Consolidée Groupe Koesio")
    st.markdown("Vue d'ensemble comparative de la performance de rapprochement inter-filiales.")
    
    # Indicateurs de performance Groupe
    g1, g2, g3 = st.columns(3)
    with g1:
        st.metric("Volume Total Groupe", "4 890 Factures", "+3.8%")
    with g2:
        st.metric("Taux Rapprochement Moyen Groupe", "81.2%", "+1.5%")
    with g3:
        st.metric("Gain de Temps Estimé", "124 Heures / Mois", "Grâce à l'automatisation")
        
    st.markdown("---")
    
    # Tableau de benchmark par filiale
    st.markdown("### 🏆 Benchmark Inter-Filiales")
    df_groupe = pd.DataFrame({
        "Filiale Koesio": ["Koesio CEE", "Koesio AURA", "Koesio Aquitaine", "Koesio IDF", "Koesio PACA"],
        "Factures Traitées": [1248, 1100, 950, 890, 702],
        "Rapprochement Auto (%)": ["84.5%", "82.0%", "79.8%", "78.5%", "81.0%"],
        "Écarts à Traiter": [42, 51, 63, 58, 30],
        "Statut Recette / Prod": ["✅ Validé v5.5.0", "⚙️ En préparation", "✅ Production", "✅ Production", "✅ Production"]
    })
    st.dataframe(df_groupe, hide_index=True, use_container_width=True)
    
    # Graphique de répartition du volume Groupe
    fig_groupe = px.line(
        df_groupe, 
        x="Filiale Koesio", 
        y="Factures Traitées", 
        title="Volume de Factures Traitées par Filiale (Mensuel)",
        markers=True,
        color_discrete_sequence=["#2B6CB0"]
    )
    st.plotly_chart(fig_groupe, use_container_width=True)
