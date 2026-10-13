import streamlit as st
import pandas as pd
import plotly.express as px

# Configuration de la page Streamlit
st.set_page_config(
    page_title="GRIFF - Tour de Contrôle Koesio",
    page_icon="🦅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Style CSS personnalisé
st.markdown("""
    <style>
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #00539C;
        margin-bottom: 0px;
    }
    .acronym-box {
        background-color: #F0F4F8;
        border-left: 5px solid #00539C;
        padding: 12px 18px;
        border-radius: 4px;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# En-tête principal
st.markdown('<p class="main-title">🦅 GRIFF — Tour de Contrôle Koesio</p>', unsafe_allow_html=True)

# Explication de l'acronyme GRIFF
st.markdown("""
<div class="acronym-box">
    <b>GRIFF</b> : <b>G</b>estion & <b>R</b>approchement <b>I</b>ntelligent des <b>F</b>actures et <b>F</b>lux<br/>
    <small>Application d'analyse mensuelle et de contrôle automatisé des rapprochements BL / Factures entre les outils <b>Isiwork</b> et <b>Artis</b> pour l'ensemble des filiales Koesio.</small>
</div>
""", unsafe_unsafe_html=False if hasattr(st, "unsafe_html") else True)

# Sidebar : Chargement des fichiers
st.sidebar.header("📂 Ingestion des Données")
st.sidebar.markdown("Téléversez vos exports mensuels pour lancer le contrôle automatisé.")

files_isiwork = st.sidebar.file_uploader(
    "Exports Isiwork (1 ou plusieurs fichiers)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Sélectionnez un ou plusieurs fichiers d'exports Isiwork."
)

files_artis = st.sidebar.file_uploader(
    "Exports Artis (1 ou plusieurs fichiers)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Sélectionnez un ou plusieurs fichiers d'exports Artis."
)

st.sidebar.markdown("---")
st.sidebar.header("🎯 Filtres Globaux")
filiale_select = st.sidebar.selectbox(
    "Filiale Koesio",
    ["Toutes les filiales", "Koesio CEE", "Koesio AURA", "Koesio Aquitaine", "Autres filiales"]
)

mois_select = st.sidebar.selectbox(
    "Période d'analyse",
    ["Mois en cours", "Dernier mois clos", "Cumul annuel (YTD)"]
)

# Fonction de chargement et concaténation des fichiers
@st.cache_data
def load_and_merge_files(uploaded_files):
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

# Traitement des données si les fichiers sont fournis
df_isiwork = load_and_merge_files(files_isiwork) if files_isiwork else pd.DataFrame()
df_artis = load_and_merge_files(files_artis) if files_artis else pd.DataFrame()

# Navigation par Onglets
tab1, tab2 = st.tabs(["🏢 Analyse par Filiale", "🌍 Consolidation Groupe"])

# --- ONGLET 1 : ANALYSE PAR FILIALE ---
with tab1:
    st.subheader(f"📊 Analyse détaillée — {filiale_select}")
    
    if files_isiwork or files_artis:
        st.success(f" Données chargées avec succès : {len(files_isiwork)} fichier(s) Isiwork | {len(files_artis)} fichier(s) Artis")
        
        # Section KPI
        col1, col2, col3, col4 = st.st.columns(4) if hasattr(st, "st") else st.columns(4)
        with col1:
            st.metric("Total Factures Ingestées", "1 248", delta="+5.2%")
        with col2:
            st.metric("Taux de Rapprochement Auto", "84.5%", delta="+2.1%")
        with col3:
            st.metric("Écarts de Prix / Montants", "42", delta="-8", delta_color="inverse")
        with col4:
            st.metric("Dossiers à Dérapprocher / Traiter", "18", delta="À vérifier")
            
        st.markdown("---")
        
        # Modèle type d'analyse
        col_left, col_right = st.columns([2, 1])
        
        with col_left:
            st.markdown("### 🔍 Répartition des Anomalies & Rejets par Fournisseur")
            # Données factices de démonstration
            df_demo = pd.DataFrame({
                "Fournisseur": ["SHARP", "EDOX", "TECH DATA / TD SYNNEX", "HP", "Canon"],
                "Rapprochés Auto": [450, 210, 180, 120, 80],
                "Écarts / Erreurs": [15, 8, 22, 5, 2]
            })
            fig = px.bar(
                df_demo, 
                x="Fournisseur", 
                y=["Rapprochés Auto", "Écarts / Erreurs"],
                title="Status des Rapprochements par Fournisseur Clé",
                barmode="stack",
                color_discrete_sequence=["#00539C", "#E53E3E"]
            )
            st.plotly_chart(fig, use_container_width=True)
            
        with col_right:
            st.markdown("### ⚠️ Principales Causes de Non-Rapprochement")
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
            st.dataframe(causes_df, hide_index=
