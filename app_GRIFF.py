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
# 3. BARRE LATÉRALE (SIDEBAR) : INGESTION UNIQUE
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

# ------------------------------------------------------------------------------
# 4. FONCTIONS DE CHARGEMENT & NETTOYAGE LIGNE EN-TÊTE
# ------------------------------------------------------------------------------
@st.cache_data
def load_zeendoc_data(uploaded_files):
    df_list = []
    filiale_from_isiwork = None
    
    for file in uploaded_files:
        try:
            df_temp = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
            
            # Détection du nom de la filiale en cellule A1 ou dans les premières lignes
            if not df_temp.empty and pd.notna(df_temp.iloc[0, 0]):
                val_a1 = str(df_temp.iloc[0, 0]).strip()
                if "KOESIO" in val_a1.upper():
                    filiale_from_isiwork = val_a1
            
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
        df_concat = pd.concat(df_list, ignore_index=True)
        return df_concat, filiale_from_isiwork
    return pd.DataFrame(), None

@st.cache_data
def load_artis_data(uploaded_files):
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

df_isiwork, filiale_isiwork = load_zeendoc_data(files_isiwork) if files_isiwork else (pd.DataFrame(), None)
df_artis = load_artis_data(files_artis) if files_artis else pd.DataFrame()

# ------------------------------------------------------------------------------
# 5. DÉTECTION DYNAMIQUE DE LA FILIALE DEPUIS LES FICHIERS
# ------------------------------------------------------------------------------
def detect_filiale(df_art, filiale_isi):
    if not df_art.empty and 'Organisation interne' in df_art.columns:
        filiales = df_art['Organisation interne'].dropna().unique().tolist()
        if filiales:
            return ", ".join(map(str, filiales))
            
    if filiale_isi:
        return filiale_isi
        
    return "En attente de fichier..."

filiale_detectee = detect_filiale(df_artis, filiale_isiwork)

# ------------------------------------------------------------------------------
# 6. FONCTIONS DE CALCUL VENTILÉ PAR MOIS COMPTABLE
# ------------------------------------------------------------------------------
mois_fr = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

def compute_monthly_breakdown(df_art, df_isi):
    rows_ged = []
    rows_erp = []
    
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

    periods_art = df_art['Mois_Period'].dropna().unique().tolist() if 'Mois_Period' in df_art else []
    periods_isi = df_isi['Mois_Period'].dropna().unique().tolist() if 'Mois_Period' in df_isi else []
    all_periods = sorted(list(set(periods_art + periods_isi)))

    if not all_periods:
        all_periods = [None]

    for p in all_periods:
        label_mois = f"{mois_fr.get(p.month, str(p.month))} {p.year}" if p else "Mois non détecté"
        
        sub_art = df_art[df_art['Mois_Period'] == p] if p and 'Mois_Period' in df_art else df_art
        sub_isi = df_isi[df_isi['Mois_Period'] == p] if p and 'Mois_Period' in df_isi else df_isi
        
        # ARTIS (ERP)
        tot_art = len(sub_art)
        if tot_art > 0 and 'Statut' in sub_art.columns:
            art_batch = len(sub_art[sub_art['Statut'] == 'Bon à payer'])
            art_man = tot_art - art_batch
            pct_batch = (art_batch / tot_art) * 100
            pct_man = (art_man / tot_art) * 100
        else:
            art_batch, art_man, pct_batch, pct_man = 0, 0, 0.0, 0.0

        # ISIWORK (GED)
        tot_isi = len(sub_isi)
        if tot_isi > 0 and 'Statut de rapprochement' in sub_isi.columns:
            auto_c = len(sub_isi[sub_isi['Statut de rapprochement'] == 'Rapproché automatiquement'])
            
            faux_c = 0
            if 'Tampons / Notes / Conversations' in sub_isi.columns:
                notes = sub_isi['Tampons / Notes / Conversations'].fillna('')
                faux_c = len(sub_isi[notes.str.contains('FRAIS|PORT|D3EE', case=False, regex=True)])
            
            man_tot = len(sub_isi[sub_isi['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)])
            vrai_c = max(0, man_tot - faux_c)
            
            ged_pure_c = max(0, tot_art - tot_isi) if tot_art > tot_isi else 0
            attente_c = max(0, tot_isi - (auto_c + faux_c + vrai_c))
            
            pct_a = (auto_c / tot_isi) * 100
            pct_f = (faux_c / tot_isi) * 100
            pct_v = (vrai_c / tot_isi) * 100
            pct_gp = (ged_pure_c / tot_art * 100) if tot_art > 0 else 0.0
            pct_att = (attente_c / tot_isi) * 100
        else:
            auto_c, faux_c, vrai_c, ged_pure_c, attente_c = 0, 0, 0, 0, 0
            pct_a, pct_f, pct_v, pct_gp, pct_att = 0.0, 0.0, 0.0, 0.0, 0.0

        rows_ged.append({
            "Mois Comptable": label_mois,
            "Total Factures Traitées": tot_isi,
            "Rapprochement Auto": f"{pct_a:.1f}% ({auto_c})",
            "Faux Manuel (Frais/Port)": f"{pct_f:.1f}% ({faux_c})",
            "Vrai Manuel (Écarts)": f"{pct_v:.1f}% ({vrai_c})",
            "Flux GED pure (Saisie direct ERP)": f"{pct_gp:.1f}% ({ged_pure_c})",
            "En attente": f"{pct_att:.1f}% ({attente_c})"
        })

        rows_erp.append({
            "Mois Comptable": label_mois,
            "Volume Intégré Artis": tot_art,
            "Via Batch Automatique (admin)": f"{pct_batch:.1f}% ({art_batch})",
            "Saisie / Forçage Manuel ERP": f"{pct_man:.1f}% ({art_man})"
        })

    return pd.DataFrame(rows_ged), pd.DataFrame(rows_erp)

df_ged_monthly, df_erp_monthly = compute_monthly_breakdown(df_artis, df_isiwork)

if len(df_ged_monthly) == 1:
    periode_globale = df_ged_monthly["Mois Comptable"].iloc[0]
else:
    periode_globale = f"{df_ged_monthly['Mois Comptable'].iloc[0]} et {df_ged_monthly['Mois Comptable'].iloc[-1]}"

# ------------------------------------------------------------------------------
# 7. ONGLETS DE NAVIGATION
# ------------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🏢 Vision Filiale (Rapport & Mail)", "🌍 Consolidation Groupe"])

# ==============================================================================
# ONGLET 1 : VISION FILIALE
# ==============================================================================
with tab1:
    st.subheader(f"📊 Bilan de Production — {filiale_detectee}")
    
    if not df_isiwork.empty or not df_artis.empty:
        st.success(f"📅 Périodes analysées : **{periode_globale}**")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Factures Artis", f"{len(df_artis):,}".replace(',', ' '))
        with c2:
            st.metric("Total Factures Isiwork", f"{len(df_isiwork):,}".replace(',', ' '))
        with c3:
            st.metric("Mois Traités", len(df_ged_monthly))
        with c4:
            st.metric("Filiale Détectée", filiale_detectee)
            
        st.markdown("---")
        
        st.markdown("### 1. Vision GED — Typologie des flux dans Isiwork (Lignes par Mois)")
        st.dataframe(df_ged_monthly, hide_index=True, use_container_width=True)
        
        st.markdown("### 2. Vision ERP — Intégration et traitement réel dans Artis (Lignes par Mois)")
        st.dataframe(df_erp_monthly, hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### ✉️ Modèle de Mail Généré Automatiquement")
        st.info("Copiez-collez ce texte directement dans votre messagerie :")
        
        try:
            table_ged_md = df_ged_monthly.to_markdown(index=False)
            table_erp_md = df_erp_monthly.to_markdown(index=False)
        except ImportError:
            table_ged_md = df_ged_monthly.to_string(index=False)
            table_erp_md = df_erp_monthly.to_string(index=False)
        
        mail_formatted = f"""Bonjour à tous,

Vous trouverez ci-dessous l'analyse détaillée des performances du flux de rapprochement automatique des factures fournisseurs pour {filiale_detectee}, consolidée sur la période de {periode_globale}.

Afin de maintenir une analyse d'une grande rigueur opérationnelle, nous appliquons la méthodologie de lecture fine des tampons et des lignes de commentaires Isiwork. Cela nous permet de dissocier les rapprochements purement bloquants sur les produits des "Faux Manuels" (factures validées humainement à cause de frais annexes ou de port non saisis sur le BC initial) afin d'identifier les leviers précis de performance.

1. Vision GED
Typologie des flux dans Isiwork

{table_ged_md}

Légende des états:
• Rapprochement Auto: Factures rapprochées à 100% automatiquement sans aucune intervention.
• Faux Manuel: Marchandises 100% conformes, mais forçage manuel pour l'ajout de frais annexes (port, DEEE).
• Vrai Manuel: Blocage humain obligatoire suite à des écarts de prix ou de quantités sur les produits.
• Flux GED pure: Pièces saisies directement sur l'ERP sans passer par le circuit de rapprochement GED.
• En attente: Pièces en cours de validation (effet clôture mensuelle).


2. Vision ERP — Intégration et traitement réel dans Artis

{table_erp_md}


3. Focus Actions / Analyse des rapprochements "Vrais Manuels"
Sur la part des traitements bloqués pour de vrais écarts de prix ou de quantités sur les lignes de produits (hors lignes de frais), les anomalies se concentrent principalement sur trois partenaires majeurs :
• CANON FRANCE
• SHARP BUSINESS SYSTEMS
• ACTUAL SYSTEMES


4. Suivi des chantiers techniques et conclusion
• Stabilité opérationnelle de l'ERP: L'injection finale en comptabilité confirme la performance du batch automatique sans intervention utilisateur.
• Levier "Frais Annexes": L'application locale de l'article d'écart dédié aux frais de port (ZINF-FRAIS_DE_PORT) permettra de convertir automatiquement les Faux Manuels en succès machine.

Nous restons à votre entière disposition pour échanger sur ce bilan.

Bien cordialement,
Louise BRUN
Chef de projet pour KOESIO GROUPE & KAM
Potentiale
"""
        st.code(mail_formatted, language="markdown")
        
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork/Zeendoc** et **Artis** dans le panneau latéral.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION GROUPE
# ==============================================================================
with tab2:
    st.subheader("🌍 Consolidation Inter-Filiales Koesio")
    st.info("Cet onglet permettra de consolider l'ensemble des filiales lorsque plusieurs fichiers seront chargés simultanément.")
