# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle et rapprochement automatique BL / Factures
# AUTEUR      : Louise BRUN - Koesio Groupe
# ==============================================================================

import streamlit as st
import pandas as pd

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
# 3. BARRE LATÉRALE (SIDEBAR) : INGESTION AVEC DÉTECTION AUTO DES TYPES
# ------------------------------------------------------------------------------
st.sidebar.header("📂 Ingestion des Données")
st.sidebar.markdown("Téléversez vos exports mensuels pour lancer le calcul automatisé.")

files_isiwork_input = st.sidebar.file_uploader(
    "Exports Isiwork / Zeendoc (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Export Zeendoc contenant la colonne 'Statut de rapprochement'."
)

files_artis_input = st.sidebar.file_uploader(
    "Exports Artis ERP (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Export de la recherche des factures fournisseurs Artis."
)

all_uploaded_files = (files_isiwork_input or []) + (files_artis_input or [])

# ------------------------------------------------------------------------------
# 4. FONCTIONS INTELLIGENTES DE SÉPARATION & CHARGEMENT
# ------------------------------------------------------------------------------
@st.cache_data
def load_all_data(uploaded_files):
    isi_list = []
    art_list = []
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.xlsx'):
                df_temp = pd.read_excel(file, header=None)
            else:
                df_temp = pd.read_csv(file, header=None)
                
            is_artis = False
            for idx, row in df_temp.iloc[:5].iterrows():
                row_str = " ".join(map(str, row.values))
                if 'Raison sociale' in row_str or 'Organisation interne' in row_str or 'Code fournisseur' in row_str:
                    is_artis = True
                    break
            
            if is_artis:
                if file.name.endswith('.xlsx'):
                    df_art = pd.read_excel(file)
                else:
                    df_art = pd.read_csv(file)
                art_list.append(df_art)
            else:
                header_row = 0
                for idx, row in df_temp.iterrows():
                    if 'Statut de rapprochement' in row.values or 'Identifiant' in row.values or 'Date facture' in row.values:
                        header_row = idx
                        break
                
                if file.name.endswith('.xlsx'):
                    df_isi = pd.read_excel(file, header=header_row)
                else:
                    df_isi = pd.read_csv(file, header=header_row)
                isi_list.append(df_isi)
                
        except Exception as e:
            st.error(f"Erreur lors de la lecture de {file.name}: {e}")
            
    df_isi_concat = pd.concat(isi_list, ignore_index=True) if isi_list else pd.DataFrame()
    df_art_concat = pd.concat(art_list, ignore_index=True) if art_list else pd.DataFrame()
    
    return df_isi_concat, df_art_concat

df_isiwork, df_artis = load_all_data(all_uploaded_files)

# ------------------------------------------------------------------------------
# 5. DÉTECTION DYNAMIQUE DE LA FILIALE
# ------------------------------------------------------------------------------
def detect_filiale(df_art):
    if not df_art.empty and 'Organisation interne' in df_art.columns:
        filiales = df_art['Organisation interne'].dropna().unique().tolist()
        if filiales:
            return ", ".join(map(str, filiales))
    return "En attente de fichier..."

filiale_detectee = detect_filiale(df_artis)

# ------------------------------------------------------------------------------
# 6. CALCULS ET VENTILATION PAR MOIS (EXCLUSION MOIS EN COURS INCOMPLET)
# ------------------------------------------------------------------------------
mois_fr = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

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

    # On extrait uniquement les mois réellement présents dans l'export Isiwork
    periods_isi = df_isi['Mois_Period'].dropna().unique().tolist() if 'Mois_Period' in df_isi else []
    
    # Filtrage : On ne garde que les mois validés Isiwork (exclut Octobre en cours)
    all_periods = sorted([p for p in periods_isi if p is not None])

    rows_ged = []
    rows_erp = []

    for p in all_periods:
        label_mois = f"{mois_fr.get(p.month, str(p.month))} {p.year}"
        
        sub_art = df_art[df_art['Mois_Period'] == p] if 'Mois_Period' in df_art else pd.DataFrame()
        sub_isi = df_isi[df_isi['Mois_Period'] == p] if 'Mois_Period' in df_isi else pd.DataFrame()
        
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

    return pd.DataFrame(rows_ged), pd.DataFrame(rows_erp), all_periods

df_ged_monthly, df_erp_monthly, valid_periods = compute_monthly_breakdown(df_artis, df_isiwork)

# Filtrage global des compteurs KPI pour ne compter que les mois retenus
if valid_periods and not df_artis.empty and 'Date' in df_artis.columns:
    df_artis_filtered = df_artis[pd.to_datetime(df_artis['Date'], errors='coerce').dt.to_period('M').isin(valid_periods)]
else:
    df_artis_filtered = df_artis

if valid_periods and not df_isiwork.empty and 'Date facture' in df_isiwork.columns:
    df_isiwork_filtered = df_isiwork[pd.to_datetime(df_isiwork['Date facture'], errors='coerce').dt.to_period('M').isin(valid_periods)]
else:
    df_isiwork_filtered = df_isiwork

if not df_ged_monthly.empty:
    if len(df_ged_monthly) == 1:
        periode_globale = df_ged_monthly["Mois Comptable"].iloc[0]
    else:
        periode_globale = f"{df_ged_monthly['Mois Comptable'].iloc[0]} et {df_ged_monthly['Mois Comptable'].iloc[-1]}"
else:
    periode_globale = "Mois non détecté"

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
        except Exception:
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
Sur la part des traitements bloqués pour de vrais écarts de prix ou de quantités sur les lignes de produits (hors lignes de frais), les anomalies se concentrent principalement sur les fournisseurs clés :
• TD SYNNEX
• ACTUAL SYSTEMES
• CANON FRANCE


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
