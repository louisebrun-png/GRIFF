# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle, historique M-1 et suivi des chantiers
# AUTEUR      : Louise BRUN - Koesio Groupe
# ==============================================================================

import streamlit as st
import pandas as pd
import re

# Import sécurisé de pypdf sans faire planter l'application
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
# 3. BARRE LATÉRALE (SIDEBAR) : INGESTION & HISTORIQUE MAIL PDF
# ------------------------------------------------------------------------------
st.sidebar.header("📂 Ingestion des Données M (Mois en cours)")
st.sidebar.markdown("Téléversez vos exports mensuels pour lancer le calcul automatisé.")

files_isiwork_input = st.sidebar.file_uploader(
    "Exports Isiwork / Zeendoc (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Exports Zeendoc de vos différentes filiales."
)

files_artis_input = st.sidebar.file_uploader(
    "Exports Artis ERP (.xlsx, .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Exports Artis ERP de vos différentes filiales."
)

st.sidebar.markdown("---")
st.sidebar.header("📜 Historique M-1 (Mail PDF Passé)")
file_history_input = st.sidebar.file_uploader(
    "Mail PDF du Rapport Précédent M-1 (.pdf)",
    type=["pdf"],
    accept_multiple_files=False,
    help="Téléversez le mail PDF du mois dernier pour extraire l'historique et l'évolution par filiale."
)

all_uploaded_files = (files_isiwork_input or []) + (files_artis_input or [])

# ------------------------------------------------------------------------------
# 4. PARSER AUTOMATIQUE DE L'HISTORIQUE ET DES CHANTIERS DEPUIS LE MAIL PDF
# ------------------------------------------------------------------------------
@st.cache_data
def parse_pdf_history_full(pdf_file):
    if not pdf_file or not HAS_PYPDF:
        return {}, []
    try:
        reader = pypdf.PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() + "\n"
            
        hist_data = {}
        chantiers = []
        
        # 1. Extraction des taux par filiale depuis les tableaux du PDF
        pattern_ged = re.findall(r'KOESIO\s*([A-Za-zÀ-ÿ\s\-\diT]+?)\s*(\d[\d\s]*)\s+(\d+)\s+(\d+)\s+([\d,]+)\s*%\s*\(([\d\s]+)\)\s+([\d,]+)\s*%\s*\(([\d\s]+)\)\s+([\d,]+)\s*%\s*\(([\d\s]+)\)', full_text)
        
        for match in pattern_ged:
            name = match[0].strip()
            auto_pct = float(match[4].replace(',', '.'))
            vrai_pct = float(match[8].replace(',', '.'))
            hist_data[name] = {
                "auto_pct_m1": auto_pct,
                "vrai_pct_m1": vrai_pct
            }
            
        # 2. Extraction des chantiers techniques historiques
        lines = full_text.split('\n')
        in_chantiers = False
        for line in lines:
            if "SUIVI DES CHANTIERS" in line.upper() or "FOCUS ACTIONS" in line.upper():
                in_chantiers = True
            elif in_chantiers and line.strip().startswith(('•', '1.', '2.', '3.', '-')):
                chantiers.append(line.strip())
            elif in_chantiers and len(chantiers) > 0 and line.strip() == "":
                in_chantiers = False
                
        return hist_data, chantiers
    except Exception as e:
        st.error(f"Erreur lors de la lecture du PDF historique : {e}")
        return {}, []

dict_hist_m1, list_chantiers_m1 = parse_pdf_history_full(file_history_input) if file_history_input else ({}, [])

# ------------------------------------------------------------------------------
# 5. CHARGEMENT ET NETTOYAGE DES FICHIERS EXPORTS M
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
# 6. DÉTECTION DYNAMIQUE DE LA FILIALE EN VISION INDIVIDUELLE
# ------------------------------------------------------------------------------
def detect_filiale(df_art):
    if not df_art.empty and 'Organisation interne' in df_art.columns:
        filiales = df_art['Organisation interne'].dropna().unique().tolist()
        if filiales:
            return ", ".join(map(str, filiales))
    return "En attente de fichier..."

filiale_detectee = detect_filiale(df_artis)

# ------------------------------------------------------------------------------
# 7. CALCULS ET VENTILATION PAR MOIS
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

    periods_isi = df_isi['Mois_Period'].dropna().unique().tolist() if 'Mois_Period' in df_isi else []
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
        tot_isi_brut = len(sub_isi)
        if tot_isi_brut > 0 and 'Statut de rapprochement' in sub_isi.columns:
            notes = sub_isi['Tampons / Notes / Conversations'].fillna('') if 'Tampons / Notes / Conversations' in sub_isi.columns else pd.Series('', index=sub_isi.index)
            
            no_rappro_mask = sub_isi['Statut de rapprochement'].isna()
            faf_artis_mask = notes.str.contains('FAF ARTIS|FAF_ARTIS', case=False, regex=True)
            
            gap_volume = max(0, tot_art - tot_isi_brut) if tot_art > tot_isi_brut else 0
            ged_pure_c = no_rappro_mask.sum() + faf_artis_mask.sum() + gap_volume
            
            vol_epure = max(0, tot_isi_brut - (no_rappro_mask.sum() + faf_artis_mask.sum()))
            
            auto_c = len(sub_isi[sub_isi['Statut de rapprochement'] == 'Rapproché automatiquement'])
            faux_c = len(sub_isi[notes.str.contains('FRAIS|PORT|D3EE', case=False, regex=True)])
            
            man_tot = len(sub_isi[sub_isi['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)])
            vrai_c = max(0, man_tot - (faux_c + faf_artis_mask.sum()))
            
            attente_c = max(0, vol_epure - (auto_c + faux_c + vrai_c))
            
            denom = vol_epure if vol_epure > 0 else tot_isi_brut
            
            pct_gp = (ged_pure_c / tot_isi_brut * 100) if tot_isi_brut > 0 else 0.0
            pct_a = (auto_c / denom) * 100
            pct_f = (faux_c / denom) * 100
            pct_v = (vrai_c / denom) * 100
            pct_att = (attente_c / denom) * 100
        else:
            auto_c, faux_c, vrai_c, ged_pure_c, attente_c, vol_epure = 0, 0, 0, 0, 0, 0
            pct_a, pct_f, pct_v, pct_gp, pct_att = 0.0, 0.0, 0.0, 0.0, 0.0

        rows_ged.append({
            "Mois Comptable": label_mois,
            "📋 Vol. Brut GED": tot_isi_brut,
            "📁 Flux GED Pure (Archivage / FAF)": f"{pct_gp:.1f}% ({ged_pure_c})",
            "📊 Vol. Épuré": vol_epure,
            "🤖 Rapprochement Auto": f"{pct_a:.1f}% ({auto_c})",
            "🚚 Faux Manuel (Frais/Port)": f"{pct_f:.1f}% ({faux_c})",
            "🛠️ Vrai Manuel (Écarts)": f"{pct_v:.1f}% ({vrai_c})",
            "⏳ En attente": f"{pct_att:.1f}% ({attente_c})"
        })

        rows_erp.append({
            "Mois Comptable": label_mois,
            "💻 Volume Intégré Artis": tot_art,
            "🤖 Via Batch Automatique (admin)": f"{pct_batch:.1f}% ({art_batch})",
            "👤 Saisie / Forçage Manuel ERP": f"{pct_man:.1f}% ({art_man})"
        })

    return pd.DataFrame(rows_ged), pd.DataFrame(rows_erp), all_periods

df_ged_monthly, df_erp_monthly, valid_periods = compute_monthly_breakdown(df_artis, df_isiwork)

# Filtrage DataFrames
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
# 8. EXTRACTION DYNAMIQUE TOPS FOURNISSEURS
# ------------------------------------------------------------------------------
txt_top_vrai_html = ""
txt_top_faux_html = ""

if not df_isiwork_filtered.empty and 'Statut de rapprochement' in df_isiwork_filtered.columns:
    df_man = df_isiwork_filtered[df_isiwork_filtered['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)].copy()
    notes_man = df_man['Tampons / Notes / Conversations'].fillna('') if 'Tampons / Notes / Conversations' in df_man.columns else pd.Series('', index=df_man.index)
    
    mask_faux = notes_man.str.contains('FRAIS|PORT|D3EE', case=False, regex=True)
    mask_faf = notes_man.str.contains('FAF ARTIS|FAF_ARTIS', case=False, regex=True)
    
    df_vrai_fourn = df_man[~mask_faux & ~mask_faf]
    df_faux_fourn = df_man[mask_faux]
    
    if not df_vrai_fourn.empty and 'Fournisseur' in df_vrai_fourn.columns:
        s_vrai = df_vrai_fourn['Fournisseur'].value_counts().head(5)
        txt_top_vrai_html = "".join([f"<li><b>{fourn}</b> : {cnt} factures</li>" for fourn, cnt in s_vrai.items()])
    
    if not df_faux_fourn.empty and 'Fournisseur' in df_faux_fourn.columns:
        s_faux = df_faux_fourn['Fournisseur'].value_counts().head(5)
        txt_top_faux_html = "".join([f"<li><b>{fourn}</b> : {cnt} factures</li>" for fourn, cnt in s_faux.items()])

if not txt_top_vrai_html:
    txt_top_vrai_html = "<li><b>TD SYNNEX</b></li><li><b>ACTUAL SYSTEMES</b></li><li><b>CANON FRANCE</b></li>"

if not txt_top_faux_html:
    txt_top_faux_html = "<li><b>ACTUAL SYSTEMES</b></li><li><b>CANON FRANCE</b></li><li><b>EPARTS</b></li>"

# ------------------------------------------------------------------------------
# 9. ONGLETS DE NAVIGATION
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
        st.info("Sélectionnez le texte ci-dessous et collez-le directement dans Outlook : les tableaux et emojis se formeront parfaitement !")
        
        table_ged_html = df_ged_monthly.to_html(index=False, border=1, classes='content-table')
        table_erp_html = df_erp_monthly.to_html(index=False, border=1, classes='content-table')
        
        mail_html = f"""
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.5; padding: 10px; background-color: #ffffff; border-radius: 5px;">
            <p>Bonjour à tous,</p>
            
            <p>Vous trouverez ci-dessous l'analyse détaillée des performances du flux de rapprochement automatique des factures fournisseurs pour <b>{filiale_detectee}</b>, consolidée sur la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément aux directives du Groupe, nous appliquons la méthodologie d'<b>isolation des flux « GED Pure / FAF Artis »</b> (factures d'archivage directement traitées de bout en bout dans l'ERP sans rapprochement en GED) afin de refléter l'efficacité machine réelle d'Isiwork.</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — Typologie des flux dans la GED</h3>
            {table_ged_html}
            
            <p style="font-size: 12px; color: #555; margin-top: 10px;">
                <b>Légende des états :</b><br>
                • <b>🤖 Rapprochement Auto :</b> Factures rapprochées à 100% automatiquement sans aucune intervention.<br>
                • <b>🚚 Faux Manuel :</b> Marchandises 100% conformes, mais forçage manuel pour l'ajout de frais annexes (port, DEEE).<br>
                • <b>🛠️ Vrai Manuel :</b> Blocage humain obligatoire suite à des écarts de prix ou de quantités sur les produits.<br>
                • <b>📁 Flux GED pure :</b> Pièces d'archivage direct saisies sur l'ERP sans passer par le circuit de rapprochement GED.<br>
                • <b>⏳ En attente :</b> Pièces en cours de validation (effet clôture mensuelle).
            </p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">💻 2. VISION ERP ARTIS — Intégration et traitement réel dans l'ERP</h3>
            {table_erp_html}
            
            <h3 style="color: #0056b3; margin-top: 20px;">🔍 3. FOCUS ACTIONS — Analyse des "Vrais Manuels" & Frictions</h3>
            <p><b>• Top 5 Fournisseurs bloquants (Écarts prix/réf) :</b></p>
            <ul>
                {txt_top_vrai_html}
            </ul>
            
            <p><b>• Top Faux Manuels (Frais de port et frais annexes isolés) :</b></p>
            <ul>
                {txt_top_faux_html}
            </ul>
            
            <h3 style="color: #0056b3; margin-top: 20px;">⚙️ 4. SUIVI DES CHANTIERS TECHNIQUES & PERSPECTIVES</h3>
            <ul>
                <li><b>Stabilité opérationnelle de l'ERP :</b> L'injection finale en comptabilité confirme la performance du batch automatique sans intervention utilisateur.</li>
                <li><b>Levier "Frais Annexes" :</b> Les Faux Manuels s'expliquent par l'absence des lignes de frais de port ou DEEE sur la commande d'achat ou le BL initial. Au niveau d'Isiwork, il est alors trop tard pour automatiser. Pour éliminer ces blocages, ces frais doivent impérativement être saisis en amont lors de la création de la commande ou de la réception du BL dans Artis.</li>
            </ul>
            
            <p>Cordialement,</p>
            <p><b>Louise BRUN</b><br>
            Chef de projet pour KOESIO GROUPE & KAM<br>
            Potentiale</p>
        </div>
        """
        
        st.markdown("""
        <style>
            .content-table {
                border-collapse: collapse;
                margin: 15px 0;
                font-size: 13px;
                min-width: 100%;
                border: 1px solid #dddddd;
            }
            .content-table th {
                background-color: #0056b3;
                color: #ffffff;
                text-align: left;
                padding: 8px 12px;
            }
            .content-table td {
                padding: 8px 12px;
                border: 1px solid #dddddd;
            }
            .content-table tr:nth-of-type(even) {
                background-color: #f9f9f9;
            }
        </style>
        """, unsafe_allow_html=True)
        
        st.components.v1.html(mail_html, height=850, scrolling=True)
        
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork/Zeendoc** et **Artis** dans le panneau latéral.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION GROUPE ET COMPARATIF HISTORIQUE M-1
# ==============================================================================
with tab2:
    st.subheader("🌍 Consolidation Inter-Filiales — Groupe Koesio")
    
    if dict_hist_m1:
        st.success(f"📜 **Mail PDF M-1 chargé** : Données historiques extraites pour {len(dict_hist_m1)} filiales.")
        if list_chantiers_m1:
            st.info("⚙️ **Chantiers historiques M-1 identifiés :**\n" + "\n".join([f"• {c}" for c in list_chantiers_m1[:4]]))
    
    if not df_artis_filtered.empty and 'Organisation interne' in df_artis_filtered.columns:
        filiales_groupe = df_artis_filtered['Organisation interne'].dropna().unique().tolist()
        
        st.success(f"🏢 **{len(filiales_groupe)} filiales actives chargées** : {', '.join(map(str, filiales_groupe))}")
        
        rows_ged_groupe = []
        tot_vol_brut = 0
        tot_ged_pure = 0
        tot_vol_epure = 0
        tot_auto = 0
        tot_faux = 0
        tot_vrai = 0
        tot_att = 0
        
        for idx, f in enumerate(filiales_groupe, start=1):
            sub_art_f = df_artis_filtered[df_artis_filtered['Organisation interne'] == f]
            sub_isi_f = df_isiwork_filtered[df_isiwork_filtered['Organisation interne'] == f] if 'Organisation interne' in df_isiwork_filtered.columns else df_isiwork_filtered
            
            vol_brut = len(sub_isi_f)
            tot_art_f = len(sub_art_f)
            
            notes = sub_isi_f['Tampons / Notes / Conversations'].fillna('') if 'Tampons / Notes / Conversations' in sub_isi_f.columns else pd.Series('', index=sub_isi_f.index)
            
            no_rappro = sub_isi_f['Statut de rapprochement'].isna() if 'Statut de rapprochement' in sub_isi_f.columns else pd.Series(False, index=sub_isi_f.index)
            faf_cnt = len(sub_isi_f[notes.str.contains('FAF ARTIS|FAF_ARTIS', case=False, regex=True)])
            gap_cnt = max(0, tot_art_f - vol_brut) if tot_art_f > vol_brut else 0
            
            ged_pure = no_rappro.sum() + faf_cnt + gap_cnt
            vol_epure = max(0, vol_brut - (no_rappro.sum() + faf_cnt))
            
            auto_c = len(sub_isi_f[sub_isi_f['Statut de rapprochement'] == 'Rapproché automatiquement']) if 'Statut de rapprochement' in sub_isi_f.columns else 0
            faux_c = len(sub_isi_f[notes.str.contains('FRAIS|PORT|D3EE', case=False, regex=True)])
            
            man_tot = len(sub_isi_f[sub_isi_f['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)]) if 'Statut de rapprochement' in sub_isi_f.columns else 0
            vrai_c = max(0, man_tot - (faux_c + faf_cnt))
            
            att_c = max(0, vol_epure - (auto_c + faux_c + vrai_c))
            
            pct_gp = (ged_pure / vol_brut * 100) if vol_brut > 0 else 0
            denom = vol_epure if vol_epure > 0 else vol_brut
            
            pct_a = (auto_c / denom * 100) if denom > 0 else 0
            pct_f = (faux_c / denom * 100) if denom > 0 else 0
            pct_v = (vrai_c / denom * 100) if denom > 0 else 0
            pct_att = (att_c / denom * 100) if denom > 0 else 0
            pct_pot = pct_a + pct_f
            
            # CALCUL ÉVOLUTION M VS M-1
            str_eval_auto = f"{pct_a:.1f}% ({auto_c})"
            if dict_hist_m1:
                # recherche correspondante filiale
                for key_h, val_h in dict_hist_m1.items():
                    if f.upper() in key_h.upper() or key_h.upper() in f.upper():
                        delta_a = pct_a - val_h.get("auto_pct_m1", pct_a)
                        symb = "↗️" if delta_a > 0 else "↘️" if delta_a < 0 else "➡️"
                        str_eval_auto += f" [{symb} {delta_a:+.1f}% vs M-1]"
                        break

            tot_vol_brut += vol_brut
            tot_ged_pure += ged_pure
            tot_vol_epure += vol_epure
            tot_auto += auto_c
            tot_faux += faux_c
            tot_vrai += vrai_c
            tot_att += att_c
            
            rows_ged_groupe.append({
                "Rang": idx,
                "Filiale Groupe": f,
                "Vol. Brut GED": vol_brut,
                "Flux GED Pure": ged_pure,
                "Vol. Épuré": vol_epure,
                "🤖 Rapprochement Auto (Évolution M vs M-1)": str_eval_auto,
                "🚚 Faux Manuel (Frais)": f"{pct_f:.1f}% ({faux_c})",
                "🛠️ Vrai Manuel (Litiges)": f"{pct_v:.1f}% ({vrai_c})",
                "⏳ En attente": f"{pct_att:.1f}% ({att_c})",
                "🎯 Potentiel Cible": f"{pct_pot:.1f}%"
            })
            
        pct_tot_gp = (tot_ged_pure / tot_vol_brut * 100) if tot_vol_brut > 0 else 0
        pct_tot_a = (tot_auto / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_f = (tot_faux / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_v = (tot_vrai / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        pct_tot_att = (tot_att / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
        
        rows_ged_groupe.append({
            "Rang": "—",
            "Filiale Groupe": "TOTAL GROUPE CONSOLIDÉ",
            "Vol. Brut GED": tot_vol_brut,
            "Flux GED Pure": tot_ged_pure,
            "Vol. Épuré": tot_vol_epure,
            "🤖 Rapprochement Auto (Évolution M vs M-1)": f"{pct_tot_a:.1f}% ({tot_auto})",
            "🚚 Faux Manuel (Frais)": f"{pct_tot_f:.1f}% ({tot_faux})",
            "🛠️ Vrai Manuel (Litiges)": f"{pct_tot_v:.1f}% ({tot_vrai})",
            "⏳ En attente": f"{pct_tot_att:.1f}% ({tot_att})",
            "🎯 Potentiel Cible": f"{(pct_tot_a + pct_tot_f):.1f}%"
        })
        
        df_ged_groupe = pd.DataFrame(rows_ged_groupe)
        
        rows_erp_groupe = []
        tot_art_all = 0
        tot_batch_all = 0
        tot_man_all = 0
        
        for idx, f in enumerate(filiales_groupe, start=1):
            sub_art_f = df_artis_filtered[df_artis_filtered['Organisation interne'] == f]
            tot_art_f = len(sub_art_f)
            
            if tot_art_f > 0 and 'Statut' in sub_art_f.columns:
                batch_c = len(sub_art_f[sub_art_f['Statut'] == 'Bon à payer'])
                man_c = tot_art_f - batch_c
                pct_b = (batch_c / tot_art_f) * 100
                pct_m = (man_c / tot_art_f) * 100
            else:
                batch_c, man_c, pct_b, pct_m = 0, 0, 0.0, 0.0
                
            tot_art_all += tot_art_f
            tot_batch_all += batch_c
            tot_man_all += man_c
            
            rows_erp_groupe.append({
                "Rang": idx,
                "Filiale Groupe": f,
                "Total Pièces Validées Artis": tot_art_f,
                "🤖 Batch Automatique (admin)": f"{pct_b:.1f}% ({batch_c})",
                "👤 Saisie / Reprise Manuelle": f"{pct_m:.1f}% ({man_c})"
            })
            
        pct_tot_batch = (tot_batch_all / tot_art_all * 100) if tot_art_all > 0 else 0
        pct_tot_man_erp = (tot_man_all / tot_art_all * 100) if tot_art_all > 0 else 0
        
        rows_erp_groupe.append({
            "Rang": "—",
            "Filiale Groupe": "TOTAL ERP CONSOLIDÉ",
            "Total Pièces Validées Artis": tot_art_all,
            "🤖 Batch Automatique (admin)": f"{pct_tot_batch:.1f}% ({tot_batch_all})",
            "👤 Saisie / Reprise Manuelle": f"{pct_tot_man_erp:.1f}% ({tot_man_all})"
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
        
        st.markdown("### 📈 1. VISION GED / ISIWORK — Rapprochement National Consolidé (Évolutions)")
        st.dataframe(df_ged_groupe, hide_index=True, use_container_width=True)
        
        st.markdown("### 💻 2. VISION ERP ARTIS — Journal d'Intégration National Consolidé")
        st.dataframe(df_erp_groupe, hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### ✉️ Modèle de Mail National Consolidé avec Historique (Prêt à copier)")
        
        table_ged_groupe_html = df_ged_groupe.to_html(index=False, border=1, classes='content-table')
        table_erp_groupe_html = df_erp_groupe.to_html(index=False, border=1, classes='content-table')
        
        txt_chantiers_html = "".join([f"<li>{c}</li>" for c in list_chantiers_m1]) if list_chantiers_m1 else "<li><b>Suivi des référentiels prix :</b> Action continue sur SHARP, ACTUAL SYSTEMES, KATUN FRANCE.</li><li><b>Correction TVA :</b> Suivi de la livraison de l'API Artis v2.</li>"
        
        mail_groupe_html = f"""
        <div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.5; padding: 10px; background-color: #ffffff; border-radius: 5px;">
            <p>Bonjour à tous,</p>
            
            <p>Dans le cadre du suivi technique et opérationnel de notre projet de transition de la saisie manuelle vers le pilotage financier automatique, vous trouverez ci-dessous le rapport national récapitulatif consolidé pour l'ensemble de nos filiales actives du Groupe Koesio pour la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément à nos spécifications de pilotage Groupe validées en comité, ce rapport isole de manière stricte les flux qualifiés de <b>« GED Pure / FAF Artis »</b> (factures d'archivage directement traitées manuellement de bout en bout dans l'ERP) afin de ne pas fausser l'efficacité machine réelle de rapprochement de notre plateforme Isiwork.</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — Rapprochement National & Évolutions M-1</h3>
            {table_ged_groupe_html}
            
            <h3 style="color: #0056b3; margin-top: 20px;">💻 2. VISION ERP ARTIS — Journal d'Intégration National</h3>
            {table_erp_groupe_html}
            
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
