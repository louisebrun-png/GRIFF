# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle, historique M-1, suivi PA et génération de mails
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
# 5. HISTORIQUE M-1 PAR DÉFAUT ET EXTRACTION DEPUIS LE MAIL PDF
# ------------------------------------------------------------------------------
HISTORIQUE_DEFAULT_M1 = {
    "KOESIO CENTRE EST": 48.5,
    "KOESIO EST": 40.3,
    "KOESIO NORD OUEST": 34.1,
    "KOESIO NORD-OUEST": 34.1,
    "KOESIO ILE DE FRANCE": 36.8,
    "KOESIO IDF": 36.8,
    "KOESIO OCCITANIE": 36.9,
    "EOLE": 34.2,
    "2IT SOLUTIONS": 8.8,
    "KOESIO PACA": 39.2,
    "KOESIO MEDITERRANEE": 25.4,
    "TOTAL GROUPE CONSOLIDÉ": 36.6
}

@st.cache_data
def parse_pdf_history_full(pdf_file):
    """Extrait l'historique M-1 (taux auto filiale / groupe, mois précédent et chantiers) depuis le PDF."""
    hist_data = HISTORIQUE_DEFAULT_M1.copy()
    hist_month_label = "M-1"
    chantiers = [
        "Stabilité opérationnelle de l'ERP : L'injection finale en comptabilité confirme la performance du batch automatique sans intervention utilisateur.",
        "Levier \"Frais Annexes\" : Les Faux Manuels s'expliquent par l'absence des lignes de frais de port ou DEEE sur la commande d'achat ou le BL initial dans Artis.",
        "Alerte Référentiels Prix : Suivi des mises à jour tarifaires en amont sur les partenaires stratégiques."
    ]
    if not pdf_file or not HAS_PYPDF:
        return hist_data, hist_month_label, chantiers

    try:
        reader = pypdf.PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += (page.extract_text() or "") + "\n"

        if not full_text.strip():
            # PDF scanné (image) : on tente de déduire le nom de filiale/mois depuis le nom du fichier
            fname_upper = getattr(pdf_file, "name", "").upper()
            if "IDF" in fname_upper or "ILE DE FRANCE" in fname_upper:
                hist_month_label = "Septembre 2026"
            return hist_data, hist_month_label, chantiers

        # 1. Détection format Rapport Groupe (lignes par filiale : KOESIO XXX VolBrut GedPure VolEpure XX,X% (YYY))
        pattern_ged_groupe = re.findall(
            r'KOESIO\s*([A-Za-zÀ-ÿ\s\-\diT]+?)\s*(\d[\d\s]*)\s+(\d+)\s+(\d+)\s+([\d.,]+)\s*%\s*\(([\d\s]+)\)',
            full_text
        )
        for match in pattern_ged_groupe:
            name = match[0].strip().upper()
            auto_pct = float(match[4].replace(',', '.'))
            hist_data[f"KOESIO {name}"] = auto_pct

        # 2. Détection format Mail Filiale Individuel (ex: "Septembre 2026 | 356 | 11.0% (44) | 36.8% (131)")
        filiale_pdf = None
        m_fil = re.search(r'(KOESIO\s+[A-ZÀ-Ÿ\s\-]+?|EOLE|2IT\s+SOLUTIONS)(?:\s*[-–]|\s+sur\s+la\s+période|\n)', full_text)
        if m_fil:
            filiale_pdf = m_fil.group(1).strip().upper()
            if filiale_pdf == "KOESIO IDF":
                filiale_pdf = "KOESIO ILE DE FRANCE"

        pattern_mois_filiale = re.findall(
            r'(Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Septembre|Octobre|Novembre|Décembre)\s+(20\d{2}).*?([\d.,]+)\s*%\s*\(\d+\)\s+([\d.,]+)\s*%\s*\(\d+\)',
            full_text,
            flags=re.IGNORECASE | re.DOTALL
        )
        if pattern_mois_filiale:
            last_mois, last_year, _, auto_pct_str = pattern_mois_filiale[-1]
            hist_month_label = f"{last_mois.capitalize()} {last_year}"
            auto_val = float(auto_pct_str.replace(',', '.'))
            if filiale_pdf:
                hist_data[filiale_pdf] = auto_val
            else:
                hist_data["FILIALE_PDF"] = auto_val

        return hist_data, hist_month_label, chantiers
    except Exception:
        return hist_data, hist_month_label, chantiers

dict_hist_m1, hist_month_m1_label, list_chantiers_m1 = parse_pdf_history_full(file_history_input)

def get_hist_auto_for_filiale(filiale_name, compare_dict):
    """Recherche le taux M-1 correspondant à une filiale donnée."""
    if not compare_dict or not filiale_name:
        return None
    f_up = filiale_name.upper().strip()
    aliases = {
        "KOESIO ILE DE FRANCE": ["KOESIO ILE DE FRANCE", "KOESIO IDF"],
        "KOESIO NORD OUEST": ["KOESIO NORD OUEST", "KOESIO NORD-OUEST"],
        "KOESIO NORD-OUEST": ["KOESIO NORD-OUEST", "KOESIO NORD OUEST"],
    }
    candidates = aliases.get(f_up, [f_up])
    for cand in candidates:
        if cand in compare_dict:
            return compare_dict[cand]
    for key_h, val_h in compare_dict.items():
        if f_up in key_h.upper() or key_h.upper() in f_up:
            return val_h
    return compare_dict.get("FILIALE_PDF", None)

def format_tendance_auto(pct_actuel, pct_m1, label_m1="M-1"):
    """Formate le texte explicite de tendance (augmente / descend / stabilise) pour le mail filiale."""
    if pct_m1 is None:
        return f"Votre taux de rapprochement automatique s'établit à <b>{pct_actuel:.1f}%</b> ce mois-ci."
    diff = round(pct_actuel - pct_m1, 1)
    if diff > 0.5:
        return (
            f"📈 <b>En hausse (+{diff:.1f} pts vs {label_m1}) :</b> Votre taux de rapprochement automatique "
            f"<b>augmente</b>, passant de <b>{pct_m1:.1f}%</b> à <b>{pct_actuel:.1f}%</b>."
        )
    elif diff < -0.5:
        return (
            f"📉 <b>En baisse ({diff:.1f} pts vs {label_m1}) :</b> Votre taux de rapprochement automatique "
            f"<b>descend</b>, passant de <b>{pct_m1:.1f}%</b> à <b>{pct_actuel:.1f}%</b>."
        )
    else:
        return (
            f"➡️ <b>Stable ({diff:+.1f} pts vs {label_m1}) :</b> Votre taux de rapprochement automatique "
            f"<b>se stabilise</b> autour de <b>{pct_actuel:.1f}%</b> (contre <b>{pct_m1:.1f}%</b> en {label_m1})."
        )

# ------------------------------------------------------------------------------
# 6. CHARGEMENT ET DÉTECTION ROBUSTE DES FICHIERS PAR SIGNATURE ZEENDOC
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
                
                # Filtrage des lignes vides éventuelles en pied d'export Zeendoc
                if 'Identifiant' in df_isi.columns:
                    df_isi = df_isi[df_isi['Identifiant'].notna()].copy()
                else:
                    df_isi = df_isi.dropna(how='all').copy()
                    
                df_isi['Organisation_Zeendoc'] = entite_detectee
                isi_list.append(df_isi)
            else:
                if file.name.endswith('.xlsx'):
                    df_art = pd.read_excel(file)
                else:
                    df_art = pd.read_csv(file)
                
                df_art = df_art.dropna(how='all').copy()
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
def detect_filiale(df_art, df_isi):
    if not df_art.empty and 'Organisation interne' in df_art.columns:
        filiales = df_art['Organisation interne'].dropna().unique().tolist()
        if filiales:
            return ", ".join(map(str, filiales))
    if not df_isi.empty and 'Organisation_Zeendoc' in df_isi.columns:
        orgs = df_isi['Organisation_Zeendoc'].dropna().unique().tolist()
        if orgs:
            return ", ".join(map(str, orgs))
    return "En attente de fichier..."

filiale_detectee = detect_filiale(df_artis, df_isiwork)

# ------------------------------------------------------------------------------
# 8. FILTRAGE STRICT DES MOIS CLÔTURÉS (EXCLUSION DU MOIS EN COURS)
# ------------------------------------------------------------------------------
mois_fr = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

def get_closed_periods(df_art, df_isi):
    p_art, p_isi = [], []
    if not df_art.empty and 'Date' in df_art.columns:
        p_art = pd.to_datetime(df_art['Date'], errors='coerce').dt.to_period('M').dropna().unique().tolist()
    if not df_isi.empty and 'Date facture' in df_isi.columns:
        p_isi = pd.to_datetime(df_isi['Date facture'], errors='coerce').dt.to_period('M').dropna().unique().tolist()
    
    all_periods = sorted(list(set(p_art + p_isi)))
    
    current_ongoing_period = pd.Period(pd.Timestamp.now(), freq='M')
    closed_periods = [p for p in all_periods if p < current_ongoing_period]
    
    return closed_periods if closed_periods else all_periods

all_periods_p = get_closed_periods(df_artis, df_isiwork)

if all_periods_p:
    if len(all_periods_p) == 1:
        periode_globale = f"{mois_fr.get(all_periods_p[0].month, str(all_periods_p[0].month))} {all_periods_p[0].year}"
    else:
        m_start = f"{mois_fr.get(all_periods_p[0].month, str(all_periods_p[0].month))} {all_periods_p[0].year}"
        m_end = f"{mois_fr.get(all_periods_p[-1].month, str(all_periods_p[-1].month))} {all_periods_p[-1].year}"
        periode_globale = f"{m_start} et {m_end}"
else:
    periode_globale = "Mois non détecté"

if all_periods_p:
    if not df_artis.empty and 'Date' in df_artis.columns:
        df_artis_filtered = df_artis[pd.to_datetime(df_artis['Date'], errors='coerce').dt.to_period('M').isin(all_periods_p)]
    else:
        df_artis_filtered = df_artis
        
    if not df_isiwork.empty and 'Date facture' in df_isiwork.columns:
        df_isiwork_filtered = df_isiwork[pd.to_datetime(df_isiwork['Date facture'], errors='coerce').dt.to_period('M').isin(all_periods_p)]
    else:
        df_isiwork_filtered = df_isiwork
else:
    df_artis_filtered = df_artis
    df_isiwork_filtered = df_isiwork

# ------------------------------------------------------------------------------
# 8B. ANALYSE SUIVI PA (PLATEFORME AGRÉÉE / FACTURATION ÉLECTRONIQUE)
# ------------------------------------------------------------------------------
def compute_pa_metrics(df_isi):
    """Calcule le volume PA, les doublons détectés et le Top 10 des plus gros fournisseurs PA."""
    if df_isi.empty:
        return 0, 0, pd.DataFrame(columns=["Rang", "Fournisseur", "Factures PA", "Dont Doublons"])

    col_fe = "FE - Statut de traitement des factures"
    if col_fe in df_isi.columns:
        mask_pa = df_isi[col_fe].astype(str).str.contains("PA", case=False, na=False)
    elif "Déposeur du document" in df_isi.columns:
        mask_pa = df_isi["Déposeur du document"].astype(str).str.strip().str.upper() == "FE"
    else:
        return 0, 0, pd.DataFrame(columns=["Rang", "Fournisseur", "Factures PA", "Dont Doublons"])

    df_pa = df_isi[mask_pa].copy()
    nb_pa = len(df_pa)
    if nb_pa == 0:
        return 0, 0, pd.DataFrame(columns=["Rang", "Fournisseur", "Factures PA", "Dont Doublons"])

    has_dup_col = "Statut de doublon" in df_pa.columns
    nb_doublons = int(df_pa["Statut de doublon"].notna().sum()) if has_dup_col else 0

    if "Fournisseur" in df_pa.columns:
        top10_series = df_pa["Fournisseur"].fillna("Inconnu").value_counts().head(10)
        rows_top10 = []
        for idx, (fourn, cnt) in enumerate(top10_series.items(), start=1):
            dup_f = int(df_pa[df_pa["Fournisseur"].fillna("Inconnu") == fourn]["Statut de doublon"].notna().sum()) if has_dup_col else 0
            rows_top10.append({
                "Rang": idx,
                "Fournisseur": fourn,
                "Factures PA": int(cnt),
                "Dont Doublons": dup_f
            })
        df_top10_pa = pd.DataFrame(rows_top10)
    else:
        df_top10_pa = pd.DataFrame(columns=["Rang", "Fournisseur", "Factures PA", "Dont Doublons"])

    return nb_pa, nb_doublons, df_top10_pa

# Fonction d'analyse consolidée par période
def compute_consolidation_for_period(p_period, compare_dict=None, compare_label="M-1"):
    if not df_artis.empty and 'Date' in df_artis.columns:
        df_art_p = df_artis[pd.to_datetime(df_artis['Date'], errors='coerce').dt.to_period('M') == p_period]
    else:
        df_art_p = df_artis
        
    if not df_isiwork.empty and 'Date facture' in df_isiwork.columns:
        df_isi_p = df_isiwork[pd.to_datetime(df_isiwork['Date facture'], errors='coerce').dt.to_period('M') == p_period]
    else:
        df_isi_p = df_isiwork
        
    filiales_g = df_art_p['Organisation interne'].dropna().unique().tolist() if 'Organisation interne' in df_art_p.columns else []
    if not filiales_g and not df_isi_p.empty and 'Organisation_Zeendoc' in df_isi_p.columns:
        filiales_g = df_isi_p['Organisation_Zeendoc'].dropna().unique().tolist()
    
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
    
    rows_ged_raw, rows_erp_raw = [], []
    tot_vol_brut, tot_ged_pure, tot_vol_epure, tot_auto, tot_faux, tot_vrai, tot_att = 0, 0, 0, 0, 0, 0, 0
    tot_art_all, tot_batch_all, tot_man_all = 0, 0, 0
    
    dict_pct_auto_out = {}

    for f in filiales_g:
        sub_art_f = df_art_p[df_art_p['Organisation interne'] == f] if 'Organisation interne' in df_art_p.columns else pd.DataFrame()
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
            "Filiale Groupe": f, "Total": tot_art_f, "batch_c": batch_c, "pct_b": pct_b, "man_c": man_c, "pct_m": pct_m
        })
        
        if 'Organisation_Zeendoc' in df_isi_p.columns:
            termes_zeendoc = CORRESPONDANCE_ZEENDOC.get(f, [f])
            mask_isi = df_isi_p['Organisation_Zeendoc'].astype(str).str.upper().apply(
                lambda x: any(term.upper() in x for term in termes_zeendoc)
            )
            sub_isi_f = df_isi_p[mask_isi].copy()
        else:
            sub_isi_f = pd.DataFrame()
            
        vol_brut = len(sub_isi_f)
        
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
        
        dict_pct_auto_out[f] = pct_a
        
        tot_vol_brut += vol_brut
        tot_ged_pure += ged_pure
        tot_vol_epure += vol_epure
        tot_auto += auto_c
        tot_faux += faux_c
        tot_vrai += vrai_c
        tot_att += att_c
        
        rows_ged_raw.append({
            "Filiale Groupe": f, "Vol. Brut GED": vol_brut, "Flux GED Pure": ged_pure, "Vol. Épuré": vol_epure,
            "pct_a": pct_a, "auto_c": auto_c, "pct_f": pct_f, "faux_c": faux_c,
            "pct_v": pct_v, "vrai_c": vrai_c, "pct_att": pct_att, "att_c": att_c, "pct_pot": pct_pot
        })

    # TRI GED : % AUTO DESC, PUIS VOL. ÉPURÉ DESC
    df_ged_temp = pd.DataFrame(rows_ged_raw)
    if not df_ged_temp.empty:
        df_ged_temp = df_ged_temp.sort_values(by=["pct_a", "Vol. Épuré"], ascending=[False, False]).reset_index(drop=True)
    
    rows_ged_groupe = []
    for idx, row in df_ged_temp.iterrows():
        f_name = row["Filiale Groupe"]
        pct_a_val = row["pct_a"]
        auto_c_val = row["auto_c"]
        
        str_eval_auto = f"{pct_a_val:.1f}% ({auto_c_val})"
        if compare_dict:
            comp_val = get_hist_auto_for_filiale(f_name, compare_dict)
            if comp_val is not None:
                delta_a = pct_a_val - comp_val
                symb = "↗️" if delta_a > 0.5 else "↘️" if delta_a < -0.5 else "➡️"
                str_eval_auto += f" [{symb} {delta_a:+.1f}% vs {compare_label}]"

        rows_ged_groupe.append({
            "Rang": idx + 1,
            "Filiale Groupe": f_name,
            "Vol. Brut GED": row["Vol. Brut GED"],
            "Flux GED Pure": f"{row['Flux GED Pure']}",
            "Vol. Épuré": row["Vol. Épuré"],
            "🤖 Rapprochement Auto": str_eval_auto,
            "🚚 Faux Manuel (Frais)": f"{row['pct_f']:.1f}% ({row['faux_c']})",
            "🛠️ Vrai Manuel (Litiges)": f"{row['pct_v']:.1f}% ({row['vrai_c']})",
            "⏳ En attente": f"{row['pct_att']:.1f}% ({row['att_c']})",
            "🎯 Potentiel Cible": f"{row['pct_pot']:.1f}%"
        })
        
    pct_tot_a = (tot_auto / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
    pct_tot_f = (tot_faux / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
    pct_tot_v = (tot_vrai / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
    pct_tot_att = (tot_att / tot_vol_epure * 100) if tot_vol_epure > 0 else 0
    
    dict_pct_auto_out["TOTAL GROUPE CONSOLIDÉ"] = pct_tot_a
    
    str_tot_auto = f"{pct_tot_a:.1f}% ({tot_auto})"
    if compare_dict:
        tot_comp = compare_dict.get("TOTAL GROUPE CONSOLIDÉ", 36.6)
        delta_tot = pct_tot_a - tot_comp
        symb_tot = "↗️" if delta_tot > 0.5 else "↘️" if delta_tot < -0.5 else "➡️"
        str_tot_auto += f" [{symb_tot} {delta_tot:+.1f}% vs {compare_label}]"

    rows_ged_groupe.append({
        "Rang": "—",
        "Filiale Groupe": "TOTAL GROUPE CONSOLIDÉ",
        "Vol. Brut GED": tot_vol_brut,
        "Flux GED Pure": f"{tot_ged_pure}",
        "Vol. Épuré": tot_vol_epure,
        "🤖 Rapprochement Auto": str_tot_auto,
        "🚚 Faux Manuel (Frais)": f"{pct_tot_f:.1f}% ({tot_faux})",
        "🛠️ Vrai Manuel (Litiges)": f"{pct_tot_v:.1f}% ({tot_vrai})",
        "⏳ En attente": f"{pct_tot_att:.1f}% ({tot_att})",
        "🎯 Potentiel Cible": f"{(pct_tot_a + pct_tot_f):.1f}%"
    })
    
    df_ged_groupe = pd.DataFrame(rows_ged_groupe)
    
    # TRI ERP : % IMPORT DESC, PUIS TOTAL PIÈCES DESC
    df_erp_temp = pd.DataFrame(rows_erp_raw)
    if not df_erp_temp.empty:
        df_erp_temp = df_erp_temp.sort_values(by=["pct_b", "Total"], ascending=[False, False]).reset_index(drop=True)
    
    rows_erp_groupe = []
    for idx, row in df_erp_temp.iterrows():
        rows_erp_groupe.append({
            "Rang": idx + 1,
            "Filiale Groupe": row["Filiale Groupe"],
            "Total Pièces Validées Artis": row["Total"],
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
    
    return df_ged_groupe, df_erp_groupe, dict_pct_auto_out

# ------------------------------------------------------------------------------
# 9. EXTRACTION DYNAMIQUE TOPS FOURNISSEURS (VRAIS MANUELS, FAUX MANUELS & PA)
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
    txt_top_vrai_html = "<li>Aucun Vrai Manuel bloquant détecté</li>"

if not txt_top_faux_html:
    txt_top_faux_html = "<li>Aucun Faux Manuel détecté</li>"

nb_factures_pa, nb_doublons_pa, df_top10_pa = compute_pa_metrics(df_isiwork_filtered)

# ------------------------------------------------------------------------------
# 10. STYLES CSS COMMUNS POUR LES MAILS HTML
# ------------------------------------------------------------------------------
MAIL_TABLE_CSS = """
<style>
    .content-table {
        border-collapse: collapse;
        margin: 10px 0;
        font-size: 11px;
        min-width: 100%;
        border: 1px solid #dddddd;
    }
    .content-table th {
        background-color: #0056b3;
        color: #ffffff;
        text-align: left;
        padding: 6px 8px;
    }
    .content-table td {
        padding: 6px 8px;
        border: 1px solid #dddddd;
    }
    .content-table tr:nth-of-type(even) {
        background-color: #f9f9f9;
    }
</style>
"""

# ------------------------------------------------------------------------------
# 11. ONGLETS DE NAVIGATION
# ------------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🏢 Vision Filiale (Rapport & Mail)", "🌍 Consolidation Groupe"])

# ==============================================================================
# ONGLET 1 : VISION FILIALE (AVEC MAIL + TENDANCE HISTORIQUE + SUIVI PA)
# ==============================================================================
with tab1:
    st.subheader(f"📊 Bilan de Production — {filiale_detectee}")
    
    if not df_isiwork.empty or not df_artis.empty:
        st.success(f"📅 Mois clôturés analysés : **{', '.join([f'{mois_fr.get(p.month, str(p.month))} {p.year}' for p in all_periods_p])}**")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Factures Artis", f"{len(df_artis_filtered):,}".replace(',', ' '))
        with c2:
            st.metric("Total Factures Isiwork", f"{len(df_isiwork_filtered):,}".replace(',', ' '))
        with c3:
            st.metric("Factures issues de PA", f"{nb_factures_pa:,}".replace(',', ' '), delta=f"{nb_doublons_pa} doublons" if nb_doublons_pa > 0 else "0 doublon", delta_color="inverse" if nb_doublons_pa > 0 else "off")
        with c4:
            st.metric("Filiale Détectée", filiale_detectee)
            
        st.markdown("---")
        
        dict_prev_filiale = dict_hist_m1
        label_prev_filiale = hist_month_m1_label
        computed_filiale_months = []
        
        def clean_filiale_df(df_table):
            """Supprime la ligne TOTAL redondante lorsqu'il n'y a qu'une seule filiale analysée."""
            if df_table is None or df_table.empty:
                return df_table
            # Si 2 lignes (1 filiale + 1 ligne TOTAL), on ne garde que la ligne de la filiale
            if len(df_table) == 2 and "Filiale Groupe" in df_table.columns:
                df_one = df_table[~df_table["Filiale Groupe"].astype(str).str.startswith("TOTAL")].copy()
                if "Rang" in df_one.columns:
                    df_one = df_one.drop(columns=["Rang"])
                return df_one
            return df_table

        for p in all_periods_p:
            lbl_m = f"{mois_fr.get(p.month, str(p.month))} {p.year}"
            df_g_m, df_e_m, dict_auto_m = compute_consolidation_for_period(p, compare_dict=dict_prev_filiale, compare_label=label_prev_filiale)
            df_g_m = clean_filiale_df(df_g_m)
            df_e_m = clean_filiale_df(df_e_m)
            
            # Récupération du taux auto actuel et précédent de la filiale
            filiale_cle = filiale_detectee.split(",")[0].strip()
            pct_auto_m = dict_auto_m.get(filiale_cle, dict_auto_m.get("TOTAL GROUPE CONSOLIDÉ", 0.0))
            pct_auto_prev = get_hist_auto_for_filiale(filiale_cle, dict_prev_filiale)
            tendance_txt = format_tendance_auto(pct_auto_m, pct_auto_prev, label_prev_filiale)
            
            computed_filiale_months.append({
                "label": lbl_m,
                "df_ged": df_g_m,
                "df_erp": df_e_m,
                "pct_auto": pct_auto_m,
                "pct_prev": pct_auto_prev,
                "label_prev": label_prev_filiale,
                "tendance_html": tendance_txt
            })
            
            dict_prev_filiale = dict_auto_m
            label_prev_filiale = lbl_m

        # Affichage des tableaux à l'écran
        if len(computed_filiale_months) == 2:
            m_a, m_b = computed_filiale_months[0], computed_filiale_months[1]
            
            st.markdown("### 📈 1. VISION GED / ISIWORK (Comparatif Côte à Côte)")
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.caption(f"**{m_a['label']}**")
                st.dataframe(m_a['df_ged'], hide_index=True, use_container_width=True)
            with col_g2:
                st.caption(f"**{m_b['label']}**")
                st.dataframe(m_b['df_ged'], hide_index=True, use_container_width=True)
                
            st.markdown("### 💻 2. VISION ERP ARTIS (Comparatif Côte à Côte)")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                st.caption(f"**{m_a['label']}**")
                st.dataframe(m_a['df_erp'], hide_index=True, use_container_width=True)
            with col_e2:
                st.caption(f"**{m_b['label']}**")
                st.dataframe(m_b['df_erp'], hide_index=True, use_container_width=True)
                
            t_ged_f1 = m_a['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_ged_f2 = m_b['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_erp_f1 = m_a['df_erp'].to_html(index=False, border=1, classes='content-table')
            t_erp_f2 = m_b['df_erp'].to_html(index=False, border=1, classes='content-table')
            
            tables_filiale_mail_html = f"""
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK ({m_a['label']} vs {m_b['label']})</h3>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" style="table-layout: fixed;">
                <tr>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m_a['label']}</h4>
                        {t_ged_f1}
                    </td>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m_b['label']}</h4>
                        {t_ged_f2}
                    </td>
                </tr>
            </table>
            <p style="margin-top: 8px;">{m_b['tendance_html']}</p>
            
            <h3 style="color: #0056b3; margin-top: 25px;">💻 2. VISION ERP ARTIS ({m_a['label']} vs {m_b['label']})</h3>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" style="table-layout: fixed;">
                <tr>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m_a['label']}</h4>
                        {t_erp_f1}
                    </td>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m_b['label']}</h4>
                        {t_erp_f2}
                    </td>
                </tr>
            </table>
            """
        else:
            m_single = computed_filiale_months[0]
            st.markdown(f"### 📈 1. VISION GED / ISIWORK — {m_single['label']}")
            st.dataframe(m_single['df_ged'], hide_index=True, use_container_width=True)
            st.info(m_single['tendance_html'].replace("<b>", "**").replace("</b>", "**"))
            
            st.markdown(f"### 💻 2. VISION ERP ARTIS — {m_single['label']}")
            st.dataframe(m_single['df_erp'], hide_index=True, use_container_width=True)
            
            t_ged_single = m_single['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_erp_single = m_single['df_erp'].to_html(index=False, border=1, classes='content-table')
            
            tables_filiale_mail_html = f"""
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — {m_single['label']}</h3>
            {t_ged_single}
            <p style="margin-top: 8px;">{m_single['tendance_html']}</p>
            
            <h3 style="color: #0056b3; margin-top: 25px;">💻 2. VISION ERP ARTIS — {m_single['label']}</h3>
            {t_erp_single}
            """

        # Affichage à l'écran de la nouvelle section 5 - Suivi PA
        st.markdown("---")
        st.markdown(f"### 📨 5. SUIVI PA (Plateforme Agréée) — {periode_globale}")
        if nb_doublons_pa > 0:
            st.warning(f"⚠️ **Attention à vos doublons :** Vous avez désormais **{nb_factures_pa} factures issues de la PA**, dont **{nb_doublons_pa} doublons détectés** (`Doublon sur index`).")
        else:
            st.info(f"ℹ️ Vous avez désormais **{nb_factures_pa} factures issues de la PA** ce mois-ci (aucun doublon détecté).")
            
        if not df_top10_pa.empty:
            st.caption("**Top 10 des plus gros fournisseurs issus de la PA :**")
            st.dataframe(df_top10_pa, hide_index=True, use_container_width=True)
            top10_pa_list_html = "".join([
                f"<li><b>{row['Fournisseur']}</b> : {row['Factures PA']} factures" + (f" <i>(dont {row['Dont Doublons']} doublons)</i>" if row['Dont Doublons'] > 0 else "") + "</li>"
                for _, row in df_top10_pa.iterrows()
            ])
        else:
            top10_pa_list_html = "<li>Aucune facture issue de la PA sur cette période.</li>"

        alerte_pa_html = (
            f"⚠️ <b>Attention à vos doublons :</b> vous avez désormais <b>{nb_factures_pa} factures issues de la PA</b> "
            f"(dont <b>{nb_doublons_pa} doublon(s) détecté(s)</b>)."
            if nb_doublons_pa > 0
            else f"Vous avez désormais <b>{nb_factures_pa} factures issues de la PA</b> sur cette période."
        )

        # Rétablissement du Mail Filiale complet (sections 1 à 5)
        st.markdown("---")
        st.markdown(f"### ✉️ Modèle de Mail Filiale — {filiale_detectee} (Prêt à copier pour Outlook)")
        
        txt_chantiers_filiale_html = "".join([f"<li>{c}</li>" for c in list_chantiers_m1])
        
        mail_filiale_html = f"""
        {MAIL_TABLE_CSS}
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #222; line-height: 1.5; padding: 14px; background-color: #ffffff; border-radius: 6px; border: 1px solid #e0e0e0;">
            <p><b>Objet :</b> Analyse de Production {filiale_detectee} - Statistiques et Performance de Rapprochement - {periode_globale}</p>
            <hr style="border: 0; border-top: 1px solid #eee; margin: 10px 0;">
            
            <p>Bonjour à tous,</p>
            
            <p>Vous trouverez ci-dessous l'analyse détaillée des performances du flux de rapprochement automatique des factures fournisseurs pour <b>{filiale_detectee}</b>, consolidée sur la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément aux directives du Groupe, nous appliquons la méthodologie d'isolation des flux <b>« GED Pure / FAF Artis »</b> (factures d'archivage directement traitées de bout en bout dans l'ERP sans rapprochement en GED) afin de refléter l'efficacité machine réelle d'Isiwork.</p>
            
            {tables_filiale_mail_html}
            
            <p style="font-size: 11px; color: #555; margin-top: 8px;">
                <i>Légende des états : <b>Rapprochement Auto</b> = Factures rapprochées à 100 % automatiquement ; <b>Faux Manuel</b> = Marchandises conformes mais forçage manuel pour frais annexes (port, DEEE) ; <b>Vrai Manuel</b> = Écarts de prix ou quantités ; <b>Flux GED Pure</b> = Pièces saisies directement dans l'ERP.</i>
            </p>
            
            <h3 style="color: #0056b3; margin-top: 22px;">🔍 3. FOCUS ACTIONS — Analyse des "Vrais Manuels" & Frictions</h3>
            <p><b>• Top 5 Fournisseurs bloquants (Écarts prix/réf — Vrais Manuels) :</b></p>
            <ul>
                {txt_top_vrai_html}
            </ul>
            <p><b>• Top 5 Faux Manuels (Frais de port et frais annexes isolés) :</b></p>
            <ul>
                {txt_top_faux_html}
            </ul>
            
            <h3 style="color: #0056b3; margin-top: 22px;">⚙️ 4. SUIVI DES CHANTIERS TECHNIQUES & PERSPECTIVES</h3>
            <ul>
                {txt_chantiers_filiale_html}
            </ul>
            
            <h3 style="color: #0056b3; margin-top: 22px;">📨 5. SUIVI PA (Facturation Électronique)</h3>
            <p>{alerte_pa_html}</p>
            <p><b>• Top 10 des plus gros fournisseurs issus de la PA :</b></p>
            <ol>
                {top10_pa_list_html}
            </ol>
            
            <p style="margin-top: 16px;">Cordialement,</p>
            <p><b>Louise BRUN</b><br>
            Chef de projet pour KOESIO GROUPE & KAM<br>
            Potentiale</p>
        </div>
        """
        
        st.components.v1.html(mail_filiale_html, height=1350, scrolling=True)
            
    else:
        st.warning("👈 Veuillez déposer vos fichiers d'export **Isiwork/Zeendoc** et **Artis** dans le panneau latéral.")

# ==============================================================================
# ONGLET 2 : CONSOLIDATION GROUPE
# ==============================================================================
with tab2:
    st.subheader("🌍 Consolidation Inter-Filiales — Groupe Koesio")
    
    if all_periods_p:
        dict_prev_month = dict_hist_m1
        label_prev_month = hist_month_m1_label
        
        computed_months = []
        
        for p in all_periods_p:
            lbl_m = f"{mois_fr.get(p.month, str(p.month))} {p.year}"
            df_g_p, df_e_p, dict_auto_p = compute_consolidation_for_period(p, compare_dict=dict_prev_month, compare_label=label_prev_month)
            
            computed_months.append({
                "label": lbl_m,
                "df_ged": df_g_p,
                "df_erp": df_e_p
            })
            
            # Chaînage pour le mois suivant
            dict_prev_month = dict_auto_p
            label_prev_month = lbl_m
            
        # AFFICHAGE CÔTE À CÔTE SI 2 MOIS
        if len(computed_months) == 2:
            m1, m2 = computed_months[0], computed_months[1]
            
            st.markdown("### 📈 1. VISION GED / ISIWORK — Rapprochement National Consolidé")
            c_ged1, c_ged2 = st.columns(2)
            with c_ged1:
                st.markdown(f"#### 📅 {m1['label']}")
                st.dataframe(m1['df_ged'], hide_index=True, use_container_width=True)
            with c_ged2:
                st.markdown(f"#### 📅 {m2['label']}")
                st.dataframe(m2['df_ged'], hide_index=True, use_container_width=True)
                
            st.markdown("---")
            
            st.markdown("### 💻 2. VISION ERP ARTIS — Journal d'Intégration National Consolidé")
            c_erp1, c_erp2 = st.columns(2)
            with c_erp1:
                st.markdown(f"#### 📅 {m1['label']}")
                st.dataframe(m1['df_erp'], hide_index=True, use_container_width=True)
            with c_erp2:
                st.markdown(f"#### 📅 {m2['label']}")
                st.dataframe(m2['df_erp'], hide_index=True, use_container_width=True)
                
            st.caption("ℹ️ **Note sur l'intégration ERP** : Certaines factures d'achat (abonnements, sous-traitance) peuvent se retrouver en saisie directe dans ARTIS, mais l'essentiel du flux doit transiter par Isiwork (ex. cas connu chez EOLE : TVA à 0%).")
            
            # MAIL HTML CÔTE À CÔTE
            t_ged_html1 = m1['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_ged_html2 = m2['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_erp_html1 = m1['df_erp'].to_html(index=False, border=1, classes='content-table')
            t_erp_html2 = m2['df_erp'].to_html(index=False, border=1, classes='content-table')
            
            mail_side_by_side_html = f"""
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — Rapprochement National ({m1['label']} vs {m2['label']})</h3>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" style="table-layout: fixed;">
                <tr>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m1['label']}</h4>
                        {t_ged_html1}
                    </td>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m2['label']}</h4>
                        {t_ged_html2}
                    </td>
                </tr>
            </table>
            
            <h3 style="color: #0056b3; margin-top: 25px;">💻 2. VISION ERP ARTIS — Journal d'Intégration National ({m1['label']} vs {m2['label']})</h3>
            <table width="100%" border="0" cellspacing="10" cellpadding="0" style="table-layout: fixed;">
                <tr>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m1['label']}</h4>
                        {t_erp_html1}
                    </td>
                    <td valign="top" width="50%">
                        <h4 style="color: #333; margin-bottom: 5px;">📅 {m2['label']}</h4>
                        {t_erp_html2}
                    </td>
                </tr>
            </table>
            """
        else:
            # Cas 1 seul mois
            m = computed_months[0]
            st.markdown(f"### 📈 1. VISION GED / ISIWORK — {m['label']}")
            st.dataframe(m['df_ged'], hide_index=True, use_container_width=True)
            st.markdown(f"### 💻 2. VISION ERP ARTIS — {m['label']}")
            st.dataframe(m['df_erp'], hide_index=True, use_container_width=True)
            
            t_ged_html = m['df_ged'].to_html(index=False, border=1, classes='content-table')
            t_erp_html = m['df_erp'].to_html(index=False, border=1, classes='content-table')
            
            mail_side_by_side_html = f"""
            <h3 style="color: #0056b3; margin-top: 20px;">📈 1. VISION GED / ISIWORK — {m['label']}</h3>
            {t_ged_html}
            <h3 style="color: #0056b3; margin-top: 20px;">💻 2. VISION ERP ARTIS — {m['label']}</h3>
            {t_erp_html}
            """
            
        st.markdown("---")
        st.markdown("### ✉️ Modèle de Mail National Consolidé (Prêt à copier pour Outlook)")
        
        txt_chantiers_html = "".join([f"<li>{c}</li>" for c in list_chantiers_m1])
        
        alerte_pa_groupe_html = (
            f"⚠️ <b>Attention à vos doublons :</b> vous avez désormais <b>{nb_factures_pa} factures issues de la PA</b> "
            f"(dont <b>{nb_doublons_pa} doublon(s) détecté(s)</b>)."
            if nb_doublons_pa > 0
            else f"Vous avez désormais <b>{nb_factures_pa} factures issues de la PA</b> sur cette période."
        )
        top10_pa_groupe_html = "".join([
            f"<li><b>{row['Fournisseur']}</b> : {row['Factures PA']} factures" + (f" <i>(dont {row['Dont Doublons']} doublons)</i>" if row['Dont Doublons'] > 0 else "") + "</li>"
            for _, row in df_top10_pa.iterrows()
        ]) if not df_top10_pa.empty else "<li>Aucune facture issue de la PA sur cette période.</li>"
        
        mail_groupe_html = f"""
        {MAIL_TABLE_CSS}
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #222; line-height: 1.5; padding: 10px; background-color: #ffffff; border-radius: 5px;">
            <p>Bonjour à tous,</p>
            
            <p>Dans le cadre du suivi technique et opérationnel de notre projet de transition de la saisie manuelle vers le pilotage financier automatique, vous trouverez ci-dessous le rapport national récapitulatif comparatif pour l'ensemble de nos filiales actives du Groupe Koesio pour la période de <b>{periode_globale}</b>.</p>
            
            <p>Conformément à nos spécifications de pilotage Groupe validées en comité, ce rapport isole de manière stricte les flux qualifiés de <b>« GED Pure / FAF Artis »</b> afin de ne pas fausser l'efficacité machine réelle de rapprochement de notre plateforme Isiwork.</p>
            
            {mail_side_by_side_html}
            
            <p style="font-size: 11px; color: #555; margin-top: 10px;">
                <i>* Note : Certaines factures d'achat (abonnements, sous-traitance) peuvent se retrouver en saisie directe dans ARTIS, mais le flux doit être majoritairement issu d'Isiwork. Si ce n'est pas le cas, une analyse des pièces saisies en direct s'impose (ex. cas connu chez EOLE : TVA à 0%).</i>
            </p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">🔍 3. FOCUS ACTIONS — Enseignements Clés du Groupe</h3>
            <p><b>• Alerte Référentiels Prix :</b> L'analyse des blocages « Vrais Manuels » montre que la charge de traitement réside principalement sur un problème de mise à jour tarifaire en amont sur nos partenaires stratégiques (TD SYNNEX, ACTUAL SYSTEMES, CANON FRANCE, SHARP).</p>
            
            <h3 style="color: #0056b3; margin-top: 20px;">⚙️ 4. SUIVI DES CHANTIERS TECHNIQUES & PERSPECTIVES (Rappel M-1 & Avancées)</h3>
            <ul>
                {txt_chantiers_html}
            </ul>
            
            <h3 style="color: #0056b3; margin-top: 20px;">📨 5. SUIVI PA (Facturation Électronique)</h3>
            <p>{alerte_pa_groupe_html}</p>
            <p><b>• Top 10 des plus gros fournisseurs issus de la PA :</b></p>
            <ol>
                {top10_pa_groupe_html}
            </ol>
            
            <p>Cordialement,</p>
            <p><b>Louise BRUN</b><br>
            Chef de projet pour KOESIO GROUPE & KAM<br>
            Potentiale</p>
        </div>
        """
        
        st.components.v1.html(mail_groupe_html, height=1450, scrolling=True)
        
    else:
        st.warning("👈 Téléversez les fichiers de vos différentes filiales dans la barre latérale pour afficher la consolidation groupe.")
