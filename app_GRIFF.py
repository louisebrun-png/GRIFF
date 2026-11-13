# ==============================================================================
# APPLICATION : GRIFF — Tour de Contrôle Koesio
# DESCRIPTION : Analyse mensuelle, historique M-1 (PDF), suivi PA et génération de mails
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

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import pytesseract
    from pdf2image import convert_from_bytes
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

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
    help="Téléversez le rapport PDF du mois précédent (Groupe ou Filiale) pour extraire l'historique M-1."
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
# 5. RÉFÉRENTIEL EXACT DES CORRESPONDANCES FILIALES (ZEENDOC & ARTIS)
# ------------------------------------------------------------------------------
CORRESPONDANCE_ZEENDOC = {
    "KOESIO NORD OUEST": ["KOESIO NORD-OUEST", "KOESIO NORD OUEST", "NORD OUEST", "NORD-OUEST", "KNO"],
    "KOESIO OUEST": ["KOESIO OUEST", "OUEST"],
    "KOESIO ILE DE FRANCE": ["KOESIO IDF", "KOESIO|Rappro Artis IDF", "KOESIO ILE DE FRANCE", "KOESIO ÎLE-DE-FRANCE", "ILE DE FRANCE", "ÎLE-DE-FRANCE", "IDF"],
    "KOESIO AQUITAINE": ["JEAPI Aquitaine", "KOESIO AQUITAINE", "AQUITAINE", "EX GTO"],
    "KOESIO OCCITANIE": ["BCMP", "KOESIO OCCITANIE", "OCCITANIE", "OCC"],
    "KOESIO CENTRE EST": ["Esus", "KOESIO CENTRE EST", "CENTRE EST", "CEE"],
    "2iT solutions": ["RHONE ALPES REPRO INFORMATIQUE", "RHÔNE ALPES REPRO INFORMATIQUE", "2iT solutions", "2IT SOLUTIONS", "KOESIO 2IT SOLUTIONS", "2IT"],
    "KOESIO MEDITERRANEE": ["KOESIO MED", "KOESIO MEDITERRANEE", "KOESIO MÉDITERRANÉE", "MEDITERRANEE", "MÉDITERRANÉE"],
    "KOESIO PACA": ["KOESIO PACA", "PACA"],
    "KOESIO EST": ["KOESIO EST"],
    "EOLE": ["KOESIO EOLE", "EOLE"]
}

def canonicalize_filiale(raw_name):
    """Mappe tout libellé (Artis, Zeendoc, PDF) vers le nom officiel de la filiale dans CORRESPONDANCE_ZEENDOC."""
    if pd.isna(raw_name) or not str(raw_name).strip():
        return "Inconnue"
    s_clean = str(raw_name).strip()
    s_up = s_clean.upper()

    # Vérification exacte ou par alias (les libellés les plus spécifiques d'abord)
    ordered_keys = [
        "KOESIO NORD OUEST",
        "KOESIO ILE DE FRANCE",
        "KOESIO CENTRE EST",
        "KOESIO AQUITAINE",
        "KOESIO OCCITANIE",
        "KOESIO MEDITERRANEE",
        "KOESIO PACA",
        "KOESIO OUEST",
        "KOESIO EST",
        "2iT solutions",
        "EOLE"
    ]
    for canon_key in ordered_keys:
        aliases = CORRESPONDANCE_ZEENDOC.get(canon_key, [canon_key])
        for alias in aliases:
            if alias.upper() in s_up:
                return canon_key
    return s_clean

# ------------------------------------------------------------------------------
# 5B. LECTURE DYNAMIQUE DE L'HISTORIQUE M-1 DEPUIS LE MAIL PDF (AUCUNE VALEUR EN DUR)
# ------------------------------------------------------------------------------
def extract_text_from_pdf(pdf_file):
    """Extrait le texte d'un PDF via pypdf, pdfplumber, PyMuPDF ou OCR fallback."""
    if pdf_file is None:
        return "", False

    pdf_bytes = pdf_file.getvalue()
    text_out = ""

    # 1. pypdf
    if HAS_PYPDF:
        try:
            import io
            reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
            text_out = "\n".join((p.extract_text() or "") for p in reader.pages)
        except Exception:
            text_out = ""

    # 2. pdfplumber fallback
    if not text_out.strip() and HAS_PDFPLUMBER:
        try:
            import io
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                text_out = "\n".join((p.extract_text() or "") for p in pdf.pages)
        except Exception:
            pass

    # 3. PyMuPDF (fitz) fallback
    if not text_out.strip() and HAS_PYMUPDF:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            text_out = "\n".join((page.get_text("text") or "") for page in doc)
        except Exception:
            pass

    # 4. OCR Tesseract si PDF imprimé en mode image (ex: "Print to PDF" navigateur)
    is_image_only = False
    if not text_out.strip():
        is_image_only = True
        if HAS_OCR:
            try:
                images = convert_from_bytes(pdf_bytes, dpi=200)
                ocr_texts = [pytesseract.image_to_string(img, lang="fra") for img in images]
                text_out = "\n".join(ocr_texts)
            except Exception:
                pass

    return text_out, is_image_only

@st.cache_data
def parse_pdf_history_full(pdf_file):
    """
    Parse exclusivement le PDF historique téléversé (aucun taux hardcodé).
    Supporte :
      - Format Rapport National Consolidé Groupe (multi-filiales, mono ou bi-mois, ex. Septembre 2026)
      - Format Mail Filiale Individuel (historique mensuel d'une filiale)
    """
    hist_data = {}
    hist_month_label = "M-1"
    chantiers = []
    parse_status = "no_pdf"  # "no_pdf", "ok", "image_pdf_no_ocr", "empty_match"

    if not pdf_file:
        return hist_data, hist_month_label, chantiers, parse_status

    full_text, is_image_only = extract_text_from_pdf(pdf_file)

    if not full_text.strip():
        parse_status = "image_pdf_no_ocr" if is_image_only else "empty_match"
        return hist_data, hist_month_label, chantiers, parse_status

    # Normalisation des espaces multiples et retours à la ligne dans les noms de filiales coupés sur 2 lignes
    # Ex: "KOESIO ILE DE\nFRANCE" -> "KOESIO ILE DE FRANCE"
    text_norm = re.sub(r'([A-ZÀ-Ÿ])\n([A-ZÀ-Ÿ])', r'\1 \2', full_text)

    # Détection du dernier mois mentionné dans l'en-tête ou les colonnes (ex: "Septembre 2026")
    mois_list = re.findall(
        r'(Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Aout|Septembre|Octobre|Novembre|Décembre)\s+(20\d{2})',
        text_norm,
        flags=re.IGNORECASE
    )
    if mois_list:
        # On prend le mois le plus récent présent dans le rapport (avant le mail cité éventuel)
        top_section = text_norm.split("2. VISION ERP")[0] if "2. VISION ERP" in text_norm else text_norm[:3500]
        mois_top = re.findall(
            r'(Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Aout|Septembre|Octobre|Novembre|Décembre)\s+(20\d{2})',
            top_section,
            flags=re.IGNORECASE
        )
        chosen_m, chosen_y = mois_top[-1] if mois_top else mois_list[0]
        chosen_m = "Août" if chosen_m.lower() == "aout" else chosen_m.capitalize()
        hist_month_label = f"{chosen_m} {chosen_y}"

    # 1. Extraction des lignes de filiales (Format Rapport National Consolidé)
    # On isole la section "1. VISION GED" avant le mail cité antérieur (ex: "Le mar. 25 août")
    sec_ged = text_norm.split("2. VISION ERP")[0] if "2. VISION ERP" in text_norm else text_norm
    if "Le " in sec_ged and " a écrit" in sec_ged:
        sec_ged = sec_ged.split(" a écrit")[0]

    # Motif capturant chaque nom de filiale suivi de ses blocs de taux "XX.X% (NNN)"
    entity_regex = (
        r'(KOESIO\s+(?:ILE\s+DE\s+FRANCE|ÎLE-DE-FRANCE|CENTRE\s+EST|NORD\s+OUEST|NORD-OUEST|'
        r'MEDITERRANEE|MÉDITERRANÉE|OCCITANIE|AQUITAINE|OUEST|PACA|EST|EOLE)|'
        r'2iT\s*solutions|2IT\s*SOLUTIONS|EOLE|TOTAL\s+GROUPE\s+CONSOLID[EÉ])'
    )

    for m_ent in re.finditer(entity_regex, sec_ged, flags=re.IGNORECASE):
        raw_ent = re.sub(r'\s+', ' ', m_ent.group(1)).strip()
        canon_ent = "TOTAL GROUPE CONSOLIDÉ" if "TOTAL" in raw_ent.upper() else canonicalize_filiale(raw_ent)
        # Fenêtre de texte suivant le nom de la filiale sur la même ligne/bloc
        window = sec_ged[m_ent.end():m_ent.end() + 260]
        # Arrêt avant la prochaine filiale éventuelle
        next_m = re.search(entity_regex, window, flags=re.IGNORECASE)
        if next_m:
            window = window[:next_m.start()]

        # Tous les pourcentages suivis d'un volume entre parenthèses : ex "53.7% (359)" et "37.2% (163)"
        pct_matches = re.findall(r'(\d{1,3}[.,]\d)\s*%\s*\(\s*\d[\d\s]*\)', window)
        if pct_matches:
            # Si tableau comparatif 2 mois (Août et Septembre), le 1er bloc = M-2 et le 2ème bloc = M-1 (le plus récent)
            # Chaque bloc commence par le % Rapprochement Auto suivi du % Faux Manuel, % Vrai Manuel...
            # En bi-mois, le dernier taux auto de la ligne correspond au mois le plus récent (Septembre)
            if len(pct_matches) >= 6:
                # 2 mois côte à côte (4 ou 5 taux par mois -> le taux auto de M-1 est la 5ème ou 6ème occurrence)
                val_auto = float(pct_matches[5].replace(',', '.'))
            elif len(pct_matches) >= 5:
                val_auto = float(pct_matches[4].replace(',', '.'))
            else:
                val_auto = float(pct_matches[0].replace(',', '.'))
            hist_data[canon_ent] = val_auto

    # 2. Extraction Format Mail Filiale Individuel (ex: "Septembre 2026 | 356 | 11.0% (44) | 36.8% (131)")
    if not hist_data:
        filiale_pdf = None
        m_fil = re.search(entity_regex, sec_ged, flags=re.IGNORECASE)
        if m_fil and "TOTAL" not in m_fil.group(1).upper():
            filiale_pdf = canonicalize_filiale(re.sub(r'\s+', ' ', m_fil.group(1)).strip())

        pattern_mois_filiale = re.findall(
            r'(Janvier|Février|Mars|Avril|Mai|Juin|Juillet|Août|Aout|Septembre|Octobre|Novembre|Décembre)\s+(20\d{2}).*?(\d{1,3}[.,]\d)\s*%\s*\(\d+\)\s+(\d{1,3}[.,]\d)\s*%\s*\(\d+\)',
            sec_ged,
            flags=re.IGNORECASE | re.DOTALL
        )
        if pattern_mois_filiale:
            last_m, last_y, _, auto_str = pattern_mois_filiale[-1]
            last_m = "Août" if last_m.lower() == "aout" else last_m.capitalize()
            hist_month_label = f"{last_m} {last_y}"
            auto_val = float(auto_str.replace(',', '.'))
            if filiale_pdf:
                hist_data[filiale_pdf] = auto_val
            else:
                hist_data["FILIALE_PDF"] = auto_val

    # 3. Extraction dynamique de la section 4 (Suivi des chantiers techniques) depuis le PDF
    m_sec4 = re.search(r'4\.\s*SUIVI DES CHANTIERS[^\n]*\n(.*?)(?:Cordialement|Louise BRUN|https://|$)', text_norm, flags=re.IGNORECASE | re.DOTALL)
    if m_sec4:
        raw_lines = [l.strip(" •-\t") for l in m_sec4.group(1).splitlines() if l.strip()]
        chantiers = [l for l in raw_lines if len(l) > 10 and not l.startswith("http")][:5]

    parse_status = "ok" if hist_data else "empty_match"
    return hist_data, hist_month_label, chantiers, parse_status

dict_hist_m1, hist_month_m1_label, list_chantiers_m1, pdf_parse_status = parse_pdf_history_full(file_history_input)

def get_hist_auto_for_filiale(filiale_name, compare_dict):
    """Recherche le taux M-1 extrait du PDF pour une filiale donnée (sans valeur en dur)."""
    if not compare_dict or not filiale_name:
        return None
    canon = canonicalize_filiale(filiale_name)
    if canon in compare_dict:
        return compare_dict[canon]
    for key_h, val_h in compare_dict.items():
        if canon.upper() == key_h.upper():
            return val_h
    return compare_dict.get("FILIALE_PDF", None)

def format_tendance_auto(pct_actuel, pct_m1, label_m1="M-1"):
    """Indique explicitement à la filiale si son taux automatique augmente, descend ou se stabilise."""
    if pct_m1 is None:
        return f"Votre taux de rapprochement automatique s'établit à <b>{pct_actuel:.1f}%</b> ce mois-ci <i>(aucun taux historique M-1 détecté dans le PDF)</i>."
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
                
                if 'Identifiant' in df_isi.columns:
                    df_isi = df_isi[df_isi['Identifiant'].notna()].copy()
                else:
                    df_isi = df_isi.dropna(how='all').copy()
                    
                df_isi['Organisation_Zeendoc'] = entite_detectee
                df_isi['Filiale_Canonique'] = canonicalize_filiale(entite_detectee)
                isi_list.append(df_isi)
            else:
                if file.name.endswith('.xlsx'):
                    df_art = pd.read_excel(file)
                else:
                    df_art = pd.read_csv(file)
                
                df_art = df_art.dropna(how='all').copy()
                if 'Organisation interne' in df_art.columns:
                    df_art['Organisation interne'] = df_art['Organisation interne'].apply(canonicalize_filiale)
                    
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
    if not df_isi.empty and 'Filiale_Canonique' in df_isi.columns:
        orgs = df_isi['Filiale_Canonique'].dropna().unique().tolist()
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

# ------------------------------------------------------------------------------
# 8C. ANALYSE CONSOLIDÉE PAR PÉRIODE
# ------------------------------------------------------------------------------
def compute_consolidation_for_period(p_period, compare_dict=None, compare_label="M-1"):
    if not df_artis.empty and 'Date' in df_artis.columns:
        df_art_p = df_artis[pd.to_datetime(df_artis['Date'], errors='coerce').dt.to_period('M') == p_period]
    else:
        df_art_p = df_artis
        
    if not df_isiwork.empty and 'Date facture' in df_isiwork.columns:
        df_isi_p = df_isiwork[pd.to_datetime(df_isiwork['Date facture'], errors='coerce').dt.to_period('M') == p_period]
    else:
        df_isi_p = df_isiwork
        
    f_art_list = df_art_p['Organisation interne'].dropna().unique().tolist() if 'Organisation interne' in df_art_p.columns else []
    f_isi_list = df_isi_p['Filiale_Canonique'].dropna().unique().tolist() if 'Filiale_Canonique' in df_isi_p.columns else []
    filiales_g = sorted(list(set(f_art_list + f_isi_list)))
    
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
    if compare_dict and "TOTAL GROUPE CONSOLIDÉ" in compare_dict:
        tot_comp = compare_dict["TOTAL GROUPE CONSOLIDÉ"]
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

# Alerte informative dans la barre latérale sur l'état de lecture du PDF historique M-1
if file_history_input is not None:
    if pdf_parse_status == "ok":
        st.sidebar.success(f"✅ Historique M-1 extrait ({hist_month_m1_label}) : {len(dict_hist_m1)} entité(s) lue(s) dans le PDF.")
    elif pdf_parse_status == "image_pdf_no_ocr":
        st.sidebar.warning("⚠️ Ce PDF est une image scannée sans texte sélectionnable (enregistrez-le en PDF natif ou installez Tesseract OCR pour lire les chiffres).")
    else:
        st.sidebar.warning("⚠️ Aucun taux historique n'a pu être extrait de ce PDF.")

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
            st.metric(
                "Factures issues de PA",
                f"{nb_factures_pa:,}".replace(',', ' '),
                delta=f"{nb_doublons_pa} doublons" if nb_doublons_pa > 0 else "0 doublon",
                delta_color="inverse" if nb_doublons_pa > 0 else "off"
            )
        with c4:
            st.metric("Filiale Détectée", filiale_detectee)
            
        st.markdown("---")
        
        dict_prev_filiale = dict_hist_m1 if dict_hist_m1 else None
        label_prev_filiale = hist_month_m1_label
        computed_filiale_months = []
        
        def clean_filiale_df(df_table):
            """Supprime la ligne TOTAL redondante lorsqu'il n'y a qu'une seule filiale analysée."""
            if df_table is None or df_table.empty:
                return df_table
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
            st.info(m_single['tendance_html'].replace("<b>", "**").replace("</b>", "**").replace("<i>", "*").replace("</i>", "*"))
            
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

        # Mail Filiale complet (sections 1 à 5)
        st.markdown("---")
        st.markdown(f"### ✉️ Modèle de Mail Filiale — {filiale_detectee} (Prêt à copier pour Outlook)")
        
        chantiers_display = list_chantiers_m1 if list_chantiers_m1 else [
            "Stabilité opérationnelle de l'ERP et suivi du batch d'injection automatique.",
            "Résorption des Faux Manuels (frais de port et DEEE) sur les commandes et BL."
        ]
        txt_chantiers_filiale_html = "".join([f"<li>{c}</li>" for c in chantiers_display])
        
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
        dict_prev_month = dict_hist_m1 if dict_hist_m1 else None
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
        
        chantiers_groupe_display = list_chantiers_m1 if list_chantiers_m1 else [
            "Alerte Référentiels Prix : SHARP, ACTUAL SYSTEMES, KATUN FRANCE.",
            "Correctif API TVA : Suivi de la livraison globale sur Artis v2."
        ]
        txt_chantiers_html = "".join([f"<li>{c}</li>" for c in chantiers_groupe_display])
        
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
