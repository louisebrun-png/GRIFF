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
    from rapidocr_onnxruntime import RapidOCR
    HAS_RAPIDOCR = True
except ImportError:
    HAS_RAPIDOCR = False

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
st.sidebar.header("📜 Historique M-1 (Fichier Excel Maître)")
file_history_input = st.sidebar.file_uploader(
    "Fichier Historique Excel (.xlsx, .csv) ou Mail (.pdf, .eml)",
    type=["xlsx", "csv", "pdf", "eml", "txt", "html"],
    accept_multiple_files=True,
    key=f"hist_{u_key}",
    help="Téléversez votre fichier maître GRIFF_Historique_Koesio.xlsx (recommandé) ou vos mails historiques."
)
pasted_history_text = st.sidebar.text_area(
    "Ou collez le texte du mail M-1 (optionnel)",
    height=90,
    key=f"hist_txt_{u_key}",
    placeholder="Collez ici le contenu du mail M-1 si votre PDF est une image scannée..."
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
    """Mappe tout libellé (Artis, Zeendoc, PDF ou OCR) vers le nom officiel de la filiale dans CORRESPONDANCE_ZEENDOC."""
    if pd.isna(raw_name) or not str(raw_name).strip():
        return "Inconnue"
    s_clean = str(raw_name).strip()
    # Nettoyage des artefacts OCR fréquents sur les PDF imprimés depuis navigateur
    s_ocr_clean = (
        s_clean.upper()
        .replace("KOES10", "KOESIO")
        .replace("KOE$10", "KOESIO")
        .replace("KO$IO", "KOESIO")
        .replace("QUEST", "OUEST")
        .replace("217 S8LUTIONS", "2IT SOLUTIONS")
        .replace("217 SOLUTIONS", "2IT SOLUTIONS")
    )

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
            if alias.upper() in s_ocr_clean:
                return canon_key
    return s_clean

# ------------------------------------------------------------------------------
# 5B. LECTURE DYNAMIQUE DE L'HISTORIQUE M-1 DEPUIS LE MAIL PDF (AUCUNE VALEUR EN DUR)
# ------------------------------------------------------------------------------
def extract_text_from_pdf(pdf_file):
    """Extrait le texte d'un PDF via pypdf, pdfplumber, PyMuPDF ou OCR (PyMuPDF+Tesseract / RapidOCR) sur les 2 premières pages."""
    if pdf_file is None:
        return "", False

    pdf_bytes = pdf_file.getvalue()
    text_out = ""

    # 1. pypdf (2 premières pages)
    if HAS_PYPDF:
        try:
            import io
            reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
            text_out = "\n".join((p.extract_text() or "") for p in reader.pages[:2])
        except Exception:
            text_out = ""

    # 2. pdfplumber fallback
    if not text_out.strip() and HAS_PDFPLUMBER:
        try:
            import io
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                text_out = "\n".join((p.extract_text() or "") for p in pdf.pages[:2])
        except Exception:
            pass

    # 3. PyMuPDF (fitz) fallback
    if not text_out.strip() and HAS_PYMUPDF:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            text_out = "\n".join((page.get_text("text") or "") for page in doc[:2])
        except Exception:
            pass

    # 4. OCR rapide sur les 2 premières pages uniquement (économie RAM + vitesse x4 sur Streamlit Cloud)
    is_image_only = False
    if not text_out.strip():
        is_image_only = True
        # 4a. PyMuPDF + Tesseract (système Linux via packages.txt : tesseract-ocr + tesseract-ocr-fra)
        if HAS_PYMUPDF:
            try:
                import pytesseract
                from PIL import Image
                import io
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                ocr_texts = []
                for page in doc[:2]:
                    pix = page.get_pixmap(dpi=150)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    try:
                        txt_p = pytesseract.image_to_string(img, lang="fra")
                    except Exception:
                        txt_p = pytesseract.image_to_string(img)
                    ocr_texts.append(txt_p)
                text_out = "\n".join(ocr_texts)
            except Exception:
                pass
        # 4b. RapidOCR + PyMuPDF fallback
        if not text_out.strip() and HAS_RAPIDOCR and HAS_PYMUPDF:
            try:
                ocr_engine = RapidOCR()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                ocr_lines = []
                for page in doc[:2]:
                    pix = page.get_pixmap(dpi=150)
                    result, _ = ocr_engine(pix.tobytes("png"))
                    if result:
                        ocr_lines.append("\n".join([item[1] for item in result]))
                text_out = "\n".join(ocr_lines)
            except Exception:
                pass
        # 4c. pdf2image + Tesseract fallback
        if not text_out.strip() and HAS_OCR:
            try:
                images = convert_from_bytes(pdf_bytes, dpi=150, first_page=1, last_page=2)
                ocr_texts = [pytesseract.image_to_string(img) for img in images]
                text_out = "\n".join(ocr_texts)
            except Exception:
                pass

    return text_out, is_image_only

def parse_single_history_text(full_text):
    """
    Extrait les taux M-1 depuis le texte d'un mail historique (Filiale individuelle ou Groupe consolidé).
    Gère l'invariant mathématique des colonnes GED :
      - 5 pourcentages sur la ligne = [Flux GED Pure] + [Auto, Faux, Vrai, Attente] -> Auto = index 1
      - 4 pourcentages sur la ligne = [Auto, Faux, Vrai, Attente] -> Auto = index 0
    """
    local_hist = {}
    local_label = None
    local_chantiers = []

    if not full_text or not full_text.strip():
        return local_hist, local_label, local_chantiers

    text_norm = re.sub(r'([A-ZÀ-Ÿ0-9$])\n([A-ZÀ-Ÿ0-9])', r'\1 \2', full_text)

    # Si le mail contient un erratum / correctif plus bas (cas KOESIO MEDITERRANEE), on privilégie la section corrigée
    if "voici le correctif" in text_norm.lower():
        idx_corr = text_norm.lower().rfind("voici le correctif")
        text_norm = text_norm[idx_corr:]
    else:
        # Démarrer au premier "Bonjour" pour ignorer les dates d'en-tête Gmail (ex: "23 octobre 2026")
        if "Bonjour" in text_norm:
            text_norm = text_norm[text_norm.find("Bonjour"):]
        # Couper les anciens mails cités en bas de fil (ex: "Le mar. 25 août 2026 ... a écrit :")
        if " a écrit" in text_norm:
            text_norm = text_norm.split(" a écrit")[0]

    mois_to_num = {
        "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4,
        "mai": 5, "juin": 6, "juillet": 7, "août": 8, "aout": 8,
        "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12
    }
    num_to_mois = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
        7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }

    # On isole la section 1 (Vision GED) avant la section 2 (Vision ERP)
    sec_ged = re.split(r'2[.,\s]+VISION\s+ERP', text_norm, flags=re.IGNORECASE)[0]

    entity_regex = (
        r'((?:KOESIO|KOES10|KOE\$10|KO\$IO)\s+(?:ILE\s+DE\s+FRANCE|ÎLE-DE-FRANCE|IDF|CENTRE\s+EST|CEE|'
        r'NORD\s+(?:OUEST|QUEST)|NORD-(?:OUEST|QUEST)|KNO|MEDITERRANEE|MÉDITERRANÉE|MED|OCCITANIE|OCC|'
        r'AQUITAINE|AQU|(?:OUEST|QUEST)|PACA|EST|EOLE)|'
        r'(?:2iT|2IT|217)\s*(?:solutions|s8lutions)|EOLE|TOTAL\s+GROUPE\s+CONSOLID[EÉ])'
    )

    mois_found = re.findall(
        r'(Janvier|Février|Fevrier|Mars|Avril|Mai|Juin|Juillet|Août|Aout|Septembre|Octobre|Novembre|Décembre|Decembre)\s+(20\d{2})',
        sec_ged,
        flags=re.IGNORECASE
    )
    unique_periods = sorted(list({
        (int(y_s), mois_to_num[m_s.lower()])
        for m_s, y_s in mois_found
        if m_s.lower() in mois_to_num
    }))

    # 1. Format Mail Filiale Individuel (lignes mensuelles : "Septembre 2026 | ...")
    is_individual_mail = ("MOIS COMPTABLE" in sec_ged.upper()) or ("TOTAL GROUPE" not in sec_ged.upper())
    if is_individual_mail:
        filiale_pdf = None
        m_fil = re.search(entity_regex, full_text, flags=re.IGNORECASE)
        if m_fil and "TOTAL" not in m_fil.group(1).upper():
            filiale_pdf = canonicalize_filiale(re.sub(r'\s+', ' ', m_fil.group(1)).strip())

        # On repère chaque ligne commençant par un Mois + Année dans sec_ged
        month_matches = list(re.finditer(
            r'(Janvier|Février|Fevrier|Mars|Avril|Mai|Juin|Juillet|Août|Aout|Septembre|Octobre|Novembre|Décembre|Decembre)\s+(20\d{2})',
            sec_ged,
            flags=re.IGNORECASE
        ))
        best_period = (-1, -1)
        best_auto = None

        for idx_mm, mm in enumerate(month_matches):
            m_str, y_str = mm.group(1), mm.group(2)
            m_num = mois_to_num.get(m_str.lower(), 0)
            per = (int(y_str), m_num)

            end_w = month_matches[idx_mm + 1].start() if idx_mm + 1 < len(month_matches) else min(len(sec_ged), mm.end() + 220)
            row_window = sec_ged[mm.end():end_w]

            # Extraction de tous les pourcentages "XX.X%" sur la ligne mensuelle (indépendamment des retours à la ligne OCR sur "(NNN)")
            pct_tokens = [float(x.replace(',', '.')) for x in re.findall(r'(\d{1,3}[.,]\d)\s*%', row_window)]
            if len(pct_tokens) >= 2:
                # Si 5 pourcentages (Flux GED Pure en % + Auto/Faux/Vrai/Attente), Auto est à l'index 1
                # Si 4 pourcentages (Flux GED Pure en entier + Auto/Faux/Vrai/Attente), Auto est à l'index 0
                auto_val = pct_tokens[1] if len(pct_tokens) >= 5 else pct_tokens[0]
                if per >= best_period:
                    best_period = per
                    best_auto = auto_val

        if best_auto is not None and best_period[0] > 0:
            local_label = f"{num_to_mois[best_period[1]]} {best_period[0]}"
            if filiale_pdf:
                local_hist[filiale_pdf] = best_auto
            else:
                local_hist["FILIALE_PDF"] = best_auto

    # 2. Format Rapport National Consolidé Groupe (multi-filiales)
    if not local_hist:
        older_month_name = None
        if unique_periods:
            latest_y, latest_m = unique_periods[-1]
            local_label = f"{num_to_mois[latest_m]} {latest_y}"
            if len(unique_periods) >= 2:
                _, older_m = unique_periods[-2]
                older_month_name = num_to_mois[older_m]

        all_matches = list(re.finditer(entity_regex, sec_ged, flags=re.IGNORECASE))
        candidates_by_entity = {}

        for idx_m, m_ent in enumerate(all_matches):
            raw_ent = re.sub(r'\s+', ' ', m_ent.group(1)).strip()
            canon_ent = "TOTAL GROUPE CONSOLIDÉ" if "TOTAL" in raw_ent.upper() else canonicalize_filiale(raw_ent)

            end_pos = all_matches[idx_m + 1].start() if idx_m + 1 < len(all_matches) else min(len(sec_ged), m_ent.end() + 250)
            window = sec_ged[m_ent.end():end_pos]

            m_auto = re.search(r'(\d{1,3}[.,]\d)\s*%\s*\(\s*\d[\d\s]*\)', window)
            if not m_auto:
                continue

            val_auto = float(m_auto.group(1).replace(',', '.'))
            has_vs_older = bool(older_month_name and re.search(rf'vs\s+{older_month_name}', window, flags=re.IGNORECASE))
            candidates_by_entity.setdefault(canon_ent, []).append({
                "val": val_auto,
                "has_vs_older": has_vs_older
            })

        is_two_month_report = len(unique_periods) >= 2
        for canon_ent, entries in candidates_by_entity.items():
            if is_two_month_report:
                right_entries = [e for e in entries if e["has_vs_older"]]
                if right_entries:
                    local_hist[canon_ent] = right_entries[-1]["val"]
                elif len(entries) >= 2:
                    local_hist[canon_ent] = entries[-1]["val"]
            else:
                local_hist[canon_ent] = entries[-1]["val"]

    # 3. Extraction des chantiers techniques (section 4)
    m_sec4 = re.search(r'4\.\s*SUIVI DES CHANTIERS[^\n]*\n(.*?)(?:Cordialement|Louise BRUN|https://|$)', text_norm, flags=re.IGNORECASE | re.DOTALL)
    if m_sec4:
        raw_lines = [l.strip(" •-\t") for l in m_sec4.group(1).splitlines() if l.strip()]
        local_chantiers = [l for l in raw_lines if len(l) > 10 and not l.startswith("http")][:5]

    return local_hist, local_label, local_chantiers


def parse_excel_history(file_obj, target_m1_period=None):
    """Lit l'onglet Historique_Mensuel du fichier Excel et extrait UNIQUEMENT le mois M-1 (relatif au mois analysé)."""
    import io
    raw_bytes = file_obj.getvalue()
    fname = getattr(file_obj, "name", "").lower()
    if fname.endswith(".xlsx"):
        xl = pd.ExcelFile(io.BytesIO(raw_bytes))
        sheet_to_read = "Historique_Mensuel" if "Historique_Mensuel" in xl.sheet_names else xl.sheet_names[0]
        # Détection automatique de la ligne d'en-tête ("Filiale Groupe")
        df_raw = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet_to_read, header=None, nrows=10)
        hdr_idx = 0
        for r_i, row in df_raw.iterrows():
            if any("Filiale" in str(v) for v in row.values if pd.notna(v)):
                hdr_idx = r_i
                break
        df_h = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet_to_read, header=hdr_idx)
    else:
        df_h = pd.read_csv(io.BytesIO(raw_bytes))

    df_h = df_h.dropna(how="all").copy()
    if "Filiale Groupe" not in df_h.columns:
        return {}, None

    df_h = df_h[df_h["Filiale Groupe"].notna()].copy()
    df_h["Filiale_Canon"] = df_h["Filiale Groupe"].apply(canonicalize_filiale)

    # Identification de la période la plus récente dans le fichier historique
    mois_map_inv = {
        "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4,
        "mai": 5, "juin": 6, "juillet": 7, "août": 8, "aout": 8,
        "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12
    }
    num_to_mois_fr = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
        7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }

    # Construction d'une colonne normalisée YYYY-MM pour filtrer strictement M-1
    if "Période (YYYY-MM)" in df_h.columns and df_h["Période (YYYY-MM)"].notna().any():
        df_h["_per_sort"] = df_h["Période (YYYY-MM)"].astype(str).str.strip().str[:7]
    elif "Mois" in df_h.columns and "Année" in df_h.columns:
        m_num_s = df_h["Mois"].astype(str).str.strip().str.lower().map(mois_map_inv).fillna(0).astype(int)
        y_num_s = pd.to_numeric(df_h["Année"], errors="coerce").fillna(0).astype(int)
        df_h["_per_sort"] = [f"{y:04d}-{m:02d}" if (y > 0 and m > 0) else "" for y, m in zip(y_num_s, m_num_s)]
    else:
        df_h["_per_sort"] = ""

    valid_pers = sorted([p for p in df_h["_per_sort"].dropna().unique() if len(p) == 7])
    if valid_pers:
        if target_m1_period is not None:
            m1_str = str(target_m1_period)  # ex: "2026-09" si mois analysé = 2026-10
            if m1_str in valid_pers:
                chosen_per = m1_str
            else:
                # Dernier mois disponible strictement antérieur au mois analysé (<= M-1)
                earlier_pers = [p for p in valid_pers if p <= m1_str]
                chosen_per = earlier_pers[-1] if earlier_pers else valid_pers[-1]
        else:
            chosen_per = valid_pers[-1]

        df_latest = df_h[df_h["_per_sort"] == chosen_per].copy()
        try:
            y_p, m_p = chosen_per.split("-")
            lbl_out = f"{num_to_mois_fr[int(m_p)]} {int(y_p)}"
        except Exception:
            lbl_out = chosen_per
    else:
        df_latest = df_h.copy()
        lbl_out = "M-1"

    # Recherche de la colonne Taux Rapprochement Auto (%)
    col_taux_auto = None
    for c in df_latest.columns:
        if "taux" in str(c).lower() and "auto" in str(c).lower() and "artis" not in str(c).lower():
            col_taux_auto = c
            break
    if col_taux_auto is None:
        return {}, None

    out_dict = {}
    for _, r in df_latest.iterrows():
        fn = r["Filiale_Canon"]
        val = pd.to_numeric(r[col_taux_auto], errors="coerce")
        if pd.notna(val) and fn and fn != "Inconnue":
            out_dict[fn] = float(val)

    # Calcul exact du taux consolidé Groupe M-1 pondéré par Vol. Épuré GED si disponible
    if "Nb Auto GED" in df_latest.columns and "Vol. Épuré GED" in df_latest.columns:
        tot_auto_m1 = pd.to_numeric(df_latest["Nb Auto GED"], errors="coerce").sum()
        tot_epure_m1 = pd.to_numeric(df_latest["Vol. Épuré GED"], errors="coerce").sum()
        if tot_epure_m1 > 0:
            out_dict["TOTAL GROUPE CONSOLIDÉ"] = round(float(tot_auto_m1 / tot_epure_m1 * 100), 1)

    return out_dict, lbl_out


@st.cache_data
def parse_pdf_history_full(pdf_files, pasted_text="", target_m1_period=None):
    """
    Parse le fichier Excel maître GRIFF_Historique_Koesio.xlsx (prioritaire) ou un/plusieurs PDF/mails historiques.
    """
    hist_data = {}
    hist_month_label = "M-1"
    chantiers = []
    files_list = pdf_files if isinstance(pdf_files, list) else ([pdf_files] if pdf_files else [])

    if not files_list and not (pasted_text and pasted_text.strip()):
        return hist_data, hist_month_label, chantiers, "no_pdf"

    any_image_only = False
    labels_found = []

    for pf in files_list:
        fname_low = getattr(pf, "name", "").lower()
        if fname_low.endswith((".xlsx", ".csv")):
            try:
                h_xl, lbl_xl = parse_excel_history(pf, target_m1_period=target_m1_period)
                hist_data.update(h_xl)
                if lbl_xl:
                    labels_found.append(lbl_xl)
            except Exception as e:
                st.sidebar.error(f"Erreur lecture historique Excel : {e}")
            continue
        if fname_low.endswith((".eml", ".txt", ".html")):
            raw_bytes = pf.getvalue()
            try:
                txt = raw_bytes.decode("utf-8")
            except Exception:
                txt = raw_bytes.decode("latin-1", errors="ignore")
            if fname_low.endswith(".eml"):
                import email
                from bs4 import BeautifulSoup
                msg = email.message_from_string(txt)
                parts_txt = []
                for part in msg.walk():
                    ctype = part.get_content_type()
                    if ctype in ("text/plain", "text/html"):
                        payload = part.get_payload(decode=True)
                        if payload:
                            decoded = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")
                            if ctype == "text/html":
                                decoded = BeautifulSoup(decoded, "html.parser").get_text(" ")
                            parts_txt.append(decoded)
                txt = "\n".join(parts_txt) if parts_txt else txt
            is_img = False
        else:
            txt, is_img = extract_text_from_pdf(pf)
        if not txt.strip():
            if is_img:
                any_image_only = True
            continue
        h_part, lbl_part, ch_part = parse_single_history_text(txt)
        hist_data.update(h_part)
        if lbl_part:
            labels_found.append(lbl_part)
        if ch_part and not chantiers:
            chantiers = ch_part

    if pasted_text and pasted_text.strip():
        h_part, lbl_part, ch_part = parse_single_history_text(pasted_text)
        hist_data.update(h_part)
        if lbl_part:
            labels_found.append(lbl_part)
        if ch_part and not chantiers:
            chantiers = ch_part

    if labels_found:
        hist_month_label = labels_found[-1]

    if hist_data:
        parse_status = "ok"
    elif any_image_only:
        parse_status = "image_pdf_no_ocr"
    else:
        parse_status = "empty_match"

    return hist_data, hist_month_label, chantiers, parse_status

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
target_m1_p = (all_periods_p[0] - 1) if all_periods_p else None
dict_hist_m1, hist_month_m1_label, list_chantiers_m1, pdf_parse_status = parse_pdf_history_full(
    file_history_input, pasted_history_text, target_m1_period=target_m1_p
)

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
    if compare_dict:
        tot_comp = compare_dict.get("TOTAL GROUPE CONSOLIDÉ")
        if tot_comp is None:
            # Si seuls des mails individuels filiales ont été chargés, on calcule la moyenne des taux M-1 des filiales présentes
            matched_vals = [get_hist_auto_for_filiale(fn, compare_dict) for fn in filiales_g]
            matched_vals = [v for v in matched_vals if v is not None]
            if matched_vals:
                tot_comp = sum(matched_vals) / len(matched_vals)
        if tot_comp is not None:
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
if file_history_input or (pasted_history_text and pasted_history_text.strip()):
    if pdf_parse_status == "ok":
        details_str = " | ".join([f"{k}: {v:.1f}%" for k, v in dict_hist_m1.items()])
        st.sidebar.success(f"✅ Historique {hist_month_m1_label} extrait ({len(dict_hist_m1)} filiale(s)) : {details_str}")
    elif pdf_parse_status == "image_pdf_no_ocr":
        st.sidebar.warning("⚠️ PDF image détecté sans moteur OCR actif : vérifiez `requirements.txt` (`pymupdf` + `rapidocr-onnxruntime`) ou collez le texte du mail ci-dessus.")
    else:
        st.sidebar.warning("⚠️ Aucun taux historique n'a pu être extrait de ce PDF.")


def build_monthly_history_export_excel(all_periods_list):
    """Génère un fichier Excel prêt à copier-coller dans GRIFF_Historique_Koesio.xlsx pour le(s) mois analysé(s)."""
    import io
    rows_out = []
    for p_per in all_periods_list:
        m_name = mois_fr.get(p_per.month, str(p_per.month))
        y_num = p_per.year
        per_str = f"{y_num}-{p_per.month:02d}"

        df_art_p = df_artis[pd.to_datetime(df_artis['Date'], errors='coerce').dt.to_period('M') == p_per] if (not df_artis.empty and 'Date' in df_artis.columns) else df_artis
        df_isi_p = df_isiwork[pd.to_datetime(df_isiwork['Date facture'], errors='coerce').dt.to_period('M') == p_per] if (not df_isiwork.empty and 'Date facture' in df_isiwork.columns) else df_isiwork

        f_art_l = df_art_p['Organisation interne'].dropna().unique().tolist() if 'Organisation interne' in df_art_p.columns else []
        f_isi_l = df_isi_p['Filiale_Canonique'].dropna().unique().tolist() if 'Filiale_Canonique' in df_isi_p.columns else []
        fils = sorted(list(set(f_art_l + f_isi_l)))

        for fn in fils:
            sub_a = df_art_p[df_art_p['Organisation interne'] == fn] if 'Organisation interne' in df_art_p.columns else pd.DataFrame()
            tot_a = len(sub_a)
            batch_c = len(sub_a[sub_a['Statut'].notna() & (sub_a['Statut'] != '')]) if (tot_a > 0 and 'Statut' in sub_a.columns) else 0
            man_a = tot_a - batch_c
            pct_ba = round(batch_c / tot_a * 100, 1) if tot_a > 0 else 0.0
            pct_ma = round(man_a / tot_a * 100, 1) if tot_a > 0 else 0.0

            if 'Organisation_Zeendoc' in df_isi_p.columns:
                termes = CORRESPONDANCE_ZEENDOC.get(fn, [fn])
                mask_i = df_isi_p['Organisation_Zeendoc'].astype(str).str.upper().apply(lambda x: any(t.upper() in x for t in termes))
                sub_i = df_isi_p[mask_i].copy()
            else:
                sub_i = pd.DataFrame()

            vol_b = len(sub_i)
            if vol_b > 0 and 'Statut de rapprochement' in sub_i.columns:
                auto_c = len(sub_i[sub_i['Statut de rapprochement'] == 'Rapproché automatiquement'])
                not_auto = sub_i['Statut de rapprochement'] != 'Rapproché automatiquement'
                notes_s = sub_i['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in sub_i.columns else pd.Series('', index=sub_i.index)
                exp_s = sub_i['Date export ERP'].fillna('').astype(str) if 'Date export ERP' in sub_i.columns else pd.Series('', index=sub_i.index)
                is_faf = notes_s.str.contains('FAF|TRAIT|ARTIS|ARCHIV|DÉJÀ TRAITÉ', case=False, regex=True) | exp_s.str.contains('FAF|TRAIT|ARTIS|ARCHIV|DÉJÀ TRAITÉ', case=False, regex=True) | sub_i['Statut de rapprochement'].isna()
                ged_p = int((not_auto & is_faf).sum())
                vol_ep = max(0, vol_b - ged_p)
                sub_ep = sub_i[~(not_auto & is_faf)].copy()
                df_m = sub_ep[sub_ep['Statut de rapprochement'].str.contains('manuellement', case=False, na=False)].copy()
                n_m = df_m['Tampons / Notes / Conversations'].fillna('').astype(str) if 'Tampons / Notes / Conversations' in df_m.columns else pd.Series('', index=df_m.index)
                faux_c = len(df_m[n_m.str.contains('FRAIS|PORT|D3EE|DEEE', case=False, regex=True)]) if not df_m.empty else 0
                vrai_c = max(0, len(df_m) - faux_c)
                att_c = max(0, vol_ep - (auto_c + faux_c + vrai_c))
            else:
                auto_c, ged_p, vol_ep, faux_c, vrai_c, att_c = 0, 0, 0, 0, 0, 0

            pa_cnt, pa_dup, _ = compute_pa_metrics(sub_i)
            denom = vol_ep if vol_ep > 0 else 1
            pct_a = round(auto_c / denom * 100, 1) if vol_ep > 0 else 0.0
            pct_f = round(faux_c / denom * 100, 1) if vol_ep > 0 else 0.0
            pct_v = round(vrai_c / denom * 100, 1) if vol_ep > 0 else 0.0
            pct_att = round(att_c / denom * 100, 1) if vol_ep > 0 else 0.0

            rows_out.append({
                "Mois": m_name,
                "Année": y_num,
                "Période (YYYY-MM)": per_str,
                "Filiale Groupe": fn,
                "Vol. Brut GED": vol_b,
                "Flux GED Pure": ged_p,
                "Vol. Épuré GED": vol_ep,
                "Nb Auto GED": auto_c,
                "Taux Rapprochement Auto (%)": pct_a,
                "Nb Faux Manuel": faux_c,
                "Taux Faux Manuel (%)": pct_f,
                "Nb Vrai Manuel": vrai_c,
                "Taux Vrai Manuel (%)": pct_v,
                "Nb En Attente": att_c,
                "Taux En Attente (%)": pct_att,
                "Potentiel Cible (%)": round(pct_a + pct_f, 1),
                "Total Pièces Artis": tot_a,
                "Nb Batch Auto Artis": batch_c,
                "Taux Batch Auto Artis (%)": pct_ba,
                "Nb Manuel Artis": man_a,
                "Taux Manuel Artis (%)": pct_ma,
                "Factures PA": pa_cnt,
                "Doublons PA": pa_dup
            })

    df_exp = pd.DataFrame(rows_out)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False, sheet_name="Nouvelles_Lignes_Historique")
    return buf.getvalue()

if all_periods_p and (not df_isiwork.empty or not df_artis.empty):
    st.sidebar.markdown("---")
    st.sidebar.header("💾 Alimenter l'Historique Excel")
    st.sidebar.download_button(
        label="📥 Exporter les lignes du mois (.xlsx)",
        data=build_monthly_history_export_excel(all_periods_p),
        file_name=f"Export_Historique_GRIFF_{all_periods_p[-1]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Téléchargez les lignes du mois analysé pour les copier-coller à la suite de votre fichier maître GRIFF_Historique_Koesio.xlsx."
    )

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
