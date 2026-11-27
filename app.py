from collections import Counter
from datetime import date, datetime
import io
import json
import os
import re
import zlib
from zoneinfo import ZoneInfo
import numpy as np
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import streamlit as st

LOCAL_TZ = ZoneInfo("Europe/Istanbul")


def now_local() -> datetime:
    return datetime.now(LOCAL_TZ)


def tr_lower(text: str) -> str:
    return (
        str(text or "")
        .replace("İ", "i")
        .replace("I", "ı")
        .lower()
    )


st.set_page_config(
    page_title="İzlenim Paylaşım Paneli",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded",
)

BG_DARK = "#0B111E"
CARD_DARK = "#131C2E"
CARD_BORDER = "#24344D"
TEAL = "#22D3EE"
TEAL_SOFT = "#38BDF8"
PURPLE = "#8B5CF6"
PURPLE_LIGHT = "#A78BFA"
TEXT_LIGHT = "#F8FAFC"
TEXT_MUTED = "#94A3B8"
POS_COLOR = "#10B981"
NEU_COLOR = "#94A3B8"
NEG_COLOR = "#F87171"

st.markdown(
    f"""
    <style>
        .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"] {{
            background-color: {BG_DARK} !important;
            color: {TEXT_LIGHT} !important;
        }}
        [data-testid="stHeader"] {{
            background-color: {BG_DARK} !important;
            color: {TEXT_LIGHT} !important;
        }}
        [data-testid="stToolbar"] button,
        [data-testid="stToolbar"] svg {{
            color: {TEXT_LIGHT} !important;
            fill: {TEXT_LIGHT} !important;
        }}
        [data-testid="stPopover"] > div,
        [data-testid="stPopoverBody"],
        div[role="dialog"],
        ul[role="listbox"],
        ul[role="menu"] {{
            background-color: {CARD_DARK} !important;
            color: {TEXT_LIGHT} !important;
            border: 1px solid {CARD_BORDER} !important;
        }}
        [data-testid="stPopoverBody"] *,
        div[role="dialog"] *,
        ul[role="listbox"] *,
        ul[role="menu"] * {{
            color: {TEXT_LIGHT} !important;
        }}
        [data-baseweb="popover"],
        [data-baseweb="popover"] > div,
        [data-baseweb="menu"],
        [role="menu"],
        [role="menuitem"],
        [role="menuitemradio"] {{
            background-color: {CARD_DARK} !important;
            color: {TEXT_LIGHT} !important;
        }}
        [data-baseweb="popover"] *,
        [role="menu"] * {{
            color: {TEXT_LIGHT} !important;
        }}
        [data-testid="stSidebar"] {{
            background-color: {CARD_DARK} !important;
            border-right: 1px solid {CARD_BORDER};
            color: {TEXT_LIGHT} !important;
        }}
        .stApp label,
        .stApp label p,
        .stApp [data-testid="stWidgetLabel"] p,
        .stApp [data-testid="stMarkdownContainer"] p,
        .stApp [data-testid="stMarkdownContainer"] li,
        .stApp [data-testid="stCaptionContainer"],
        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5 {{
            color: {TEXT_LIGHT} !important;
        }}
        .stApp code {{
            background-color: #1E293B !important;
            color: {TEAL} !important;
            border: 1px solid {CARD_BORDER};
        }}
        [data-testid="stForm"] {{
            background-color: {CARD_DARK} !important;
            border: 1px solid {CARD_BORDER} !important;
            border-radius: 12px;
            padding: 1.25rem;
        }}
        .stTextInput > div > div > input,
        .stTextArea > div > div > textarea,
        .stSelectbox > div > div > div,
        .stMultiSelect > div > div > div,
        .stNumberInput > div > div > input,
        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stDateInput"] input {{
            background-color: #0F172A !important;
            color: {TEXT_LIGHT} !important;
            -webkit-text-fill-color: {TEXT_LIGHT} !important;
            border-color: {CARD_BORDER} !important;
            border-radius: 8px !important;
            min-height: 40px !important;
        }}
        .stTextInput input::placeholder,
        .stTextArea textarea::placeholder,
        div[data-testid="stDateInput"] input::placeholder {{
            color: {TEXT_MUTED} !important;
            -webkit-text-fill-color: {TEXT_MUTED} !important;
            opacity: 1 !important;
        }}
        div[data-testid="stTextInput"] div[data-testid="InputInstructions"] {{
            display: none !important;
        }}
        div[data-testid="stDateInput"],
        div[data-testid="stDateInput"] > div,
        div[data-testid="stDateInput"] > div > div {{
            width: 100% !important;
            max-width: 100% !important;
        }}
        div[data-testid="stDateInput"] svg {{
            color: {TEXT_MUTED} !important;
            fill: {TEXT_MUTED} !important;
        }}
        .stButton > button,
        .stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {{
            background: linear-gradient(90deg, {TEAL_SOFT} 0%, {PURPLE} 100%) !important;
            color: #FFFFFF !important;
            font-weight: 600 !important;
            border: none !important;
            border-radius: 8px;
            padding: 0.5rem 1rem;
        }}
        .stButton > button:hover,
        .stDownloadButton > button:hover,
        [data-testid="stFormSubmitButton"] > button:hover {{
            opacity: 0.92;
            color: #FFFFFF !important;
        }}
        .panel-header {{
            background: linear-gradient(120deg, {CARD_DARK} 0%, #1B1638 60%, #0E293B 100%);
            border: 1px solid {CARD_BORDER};
            border-left: 5px solid {TEAL};
            border-radius: 12px;
            padding: 0.95rem 1.3rem;
            height: 132px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .panel-brand {{
            font-size: 1.34rem;
            font-weight: 800;
            letter-spacing: -0.4px;
            color: {TEXT_LIGHT};
        }}
        .mood-corner-card {{
            background: linear-gradient(135deg, #0D2A29 0%, #132238 55%, #1A1838 100%);
            border: 1.5px solid rgba(16, 185, 129, 0.55);
            box-shadow: 0 8px 28px rgba(16, 185, 129, 0.16), inset 0 1px 0 rgba(255, 255, 255, 0.08);
            border-radius: 12px;
            padding: 0.85rem 1.15rem;
            height: 132px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .mood-top-row {{
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }}
        .mood-svg-wrap {{
            width: 46px;
            height: 46px;
            flex-shrink: 0;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(16, 185, 129, 0.14);
            border: 1px solid rgba(16, 185, 129, 0.45);
        }}
        .mood-title {{
            font-size: 1.05rem;
            font-weight: 800;
            line-height: 1.15;
            letter-spacing: -0.2px;
        }}
        .mood-sub {{
            font-size: 0.75rem;
            color: #CBD5E1;
            margin-top: 0.14rem;
        }}
        .mood-pct-row {{
            display: flex;
            justify-content: space-between;
            gap: 0.35rem;
            margin-top: 0.52rem;
            font-size: 0.75rem;
            font-weight: 700;
        }}
        .mood-bar-track {{
            width: 100%;
            height: 7px;
            border-radius: 999px;
            overflow: hidden;
            display: flex;
            background: #0F172A;
            margin-top: 0.32rem;
            border: 1px solid rgba(148, 163, 184, 0.2);
        }}
        .metric-card {{
            background-color: {CARD_DARK};
            border: 1px solid {CARD_BORDER};
            border-radius: 10px;
            padding: 1rem;
            text-align: left;
            height: 100%;
        }}
        .metric-label {{
            color: {TEXT_MUTED};
            font-size: 0.85rem;
            margin-bottom: 0.25rem;
        }}
        .metric-value {{
            color: {TEXT_LIGHT};
            font-size: 1.55rem;
            font-weight: 700;
        }}
        .metric-sub {{
            color: {TEAL};
            font-size: 0.8rem;
            margin-top: 0.2rem;
        }}
        .detail-box {{
            background-color: {CARD_DARK};
            border: 1px solid {CARD_BORDER};
            border-radius: 10px;
            padding: 1rem 1.15rem;
            margin-bottom: 0.85rem;
        }}
        .action-showcase-banner {{
            background: linear-gradient(125deg, #1A1636 0%, #13223A 65%, #1E1B4B 100%);
            border: 1.5px solid rgba(167, 139, 250, 0.55);
            border-left: 6px solid {PURPLE_LIGHT};
            box-shadow: 0 8px 24px rgba(139, 92, 246, 0.14);
            border-radius: 12px;
            padding: 0.95rem 1.2rem;
            margin-bottom: 0.75rem;
        }}
        .action-note-box {{
            background: linear-gradient(120deg, #161F33 0%, #191936 100%);
            border: 1px solid rgba(167, 139, 250, 0.35);
            border-left: 5px solid {TEAL};
            border-radius: 10px;
            padding: 0.9rem 1.15rem;
            margin-bottom: 0.7rem;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.22);
        }}
        .action-note-box-done {{
            background: linear-gradient(120deg, #0E2623 0%, #141E32 100%);
            border: 1px solid rgba(16, 185, 129, 0.55);
            border-left: 5px solid {POS_COLOR};
            border-radius: 10px;
            padding: 0.95rem 1.15rem;
            margin-bottom: 0.7rem;
        }}
        .criteria-card {{
            background-color: {CARD_DARK};
            border: 1px solid {CARD_BORDER};
            border-top: 4px solid {TEAL};
            border-radius: 10px;
            padding: 0.95rem 1.05rem;
            margin-bottom: 0.75rem;
            min-height: 165px;
        }}
        .criteria-card-purple {{
            background-color: {CARD_DARK};
            border: 1px solid {CARD_BORDER};
            border-top: 4px solid {PURPLE};
            border-radius: 10px;
            padding: 0.95rem 1.05rem;
            margin-bottom: 0.75rem;
            min-height: 165px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

AKSIYON_PLANI_MAP = {
    "1001": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1005": "Analist rolüne geçiş yapmak istiyor, yöneticisi ile",
    "1011": "lokasyon değişikliği talebi bulunuyor.",
    "1014": "istisna talebi bulunuyor, gmy onayı alınacak.",
    "1020": "ücret iyileştirme talebi bulunuyor, direktörü ile",
    "1029": "iş yoğunluğu nedeniyle destek ihtiyacı bulunduğunu",
    "1036": "eğitim seçerken sorun yaşamakta, akademi ekibi ile",
    "1041": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1047": "Analist rolüne geçiş yapmak istiyor, yöneticisi ile",
    "1056": "lokasyon değişikliği talebi bulunuyor.",
    "1059": "istisna talebi bulunuyor, gmy onayı alınacak.",
    "1063": "iş yoğunluğu nedeniyle destek ihtiyacı bulunduğunu",
    "1072": "ücret iyileştirme talebi bulunuyor, direktörü ile",
    "1078": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1082": "iş yoğunluğu nedeniyle destek ihtiyacı bulunduğunu",
    "1088": "ücret iyileştirme talebi bulunuyor, direktörü ile",
    "1095": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1099": "ücret iyileştirme talebi bulunuyor, direktörü ile",
    "1105": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1110": "Anahtar role geçiş talebi bulunuyor, değerlendirilecek.",
    "1113": "iş yoğunluğu nedeniyle destek ihtiyacı bulunduğunu",
    "1117": "Ekip değişikliği talebi bulunmakta, toplantı",
    "1124": "Anahtar role geçiş talebi bulunuyor, değerlendirilecek.",
}

LOKASYON_OPTIONS = [
    "İSTANBUL TEKNOKENT",
    "TUTOM - ARGE",
    "ANKARA TEKNOKENT",
    "TUTOM",
    "İŞ KULE - 3",
]

LOKASYON_BY_SICIL = {
    "1001": "ANKARA TEKNOKENT",
    "1002": "İSTANBUL TEKNOKENT",
    "1003": "ANKARA TEKNOKENT",
    "1004": "İSTANBUL TEKNOKENT",
    "1005": "ANKARA TEKNOKENT",
    "1006": "İSTANBUL TEKNOKENT",
    "1007": "ANKARA TEKNOKENT",
    "1008": "ANKARA TEKNOKENT",
    "1009": "ANKARA TEKNOKENT",
    "1010": "ANKARA TEKNOKENT",
    "1011": "İSTANBUL TEKNOKENT",
    "1012": "TUTOM - ARGE",
    "1013": "ANKARA TEKNOKENT",
    "1014": "TUTOM",
    "1015": "TUTOM - ARGE",
    "1016": "İSTANBUL TEKNOKENT",
    "1017": "TUTOM - ARGE",
    "1018": "TUTOM - ARGE",
    "1019": "İSTANBUL TEKNOKENT",
    "1020": "İSTANBUL TEKNOKENT",
    "1021": "İSTANBUL TEKNOKENT",
    "1022": "İSTANBUL TEKNOKENT",
    "1023": "TUTOM",
    "1024": "TUTOM - ARGE",
    "1025": "TUTOM - ARGE",
    "1026": "TUTOM - ARGE",
    "1027": "İSTANBUL TEKNOKENT",
    "1028": "TUTOM",
    "1029": "İSTANBUL TEKNOKENT",
    "1030": "İSTANBUL TEKNOKENT",
    "1031": "İSTANBUL TEKNOKENT",
    "1032": "İSTANBUL TEKNOKENT",
    "1033": "İSTANBUL TEKNOKENT",
    "1034": "TUTOM - ARGE",
    "1035": "İSTANBUL TEKNOKENT",
    "1036": "İSTANBUL TEKNOKENT",
    "1037": "İSTANBUL TEKNOKENT",
    "1038": "TUTOM - ARGE",
    "1039": "İSTANBUL TEKNOKENT",
    "1040": "TUTOM",
    "1041": "İSTANBUL TEKNOKENT",
    "1042": "TUTOM - ARGE",
    "1043": "TUTOM - ARGE",
    "1044": "İSTANBUL TEKNOKENT",
    "1045": "İSTANBUL TEKNOKENT",
    "1046": "TUTOM - ARGE",
    "1047": "ANKARA TEKNOKENT",
    "1048": "TUTOM - ARGE",
    "1049": "TUTOM - ARGE",
    "1050": "TUTOM",
    "1051": "TUTOM - ARGE",
    "1052": "TUTOM - ARGE",
    "1053": "TUTOM - ARGE",
    "1054": "İSTANBUL TEKNOKENT",
    "1055": "TUTOM - ARGE",
    "1056": "TUTOM - ARGE",
    "1057": "İSTANBUL TEKNOKENT",
    "1058": "İSTANBUL TEKNOKENT",
    "1059": "İSTANBUL TEKNOKENT",
    "1060": "İSTANBUL TEKNOKENT",
    "1061": "İSTANBUL TEKNOKENT",
    "1062": "ANKARA TEKNOKENT",
    "1063": "TUTOM - ARGE",
    "1064": "TUTOM",
    "1065": "ANKARA TEKNOKENT",
    "1066": "İSTANBUL TEKNOKENT",
    "1067": "İSTANBUL TEKNOKENT",
    "1068": "İSTANBUL TEKNOKENT",
    "1069": "İSTANBUL TEKNOKENT",
    "1070": "İSTANBUL TEKNOKENT",
    "1071": "ANKARA TEKNOKENT",
    "1072": "TUTOM - ARGE",
    "1073": "İSTANBUL TEKNOKENT",
    "1074": "İSTANBUL TEKNOKENT",
    "1075": "İSTANBUL TEKNOKENT",
    "1076": "ANKARA TEKNOKENT",
    "1077": "TUTOM - ARGE",
    "1078": "TUTOM - ARGE",
    "1079": "TUTOM - ARGE",
    "1080": "ANKARA TEKNOKENT",
    "1081": "İŞ KULE - 3",
    "1082": "TUTOM - ARGE",
    "1083": "TUTOM - ARGE",
    "1084": "İSTANBUL TEKNOKENT",
    "1085": "ANKARA TEKNOKENT",
    "1086": "TUTOM - ARGE",
    "1087": "TUTOM - ARGE",
    "1088": "İSTANBUL TEKNOKENT",
    "1089": "İSTANBUL TEKNOKENT",
    "1090": "İSTANBUL TEKNOKENT",
    "1091": "TUTOM - ARGE",
    "1092": "İSTANBUL TEKNOKENT",
    "1093": "TUTOM - ARGE",
    "1094": "TUTOM - ARGE",
    "1095": "İSTANBUL TEKNOKENT",
    "1096": "İSTANBUL TEKNOKENT",
    "1097": "İSTANBUL TEKNOKENT",
    "1098": "TUTOM - ARGE",
    "1099": "ANKARA TEKNOKENT",
    "1100": "İSTANBUL TEKNOKENT",
    "1101": "İSTANBUL TEKNOKENT",
    "1102": "İSTANBUL TEKNOKENT",
    "1103": "İSTANBUL TEKNOKENT",
    "1104": "TUTOM - ARGE",
    "1105": "TUTOM - ARGE",
    "1106": "TUTOM - ARGE",
    "1107": "TUTOM - ARGE",
    "1108": "İŞ KULE - 3",
    "1109": "ANKARA TEKNOKENT",
    "1110": "ANKARA TEKNOKENT",
    "1111": "ANKARA TEKNOKENT",
    "1112": "ANKARA TEKNOKENT",
    "1113": "İSTANBUL TEKNOKENT",
    "1114": "ANKARA TEKNOKENT",
    "1115": "ANKARA TEKNOKENT",
    "1116": "TUTOM - ARGE",
    "1117": "TUTOM - ARGE",
    "1118": "TUTOM - ARGE",
    "1119": "TUTOM - ARGE",
    "1120": "ANKARA TEKNOKENT",
    "1121": "İSTANBUL TEKNOKENT",
    "1122": "TUTOM - ARGE",
    "1123": "ANKARA TEKNOKENT",
    "1124": "İSTANBUL TEKNOKENT",
    "1125": "İSTANBUL TEKNOKENT",
    "1126": "İSTANBUL TEKNOKENT",
}


def _normalize_location_value(raw_loc: object, sicil_no: str = "", fallback_idx: int = 0) -> str:
    s_clean = str(sicil_no or "").strip()
    if s_clean in LOKASYON_BY_SICIL:
        return LOKASYON_BY_SICIL[s_clean]

    txt = " ".join(str(raw_loc or "").split()).strip()
    low = tr_lower(txt)
    if "ankara" in low and "teknokent" in low:
        return "ANKARA TEKNOKENT"
    if "istanbul" in low and "teknokent" in low:
        return "İSTANBUL TEKNOKENT"
    if "tutom" in low and "arge" in low:
        return "TUTOM - ARGE"
    if "tutom" in low:
        return "TUTOM"
    if "kule" in low:
        return "İŞ KULE - 3"
    if txt in LOKASYON_OPTIONS:
        return txt
    return LOKASYON_OPTIONS[fallback_idx % len(LOKASYON_OPTIONS)]


def _is_completed_flag(val: object) -> bool:
    if isinstance(val, bool):
        return val
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return False
    return tr_lower(str(val).strip()) in {
        "true",
        "1",
        "evet",
        "yes",
        "tamamlandi",
        "tamamlandı",
        "done",
        "kapalı",
        "kapali",
    }


def _has_completion_note(note_val: object) -> bool:
    if note_val is None or (isinstance(note_val, float) and np.isnan(note_val)):
        return False
    return len(str(note_val).strip()) >= 2


def render_corner_mood_widget(df_source: pd.DataFrame, highlight_positive: bool = True) -> str:
    highlight_style = (
        "background: linear-gradient(135deg, #0B332E 0%, #11293E 55%, #1C1740 100%);"
        "border: 2px solid #10B981;"
        "box-shadow: 0 10px 32px rgba(16, 185, 129, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.14);"
        if highlight_positive
        else ""
    )

    if df_source is None or df_source.empty or "duygu_durumu" not in df_source.columns:
        svg_empty = f"""
        <svg width="30" height="30" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14" stroke="{TEXT_MUTED}" stroke-width="2.2"/>
            <circle cx="13" cy="15" r="1.8" fill="{TEXT_MUTED}"/>
            <circle cx="23" cy="15" r="1.8" fill="{TEXT_MUTED}"/>
            <path d="M12.5 23H23.5" stroke="{TEXT_MUTED}" stroke-width="2.2" stroke-linecap="round"/>
        </svg>
        """
        return f"""
        <div class="mood-corner-card" style="{highlight_style}">
            <div class="mood-top-row">
                <div class="mood-svg-wrap">{svg_empty}</div>
                <div>
                    <div class="mood-title" style="color:{TEXT_MUTED};">Genel Durum: Veri Yok</div>
                    <div class="mood-sub">Filtrede kayıt bulunamadı</div>
                </div>
            </div>
            <div class="mood-pct-row">
                <span style="color:{POS_COLOR};">Olumlu %0.0</span>
                <span style="color:{NEU_COLOR};">Nötr %0.0</span>
                <span style="color:{NEG_COLOR};">Olumsuz %0.0</span>
            </div>
            <div class="mood-bar-track"></div>
        </div>
        """

    s_duygu = df_source["duygu_durumu"].astype(str).str.strip()
    total = len(s_duygu)
    pos_cnt = int((s_duygu == "Olumlu").sum())
    neg_cnt = int((s_duygu == "Olumsuz").sum())
    neu_cnt = max(0, total - pos_cnt - neg_cnt)

    pos_pct = round((pos_cnt / total) * 100, 1) if total > 0 else 0.0
    neg_pct = round((neg_cnt / total) * 100, 1) if total > 0 else 0.0
    neu_pct = round(max(0.0, 100.0 - pos_pct - neg_pct), 1) if total > 0 else 0.0

    if pos_cnt >= neg_cnt and pos_cnt >= neu_cnt:
        durum_baslik = "Genel Olumlu"
        durum_renk = "#34D399" if highlight_positive else POS_COLOR
        vurgu_yuzde = pos_pct
        svg_face = f"""
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14" fill="rgba(16, 185, 129, 0.22)" stroke="{durum_renk}" stroke-width="2.4"/>
            <circle cx="13" cy="14.5" r="2.0" fill="{durum_renk}"/>
            <circle cx="23" cy="14.5" r="2.0" fill="{durum_renk}"/>
            <path d="M11.5 20.5C13.2 23.6 15.5 25.0 18 25.0C20.5 25.0 22.8 23.6 24.5 20.5" stroke="{durum_renk}" stroke-width="2.5" stroke-linecap="round"/>
        </svg>
        """
    elif neg_cnt > pos_cnt and neg_cnt >= neu_cnt:
        durum_baslik = "Genel Olumsuz"
        durum_renk = NEG_COLOR
        vurgu_yuzde = neg_pct
        svg_face = f"""
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14" fill="rgba(248, 113, 113, 0.16)" stroke="{NEG_COLOR}" stroke-width="2.2"/>
            <circle cx="13" cy="15" r="1.9" fill="{NEG_COLOR}"/>
            <circle cx="23" cy="15" r="1.9" fill="{NEG_COLOR}"/>
            <path d="M12 24.2C13.5 21.4 15.6 20 18 20C20.4 20 22.5 21.4 24 24.2" stroke="{NEG_COLOR}" stroke-width="2.3" stroke-linecap="round"/>
        </svg>
        """
    else:
        durum_baslik = "Genel Nötr"
        durum_renk = TEAL
        vurgu_yuzde = neu_pct
        svg_face = f"""
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14" fill="rgba(34, 211, 238, 0.14)" stroke="{TEAL}" stroke-width="2.2"/>
            <circle cx="13" cy="15" r="1.9" fill="{TEAL}"/>
            <circle cx="23" cy="15" r="1.9" fill="{TEAL}"/>
            <path d="M12.2 22.2H23.8" stroke="{TEAL}" stroke-width="2.3" stroke-linecap="round"/>
        </svg>
        """

    return f"""
    <div class="mood-corner-card" style="{highlight_style}">
        <div class="mood-top-row">
            <div class="mood-svg-wrap">{svg_face}</div>
            <div>
                <div class="mood-title" style="color:{durum_renk};">{durum_baslik} <span style="color:{TEXT_LIGHT}; font-size:0.95rem;">(%{vurgu_yuzde:.1f})</span></div>
                <div class="mood-sub">{total} görüşme üzerinden genel izlenim nabzı</div>
            </div>
        </div>
        <div class="mood-pct-row">
            <span style="color:#34D399;">Olumlu %{pos_pct:.1f}</span>
            <span style="color:{NEU_COLOR};">Nötr %{neu_pct:.1f}</span>
            <span style="color:{NEG_COLOR};">Olumsuz %{neg_pct:.1f}</span>
        </div>
        <div class="mood-bar-track">
            <div style="width:{pos_pct}%; background:{POS_COLOR};"></div>
            <div style="width:{neu_pct}%; background:{NEU_COLOR};"></div>
            <div style="width:{neg_pct}%; background:{NEG_COLOR};"></div>
        </div>
    </div>
    """


GMY_BIRIM_MAP = {
    "10000014-Genel Müdürlük": [
        "10000114-İnsan Kaynakları",
        "10000119-Bilgi Güvenliği ve Risk Yönetimi",
        "10000120-İç Denetim ve Kontrol",
        "10000156-Yurtdışı",
        "10000188-Kurumsal İletişim",
    ],
    "10000018-Temel Bankacılık": [
        "10000130-Finans ve İnsan Kaynakları",
        "10000132-Kredi İzleme ve Takip",
        "10000157-Müşteri Yönetimi ve Fiyatlama",
        "10000161-Ürün ve Hesap Yönetimi",
        "10000484-Servis Bankacılığı",
    ],
    "10000023-Mevduat, Kredi, Süreç ve Ödemeler": [
        "10000134-Ödemeler",
        "10000160-Ticari Krediler",
        "10000473-Varlık Yönetimi ve Yurt Dışı İşlemler",
        "10000477-Bireysel Krediler ve Bankasürans",
        "10000478-Şube ve Süreç Uygulamaları",
        "10000492-Satış ve Pazarlama",
    ],
    "10000024-Finans ve Profesyonel Hizmetler": [
        "10000141-Hizmet Yönetimi",
        "10000168-Profesyonel Hizmetler",
        "10000465-Hukuk Müşavirliği",
        "10000479-Strateji ve Finansal Yönetim",
    ],
    "10000025-Sermaye Piyasaları, Hazine ve Sektörel İş Çözümleri": [
        "10000151-Sermaye Piyasaları",
        "10000166-Sermaye Piyasaları Alım Satım Sistemleri",
        "10000167-Sermaye Piyasaları İşlem Sonrası Sistemleri",
        "10000471-Hazine",
        "10000480-Sektörel İş Çözümleri",
    ],
    "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri": [
        "10000122-Plateau Önyüz ve Süreç Uygulamaları Geliştirme Platformu",
        "10000127-Çok Kanallı Bankacılık",
        "10000180-Uluslararası Bankacılık Çözümleri",
        "10000474-Kurumsal Uygulamalar",
        "10000475-Mobil Uygulamalar",
    ],
    "10000028-Teknoloji, Yapay Zeka ve Sürekli İyileştirme": [
        "10000121-Plateau Platform ve Banka Altyapı Uygulamaları",
        "10000169-İnovasyon",
        "10000182-Veri ve Yapay Zeka Uygulamaları",
        "10000187-Deneyim Tasarımı Yetkinlik Merkezi",
        "10000464-İş Analizi Yetkinlik Merkezi",
        "10000469-Dönüşüm ve Sürekli İyileştirme",
        "10000482-Bulut Teknolojileri ve Platform Hizmetleri",
        "10000483-Yapay Zeka Yetkinlik Merkezi",
        "10000485-Proje, Portföy ve Ürün Yönetimi Yetkinlik Merkezi",
        "10000486-Mimari Yetkinlik Merkezi",
    ],
    "10000029-Kartlı Ödeme Sistemleri ve Kampanyalar": [
        "10000472-Pazarlama ve Kampanya Uygulamaları",
        "10000489-Kart Uygulamaları",
        "10000490-Kart Süreç Uygulamaları",
        "10000491-Üye İşyeri Uygulamaları",
    ],
}

DEPARTMAN_OPTIONS = [
    "Yazılım Geliştirme Departmanı",
    "Ürün ve Süreç Yönetimi Departmanı",
    "Altyapı, Bulut ve Platform Departmanı",
    "Veri ve Yapay Zeka Departmanı",
    "Kurumsal Uygulamalar ve Operasyon Departmanı",
    "Bilgi Güvenliği ve Kalite Departmanı",
]

BOLUM_OPTIONS = [
    "Ölçme ve Raporlama",
    "Ölçme Araçları",
    "Dönüşüm Atölyesi",
    "Süreç Kalite ve Güvence",
    "Geliştirici Deneyimi",
    "Çekirdek Sistemler",
    "Mobil ve Ön Yüz Geliştirme",
    "API ve Entegrasyon",
]

YONETICI_OPTIONS = [
    "Birim Yöneticisi - 1",
    "Birim Yöneticisi - 2",
    "Yetkinlik Merkezi Lideri",
    "Direktör",
    "Genel Müdür Yardımcısı",
]

IK_IS_ORTAGI_OPTIONS = [
    "Şeyma Oruç",
    "Örnek İK İş Ortağı",
    "İK İş Ortağı - Teknoloji",
    "İK İş Ortağı - Bankacılık",
]

DEFAULT_USERS = {
    "seyma": {
        "ad_soyad": "Şeyma Oruç",
        "sifre": "panel2026",
        "rol": "hrbp",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
    "hrbp_ornek": {
        "ad_soyad": "Örnek İK İş Ortağı",
        "sifre": "panel2026",
        "rol": "hrbp",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
    "admin": {
        "ad_soyad": "İK Raporlama Yöneticisi",
        "sifre": "admin2026",
        "rol": "admin",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
}

GORUSME_SORULARI = [
    ("q01_calisma_hayati", "Çalışma hayatın nasıl gidiyor?"),
    ("q02_guclu_yonler", "Sana göre olumlu/güçlü yönler neler?"),
    ("q03_ekip_rol_tatmin", "Görev aldığın ekip, rolün ve projelerin seni tatmin ediyor, sana katkı sağlıyor mu?"),
    ("q04_ekip_iyilestirme", "Ekibinde sana göre daha iyi olması gereken şeyler neler?"),
    ("q05_motivasyon_arttiran", "Kendi motivasyonunu daha çok arttıracak şeyler neler olabilir? (İş, proje, kariyer vb.)"),
    ("q06_yonetici_geri_bildirim", "Yöneticin gelişimin için zaman ayırıp düzenli geri bildirim verebiliyor mu? Görüşme sıklığı nasıl?"),
    ("q07_ym_verimlilik", "Yetkinlik Merkezi yapısının verimliliğini nasıl değerlendirirsin? (Eğitim, iş süreçleri vb.)"),
    ("q08_gmy_iletisimi", "GMY iletişimi ile ilgili değerlendirmelerin nasıl olur?"),
    ("q09_vizyon_strateji", "Sunulan vizyon ve stratejiyi büyük resimde nasıl değerlendiriyorsun?"),
    ("q10_degistirmek_istenen", "Neyi değiştirmek isterdin?"),
    ("q11_mevcut_calisma_modeli", "Mevcut çalışma modelinle ilgili ne düşünüyorsun?"),
    ("q12_ilerisi_calisma_modeli", "İlerisi için nasıl bir çalışma modeli olursa daha verimli çalışacağına inanıyorsun?"),
    ("q13_harmony", "Harmony ile ilgili görüşlerin nasıl, çalışma sürecine olan etkisini nasıl değerlendirirsin?"),
    ("q14_toplandigimiza_degsin", '"Toplandığımıza Değsin" temasının ne kadar işlendiğini düşünüyorsun?'),
    ("q15_karar_alma_surecleri", "Karar alma süreçlerini hem ekip içi işleyişler hem de üst yönetim seviyesinde nasıl değerlendirirsin?"),
]

IK_BASLIKLARI = [
    ("ik_kariyer", "İşe Alım ve Kariyer Yönetimi"),
    ("ik_deneyim_marka", "Çalışan Deneyimi ve İşveren Markası Tanıtımı"),
    ("ik_ucret_merkezi", "Ücretlendirme ve Çalışan Merkezi"),
    ("ik_akademi", "Akademi"),
]

TUM_METIN_SORULARI = GORUSME_SORULARI + IK_BASLIKLARI + [("q_kapanis", "Kapanış / Ek Değerlendirmeler")]

PUAN_BASLIKLARI = [
    ("p_genel_memnuniyet", "Genel Memnuniyet"),
    ("p_ekip_motivasyonu", "Ekip Motivasyonu"),
    ("p_kendi_motivasyonu", "Kendi Motivasyonu"),
    ("p_is_yasam_dengesi", "İş Yaşam Dengesi"),
    ("p_1_yonetici_iletisimi", "1. Yönetici İletişimi"),
    ("p_ym_degerlendirme", "Yetkinlik Merkezi Değerlendirmesi"),
    ("p_direktor_iletisimi", "Direktör İletişimi"),
    ("p_gmy_iletisimi", "GMY İletişimi"),
    ("p_ucret_yan_hak", "Ücret Yan Haklar Motivasyonu"),
    ("p_ik_iletisimi", "İK İletişimi"),
    ("p_yz_kullanim", "Yapay Zeka Kullanım Sıklığı (CodeMantis, Amigo, Asistan vb.)"),
]

SCORE_ALIAS_MAP = {
    "p_genel_memnuniyet": [
        "p_genel_memnuniyet",
        "Genel Memnuniyet",
        "Genel Memnuniyet Puanı",
        "Softtech Memnuniyeti",
        "Softtech'e Aidiyet",
        "Aidiyet",
        "st Memnuniyeti",
        "Memnuniyet",
        "Genel Memnuniyet (1-5)",
    ],
    "p_ekip_motivasyonu": [
        "p_ekip_motivasyonu",
        "Ekip Motivasyonu",
        "Ekip Motivasyonu Puanı",
        "Ekip Motivasyonu (1-5)",
    ],
    "p_kendi_motivasyonu": [
        "p_kendi_motivasyonu",
        "Kendi Motivasyonu",
        "Bireysel Motivasyon",
        "Kendi Motivasyonu Puanı",
        "Kendi Motivasyonu (1-5)",
    ],
    "p_is_yasam_dengesi": [
        "p_is_yasam_dengesi",
        "İş Yaşam Dengesi",
        "İş-Yaşam Dengesi",
        "İş Yaşam Dengesi Puanı",
        "İş Yaşam Dengesi (1-5)",
    ],
    "p_1_yonetici_iletisimi": [
        "p_1_yonetici_iletisimi",
        "1. Yönetici İletişimi",
        "1.Yönetici İletişimi",
        "Yönetici İletişimi",
        "Birinci Yönetici İletişimi",
        "1. Yönetici İletişimi Puanı",
        "1. Yönetici İletişimi (1-5)",
    ],
    "p_ym_degerlendirme": [
        "p_ym_degerlendirme",
        "Yetkinlik Merkezi Değerlendirmesi",
        "YM Değerlendirmesi",
    ],
    "p_direktor_iletisimi": [
        "p_direktor_iletisimi",
        "Direktör İletişimi",
        "Direktörüyle İletişim",
    ],
    "p_gmy_iletisimi": [
        "p_gmy_iletisimi",
        "GMY İletişimi",
        "Üst Yönetim (GMY) İletişimi",
        "Üst Yönetim İletişimi",
    ],
    "p_ucret_yan_hak": [
        "p_ucret_yan_hak",
        "Ücret Yan Haklar Motivasyonu",
        "Ücret Yan Haklar",
        "Ücret Motivasyonu",
    ],
    "p_ik_iletisimi": [
        "p_ik_iletisimi",
        "İK İletişimi",
        "İK İletişim",
    ],
    "p_yz_kullanim": [
        "p_yz_kullanim",
        "Yapay Zeka Kullanım Sıklığı (CodeMantis, Amigo, Asistan vb.)",
        "Yapay Zeka Kullanım Sıklığı",
        "Yapay Zeka Kullanımı",
        "YZ Kullanım Sıklığı",
        "YZ Kullanımı",
        "Yapay Zeka Kullanım Sıklığı (CodeMantis, Amigo, Asistan vb)",
        "YZ Kullanım Sıklığı (CodeMantis, Amigo, Asistan vb.)",
    ],
}

HALF_STEP_SCORE_OPTIONS = ["Boş", 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0]


def _parse_date_safe(val: object, fallback: date | None = None) -> date:
    default_d = fallback if fallback is not None else now_local().date()
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if val is None:
        return default_d
    s = str(val).strip()
    if not s:
        return default_d
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return default_d


def _format_tr_date_input(raw_val: object) -> str:
    if raw_val is None:
        return ""
    s = str(raw_val).strip()
    if not s:
        return ""
    parsed = _parse_date_safe(s, fallback=None)
    if parsed is not None:
        return parsed.strftime("%d.%m.%Y")
    return s


def _snap_half_step(raw_val: object) -> float | None:
    if raw_val is None or raw_val == "":
        return None
    if isinstance(raw_val, str) and tr_lower(raw_val.strip()) in ("boş", "bos", "none", "nan", "-"):
        return None
    try:
        if isinstance(raw_val, float) and np.isnan(raw_val):
            return None
        match = re.search(r"(\d+(?:[.,]\d+)?)", str(raw_val).strip())
        if not match:
            return None
        f_val = float(match.group(1).replace(",", "."))
        if np.isnan(f_val) or f_val <= 0:
            return None
        clamped = max(1.0, min(5.0, f_val))
        snapped = round(clamped * 2.0) / 2.0
        return float(max(1.0, min(5.0, snapped)))
    except Exception:
        return None


def _norm_key(k: str) -> str:
    s = tr_lower(str(k or ""))
    return re.sub(r"[^a-zçğıöşü0-9]+", "", s)


def _extract_score_from_record(r: dict, p_key: str, p_label: str) -> float | None:
    candidates = list(SCORE_ALIAS_MAP.get(p_key, [p_key, p_label]))
    for c_key in candidates:
        if c_key in r:
            sc = _snap_half_step(r.get(c_key))
            if sc is not None:
                return sc

    norm_targets = {_norm_key(c) for c in candidates if c}
    for r_k, r_v in r.items():
        if _norm_key(str(r_k)) in norm_targets:
            sc = _snap_half_step(r_v)
            if sc is not None:
                return sc

    for r_k, r_v in r.items():
        rk_low = tr_lower(str(r_k))
        if p_key == "p_genel_memnuniyet" and ("memnuniyet" in rk_low or "aidiyet" in rk_low) and "yorum" not in rk_low:
            sc = _snap_half_step(r_v)
            if sc is not None:
                return sc
        elif p_key == "p_1_yonetici_iletisimi" and "yonetici" in rk_low and ("iletişim" in rk_low or "iletisim" in rk_low):
            sc = _snap_half_step(r_v)
            if sc is not None:
                return sc
        elif (
            p_key == "p_yz_kullanim"
            and ("yapay zeka" in rk_low or "yz " in rk_low or rk_low.startswith("yz") or "codemantis" in rk_low)
            and "yorum" not in rk_low
        ):
            sc = _snap_half_step(r_v)
            if sc is not None:
                return sc

    return None


def _normalize_scores_on_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    records = out.to_dict(orient="records")
    for p_key, p_label in PUAN_BASLIKLARI:
        col_vals = []
        for r in records:
            sc = _extract_score_from_record(r, p_key, p_label)
            col_vals.append(float(sc) if sc is not None else np.nan)
        out[p_key] = pd.to_numeric(pd.Series(col_vals, index=out.index), errors="coerce")
    return out


def _has_meaningful_action_note(note_text: object) -> bool:
    if note_text is None:
        return False
    s = str(note_text).strip()
    if not s:
        return False
    return tr_lower(s) not in {"yok", "-", "hayır", "hayir", "none", "null", "boş", "bos"}


def render_synced_score_row(p_key: str, p_label: str, default_val: float | None, scope_tag: str):
    state_key = f"score_val_{scope_tag}_{p_key}"
    bar_key = f"score_bar_{scope_tag}_{p_key}"
    num_key = f"score_num_{scope_tag}_{p_key}"
    active_key = f"score_active_{scope_tag}_{p_key}"

    if state_key not in st.session_state:
        st.session_state[state_key] = default_val

    cur_val = st.session_state[state_key]
    if bar_key not in st.session_state:
        st.session_state[bar_key] = cur_val if cur_val is not None else "Boş"
    if num_key not in st.session_state:
        st.session_state[num_key] = float(cur_val) if cur_val is not None else 3.0
    if active_key not in st.session_state:
        st.session_state[active_key] = cur_val is not None

    def _on_bar_change():
        b_val = st.session_state[bar_key]
        if b_val == "Boş":
            st.session_state[state_key] = None
            st.session_state[active_key] = False
        else:
            f_val = _snap_half_step(b_val)
            st.session_state[state_key] = f_val
            st.session_state[active_key] = f_val is not None
            if f_val is not None:
                st.session_state[num_key] = f_val

    def _on_num_change():
        n_val = _snap_half_step(st.session_state[num_key])
        if n_val is None:
            n_val = 3.0
        st.session_state[num_key] = n_val
        st.session_state[state_key] = n_val
        st.session_state[active_key] = True
        st.session_state[bar_key] = n_val

    def _on_toggle_change():
        if not st.session_state[active_key]:
            st.session_state[state_key] = None
            st.session_state[bar_key] = "Boş"
        else:
            f_val = _snap_half_step(st.session_state[num_key]) or 3.0
            st.session_state[num_key] = f_val
            st.session_state[state_key] = f_val
            st.session_state[bar_key] = f_val

    st.markdown(f"**{p_label}**")
    c_bar, c_num, c_chk = st.columns([4.5, 1.4, 1.1], vertical_alignment="center")

    with c_bar:
        st.select_slider(
            f"{p_label} Bar",
            options=HALF_STEP_SCORE_OPTIONS,
            key=bar_key,
            on_change=_on_bar_change,
            label_visibility="collapsed",
        )

    with c_num:
        st.number_input(
            f"{p_label} Manuel",
            min_value=1.0,
            max_value=5.0,
            step=0.5,
            format="%.1f",
            key=num_key,
            on_change=_on_num_change,
            label_visibility="collapsed",
        )

    with c_chk:
        st.checkbox(
            "Puanlı",
            key=active_key,
            on_change=_on_toggle_change,
        )

    return st.session_state[state_key]


COLUMN_LABELS = {
    "kayit_id": "Kayıt ID",
    "sicil_no": "Sicil",
    "ad_soyad": "Ad Soyad",
    "gorusme_tarihi": "Görüşme Tarihi",
    "kayit_zamani": "Kayıt Zamanı",
    "donem": "Dönem",
    "hrbp_kullanici": "İK Sorumlusu",
    "ik_is_ortagi": "İK İş Ortağı",
    "gmy": "Genel Müdür Yardımcılığı",
    "departman": "Departman",
    "birim": "Birim",
    "bolum": "Bölüm",
    "yonetici_bilgisi": "Yönetici Bilgisi",
    "lokasyon": "Lokasyon",
    "ise_giris_tarihi": "İşe Giriş Tarihi",
    "katilimci_turu": "Katılımcı Tipi",
    "calisan_rumuzu": "Katılımcı Kodu / Adı",
    "calisma_modeli": "Çalışma Modeli",
    "duygu_durumu": "Duygu Durumu",
    "ana_tema": "Öne Çıkan Ana Tema",
    "aksiyon_ihtiyaci": "Aksiyon planı gerekli mi?",
    "aksiyon_seviyesi": "Aksiyon Seviyesi",
    "aksiyon_notu": "Aksiyon Planı / Detay",
    "aksiyon_tamamlandi": "Aksiyon Tamamlandı",
    "aksiyon_tamamlanma_tarihi": "Aksiyon Kapatılma Tarihi",
    "aksiyon_tamamlanma_notu": "Aksiyon Kapatma Yorumu / Notu",
    "q_kapanis": "Kapanış / Ek Not",
    **dict(GORUSME_SORULARI),
    **dict(IK_BASLIKLARI),
    **dict(PUAN_BASLIKLARI),
}

NEG_SIGNALS = [
    "yavaş", "geç", "sorun", "sıkıntı", "eksik", "yük", "yoğun", "zor", "karmaşık",
    "belirsiz", "yetersiz", "azalt", "kısıt", "olumsuz", "düşük", "problem",
    "zaman alıyor", "bekleniyor", "iyileşmeye açık", "sadeleştir", "şeffaf olabilir",
    "karne yükü", "yetki", "onay", "uzun sürüyor", "kopuk", "yetersiz kalıyor",
    "artırılmalı", "iyileştirilmeli", "net değil",
]

POS_SIGNALS = [
    "güçlü", "iyi", "olumlu", "memnun", "destekleyici", "başarılı", "verimli",
    "net", "açık", "kolay", "hızlı", "değerli", "faydalı", "yüksek", "yapıcı",
    "güven", "ulaşılabilir", "tatmin", "sorun yok", "akıcı", "uyumlu",
    "destekliyor", "katkı sağlıyor", "düzenli", "memnunum", "keyifli", "rahat",
    "sağlıklı", "esnek", "şeffaf", "yardımlaşma", "ekip ruhu", "sinerji",
]

TR_STOPWORDS = {
    "ve", "veya", "ile", "için", "gibi", "kadar", "daha", "çok", "az", "bir", "bu", "şu", "o",
    "da", "de", "ki", "mi", "mı", "mu", "mü", "ama", "fakat", "ancak", "lakin", "çünkü", "yani",
    "en", "her", "bazı", "tüm", "bütün", "hem", "ne", "nasıl", "neden", "niye", "var", "yok",
    "olan", "olarak", "olması", "olduğunu", "oluyor", "olabilir", "olsa", "olsun", "olmak",
    "ben", "sen", "biz", "siz", "onlar", "benim", "bizim", "kendi", "bana", "bize", "beni",
    "süreç", "süreçler", "süreçleri", "konu", "konuda", "konular", "alan", "alanlar", "şekilde",
    "ilgili", "yönelik", "göre", "üzerine", "tarafında", "tarafından", "açısından", "noktasında",
    "şuan", "şu", "an", "genel", "olarak", "bence", "açıkçası", "özellikle", "gerçekten",
}

SEMANTIC_EQUIVALENTS = [
    (r"\b(uzaktan|remote|evden|home\s*office)\b", "uzaktan_calisma"),
    (r"\b(hibrit|ofis\s*gün|ofise\s*gel|ofis\s*sayı)\b", "hibrit_ofis"),
    (r"\b(terfi|unvan|kariyer\s*yolu|kariyer\s*ilerle|pozisyon\s*değiş)\b", "kariyer_terfi"),
    (r"\b(ücret|maaş|zam|prim|yan\s*hak|enflasyon|gelir)\b", "ucret_yan_hak"),
    (r"\b(geri\s*bildirim|feedback|1-1|bire\s*bir|birebir)\b", "geri_bildirim"),
    (r"\b(yetki|onay\s*sürec|onay\s*mekanizma|bürokrasi|karar\s*alma|karar\s*sürec)\b", "karar_yetki_onay"),
    (r"\b(yetkinlik\s*merkezi|ym\b|matris\s*yapı)\b", "yetkinlik_merkezi"),
    (r"\b(iş\s*yük|yoğunluk|mesai|iş\s*yaşam|iş-yaşam|denge)\b", "is_yuku_denge"),
    (r"\b(toplantı|toplandığımıza\s*değsin|gereksiz\s*toplantı)\b", "toplanti_verimlilik"),
    (r"\b(harmony|karne|performans\s*değerlendir|hedef\s*gir)\b", "harmony_karne"),
    (r"\b(yapay\s*zeka|codemantis|amigo|asistan|\byz\b|\bai\b)\b", "yapay_zeka_arac"),
    (r"\b(eğitim|akademi|gelişim|teknik\s*eğitim|mentor)\b", "egitim_akademi"),
    (r"\b(ekip\s*içi|takım\s*arkadaş|yardımlaş|iş\s*birliği|iletişim\s*güçlü)\b", "ekip_iletisim"),
    (r"\b(yönetici|direktör|gmy|üst\s*yönetim|lider)\b", "yonetim_iletisim"),
    (r"\b(proje|ürün|teknoloji|sorumluluk|katkı)\b", "proje_rol_tatmin"),
]

CANONICAL_TOPIC_LABELS = {
    "uzaktan_calisma": "Uzaktan / evden çalışma esnekliği ve gün sayısı",
    "hibrit_ofis": "Hibrit çalışma düzeni ve ofis günleri",
    "kariyer_terfi": "Kariyer gelişimi, rol değişimi ve terfi süreçleri",
    "ucret_yan_hak": "Ücret artışı, prim ve yan hak beklentileri",
    "geri_bildirim": "Yönetici ile düzenli birebir görüşme ve geri bildirim",
    "karar_yetki_onay": "Karar alma hızı, yetki devri ve onay süreçleri",
    "yetkinlik_merkezi": "Yetkinlik Merkezi yapısı ve işleyişi",
    "is_yuku_denge": "İş yükü yoğunluğu ve iş-yaşam dengesi",
    "toplanti_verimlilik": "Toplantı yoğunluğu ve 'Toplandığımıza Değsin' pratiği",
    "harmony_karne": "Harmony / karne süreci ve operasyonel yükü",
    "yapay_zeka_arac": "Yapay zeka araçlarının (CodeMantis, Amigo vb.) kullanımı",
    "egitim_akademi": "Eğitim, Akademi ve teknik gelişim imkânları",
    "ekip_iletisim": "Ekip içi iletişim, yardımlaşma ve dayanışma",
    "yonetim_iletisim": "Yönetici, Direktör ve GMY iletişimi",
    "proje_rol_tatmin": "Proje çeşitliliği, rol tatmini ve teknik katkı",
}


@st.cache_data(show_spinner=False)
def _build_employee_directory_cached(records_json: str) -> dict[str, dict]:
    df_all = pd.read_json(io.StringIO(records_json), orient="records") if records_json else pd.DataFrame()
    directory = {}
    all_gmy_birim = [(g, b) for g, bs in GMY_BIRIM_MAP.items() for b in bs]

    if df_all is not None and not df_all.empty:
        for idx, row in df_all.iterrows():
            sicil = str(row.get("sicil_no", "") or "").strip()
            if not sicil:
                kid = str(row.get("kayit_id", "") or "")
                if "-" in kid:
                    sicil = kid.split("-")[-1].strip()
            if not sicil:
                continue
            g_val = str(row.get("gmy", "") or all_gmy_birim[idx % len(all_gmy_birim)][0])
            b_val = str(row.get("birim", "") or all_gmy_birim[idx % len(all_gmy_birim)][1])
            ad_val = str(row.get("ad_soyad", "") or row.get("calisan_rumuzu", "") or f"Kayıt {sicil}").split("(")[0].strip()
            yon_val = str(row.get("yonetici_bilgisi", "") or YONETICI_OPTIONS[idx % len(YONETICI_OPTIONS)])
            directory[sicil] = {
                "sicil_no": sicil,
                "ad_soyad": ad_val,
                "gmy": g_val,
                "departman": str(row.get("departman", "") or DEPARTMAN_OPTIONS[idx % len(DEPARTMAN_OPTIONS)]),
                "birim": b_val,
                "bolum": str(row.get("bolum", "") or BOLUM_OPTIONS[idx % len(BOLUM_OPTIONS)]),
                "yonetici_bilgisi": yon_val,
                "lokasyon": _normalize_location_value(row.get("lokasyon", ""), sicil_no=sicil, fallback_idx=idx),
                "ise_giris_tarihi": _format_tr_date_input(row.get("ise_giris_tarihi", "")) or f"15.0{(idx % 8) + 1}.202{(idx % 5) + 1}",
                "ik_is_ortagi": str(row.get("ik_is_ortagi", "") or IK_IS_ORTAGI_OPTIONS[idx % len(IK_IS_ORTAGI_OPTIONS)]),
            }

    for idx_s, (sic_k, loc_v) in enumerate(LOKASYON_BY_SICIL.items()):
        if sic_k not in directory:
            g_m, b_m = all_gmy_birim[idx_s % len(all_gmy_birim)]
            directory[sic_k] = {
                "sicil_no": sic_k,
                "ad_soyad": f"Çalışan {int(sic_k) - 1000 if sic_k.isdigit() and int(sic_k) >= 1001 else sic_k}",
                "gmy": g_m,
                "departman": DEPARTMAN_OPTIONS[idx_s % len(DEPARTMAN_OPTIONS)],
                "birim": b_m,
                "bolum": BOLUM_OPTIONS[idx_s % len(BOLUM_OPTIONS)],
                "yonetici_bilgisi": YONETICI_OPTIONS[idx_s % len(YONETICI_OPTIONS)],
                "lokasyon": loc_v,
                "ise_giris_tarihi": f"15.0{(idx_s % 8) + 1}.202{(idx_s % 5) + 1}",
                "ik_is_ortagi": IK_IS_ORTAGI_OPTIONS[idx_s % len(IK_IS_ORTAGI_OPTIONS)],
            }
        else:
            directory[sic_k]["lokasyon"] = loc_v

    return directory


def get_employee_directory(df_all: pd.DataFrame) -> dict[str, dict]:
    return _build_employee_directory_cached(_df_to_json_safe(df_all))


def normalize_turkish_text(text: str) -> str:
    low = tr_lower(text).replace("’", "'").replace("“", '"').replace("”", '"')
    for pattern, canonical in SEMANTIC_EQUIVALENTS:
        low = re.sub(pattern, f" {canonical} ", low)

    low = re.sub(r"[^a-zçğıöşü0-9_\s]", " ", low)
    tokens = []
    for tok in low.split():
        if tok in TR_STOPWORDS or len(tok) < 3:
            continue
        if "_" not in tok and len(tok) > 5:
            tok = tok[:6]
        tokens.append(tok)
    return " ".join(tokens)


def extract_canonical_topics(text: str) -> list[str]:
    low = tr_lower(text)
    matched = []
    for pattern, canonical in SEMANTIC_EQUIVALENTS:
        if re.search(pattern, low):
            matched.append(canonical)
    return matched


def split_clauses(text: str) -> list[str]:
    if not text or not isinstance(text, str):
        return []
    cleaned = text.strip()
    if not cleaned:
        return []
    parts = re.split(r"[.!?\n;]+|,\s+(?=[A-ZÇĞİÖŞÜa-zçğıöşü]{3,})", cleaned)
    result = []
    for p in parts:
        item = re.sub(r"\s+", " ", p).strip(" -•\t")
        if len(item) >= 4:
            result.append(item)
    return result if result else [cleaned]


def classify_clause(
    clause: str,
    row_duygu: str = "",
    default_dir: str = "olumlu",
    col_key: str = "",
) -> str:
    low = tr_lower(clause)

    if col_key == "q02_guclu_yonler":
        strong_neg = ["sorun", "sıkıntı", "eksik", "yetersiz", "problem", "kopuk", "düşük", "kötü"]
        if not any(w in low for w in strong_neg):
            return "olumlu"

    pos_hits = sum(1 for w in POS_SIGNALS if tr_lower(w) in low)
    neg_hits = sum(1 for w in NEG_SIGNALS if tr_lower(w) in low)

    if pos_hits > neg_hits:
        return "olumlu"
    if neg_hits > pos_hits:
        return "olumsuz"
    if row_duygu == "Olumlu":
        return "olumlu"
    if row_duygu == "Olumsuz":
        return "olumsuz"
    return default_dir


def cluster_similar_statements(
    records: list[tuple[int, str]],
    top_n: int = 3,
    sim_threshold: float = 0.34,
    max_items: int = 200,
) -> list[str]:
    if not records:
        return []

    filtered_records = []
    norm_docs = []
    for p_id, raw_txt in records:
        clean_txt = re.sub(r"\s+", " ", raw_txt).strip()
        if len(clean_txt) < 4:
            continue
        norm = normalize_turkish_text(clean_txt)
        filtered_records.append((p_id, clean_txt))
        norm_docs.append(norm if norm.strip() else tr_lower(clean_txt))

    if not filtered_records:
        return []

    if len(filtered_records) > max_items:
        filtered_records = filtered_records[:max_items]
        norm_docs = norm_docs[:max_items]
    n = len(filtered_records)
    if n == 1:
        return [f"{filtered_records[0][1]} (1 kişi)"]
    if n == 2:
        p0, t0 = filtered_records[0]
        p1, t1 = filtered_records[1]
        if p0 != p1 and (set(extract_canonical_topics(t0)) & set(extract_canonical_topics(t1))):
            return [f'"{t0}" (2 kişi benzer görüş bildirdi)']
        return [t0, t1]

    try:
        vec = TfidfVectorizer(ngram_range=(1, 2), min_df=1, sublinear_tf=True)
        tfidf_mat = vec.fit_transform(norm_docs)
        sim_mat = cosine_similarity(tfidf_mat)
    except Exception:
        sim_mat = np.eye(n)

    topics_per_item = [set(extract_canonical_topics(txt)) for _, txt in filtered_records]
    for i in range(n):
        if not topics_per_item[i]:
            continue
        for j in range(i + 1, n):
            if topics_per_item[i] & topics_per_item[j]:
                boosted = max(sim_mat[i, j], 0.55)
                sim_mat[i, j] = boosted
                sim_mat[j, i] = boosted

    unassigned = set(range(n))
    clusters = []

    while unassigned:
        best_center = max(
            unassigned,
            key=lambda idx: sum(1 for j in unassigned if sim_mat[idx, j] >= sim_threshold),
        )
        members = [j for j in list(unassigned) if sim_mat[best_center, j] >= sim_threshold]
        for m in members:
            unassigned.remove(m)
        clusters.append(members)

    ranked_clusters = []
    for member_idxs in clusters:
        unique_people = {filtered_records[idx][0] for idx in member_idxs}
        person_count = len(unique_people)

        candidate_texts = [filtered_records[idx][1] for idx in member_idxs]
        rep_text = min(
            candidate_texts,
            key=lambda t: (abs(len(t) - 65), -len(t)),
        )

        topic_counter = Counter(
            t_key for idx in member_idxs for t_key in topics_per_item[idx]
        )
        tema_etiketi = ""
        if topic_counter:
            top_topic, top_t_cnt = topic_counter.most_common(1)[0]
            if top_t_cnt >= 2 and top_topic in CANONICAL_TOPIC_LABELS:
                tema_etiketi = f"[{CANONICAL_TOPIC_LABELS[top_topic]}] "

        ranked_clusters.append((person_count, len(member_idxs), tema_etiketi, rep_text))

    ranked_clusters.sort(key=lambda x: (-x[0], -x[1], len(x[3])))

    out = []
    for person_count, _, tema_etiketi, rep_text in ranked_clusters[:top_n]:
        if person_count > 1:
            out.append(f'{tema_etiketi}"{rep_text}" ({person_count} kişi benzer görüş bildirdi)')
        else:
            out.append(f"{tema_etiketi}{rep_text}")
    return out


def _df_fingerprint(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "empty_v3"
    try:
        payload = pd.util.hash_pandas_object(df, index=True).values.tobytes()
        return f"{zlib.crc32(payload)}_v3"
    except Exception:
        return f"{len(df)}_{list(df.columns)}_v3"


def _df_to_json_safe(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "[]"
    return df.to_json(orient="records", force_ascii=False)


def _generate_narrative_unit_summary(unit_name: str, df_b: pd.DataFrame, top_pos: list[str], top_neg: list[str]) -> str:
    if df_b.empty:
        return f"{unit_name} birimi için henüz kayıtlı görüşme yorumu bulunmamaktadır."

    n_kisi = len(df_b)
    mem_series = pd.to_numeric(df_b.get("p_genel_memnuniyet", pd.Series(dtype=float)), errors="coerce").dropna()
    yon_series = pd.to_numeric(df_b.get("p_1_yonetici_iletisimi", pd.Series(dtype=float)), errors="coerce").dropna()
    yz_series = pd.to_numeric(df_b.get("p_yz_kullanim", pd.Series(dtype=float)), errors="coerce").dropna()

    mem_ort = round(float(mem_series.mean()), 2) if not mem_series.empty else None
    yon_ort = round(float(yon_series.mean()), 2) if not yon_series.empty else None
    yz_ort = round(float(yz_series.mean()), 2) if not yz_series.empty else None

    paragraflar = []
    if mem_ort is not None and mem_ort >= 4.0:
        paragraflar.append(
            f"{unit_name} biriminde ({n_kisi} görüşme) genel memnuniyetin yüksek olduğu ({mem_ort}/5.00) ve ekip içi iletişim ile yardımlaşmanın olumlu değerlendirildiği görülmektedir."
        )
    elif mem_ort is not None:
        paragraflar.append(
            f"{unit_name} biriminde ({n_kisi} görüşme) genel memnuniyet ortalaması {mem_ort}/5.00 seviyesinde olup iş ve süreç deneyimleri karma bir görünüm sergilemektedir."
        )
    else:
        paragraflar.append(f"{unit_name} biriminde toplam {n_kisi} 1-1 izlenim paylaşım görüşmesi gerçekleştirilmiştir.")

    if top_pos:
        pos_clean = "; ".join([p.replace('"', "") for p in top_pos[:2]])
        paragraflar.append(f"Öne çıkan güçlü ve olumlu başlıklarda: {pos_clean} ifadeleri dikkat çekmektedir.")

    if top_neg:
        neg_clean = "; ".join([n.replace('"', "") for n in top_neg[:2]])
        paragraflar.append(f"Gelişime açık ve iyileştirme beklenen alanlarda ise: {neg_clean} konuları ön plana çıkmaktadır.")

    if yon_ort is not None:
        if yon_ort >= 4.0:
            paragraflar.append(
                f"Birinci yönetici iletişimi oldukça güçlü ve tutarlı değerlendirilmektedir ({yon_ort}/5.00); ihtiyaç duyulduğunda ulaşılabilirlik ve düzenli geri bildirim olumlu öne çıkmaktadır."
            )
        else:
            paragraflar.append(
                f"Yönetici iletişimi ve düzenli geri bildirim sıklığı ({yon_ort}/5.00) gelişime açık bir alan olarak takip edilmelidir."
            )

    if yz_ort is not None and yz_ort >= 3.5:
        paragraflar.append(
            f"Yapay zeka araçlarının (CodeMantis, Amigo, Asistan vb.) kullanım sıklığı ({yz_ort}/5.00) geliştirme hızını ve verimliliği destekleyen unsurlar arasında yer almaktadır."
        )

    return " ".join(paragraflar)


def _analyze_department_impl(df_scoped: pd.DataFrame, secili_gmy: str) -> dict:
    if secili_gmy != "Tümü" and "gmy" in df_scoped.columns and not df_scoped.empty:
        df_dept = df_scoped[df_scoped["gmy"] == secili_gmy].copy()
        hedef_gmy_adlari = [secili_gmy]
    else:
        df_dept = df_scoped.copy()
        hedef_gmy_adlari = (
            sorted(df_dept["gmy"].dropna().unique().tolist())
            if not df_dept.empty and "gmy" in df_dept.columns
            else list(GMY_BIRIM_MAP.keys())
        )

    dept_pos_all = []
    dept_neg_all = []
    text_cols = [k for k, _ in TUM_METIN_SORULARI if k in df_dept.columns]

    for person_idx, (_, row) in enumerate(df_dept.iterrows()):
        row_duygu = str(row.get("duygu_durumu", "") or "")
        for tc in text_cols:
            val = row.get(tc, "")
            if pd.notna(val) and str(val).strip():
                def_dir = "olumsuz" if tc in ("q04_ekip_iyilestirme", "q10_degistirmek_istenen") else "olumlu"
                for c_item in split_clauses(str(val)):
                    if classify_clause(c_item, row_duygu, default_dir=def_dir, col_key=tc) == "olumsuz":
                        dept_neg_all.append((person_idx, c_item))
                    else:
                        dept_pos_all.append((person_idx, c_item))

    birim_listesi = []
    for g_ad in hedef_gmy_adlari:
        katalog_birimler = GMY_BIRIM_MAP.get(g_ad, [])
        veri_birimleri = (
            df_dept[df_dept["gmy"] == g_ad]["birim"].dropna().unique().tolist()
            if not df_dept.empty and "gmy" in df_dept.columns and "birim" in df_dept.columns
            else []
        )
        tum_birimler = []
        for b in katalog_birimler + veri_birimleri:
            if b not in tum_birimler:
                tum_birimler.append(b)

        for b_ad in tum_birimler:
            df_b = (
                df_dept[df_dept["birim"] == b_ad]
                if not df_dept.empty and "birim" in df_dept.columns
                else pd.DataFrame()
            )
            b_pos = []
            b_neg = []
            for p_idx, (_, row) in enumerate(df_b.iterrows()):
                row_duygu = str(row.get("duygu_durumu", "") or "")
                for tc in text_cols:
                    val = row.get(tc, "")
                    if pd.notna(val) and str(val).strip():
                        def_dir = "olumsuz" if tc in ("q04_ekip_iyilestirme", "q10_degistirmek_istenen") else "olumlu"
                        for c_item in split_clauses(str(val)):
                            if classify_clause(c_item, row_duygu, default_dir=def_dir, col_key=tc) == "olumsuz":
                                b_neg.append((p_idx, c_item))
                            else:
                                b_pos.append((p_idx, c_item))

            top_b_pos = cluster_similar_statements(b_pos, top_n=2)
            top_b_neg = cluster_similar_statements(b_neg, top_n=2)
            olumsuz_kisi = (
                int((df_b["duygu_durumu"] == "Olumsuz").sum())
                if not df_b.empty and "duygu_durumu" in df_b.columns
                else 0
            )
            aksiyon_kisi = (
                int((df_b["aksiyon_ihtiyaci"] == "Evet").sum())
                if not df_b.empty and "aksiyon_ihtiyaci" in df_b.columns
                else 0
            )

            yon_series = (
                pd.to_numeric(df_b["p_1_yonetici_iletisimi"], errors="coerce").dropna()
                if not df_b.empty and "p_1_yonetici_iletisimi" in df_b.columns
                else pd.Series(dtype=float)
            )
            yon_skor = round(float(yon_series.mean()), 2) if not yon_series.empty else None

            if df_b.empty:
                durum_etiketi = "Kayıt Yok"
            elif top_b_neg or olumsuz_kisi > 0 or aksiyon_kisi > 0:
                durum_etiketi = "Dikkat / İyileştirme Alanı Var"
            else:
                durum_etiketi = "Genel Olumlu / Sorun Görünmüyor"

            ozet_metin = _generate_narrative_unit_summary(b_ad, df_b, top_b_pos, top_b_neg)

            birim_listesi.append(
                {
                    "gmy": g_ad,
                    "birim": b_ad,
                    "gorusme_sayisi": len(df_b),
                    "durum": durum_etiketi,
                    "yonetici_iletisim_puani": yon_skor,
                    "olumlu": top_b_pos or (["Kayıt bulunmuyor"] if df_b.empty else ["Belirgin olumlu ifade yok"]),
                    "olumsuz": top_b_neg or (["Kayıt bulunmuyor"] if df_b.empty else ["Öne çıkan olumsuz ifade yok"]),
                    "genel_ozet_metin": ozet_metin,
                }
            )

    return {
        "gorusme_sayisi": len(df_dept),
        "olumlu": cluster_similar_statements(dept_pos_all, top_n=4) or ["Yeterli yorum bulunmuyor"],
        "olumsuz": cluster_similar_statements(dept_neg_all, top_n=4) or ["Öne çıkan olumsuz ifade bulunmuyor"],
        "birimler": birim_listesi,
    }


@st.cache_data(show_spinner=False)
def _cached_department_analysis(df_key: str, records_json: str, secili_gmy: str) -> dict:
    df_scoped = pd.read_json(io.StringIO(records_json), orient="records") if records_json else pd.DataFrame()
    return _analyze_department_impl(df_scoped, secili_gmy)


@st.cache_data(show_spinner=False)
def _cached_top_short_bullets(
    df_key: str,
    records_json: str,
    cols: tuple[str, ...],
    top_n: int = 4,
    direction: str | None = None,
) -> list[str]:
    df_view = pd.read_json(io.StringIO(records_json), orient="records") if records_json else pd.DataFrame()
    if df_view.empty:
        return []
    recs = []
    for p_idx, (_, r_item) in enumerate(df_view.iterrows()):
        r_duygu = str(r_item.get("duygu_durumu", "") or "")
        for c_key in cols:
            if c_key not in df_view.columns:
                continue
            raw_v = r_item.get(c_key, "")
            if pd.isna(raw_v) or not str(raw_v).strip():
                continue
            def_dir = "olumsuz" if c_key in ("q04_ekip_iyilestirme", "q10_degistirmek_istenen") else "olumlu"
            for cl in split_clauses(str(raw_v)):
                cl_dir = classify_clause(cl, r_duygu, default_dir=def_dir, col_key=c_key)
                if direction is None or cl_dir == direction:
                    recs.append((p_idx, cl))
    return cluster_similar_statements(recs, top_n=top_n, max_items=180)


@st.cache_data(show_spinner=False)
def _cached_hybrid_percentages(df_key: str, records_json: str) -> list[tuple[str, int, float]]:
    df_view = pd.read_json(io.StringIO(records_json), orient="records") if records_json else pd.DataFrame()
    if df_view.empty or "q12_ilerisi_calisma_modeli" not in df_view.columns:
        return []
    buckets = Counter()
    total_ans = 0
    for val in df_view["q12_ilerisi_calisma_modeli"].dropna():
        txt = str(val).strip()
        if not txt:
            continue
        total_ans += 1
        low = tr_lower(txt)
        if any(k in low for k in ["tam uzaktan", "tamamen uzaktan", "full remote", "sadece evden", "evden çalış"]):
            buckets["Tam Uzaktan / Evden"] += 1
        elif any(k in low for k in ["ofis", "tam zamanlı ofis", "her gün ofis"]):
            buckets["Ofis Ağırlıklı / Tam Ofis"] += 1
        elif any(k in low for k in ["hibrit", "esnek", "gün", "haftada", "uzaktan", "denge"]):
            buckets["Hibrit / Esnek Model"] += 1
        else:
            buckets["Diğer / Koşullu Görüş"] += 1
    if total_ans == 0:
        return []
    return [(b_name, b_cnt, round((b_cnt / total_ans) * 100, 1)) for b_name, b_cnt in buckets.most_common()]


def _compute_4_team_criteria(df_view: pd.DataFrame) -> list[dict]:
    def _mean_col(col_name: str):
        if df_view.empty or col_name not in df_view.columns:
            return None
        s = pd.to_numeric(df_view[col_name], errors="coerce").dropna()
        return round(float(s.mean()), 2) if not s.empty else None

    def _face(score):
        if score is None:
            return "😐 Veri Yok"
        if score >= 4.0:
            return "😄 Çok İyi"
        if score >= 3.2:
            return "🙂 İyi"
        return "😟 Dikkat"

    ekip_m = _mean_col("p_ekip_motivasyonu")
    kendi_m = _mean_col("p_kendi_motivasyonu")
    denge_p = _mean_col("p_is_yasam_dengesi")
    yon_p = _mean_col("p_1_yonetici_iletisimi")

    g_pos, g_tot = 0, 0
    if not df_view.empty and "q03_ekip_rol_tatmin" in df_view.columns:
        for val in df_view["q03_ekip_rol_tatmin"].dropna():
            txt = str(val).strip()
            if txt:
                g_tot += 1
                if classify_clause(txt, col_key="q03_ekip_rol_tatmin") == "olumlu":
                    g_pos += 1

    return [
        {
            "baslik": "1. Görev Alanı & Ekip Değerlendirmesi",
            "aciklama": "Ekip birliktelik duygusu, rol ve proje tatmini",
            "gosterge": f"Rol/Ekip Tatmini Olumlu: {g_pos}/{g_tot} yanıt",
            "durum": "😄 Güçlü" if (g_tot > 0 and g_pos / g_tot >= 0.65) else "🙂 Dengeli",
        },
        {
            "baslik": "2. Bireysel / Ekip Motivasyonu",
            "aciklama": "Ekip üyelerinin motivasyon ve işe tutku seviyesi",
            "gosterge": f"Ekip: {ekip_m or '-'}/5 | Bireysel: {kendi_m or '-'}/5",
            "durum": _face(ekip_m),
        },
        {
            "baslik": "3. İş / Yaşam Dengesi",
            "aciklama": "Çalışma yoğunluğu ve sosyal hayat dengesi",
            "gosterge": f"İş-Yaşam Dengesi: {denge_p or '-'}/5.00",
            "durum": _face(denge_p),
        },
        {
            "baslik": "4. Geri Bildirim & Yönetici İletişimi",
            "aciklama": "Düzenli 1-1 ve yapıcı geri bildirim süreçleri",
            "gosterge": f"1. Yönetici İletişimi: {yon_p or '-'}/5.00",
            "durum": _face(yon_p),
        },
    ]


def get_gsheet_client():
    try:
        sheet_url = st.secrets.get("SPREADSHEET_URL", "")
        gcp_info = st.secrets.get("gcp_service_account", None)
        if not sheet_url or "SENIN_SHEET_ID_BURAYA" in sheet_url or not gcp_info:
            return None, None
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(dict(gcp_info), scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(sheet_url)
        return client, spreadsheet
    except Exception:
        return None, None


def _map_records_to_df(records: list[dict]) -> pd.DataFrame:
    all_pairs = [(g, b) for g, bs in GMY_BIRIM_MAP.items() for b in bs]

    def _clean(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return ""
        return " ".join(str(v).split()).strip()

    rows_out = []
    for idx, r in enumerate(records):
        g_mapped, b_mapped = all_pairs[idx % len(all_pairs)]
        g_val = _clean(r.get("gmy") or r.get("GMY") or r.get("Genel Müdür Yardımcılığı") or r.get("GMY / Domain")) or g_mapped
        b_val = _clean(r.get("birim") or r.get("Birim") or r.get("Birim / Ekip")) or b_mapped
        dept_val = _clean(r.get("departman") or r.get("Departman")) or DEPARTMAN_OPTIONS[idx % len(DEPARTMAN_OPTIONS)]
        bolum_val = _clean(r.get("bolum") or r.get("Bölüm")) or BOLUM_OPTIONS[idx % len(BOLUM_OPTIONS)]
        yon_val = _clean(r.get("yonetici_bilgisi") or r.get("Yönetici Bilgisi")) or YONETICI_OPTIONS[idx % len(YONETICI_OPTIONS)]
        sicil_val = _clean(r.get("sicil_no") or r.get("Sicil") or r.get("Sicil No")) or str(1001 + idx)
        raw_lok = _clean(r.get("lokasyon") or r.get("Lokasyon"))
        lok_val = _normalize_location_value(raw_lok, sicil_no=sicil_val, fallback_idx=idx)
        ise_giris_val = _format_tr_date_input(r.get("ise_giris_tarihi") or r.get("İşe Giriş Tarihi")) or f"15.0{(idx % 8) + 1}.202{(idx % 5) + 1}"
        ik_ortak_val = _clean(r.get("ik_is_ortagi") or r.get("İK İş Ortağı")) or IK_IS_ORTAGI_OPTIONS[idx % len(IK_IS_ORTAGI_OPTIONS)]
        ad_soyad_val = _clean(r.get("ad_soyad") or r.get("Ad Soyad")) or f"Kayıt {idx + 1}"
        g_tarih = _clean(r.get("gorusme_tarihi") or r.get("Görüşme Tarihi") or r.get("Görüşme Tarihi (Ay/Gün/Yıl)")) or "2026-06-04"
        kayit_zamani_val = _clean(r.get("kayit_zamani") or r.get("Kayıt Zamanı")) or f"{g_tarih} 10:00"

        mem = _extract_score_from_record(r, "p_genel_memnuniyet", "Genel Memnuniyet")
        duygu_raw = _clean(r.get("duygu_durumu") or r.get("Duygu Durumu"))
        if duygu_raw in ["Olumlu", "Nötr", "Olumsuz"]:
            duygu = duygu_raw
        elif mem is None:
            duygu = "Nötr"
        elif mem >= 4.0:
            duygu = "Olumlu"
        elif mem < 3.0:
            duygu = "Olumsuz"
        else:
            duygu = "Nötr"

        q10_txt = _clean(r.get("q10_degistirmek_istenen") or r.get("st'te neyi değiştirmek isterdin?"))

        pdf_detay = AKSIYON_PLANI_MAP.get(sicil_val)
        aksiyon_var_raw = _clean(
            r.get("aksiyon_ihtiyaci")
            or r.get("Aksiyon planı gerekli mi?")
            or r.get("Aksiyon İhtiyacı")
        )
        aksiyon_not_raw = _clean(
            r.get("aksiyon_notu")
            or r.get("Aksiyon Planı / Detay")
            or r.get("Aksiyon Planı")
            or r.get("Detay")
            or r.get("Aksiyon Notu")
        )

        if pdf_detay is not None:
            aksiyon_var = "Evet"
            aksiyon_not_final = aksiyon_not_raw if _has_meaningful_action_note(aksiyon_not_raw) else pdf_detay
        elif aksiyon_var_raw in ["Evet", "Hayır"]:
            aksiyon_var = aksiyon_var_raw
            aksiyon_not_final = aksiyon_not_raw if (aksiyon_var == "Evet" and _has_meaningful_action_note(aksiyon_not_raw)) else ""
        else:
            aksiyon_var = "Hayır"
            aksiyon_not_final = ""

        tamamlanma_not_val = _clean(
            r.get("aksiyon_tamamlanma_notu")
            or r.get("Aksiyon Kapatma Yorumu / Notu")
            or r.get("Aksiyon Nasıl Tamamlandı (Not)")
        )
        tamamlanma_tarih_val = _format_tr_date_input(
            r.get("aksiyon_tamamlanma_tarihi")
            or r.get("Aksiyon Kapatılma Tarihi")
            or r.get("Aksiyon Tamamlanma Tarihi")
        )
        tamamlandi_raw = r.get("aksiyon_tamamlandi") if "aksiyon_tamamlandi" in r else r.get("Aksiyon Tamamlandı")
        aksiyon_tamamlandi = (
            _is_completed_flag(tamamlandi_raw) and _has_completion_note(tamamlanma_not_val)
            if aksiyon_var == "Evet"
            else False
        )
        if not aksiyon_tamamlandi:
            tamamlanma_tarih_val = ""

        yz_yorum = _clean(r.get("Yapay Zeka Kullanım Sıklığı Yorum(CodeMantis, Amigo, Asistan vb.)"))
        kapanis_txt = _clean(r.get("q_kapanis") or r.get("Paylaşmak istediğin veya eklemek istediğin farklı konu var mı?"))
        if yz_yorum:
            kapanis_txt = f"{kapanis_txt} | YZ Yorum: {yz_yorum}".strip(" |")

        rows_out.append(
            {
                "kayit_id": _clean(r.get("kayit_id")) or f"2026-{sicil_val}",
                "sicil_no": sicil_val,
                "ad_soyad": ad_soyad_val,
                "gorusme_tarihi": g_tarih,
                "kayit_zamani": kayit_zamani_val,
                "donem": _clean(r.get("donem") or r.get("Dönem")) or "2026 - 2. Dönem",
                "hrbp_kullanici": _clean(r.get("hrbp_kullanici")) or "seyma",
                "ik_is_ortagi": ik_ortak_val,
                "gmy": g_val,
                "departman": dept_val,
                "birim": b_val,
                "bolum": bolum_val,
                "yonetici_bilgisi": yon_val,
                "lokasyon": lok_val,
                "ise_giris_tarihi": ise_giris_val,
                "katilimci_turu": _clean(r.get("katilimci_turu")) or "Katılımcı",
                "calisan_rumuzu": f"{ad_soyad_val} ({_clean(r.get('Birim')) or b_val.split('-', 1)[-1]})",
                "calisma_modeli": _clean(r.get("calisma_modeli")) or "Hibrit",
                "duygu_durumu": duygu,
                "ana_tema": _clean(r.get("ana_tema"))
                or (
                    "Çalışma Modeli / Hibrit / İş-Yaşam"
                    if any(w in tr_lower(q10_txt) for w in ["uzaktan", "remote", "ofis", "çalışma"])
                    else "Güven ve Kültür"
                ),
                "aksiyon_ihtiyaci": aksiyon_var,
                "aksiyon_seviyesi": _clean(r.get("aksiyon_seviyesi")) or ("Birim / YM" if aksiyon_var == "Evet" else "Yok"),
                "aksiyon_notu": aksiyon_not_final,
                "aksiyon_tamamlandi": aksiyon_tamamlandi,
                "aksiyon_tamamlanma_tarihi": tamamlanma_tarih_val,
                "aksiyon_tamamlanma_notu": tamamlanma_not_val,
                "q_kapanis": kapanis_txt,
                **{
                    k: _clean(r.get(k) or r.get(lbl))
                    for k, lbl in (GORUSME_SORULARI + IK_BASLIKLARI)
                },
                **{
                    k: _extract_score_from_record(r, k, lbl)
                    for k, lbl in PUAN_BASLIKLARI
                },
            }
        )
    return _normalize_scores_on_df(pd.DataFrame(rows_out))


def _sync_action_plan_to_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = _normalize_scores_on_df(df.copy())
    if "aksiyon_tamamlandi" not in out.columns:
        out["aksiyon_tamamlandi"] = False
    if "aksiyon_tamamlanma_tarihi" not in out.columns:
        out["aksiyon_tamamlanma_tarihi"] = ""
    if "aksiyon_tamamlanma_notu" not in out.columns:
        out["aksiyon_tamamlanma_notu"] = ""
    if "lokasyon" not in out.columns:
        out["lokasyon"] = ""

    for idx, row in out.iterrows():
        sic = str(row.get("sicil_no", "") or "").strip()
        out.at[idx, "lokasyon"] = _normalize_location_value(row.get("lokasyon", ""), sicil_no=sic, fallback_idx=idx)
        done_note = str(row.get("aksiyon_tamamlanma_notu", "") or "").strip()
        done_flag = _is_completed_flag(row.get("aksiyon_tamamlandi", False)) and _has_completion_note(done_note)

        if sic in AKSIYON_PLANI_MAP:
            out.at[idx, "aksiyon_ihtiyaci"] = "Evet"
            cur_note = str(row.get("aksiyon_notu", "") or "").strip()
            if not _has_meaningful_action_note(cur_note):
                out.at[idx, "aksiyon_notu"] = AKSIYON_PLANI_MAP[sic]
            out.at[idx, "aksiyon_tamamlandi"] = done_flag
            out.at[idx, "aksiyon_tamamlanma_notu"] = done_note
            out.at[idx, "aksiyon_tamamlanma_tarihi"] = (
                _format_tr_date_input(row.get("aksiyon_tamamlanma_tarihi", "")) if done_flag else ""
            )
        else:
            has_other_aks = str(row.get("aksiyon_ihtiyaci", "")).strip() == "Evet"
            out.at[idx, "aksiyon_ihtiyaci"] = "Evet" if has_other_aks else "Hayır"
            out.at[idx, "aksiyon_tamamlandi"] = done_flag if has_other_aks else False
            out.at[idx, "aksiyon_tamamlanma_tarihi"] = (
                _format_tr_date_input(row.get("aksiyon_tamamlanma_tarihi", ""))
                if (has_other_aks and done_flag)
                else ""
            )
    return out


@st.cache_data(show_spinner=False)
def _load_file_sample_cached() -> str:
    base_dir = os.path.dirname(globals().get("__file__", ""))
    for cand in [
        "records_126_compact.json",
        "records_126.json",
        os.path.join(base_dir, "records_126_compact.json") if base_dir else "",
    ]:
        if cand and os.path.exists(cand):
            with open(cand, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, list) and len(loaded) > 0:
                df_m = _sync_action_plan_to_df(_map_records_to_df(loaded))
                return df_m.to_json(orient="records", force_ascii=False)
    return "[]"


def get_sample_data() -> pd.DataFrame:
    raw_json = _load_file_sample_cached()
    if not raw_json or raw_json == "[]":
        return pd.DataFrame([])
    return _sync_action_plan_to_df(pd.read_json(io.StringIO(raw_json), orient="records"))


def load_records() -> tuple[pd.DataFrame, str]:
    if "local_df" not in st.session_state:
        st.session_state.local_df = get_sample_data()
    else:
        st.session_state.local_df = _sync_action_plan_to_df(st.session_state.local_df)

    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            ws = spreadsheet.worksheet("gorusmeler")
            data = ws.get_all_records()
            if data:
                df = _normalize_scores_on_df(pd.DataFrame(data))
                df = _sync_action_plan_to_df(df)
                return df, f"Google Sheets ({len(df)} Kayıt)"
            return pd.DataFrame(columns=st.session_state.local_df.columns), "Google Sheets (0 Kayıt)"
        except Exception as e:
            return st.session_state.local_df.copy(), f"Yerel Havuz ({len(st.session_state.local_df)} Kayıt) • {e}"

    return st.session_state.local_df.copy(), f"Yerel Havuz ({len(st.session_state.local_df)} Kayıt)"


def update_action_completion_record(
    kayit_id: str,
    sicil_no: str,
    completed: bool,
    completion_date: str,
    completion_note: str,
) -> tuple[bool, str]:
    clean_note = (completion_note or "").strip()
    clean_date = _format_tr_date_input(completion_date) if completed else ""

    if completed and not _has_completion_note(clean_note):
        return False, "Aksiyonu kapatmak için nasıl tamamlandığına dair not girilmelidir."

    df_curr = st.session_state.local_df.copy()
    matched = False
    if not df_curr.empty:
        if kayit_id and "kayit_id" in df_curr.columns and (df_curr["kayit_id"].astype(str) == str(kayit_id)).any():
            mask = df_curr["kayit_id"].astype(str) == str(kayit_id)
            df_curr.loc[mask, "aksiyon_tamamlandi"] = bool(completed)
            df_curr.loc[mask, "aksiyon_tamamlanma_tarihi"] = clean_date
            df_curr.loc[mask, "aksiyon_tamamlanma_notu"] = clean_note
            matched = True
        elif sicil_no and "sicil_no" in df_curr.columns and (df_curr["sicil_no"].astype(str) == str(sicil_no)).any():
            mask = df_curr["sicil_no"].astype(str) == str(sicil_no)
            df_curr.loc[mask, "aksiyon_tamamlandi"] = bool(completed)
            df_curr.loc[mask, "aksiyon_tamamlanma_tarihi"] = clean_date
            df_curr.loc[mask, "aksiyon_tamamlanma_notu"] = clean_note
            matched = True
    if matched:
        st.session_state.local_df = df_curr

    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            ws = spreadsheet.worksheet("gorusmeler")
            headers = ws.row_values(1)
            for extra_col in ["aksiyon_tamamlandi", "aksiyon_tamamlanma_tarihi", "aksiyon_tamamlanma_notu"]:
                if extra_col not in headers and headers:
                    ws.update_cell(1, len(headers) + 1, extra_col)
                    headers.append(extra_col)
            if "kayit_id" in headers:
                id_idx = headers.index("kayit_id")
                done_idx = headers.index("aksiyon_tamamlandi") + 1
                dt_idx = headers.index("aksiyon_tamamlanma_tarihi") + 1
                nt_idx = headers.index("aksiyon_tamamlanma_notu") + 1
                all_vals = ws.get_all_values()
                for r_i, r_vals in enumerate(all_vals[1:], start=2):
                    if len(r_vals) > id_idx and str(r_vals[id_idx]).strip() == str(kayit_id).strip():
                        ws.update_cell(r_i, done_idx, "Evet" if completed else "Hayır")
                        ws.update_cell(r_i, dt_idx, clean_date)
                        ws.update_cell(r_i, nt_idx, clean_note)
                        break
        except Exception:
            pass

    return True, "Aksiyon kapatma bilgisi kaydedildi."


def save_or_update_record(record: dict, edit_kayit_id: str | None = None) -> tuple[bool, str]:
    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            try:
                ws = spreadsheet.worksheet("gorusmeler")
            except Exception:
                ws = spreadsheet.add_worksheet(title="gorusmeler", rows=1000, cols=75)
            headers = ws.row_values(1)
            if not headers:
                headers = list(record.keys())
                ws.append_row(headers)

            if edit_kayit_id and "kayit_id" in headers:
                all_vals = ws.get_all_values()
                id_col_idx = headers.index("kayit_id")
                for r_idx, row_vals in enumerate(all_vals[1:], start=2):
                    if len(row_vals) > id_col_idx and row_vals[id_col_idx] == edit_kayit_id:
                        updated_row = [record.get(h, "") for h in headers]
                        ws.update(f"A{r_idx}", [updated_row])
                        return True, f"Kayıt ({edit_kayit_id}) Google Sheets üzerinde güncellendi."

            row_values = [record.get(h, "") for h in headers]
            ws.append_row(row_values)
            return True, "Yeni görüşme kaydı Google Sheets'e eklendi."
        except Exception as e:
            return False, f"Google Sheets yazma hatası: {e}"

    df_curr = st.session_state.local_df.copy()
    if edit_kayit_id and not df_curr.empty and "kayit_id" in df_curr.columns and (df_curr["kayit_id"] == edit_kayit_id).any():
        idx_target = df_curr[df_curr["kayit_id"] == edit_kayit_id].index[0]
        for k, v in record.items():
            df_curr.at[idx_target, k] = v
        st.session_state.local_df = _normalize_scores_on_df(df_curr)
        return True, f"Kayıt ({edit_kayit_id}) başarıyla güncellendi."

    new_row = pd.DataFrame([record])
    st.session_state.local_df = _normalize_scores_on_df(pd.concat([df_curr, new_row], ignore_index=True))
    return True, "Yeni görüşme kaydı havuza eklendi."


def build_excel_template() -> bytes:
    bio = io.BytesIO()
    cols = [
        "Sicil",
        "Ad Soyad",
        "Görüşme Tarihi",
        "Dönem",
        "Genel Müdür Yardımcılığı",
        "Departman",
        "Birim",
        "Bölüm",
        "Yönetici Bilgisi",
        "Lokasyon",
        "İşe Giriş Tarihi",
        "İK İş Ortağı",
        "Çalışma Modeli",
        "Duygu Durumu",
        "Aksiyon planı gerekli mi?",
        "Aksiyon Planı / Detay",
        "Aksiyon Tamamlandı",
        "Aksiyon Kapatılma Tarihi",
        "Aksiyon Kapatma Yorumu / Notu",
    ] + [q_lbl for _, q_lbl in TUM_METIN_SORULARI] + [p_lbl for _, p_lbl in PUAN_BASLIKLARI]

    sample_row = {
        "Sicil": "1001",
        "Ad Soyad": "Çalışan 1",
        "Görüşme Tarihi": now_local().date().strftime("%d.%m.%Y"),
        "Dönem": "2026 - 2. Dönem",
        "Genel Müdür Yardımcılığı": "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri",
        "Departman": "Yazılım Geliştirme Departmanı",
        "Birim": "10000474-Kurumsal Uygulamalar",
        "Bölüm": "Ölçme ve Raporlama",
        "Yönetici Bilgisi": "Birim Yöneticisi - 1",
        "Lokasyon": "ANKARA TEKNOKENT",
        "İşe Giriş Tarihi": "14.03.2022",
        "İK İş Ortağı": "Şeyma Oruç",
        "Çalışma Modeli": "Hibrit",
        "Duygu Durumu": "Olumlu",
        "Aksiyon planı gerekli mi?": "Evet",
        "Aksiyon Planı / Detay": "Ekip değişikliği talebi bulunmakta, toplantı",
        "Aksiyon Tamamlandı": "Hayır",
        "Aksiyon Kapatılma Tarihi": "",
        "Aksiyon Kapatma Yorumu / Notu": "",
        "Genel Memnuniyet": 4.5,
        "Ekip Motivasyonu": 4.0,
    }
    df_tmpl = pd.DataFrame([sample_row], columns=cols)
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_tmpl.to_excel(writer, sheet_name="Görüşme Şablonu", index=False)
    bio.seek(0)
    return bio.getvalue()


def import_excel_records(uploaded_file) -> tuple[int, str]:
    try:
        df_up = pd.read_excel(uploaded_file)
        if df_up.empty:
            return 0, "Yüklenen Excel dosyası boş."

        reverse_labels = {v: k for k, v in COLUMN_LABELS.items()}
        records_raw = []
        for _, row in df_up.iterrows():
            mapped_row = {}
            for col_name in df_up.columns:
                c_str = str(col_name).strip()
                internal_key = reverse_labels.get(c_str, c_str)
                mapped_row[internal_key] = row[col_name]
            records_raw.append(mapped_row)

        df_mapped = _sync_action_plan_to_df(_map_records_to_df(records_raw))
        if df_mapped.empty:
            return 0, "Geçerli satır dönüştürülemedi."

        st.session_state.local_df = pd.concat([st.session_state.local_df, df_mapped], ignore_index=True)
        return len(df_mapped), f"{len(df_mapped)} görüşme kaydı Excel'den havuza başarıyla aktarıldı."
    except Exception as e:
        return 0, f"Excel okuma hatası: {e}"


def build_excel_export(df_raw: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    df_clean = _sync_action_plan_to_df(df_raw)
    if "aksiyon_tamamlandi" in df_clean.columns:
        df_clean["aksiyon_tamamlandi"] = df_clean["aksiyon_tamamlandi"].apply(
            lambda x: "Evet" if _is_completed_flag(x) else "Hayır"
        )

    ordered_cols = [c for c in COLUMN_LABELS.keys() if c in df_clean.columns]
    extra_cols = [c for c in df_clean.columns if c not in ordered_cols]
    df_clean = df_clean[ordered_cols + extra_cols]
    df_export = df_clean.rename(columns=COLUMN_LABELS)

    puan_cols = [c for c, _ in PUAN_BASLIKLARI if c in df_raw.columns]
    ozet_rows = []
    if not df_raw.empty and "birim" in df_raw.columns:
        for (gmy_ad, birim_ad), grp in df_raw.groupby(["gmy", "birim"], dropna=False):
            num_stack = (
                grp[puan_cols].apply(pd.to_numeric, errors="coerce").stack()
                if puan_cols
                else pd.Series(dtype=float)
            )
            ort_skor = round(float(num_stack.mean()), 2) if not num_stack.empty else None
            aksiyon_sayisi = int((grp["aksiyon_ihtiyaci"] == "Evet").sum()) if "aksiyon_ihtiyaci" in grp.columns else 0
            tamamlanan_sayisi = (
                int(grp["aksiyon_tamamlandi"].apply(_is_completed_flag).sum())
                if "aksiyon_tamamlandi" in grp.columns
                else 0
            )
            olumlu_sayisi = int((grp["duygu_durumu"] == "Olumlu").sum()) if "duygu_durumu" in grp.columns else 0
            olumsuz_sayisi = int((grp["duygu_durumu"] == "Olumsuz").sum()) if "duygu_durumu" in grp.columns else 0
            ozet_rows.append(
                {
                    "Genel Müdür Yardımcılığı": gmy_ad,
                    "Birim": birim_ad,
                    "Görüşme Sayısı": len(grp),
                    "Genel Puan Ortalaması (1-5)": ort_skor,
                    "Olumlu Görüşme": olumlu_sayisi,
                    "Olumsuz Görüşme": olumsuz_sayisi,
                    "Aksiyon İhtiyacı Olan": aksiyon_sayisi,
                    "Kapatılan Aksiyon": tamamlanan_sayisi,
                }
            )
    df_birim_ozet = pd.DataFrame(ozet_rows) if ozet_rows else pd.DataFrame(
        columns=["Genel Müdür Yardımcılığı", "Birim", "Görüşme Sayısı", "Genel Puan Ortalaması (1-5)", "Olumlu Görüşme", "Olumsuz Görüşme", "Aksiyon İhtiyacı Olan", "Kapatılan Aksiyon"]
    )

    kriter_rows = []
    for col_k, lbl_k in PUAN_BASLIKLARI:
        if col_k in df_raw.columns:
            s_num = pd.to_numeric(df_raw[col_k], errors="coerce").dropna()
            kriter_rows.append(
                {
                    "Değerlendirme Kriteri": lbl_k,
                    "Puanlanan Kişi Sayısı": int(len(s_num)),
                    "Ortalama Puan (1-5)": round(float(s_num.mean()), 2) if not s_num.empty else None,
                    "En Düşük Puan": round(float(s_num.min()), 2) if not s_num.empty else None,
                    "En Yüksek Puan": round(float(s_num.max()), 2) if not s_num.empty else None,
                }
            )
    df_kriter_ozet = pd.DataFrame(kriter_rows)

    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_export.to_excel(writer, sheet_name="Görüşme Kayıtları", index=False)
        df_birim_ozet.to_excel(writer, sheet_name="Birim Bazlı Özet", index=False)
        df_kriter_ozet.to_excel(writer, sheet_name="Kriter Ortalamaları", index=False)

        wb = writer.book
        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        body_font = Font(name="Calibri", size=10, color="0F172A")
        thin_side = Side(style="thin", color="CBD5E1")
        cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for ws_name in wb.sheetnames:
            ws = wb[ws_name]
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.row_dimensions[1].height = 30
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = cell_border

            for row_idx in range(2, ws.max_row + 1):
                max_lines = 1
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = body_font
                    cell.border = cell_border
                    if row_idx % 2 == 0:
                        cell.fill = zebra_fill
                    val_str = str(cell.value) if cell.value is not None else ""
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if isinstance(cell.value, float):
                            cell.number_format = "0.00"
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                        est_lines = max(1, (len(val_str) // 45) + 1)
                        if est_lines > max_lines:
                            max_lines = est_lines
                ws.row_dimensions[row_idx].height = min(90, max(22, max_lines * 15))

            for col_idx in range(1, ws.max_column + 1):
                col_letter = get_column_letter(col_idx)
                header_text = str(ws.cell(row=1, column=col_idx).value or "")
                ws.column_dimensions[col_letter].width = min(44, max(16, len(header_text) + 5))

    bio.seek(0)
    return bio.getvalue()


def build_pdf_summary_report(df_filtered: pd.DataFrame, secili_gmy: str, secili_donem: str, dept_analiz: dict) -> bytes:
    from matplotlib.backends.backend_pdf import PdfPages
    import matplotlib.pyplot as plt
    import textwrap

    bio = io.BytesIO()
    with PdfPages(bio) as pdf:
        fig, ax = plt.subplots(figsize=(11.69, 8.27))
        fig.patch.set_facecolor("#0B111E")
        ax.set_facecolor("#0B111E")
        ax.axis("off")

        ax.text(0.04, 0.94, "İzlenim Paylaşım Değerlendirmesi — Yönetici Özet Raporu", color="#22D3EE", fontsize=17, fontweight="bold")
        ax.text(0.04, 0.90, f"Kapsam: {secili_gmy}   |   Dönem: {secili_donem}   |   Görüşme Sayısı: {len(df_filtered)}", color="#94A3B8", fontsize=11)

        aksiyon_sayisi = int((df_filtered["aksiyon_ihtiyaci"] == "Evet").sum()) if not df_filtered.empty and "aksiyon_ihtiyaci" in df_filtered.columns else 0
        olumsuz_sayisi = int((df_filtered["duygu_durumu"] == "Olumsuz").sum()) if not df_filtered.empty and "duygu_durumu" in df_filtered.columns else 0
        ax.text(0.04, 0.84, f"• Toplam Görüşme: {len(df_filtered)}   • Aksiyon Planı Sayısı: {aksiyon_sayisi}   • Olumsuz Duygu Bildirimi: {olumsuz_sayisi}", color="#F8FAFC", fontsize=11, fontweight="bold")

        y_pos = 0.77
        ax.text(0.04, y_pos, "Departman Geneli — Öne Çıkan Olumlu Yönler:", color="#10B981", fontsize=12, fontweight="bold")
        y_pos -= 0.035
        for item in (dept_analiz.get("olumlu") or [])[:4]:
            wrapped = textwrap.fill(f"• {item}", width=105)
            ax.text(0.05, y_pos, wrapped, color="#F8FAFC", fontsize=9.5, va="top")
            y_pos -= 0.032 * (wrapped.count("\n") + 1)

        y_pos -= 0.02
        ax.text(0.04, y_pos, "Departman Geneli — Öne Çıkan Gelişim / Olumsuz Yönler:", color="#F87171", fontsize=12, fontweight="bold")
        y_pos -= 0.035
        for item in (dept_analiz.get("olumsuz") or [])[:4]:
            wrapped = textwrap.fill(f"• {item}", width=105)
            ax.text(0.05, y_pos, wrapped, color="#F8FAFC", fontsize=9.5, va="top")
            y_pos -= 0.032 * (wrapped.count("\n") + 1)

        pdf.savefig(fig)
        plt.close(fig)

        aktif_birimler = [b for b in (dept_analiz.get("birimler") or []) if b.get("gorusme_sayisi", 0) > 0]
        if not aktif_birimler:
            aktif_birimler = (dept_analiz.get("birimler") or [])[:6]

        for chunk_i in range(0, max(1, len(aktif_birimler)), 3):
            fig2, ax2 = plt.subplots(figsize=(11.69, 8.27))
            fig2.patch.set_facecolor("#0B111E")
            ax2.set_facecolor("#0B111E")
            ax2.axis("off")

            ax2.text(0.04, 0.94, "Birimler Arası Detaylı Değerlendirme (Genel Metin Özeti)", color="#22D3EE", fontsize=15, fontweight="bold")
            y_b = 0.88
            for b_item in aktif_birimler[chunk_i : chunk_i + 3]:
                yon_txt = f" | 1. Yönetici İletişimi: {b_item['yonetici_iletisim_puani']}/5.00" if b_item.get("yonetici_iletisim_puani") else ""
                ax2.text(0.04, y_b, f"{b_item['birim']} ({b_item['gorusme_sayisi']} görüşme{yon_txt})", color="#A78BFA", fontsize=11.5, fontweight="bold")
                y_b -= 0.03
                ozet_wrapped = textwrap.fill(b_item.get("genel_ozet_metin", ""), width=112)
                ax2.text(0.04, y_b, ozet_wrapped, color="#F8FAFC", fontsize=9.2, va="top")
                y_b -= 0.026 * (ozet_wrapped.count("\n") + 2)

            pdf.savefig(fig2)
            plt.close(fig2)

    bio.seek(0)
    return bio.getvalue()


@st.cache_data(show_spinner=False)
def _cached_pptx_bytes(df_key: str, records_json: str, secili_gmy: str, secili_donem: str) -> bytes:
    df_filtered = pd.read_json(io.StringIO(records_json), orient="records")
    return _build_pptx_impl(df_filtered, secili_gmy, secili_donem)


def _build_pptx_impl(df_filtered: pd.DataFrame, secili_gmy: str, secili_donem: str) -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    c_bg = RGBColor(11, 17, 30)
    c_card = RGBColor(21, 31, 51)
    c_teal = RGBColor(34, 211, 238)
    c_purple = RGBColor(139, 92, 246)
    c_white = RGBColor(248, 250, 252)
    c_muted = RGBColor(148, 163, 184)

    def set_bg(slide, color=c_bg):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, slide_num: int, total_slides: int, title: str, subtitle: str = ""):
        set_bg(slide)
        bar = slide.shapes.add_shape(1, Inches(0.6), Inches(0.45), Inches(0.12), Inches(0.55))
        bar.fill.solid()
        bar.fill.fore_color.rgb = c_teal if slide_num % 2 == 1 else c_purple
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.85), Inches(0.38), Inches(10.5), Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = c_white

        if subtitle:
            p2 = tf.add_paragraph()
            p2.text = subtitle
            p2.font.size = Pt(11)
            p2.font.color.rgb = c_muted

        ft = slide.shapes.add_textbox(Inches(0.6), Inches(7.0), Inches(12.1), Inches(0.35))
        fp = ft.text_frame.paragraphs[0]
        fp.text = f"İzlenim Paylaşım Değerlendirmesi  |  Slayt {slide_num}/{total_slides}"
        fp.font.size = Pt(9.5)
        fp.font.color.rgb = c_muted

    def add_card(slide, left, top, width, height, heading, body_lines, accent=c_teal, font_pt=10.5):
        shp = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
        shp.fill.solid()
        shp.fill.fore_color.rgb = c_card
        shp.line.color.rgb = accent
        shp.line.width = Pt(1.25)

        tf = shp.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.14)

        p0 = tf.paragraphs[0]
        p0.text = heading
        p0.font.size = Pt(13.0)
        p0.font.bold = True
        p0.font.color.rgb = accent
        p0.space_after = Pt(5)

        for line in body_lines:
            p = tf.add_paragraph()
            p.text = f"• {line}" if not line.startswith("•") else line
            p.font.size = Pt(font_pt)
            p.font.color.rgb = c_white
            p.space_after = Pt(3.5)

    dept_analiz = _analyze_department_impl(df_filtered, secili_gmy)
    ekip_4 = _compute_4_team_criteria(df_filtered)
    aktif_birimler = [b for b in dept_analiz["birimler"] if b["gorusme_sayisi"] > 0] or dept_analiz["birimler"][:8]
    birim_gruplari = [aktif_birimler[i : i + 4] for i in range(0, len(aktif_birimler), 4)]
    ozet_gruplari = [aktif_birimler[i : i + 2] for i in range(0, len(aktif_birimler), 2)]

    total_slides = 4 + max(1, len(birim_gruplari)) + max(1, len(ozet_gruplari))
    slide_counter = 1

    n_kisi = len(df_filtered)
    aksiyon_adet = int((df_filtered["aksiyon_ihtiyaci"] == "Evet").sum()) if "aksiyon_ihtiyaci" in df_filtered.columns and not df_filtered.empty else 0
    olumsuz_adet = int((df_filtered["duygu_durumu"] == "Olumsuz").sum()) if "duygu_durumu" in df_filtered.columns and not df_filtered.empty else 0
    olumlu_adet = int((df_filtered["duygu_durumu"] == "Olumlu").sum()) if "duygu_durumu" in df_filtered.columns and not df_filtered.empty else 0

    s1 = prs.slides.add_slide(blank_layout)
    add_header(s1, slide_counter, total_slides, "Yönetici Özeti — Duygu ve Aksiyon Durumu", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    add_card(
        s1, 0.6, 1.35, 5.9, 5.2, "Görüşme, Duygu ve Aksiyon Özeti",
        [
            f"Toplam Görüşülen Kişi: {n_kisi}",
            f"Aksiyon Planı Oluşturulan Görüşme: {aksiyon_adet}",
            f"Olumlu Duygu Bildirimi: {olumlu_adet} kişi",
            f"Olumsuz Duygu Bildirimi: {olumsuz_adet} kişi",
        ],
        c_teal,
    )
    add_card(
        s1, 6.8, 1.35, 5.9, 5.2, "Departman Geneli Öne Çıkan Bulgular",
        (dept_analiz["olumlu"][:2] + dept_analiz["olumsuz"][:2]) or ["Veri bulunmuyor"],
        c_purple,
    )

    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, slide_counter, total_slides, "İzlenim Paylaşım Puanlaması (2025 vs 2026 Karşılaştırması)", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1

    def _m(col):
        if df_filtered.empty or col not in df_filtered.columns:
            return "-"
        s = pd.to_numeric(df_filtered[col], errors="coerce").dropna()
        return f"{round(float(s.mean()), 2)}" if not s.empty else "-"

    comp_lines = [
        f"Softtech'e Aidiyet / Genel Memnuniyet: 2025 = 4.10  |  2026 = {_m('p_genel_memnuniyet')}",
        f"Ücret Motivasyonu: 2025 = 3.40  |  2026 = {_m('p_ucret_yan_hak')}",
        f"Ekip Motivasyonu: 2025 = 3.50  |  2026 = {_m('p_ekip_motivasyonu')}",
        f"Kendi Motivasyonu: 2025 = 3.70  |  2026 = {_m('p_kendi_motivasyonu')}",
        f"İş-Yaşam Dengesi: 2025 = 3.80  |  2026 = {_m('p_is_yasam_dengesi')}",
        f"Direktörüyle İletişim: 2025 = 3.40  |  2026 = {_m('p_direktor_iletisimi')}",
        f"Üst Yönetim (GMY) İletişimi: 2025 = 3.80  |  2026 = {_m('p_gmy_iletisimi')}",
    ]
    add_card(s2, 0.6, 1.35, 12.1, 5.2, "2025 ve 2026 Karşılaştırmalı Sonuçlar (1-5 Puan)", comp_lines, c_teal, font_pt=12.0)

    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, slide_counter, total_slides, "İzlenim Paylaşım Değerlendirme Kriterleri — Ekip (4 Başlık)", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    for i, c4 in enumerate(ekip_4):
        col_x = 0.6 if i % 2 == 0 else 6.8
        row_y = 1.35 + (i // 2) * 2.7
        add_card(
            s3, col_x, row_y, 5.9, 2.45, c4["baslik"],
            [c4["aciklama"], c4["gosterge"], f"Değerlendirme: {c4['durum']}"],
            c_teal if i % 2 == 0 else c_purple, font_pt=10.0,
        )

    for b_idx, b_grp in enumerate(birim_gruplari, start=1):
        s_b = prs.slides.add_slide(blank_layout)
        add_header(s_b, slide_counter, total_slides, f"Birimler Arası Değerlendirme & Yönetici İletişimi ({b_idx}/{len(birim_gruplari)})", f"{secili_gmy}  |  Dönem: {secili_donem}")
        slide_counter += 1
        for idx_in_page, b_item in enumerate(b_grp):
            col_x = 0.6 if idx_in_page % 2 == 0 else 6.8
            row_y = 1.32 + (idx_in_page // 2) * 2.72
            yon_str = f" | Yönetici İletişimi: {b_item['yonetici_iletisim_puani']}/5" if b_item.get("yonetici_iletisim_puani") else ""
            add_card(
                s_b, col_x, row_y, 5.9, 2.52,
                f"{b_item['birim']} ({b_item['gorusme_sayisi']} görüşme{yon_str})",
                [
                    f"Olumlu: {' / '.join(b_item['olumlu'])}",
                    f"Olumsuz: {' / '.join(b_item['olumsuz'])}",
                ],
                c_purple if "Dikkat" in b_item["durum"] else c_teal, font_pt=9.5,
            )

    for o_idx, o_grp in enumerate(ozet_gruplari, start=1):
        s_o = prs.slides.add_slide(blank_layout)
        add_header(s_o, slide_counter, total_slides, f"Birimler Arası Detaylı Değerlendirme — Metin Özeti ({o_idx}/{len(ozet_gruplari)})", f"{secili_gmy}  |  Dönem: {secili_donem}")
        slide_counter += 1
        for slot_i, b_item in enumerate(o_grp):
            top_y = 1.32 + slot_i * 2.75
            add_card(
                s_o, 0.6, top_y, 12.1, 2.55,
                f"{b_item['birim']} — Genel Değerlendirme Özeti",
                [b_item.get("genel_ozet_metin", "")],
                c_teal if slot_i % 2 == 0 else c_purple, font_pt=9.5,
            )

    s_last = prs.slides.add_slide(blank_layout)
    add_header(s_last, slide_counter, total_slides, "İK Önerileri ve Aksiyon Planları", f"{secili_gmy}  |  Dönem: {secili_donem}")
    aksiyonlar = df_filtered[df_filtered["aksiyon_ihtiyaci"] == "Evet"] if "aksiyon_ihtiyaci" in df_filtered.columns and not df_filtered.empty else pd.DataFrame()
    aksiyon_satirlari = []
    for _, r in aksiyonlar.head(6).iterrows():
        not_metni = str(r.get("aksiyon_notu", "") or "").strip()
        bir = str(r.get("birim", ""))
        kisi = str(r.get("ad_soyad", "") or r.get("calisan_rumuzu", ""))
        is_done = _is_completed_flag(r.get("aksiyon_tamamlandi", False))
        done_note = str(r.get("aksiyon_tamamlanma_notu", "") or "").strip()
        done_dt = str(r.get("aksiyon_tamamlanma_tarihi", "") or "").strip()
        done_tag = (
            f"[Kapatıldı{f' - {done_dt}' if done_dt else ''} | Yorum: {done_note}] "
            if (is_done and done_note)
            else ("[Açık] " if not is_done else "")
        )
        aksiyon_satirlari.append(f"{done_tag}[{bir} | {kisi}] {not_metni}")

    add_card(s_last, 0.6, 1.35, 12.1, 5.2, "Takip Edilecek Aksiyon Maddeleri", aksiyon_satirlari or ["Seçilen filtrede aksiyon planı bulunmuyor."], c_teal)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()


def render_izlenim_paylasim_module(
    df_all: pd.DataFrame,
    assigned_gmy_map: dict,
    gmy_options: list[str],
    mood_slot,
):
    allowed_birimler = [b for b_list in assigned_gmy_map.values() for b in b_list]
    df_scoped = df_all.copy()

    st.markdown("#### 🔎 Detaylı Filtreleme")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        sel_gmy_f = st.selectbox("Genel Müdür Yardımcılığı", ["Tümü"] + gmy_options, key="ip_sel_gmy")
    with f2:
        dept_vals = sorted(df_scoped["departman"].dropna().unique().tolist()) if "departman" in df_scoped.columns and not df_scoped.empty else DEPARTMAN_OPTIONS
        sel_dept_f = st.selectbox("Departman", ["Tümü"] + dept_vals, key="ip_sel_dept")
    with f3:
        if sel_gmy_f == "Tümü":
            veri_birimler = sorted(df_scoped["birim"].dropna().unique().tolist()) if "birim" in df_scoped.columns and not df_scoped.empty else []
            birim_filter_opts = ["Tümü"] + list(dict.fromkeys(allowed_birimler + veri_birimler))
        else:
            veri_gmy_birimler = (
                sorted(df_scoped[df_scoped["gmy"] == sel_gmy_f]["birim"].dropna().unique().tolist())
                if "gmy" in df_scoped.columns and "birim" in df_scoped.columns and not df_scoped.empty
                else []
            )
            birim_filter_opts = ["Tümü"] + list(dict.fromkeys(assigned_gmy_map.get(sel_gmy_f, []) + veri_gmy_birimler))
        sel_birim_f = st.selectbox("Birim", birim_filter_opts, key="ip_sel_birim")
    with f4:
        bolum_vals = sorted(df_scoped["bolum"].dropna().unique().tolist()) if "bolum" in df_scoped.columns and not df_scoped.empty else BOLUM_OPTIONS
        sel_bolum_f = st.selectbox("Bölüm", ["Tümü"] + bolum_vals, key="ip_sel_bolum")

    f5, f6, f7, f8 = st.columns(4)
    with f5:
        tarih_vals = sorted(df_scoped["gorusme_tarihi"].dropna().unique().tolist()) if "gorusme_tarihi" in df_scoped.columns and not df_scoped.empty else []
        sel_tarih_f = st.selectbox("Görüşme Tarihi", ["Tümü"] + tarih_vals, key="ip_sel_tarih")
    with f6:
        donem_vals = sorted(df_scoped["donem"].dropna().unique().tolist()) if "donem" in df_scoped.columns and not df_scoped.empty else ["2026 - 2. Dönem"]
        sel_donem_f = st.selectbox("Görüşme Dönemi", ["Tümü"] + donem_vals, key="ip_sel_donem")
    with f7:
        ik_vals = sorted(df_scoped["ik_is_ortagi"].dropna().unique().tolist()) if "ik_is_ortagi" in df_scoped.columns and not df_scoped.empty else IK_IS_ORTAGI_OPTIONS
        sel_ik_f = st.selectbox("İK İş Ortağı", ["Tümü"] + ik_vals, key="ip_sel_ik")
    with f8:
        sel_lok_f = st.selectbox("Lokasyon", ["Tümü"] + LOKASYON_OPTIONS, key="ip_sel_lok")

    df_view = df_scoped.copy()
    if not df_view.empty:
        filters = [
            ("gmy", sel_gmy_f),
            ("departman", sel_dept_f),
            ("birim", sel_birim_f),
            ("bolum", sel_bolum_f),
            ("gorusme_tarihi", sel_tarih_f),
            ("donem", sel_donem_f),
            ("ik_is_ortagi", sel_ik_f),
            ("lokasyon", sel_lok_f),
        ]
        for col_name, sel_val in filters:
            if sel_val != "Tümü" and col_name in df_view.columns:
                df_view = df_view[df_view[col_name] == sel_val]

    mood_slot.markdown(render_corner_mood_widget(df_view, highlight_positive=True), unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    puan_cols = [c for c, _ in PUAN_BASLIKLARI if c in df_view.columns]
    if not df_view.empty and puan_cols:
        stacked_scores = df_view[puan_cols].apply(pd.to_numeric, errors="coerce").stack()
        genel_ort = round(float(stacked_scores.mean()), 2) if not stacked_scores.empty else 0.0
    else:
        genel_ort = 0.0

    aksiyon_not_kayitlari = []
    if not df_view.empty and "aksiyon_ihtiyaci" in df_view.columns:
        df_aks = df_view[df_view["aksiyon_ihtiyaci"].astype(str).str.strip() == "Evet"]
        default_done_dt = now_local().date().strftime("%d.%m.%Y")
        for idx_aks, r_aks in df_aks.iterrows():
            sic_no = str(r_aks.get("sicil_no", "") or "").strip()
            not_txt = str(r_aks.get("aksiyon_notu", "") or "").strip() or AKSIYON_PLANI_MAP.get(sic_no, "")
            if _has_meaningful_action_note(not_txt):
                kisi_ad = str(r_aks.get("ad_soyad", "") or r_aks.get("calisan_rumuzu", "") or "Katılımcı").strip()
                birim_ad = str(r_aks.get("birim", "") or "").strip()
                gt_txt = _format_tr_date_input(r_aks.get("gorusme_tarihi", ""))
                kid_str = str(r_aks.get("kayit_id", f"row_{idx_aks}") or f"row_{idx_aks}")
                done_not_str = str(r_aks.get("aksiyon_tamamlanma_notu", "") or "").strip()
                done_dt_str = _format_tr_date_input(r_aks.get("aksiyon_tamamlanma_tarihi", "")) or default_done_dt
                is_done = _is_completed_flag(r_aks.get("aksiyon_tamamlandi", False)) and _has_completion_note(done_not_str)
                aksiyon_not_kayitlari.append(
                    {
                        "kayit_id": kid_str,
                        "sicil_no": sic_no,
                        "kisi": f"{kisi_ad} ({sic_no})" if sic_no else kisi_ad,
                        "birim": birim_ad,
                        "tarih": gt_txt,
                        "not": not_txt,
                        "tamamlandi": is_done,
                        "tamamlanma_tarihi": done_dt_str,
                        "tamamlanma_notu": done_not_str,
                    }
                )

    toplam_aksiyon_adet = len(aksiyon_not_kayitlari)
    kapali_aksiyon_adet = sum(1 for x in aksiyon_not_kayitlari if x["tamamlandi"])
    acik_aksiyon_adet = max(0, toplam_aksiyon_adet - kapali_aksiyon_adet)

    if not df_view.empty and "duygu_durumu" in df_view.columns:
        duygu_vc = df_view["duygu_durumu"].astype(str).str.strip().value_counts()
        olumlu_adet = int(duygu_vc.get("Olumlu", 0))
        olumsuz_adet = int(duygu_vc.get("Olumsuz", 0))
        notr_adet = max(0, len(df_view) - olumlu_adet - olumsuz_adet)
    else:
        olumlu_adet = olumsuz_adet = notr_adet = 0

    olumsuz_sub = "Olumsuz kayıt yok" if olumsuz_adet == 0 else f"Olumsuz: {olumsuz_adet} kayıt"

    for col_box, lbl, val, sub in [
        (k1, "Toplam Görüşülen Kişi", str(len(df_view)), "Seçili filtre kapsamı"),
        (k2, "Genel Puan Ortalaması", f"{genel_ort} / 5.00", "Buçuklu puanlar dahil"),
        (k3, "Toplam Oluşan Aksiyon", str(toplam_aksiyon_adet), f"Açık: {acik_aksiyon_adet} • Kapalı: {kapali_aksiyon_adet}"),
        (k4, "Duygu Dağılımı (Olumlu / Nötr)", f"{olumlu_adet} / {notr_adet}", olumsuz_sub),
    ]:
        col_box.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">{lbl}</div>
                <div class="metric-value">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    df_key = _df_fingerprint(df_view)
    df_json = _df_to_json_safe(df_view)
    dept_res = _cached_department_analysis(df_key, df_json, sel_gmy_f)

    exp_c1, exp_c2 = st.columns(2)
    with exp_c1:
        pptx_state_key = f"pptx_{df_key}_{sel_gmy_f}_{sel_donem_f}"
        if st.button("📊 Sunumu Hazırla (.pptx)", use_container_width=True):
            with st.spinner("Sunum oluşturuluyor..."):
                st.session_state[pptx_state_key] = _cached_pptx_bytes(df_key, df_json, sel_gmy_f, sel_donem_f)
        if pptx_state_key in st.session_state:
            st.download_button(
                label="📥 Sunumu İndir (.pptx)",
                data=st.session_state[pptx_state_key],
                file_name=f"Izlenim_Paylasim_{sel_donem_f.replace(' ', '_')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    with exp_c2:
        pdf_state_key = f"pdf_{df_key}_{sel_gmy_f}_{sel_donem_f}"
        if st.button("📄 Raporu PDF Olarak Hazırla (.pdf)", use_container_width=True):
            with st.spinner("PDF raporu oluşturuluyor..."):
                st.session_state[pdf_state_key] = build_pdf_summary_report(df_view, sel_gmy_f, sel_donem_f, dept_res)
        if pdf_state_key in st.session_state:
            st.download_button(
                label="📥 PDF Raporunu İndir (.pdf)",
                data=st.session_state[pdf_state_key],
                file_name=f"Izlenim_Paylasim_{sel_donem_f.replace(' ', '_')}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    if df_view.empty:
        st.info("Seçili filtrelere uygun görüşme kaydı bulunamadı.")
        return

    st.subheader("📊 Aksiyon Durumu")
    df_aks_durum = pd.DataFrame({
        "Kategori": ["Kapalı Aksiyon", "Açık Aksiyon", "Toplam Aksiyon"],
        "Adet": [kapali_aksiyon_adet, acik_aksiyon_adet, toplam_aksiyon_adet],
    })
    max_aks_x = max(5, int(toplam_aksiyon_adet * 1.18))
    fig_aks = px.bar(
        df_aks_durum,
        x="Adet",
        y="Kategori",
        text="Adet",
        orientation="h",
        color="Kategori",
        range_x=[0, max_aks_x],
        color_discrete_map={
            "Toplam Aksiyon": TEAL,
            "Açık Aksiyon": PURPLE_LIGHT,
            "Kapalı Aksiyon": POS_COLOR,
        },
    )
    fig_aks.update_traces(textposition="outside", textfont=dict(size=13, color=TEXT_LIGHT), cliponaxis=False)
    fig_aks.update_layout(
        paper_bgcolor=CARD_DARK,
        plot_bgcolor=CARD_DARK,
        font_color=TEXT_LIGHT,
        showlegend=False,
        xaxis_title="Aksiyon Adedi",
        yaxis_title="",
        height=220,
        margin=dict(l=15, r=35, t=15, b=20),
    )
    st.plotly_chart(fig_aks, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if "aksiyon_gorunum_modu" not in st.session_state:
        st.session_state.aksiyon_gorunum_modu = "acik"

    acik_liste = [x for x in aksiyon_not_kayitlari if not x["tamamlandi"]]
    kapali_liste = [x for x in aksiyon_not_kayitlari if x["tamamlandi"]]

    st.markdown(
        f"""
        <div class="action-showcase-banner">
            <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:0.5rem;">
                <div>
                    <div style="font-size:1.25rem; font-weight:800; color:#F8FAFC; letter-spacing:-0.2px;">⚡ ÖNE ÇIKAN AKSİYON PLANI</div>
                    <div style="font-size:0.88rem; color:#CBD5E1; margin-top:0.15rem;">Aşağıdaki görünümlerden Açık veya Kapatılan Aksiyonları seçerek detayları inceleyebilirsiniz</div>
                </div>
                <div style="background:rgba(167,139,250,0.2); border:1px solid rgba(167,139,250,0.6); color:#DDD6FE; font-weight:800; font-size:0.9rem; padding:0.35rem 0.75rem; border-radius:999px;">
                    {acik_aksiyon_adet} Açık • {kapali_aksiyon_adet} Kapatılan ({toplam_aksiyon_adet} Toplam)
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    btn_col_acik, btn_col_kapali = st.columns(2)
    with btn_col_acik:
        acik_aktif_tag = " [Aktif Görünüm]" if st.session_state.aksiyon_gorunum_modu == "acik" else ""
        if st.button(f"📂 Açık Aksiyonlar ({acik_aksiyon_adet}){acik_aktif_tag}", key="btn_show_open_actions", use_container_width=True):
            st.session_state.aksiyon_gorunum_modu = "acik"
            st.rerun()
    with btn_col_kapali:
        kapali_aktif_tag = " [Aktif Görünüm]" if st.session_state.aksiyon_gorunum_modu == "kapali" else ""
        if st.button(f"✅ Kapatılan Aksiyonlar ({kapali_aksiyon_adet}){kapali_aktif_tag}", key="btn_show_closed_actions", use_container_width=True):
            st.session_state.aksiyon_gorunum_modu = "kapali"
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    if st.session_state.aksiyon_gorunum_modu == "acik":
        if acik_liste:
            for item_aks in acik_liste:
                meta_parcalar = [p for p in [item_aks["kisi"], item_aks["birim"], item_aks["tarih"]] if p]
                meta_str = " • ".join(meta_parcalar)
                uid_tag = f"{item_aks['kayit_id']}_{item_aks['sicil_no']}"
                chk_key = f"aks_done_chk_{uid_tag}"
                dt_key = f"aks_done_dt_{uid_tag}"
                nt_key = f"aks_done_nt_{uid_tag}"
                err_key = f"aks_done_err_{uid_tag}"
                reset_flag_key = f"aks_reset_chk_{uid_tag}"

                if st.session_state.pop(reset_flag_key, False):
                    st.session_state[chk_key] = False
                    st.session_state[nt_key] = ""

                if chk_key not in st.session_state:
                    st.session_state[chk_key] = False
                if dt_key not in st.session_state:
                    st.session_state[dt_key] = now_local().date()
                if nt_key not in st.session_state:
                    st.session_state[nt_key] = ""

                row_txt_col, row_chk_col = st.columns([4.4, 1.3], vertical_alignment="top")
                with row_txt_col:
                    st.markdown(
                        f"""
                        <div class="action-note-box">
                            <div style="font-size:0.82rem; color:{TEAL}; font-weight:800; margin-bottom:0.3rem;">▸ {meta_str}</div>
                            <div style="color:{TEXT_LIGHT}; font-size:0.98rem; font-weight:500; line-height:1.5;">{item_aks['not']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with row_chk_col:
                    st.checkbox("Tamamlandı", key=chk_key)

                if st.session_state.get(chk_key, False):
                    f_dt_col, f_nt_col, f_btn_col = st.columns([1.2, 3.3, 1.1], vertical_alignment="bottom")
                    with f_dt_col:
                        st.date_input("Kapatılma Tarihi", key=dt_key, format="DD.MM.YYYY")
                    with f_nt_col:
                        st.text_input(
                            "Nasıl Tamamlandı? (Kapatma Yorumu Zorunlu)",
                            key=nt_key,
                            placeholder="Örn: Yönetici ve İK ile görüşüldü, aksiyon hayata geçirildi.",
                        )
                    with f_btn_col:
                        if st.button("Aksiyonu Kapat", key=f"aks_save_btn_{uid_tag}", use_container_width=True):
                            note_input_val = str(st.session_state.get(nt_key, "") or "").strip()
                            dt_input_val = _parse_date_safe(st.session_state.get(dt_key), fallback=now_local().date()).strftime("%d.%m.%Y")
                            if not _has_completion_note(note_input_val):
                                st.session_state[reset_flag_key] = True
                                st.session_state[err_key] = "Nasıl tamamlandığına dair not girmeden aksiyon kapatılamaz!"
                                st.rerun()
                            else:
                                st.session_state.pop(err_key, None)
                                st.session_state[reset_flag_key] = True
                                update_action_completion_record(
                                    kayit_id=item_aks["kayit_id"],
                                    sicil_no=item_aks["sicil_no"],
                                    completed=True,
                                    completion_date=dt_input_val,
                                    completion_note=note_input_val,
                                )
                                st.session_state.aksiyon_gorunum_modu = "kapali"
                                st.rerun()

                    if st.session_state.get(err_key):
                        st.warning(st.session_state[err_key])
        else:
            st.info("Seçili filtre kapsamında açık aksiyon bulunmuyor.")
    else:
        if kapali_liste:
            for item_aks in kapali_liste:
                meta_parcalar = [p for p in [item_aks["kisi"], item_aks["birim"], item_aks["tarih"]] if p]
                meta_str = " • ".join(meta_parcalar)
                uid_tag = f"{item_aks['kayit_id']}_{item_aks['sicil_no']}"
                kapatilma_tarihi = item_aks["tamamlanma_tarihi"] or now_local().date().strftime("%d.%m.%Y")
                kapatma_yorumu = item_aks["tamamlanma_notu"] or "-"

                c_card, c_reopen = st.columns([4.5, 1.2], vertical_alignment="center")
                with c_card:
                    st.markdown(
                        f"""
                        <div class="action-note-box-done">
                            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.4rem; margin-bottom:0.35rem;">
                                <span style="font-size:0.84rem; color:#34D399; font-weight:800;">✓ KAPATILDI • {meta_str}</span>
                                <span style="background:rgba(16,185,129,0.18); border:1px solid rgba(16,185,129,0.5); color:#6EE7B7; font-size:0.82rem; font-weight:700; padding:0.15rem 0.6rem; border-radius:999px;">
                                    Kapatılma Tarihi: {kapatilma_tarihi}
                                </span>
                            </div>
                            <div style="color:{TEXT_LIGHT}; font-size:0.96rem; font-weight:500; line-height:1.5; margin-bottom:0.5rem;">{item_aks['not']}</div>
                            <div style="padding-top:0.45rem; border-top:1px dashed rgba(52,211,153,0.35); font-size:0.9rem; color:#A7F3D0;">
                                <b>Kapatma Yorumu / Nasıl Tamamlandı:</b> {kapatma_yorumu}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with c_reopen:
                    if st.button("↩ Yeniden Aç", key=f"reopen_aks_{uid_tag}", use_container_width=True):
                        st.session_state[f"aks_reset_chk_{uid_tag}"] = True
                        update_action_completion_record(
                            kayit_id=item_aks["kayit_id"],
                            sicil_no=item_aks["sicil_no"],
                            completed=False,
                            completion_date="",
                            completion_note="",
                        )
                        st.rerun()
        else:
            st.info("Henüz kapatılan bir aksiyon bulunmuyor.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📈 İzlenim Paylaşım Puanlaması — 2025 vs 2026 Karşılaştırması")

    def _col_avg(col_k):
        if col_k not in df_view.columns:
            return 0.0
        s = pd.to_numeric(df_view[col_k], errors="coerce").dropna()
        return round(float(s.mean()), 2) if not s.empty else 0.0

    comp_rows = [
        {"Kriter": "Genel Memnuniyet / Aidiyet", "2025 Sonuçları": 4.10, "2026 Sonuçları": _col_avg("p_genel_memnuniyet")},
        {"Kriter": "Ücret Motivasyonu", "2025 Sonuçları": 3.40, "2026 Sonuçları": _col_avg("p_ucret_yan_hak")},
        {"Kriter": "Ekip Motivasyonu", "2025 Sonuçları": 3.50, "2026 Sonuçları": _col_avg("p_ekip_motivasyonu")},
        {"Kriter": "Kendi Motivasyonu", "2025 Sonuçları": 3.70, "2026 Sonuçları": _col_avg("p_kendi_motivasyonu")},
        {"Kriter": "İş-Yaşam Dengesi", "2025 Sonuçları": 3.80, "2026 Sonuçları": _col_avg("p_is_yasam_dengesi")},
        {"Kriter": "Direktörüyle İletişim", "2025 Sonuçları": 3.40, "2026 Sonuçları": _col_avg("p_direktor_iletisimi")},
        {"Kriter": "Üst Yönetim (GMY) İletişimi", "2025 Sonuçları": 3.80, "2026 Sonuçları": _col_avg("p_gmy_iletisimi")},
    ]
    df_comp = pd.DataFrame(comp_rows)
    df_comp["Fark (2026 - 2025)"] = (df_comp["2026 Sonuçları"] - df_comp["2025 Sonuçları"]).round(2)

    c_tbl, c_grp = st.columns([1.1, 1.3])
    with c_tbl:
        st.dataframe(df_comp, use_container_width=True, hide_index=True)
    with c_grp:
        df_melt = df_comp.melt(id_vars="Kriter", value_vars=["2025 Sonuçları", "2026 Sonuçları"], var_name="Yıl", value_name="Puan")
        fig_comp = px.bar(
            df_melt,
            x="Puan",
            y="Kriter",
            color="Yıl",
            barmode="group",
            orientation="h",
            range_x=[0, 5],
            color_discrete_map={"2025 Sonuçları": PURPLE_LIGHT, "2026 Sonuçları": TEAL},
        )
        fig_comp.update_layout(paper_bgcolor=CARD_DARK, plot_bgcolor=CARD_DARK, font_color=TEXT_LIGHT, height=300, margin=dict(l=10, r=10, t=15, b=15))
        st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("İzlenim Paylaşım Değerlendirme – İçerik")

    def _mean_metric(col_name: str):
        if col_name not in df_view.columns:
            return None
        s_num = pd.to_numeric(df_view[col_name], errors="coerce").dropna()
        return round(float(s_num.mean()), 2) if not s_num.empty else None

    def _face_badge(score):
        if score is None:
            return "😐 Veri Yok"
        if score >= 4.0:
            return "😄 Güçlü / Memnun"
        if score >= 3.2:
            return "🙂 Dengeli / İyi"
        return "😟 İyileştirme Gerekli"

    mem_val = _mean_metric("p_genel_memnuniyet")
    gmy_val = _mean_metric("p_gmy_iletisimi")
    ucret_val = _mean_metric("p_ucret_yan_hak")
    yz_val = _mean_metric("p_yz_kullanim")

    hybrid_dist = _cached_hybrid_percentages(df_key, df_json)
    hybrid_summary_txt = (
        " • ".join([f"{b_n}: %{p_v:.1f}" for b_n, _, p_v in hybrid_dist[:3]])
        if hybrid_dist
        else "Yanıt dağılımı için detay açın"
    )

    content_cards = [
        {"key": "card_1_softtech", "baslik": "1. Softtech Güçlü & Gelişim Alanları", "aciklama": "Şirketin güçlü ve gelişime açık yönleri", "gosterge": f"Genel Memnuniyet: {mem_val if mem_val is not None else '-'}/5.00", "durum": _face_badge(mem_val), "accent": "teal"},
        {"key": "card_2_karar", "baslik": "2. Karar Alma Süreçleri & Strateji", "aciklama": "Karar alma hızı ve üst yönetim strateji süreçleri", "gosterge": f"GMY İletişimi: {gmy_val if gmy_val is not None else '-'}/5.00", "durum": _face_badge(gmy_val), "accent": "purple"},
        {"key": "card_3_hibrit", "baslik": "3. Hibrit Ritmi", "aciklama": "İlerisi için tercih edilen çalışma modeli dağılımı", "gosterge": hybrid_summary_txt, "durum": "📊 Yüzdelik Dağılım", "accent": "teal"},
        {"key": "card_4_ucret", "baslik": "4. Ücret & Yan Haklar", "aciklama": "Ücret, prim, yan haklar ve çalışan merkezi yorumları", "gosterge": f"Ücret/Yan Hak: {ucret_val if ucret_val is not None else '-'}/5.00", "durum": _face_badge(ucret_val), "accent": "purple"},
        {"key": "card_5_yz", "baslik": "5. Yapay Zeka Adaptasyonu", "aciklama": "CodeMantis, Amigo ve Asistan kullanım düzeyi", "gosterge": f"YZ Kullanım Sıklığı: {yz_val if yz_val is not None else '-'}/5.00", "durum": _face_badge(yz_val), "accent": "teal"},
        {"key": "card_6_harmony", "baslik": "6. Harmony Adaptasyonu", "aciklama": "Departmanın Harmony ve karne süreçlerine uyumu", "gosterge": "Harmony & karne geri bildirimleri", "durum": "🔄 Süreç Görünümü", "accent": "purple"},
        {"key": "card_7_toplantilar", "baslik": "7. Toplandığımıza Değsin", "aciklama": "Toplantı verimlilik hassasiyeti ve tutumu", "gosterge": "Toplantı verimliliği değerlendirmeleri", "durum": "⏱️ Verimlilik Görünümü", "accent": "teal"},
    ]

    if "selected_content_card" not in st.session_state:
        st.session_state.selected_content_card = None

    for r_start in range(0, len(content_cards), 3):
        r_cols = st.columns(3)
        for c_pos, c_info in enumerate(content_cards[r_start : r_start + 3]):
            is_sel = st.session_state.selected_content_card == c_info["key"]
            border_col = TEAL if c_info["accent"] == "teal" else PURPLE_LIGHT
            sel_ring = f"box-shadow: 0 0 0 2px {border_col};" if is_sel else ""
            with r_cols[c_pos]:
                st.markdown(
                    f"""
                    <div class="criteria-card{'-purple' if c_info['accent']=='purple' else ''}" style="{sel_ring} min-height: 155px; margin-bottom: 0.45rem;">
                        <div style="font-weight:700; color:{border_col}; font-size:1.0rem; margin-bottom:0.3rem;">{c_info['baslik']}</div>
                        <div style="color:{TEXT_MUTED}; font-size:0.83rem; margin-bottom:0.5rem;">{c_info['aciklama']}</div>
                        <div style="font-weight:600; color:{TEXT_LIGHT}; font-size:0.88rem; margin-bottom:0.25rem;">{c_info['gosterge']}</div>
                        <div style="font-size:0.86rem; color:{POS_COLOR}; font-weight:700;">{c_info['durum']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                btn_txt = "Kapat" if is_sel else "Detayı Aç"
                if st.button(btn_txt, key=f"btn_{c_info['key']}", use_container_width=True):
                    st.session_state.selected_content_card = None if is_sel else c_info["key"]
                    st.rerun()

    active_card = st.session_state.selected_content_card
    if active_card == "card_1_softtech":
        st.markdown("#### 1. Softtech Güçlü & Gelişim Alanları — Karşılaştırmalı Kırılım")
        g_col, k_col = st.columns(2)
        with g_col:
            st.markdown("**✅ Güçlü Alanlar**")
            pos_bullets = _cached_top_short_bullets(df_key, df_json, ("q02_guclu_yonler", "q01_calisma_hayati", "q09_vizyon_strateji"), top_n=4, direction="olumlu")
            if not pos_bullets:
                pos_bullets = _cached_top_short_bullets(df_key, df_json, ("q02_guclu_yonler", "q01_calisma_hayati"), top_n=4)
            for b_item in pos_bullets:
                st.markdown(f"- {b_item}")
        with k_col:
            st.markdown("**⚠️ Gelişime Açık Alanlar**")
            neg_bullets = _cached_top_short_bullets(df_key, df_json, ("q10_degistirmek_istenen", "q04_ekip_iyilestirme", "q02_guclu_yonler"), top_n=4, direction="olumsuz")
            for b_item in neg_bullets:
                st.markdown(f"- {b_item}")
    elif active_card == "card_2_karar":
        st.markdown("#### 2. Karar Alma Süreçleri & Strateji — Öne Çıkan Yorumlar")
        for b_item in _cached_top_short_bullets(df_key, df_json, ("q15_karar_alma_surecleri", "q09_vizyon_strateji", "q08_gmy_iletisimi"), top_n=5):
            st.markdown(f"- {b_item}")
    elif active_card == "card_3_hibrit":
        st.markdown("#### 3. Hibrit Ritmi — İlerisi İçin Tercih Edilen Çalışma Modeli (%)")
        if hybrid_dist:
            df_hyb = pd.DataFrame(hybrid_dist, columns=["Çalışma Modeli Tercihi", "Kişi Sayısı", "Oran (%)"])
            st.dataframe(df_hyb, use_container_width=True, hide_index=True)
            for m_ad, m_cnt, m_pct in hybrid_dist:
                st.markdown(f"- **{m_ad}:** **%{m_pct:.1f}** ({m_cnt} yanıt)")
        else:
            st.info("Bu başlık için çalışma modeli yanıtı bulunamadı.")
    elif active_card == "card_4_ucret":
        st.markdown("#### 4. Ücret & Yan Haklar — En Çok Tekrar Eden Yorumlar")
        for b_item in _cached_top_short_bullets(df_key, df_json, ("ik_ucret_merkezi", "q05_motivasyon_arttiran"), top_n=5):
            st.markdown(f"- {b_item}")
    elif active_card == "card_5_yz":
        st.markdown("#### 5. Yapay Zeka Adaptasyonu — En Çok Tekrar Eden Yorumlar")
        for b_item in _cached_top_short_bullets(df_key, df_json, ("q_kapanis", "q07_ym_verimlilik", "q13_harmony"), top_n=5):
            st.markdown(f"- {b_item}")
    elif active_card == "card_6_harmony":
        st.markdown("#### 6. Harmony Adaptasyonu — En Çok Tekrar Eden Yorumlar")
        for b_item in _cached_top_short_bullets(df_key, df_json, ("q13_harmony",), top_n=5):
            st.markdown(f"- {b_item}")
    elif active_card == "card_7_toplantilar":
        st.markdown("#### 7. Toplandığımıza Değsin — En Çok Tekrar Eden Yorumlar")
        for b_item in _cached_top_short_bullets(df_key, df_json, ("q14_toplandigimiza_degsin",), top_n=5):
            st.markdown(f"- {b_item}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("👥 Ekip Değerlendirme Kriterleri (4 Başlık) & Öne Çıkan Yorumlar")
    ekip_4_cards = _compute_4_team_criteria(df_view)
    cols_4 = st.columns(4)
    for idx_e, e_info in enumerate(ekip_4_cards):
        cols_4[idx_e].markdown(
            f"""
            <div class="criteria-card">
                <div style="font-weight:700; color:{TEAL}; font-size:0.96rem; margin-bottom:0.3rem;">{e_info['baslik']}</div>
                <div style="color:{TEXT_MUTED}; font-size:0.82rem; margin-bottom:0.45rem;">{e_info['aciklama']}</div>
                <div style="font-weight:600; font-size:0.88rem; margin-bottom:0.2rem;">{e_info['gosterge']}</div>
                <div style="color:{PURPLE_LIGHT}; font-weight:700; font-size:0.88rem;">{e_info['durum']}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    d_col1, d_col2 = st.columns(2)
    with d_col1:
        st.markdown("**Öne Çıkan Olumlu Noktalar (En Çok Tekrar Eden)**")
        for item in dept_res.get("olumlu") or ["Kayıtlı olumlu ifade bulunmuyor"]:
            st.markdown(f"- {item}")
    with d_col2:
        st.markdown("**Öne Çıkan Olumsuz Noktalar (En Çok Tekrar Eden)**")
        for item in dept_res.get("olumsuz") or ["Kayıtlı olumsuz ifade bulunmuyor"]:
            st.markdown(f"- {item}")

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📋 Birimler Arası Detaylı Değerlendirme & Yönetici İletişim Puanları")
    birimler_all = dept_res.get("birimler") or []
    aktif_birimler = [b for b in birimler_all if b.get("gorusme_sayisi", 0) > 0]

    if aktif_birimler:
        yon_table_rows = [
            {
                "Birim": b["birim"],
                "Görüşme Sayısı": b["gorusme_sayisi"],
                "1. Yönetici İletişimi Puanı (1-5)": b["yonetici_iletisim_puani"] if b["yonetici_iletisim_puani"] is not None else "-",
                "Genel Durum": b["durum"],
            }
            for b in aktif_birimler
        ]
        st.dataframe(pd.DataFrame(yon_table_rows), use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("📝 Birim Özelinde Tüm Yorumların Genel Metin Özeti")
    gosterilecek_birimler = aktif_birimler if aktif_birimler else birimler_all[:8]
    tum_birim_goster = st.checkbox(
        f"Kayıt olmayan birimler dahil tüm birimleri göster ({len(birimler_all)} birim)",
        value=False,
        key=f"chk_all_units_{sel_gmy_f}",
    )
    if tum_birim_goster:
        gosterilecek_birimler = birimler_all

    for b_item in gosterilecek_birimler:
        yon_badge = f" • Yönetici İletişimi: {b_item['yonetici_iletisim_puani']}/5.00" if b_item.get("yonetici_iletisim_puani") else ""
        st.markdown(
            f"""
            <div class="detail-box">
                <div style="font-weight:700; color:{TEAL}; margin-bottom:0.4rem;">
                    {b_item['birim']} &nbsp;•&nbsp; {b_item['gorusme_sayisi']} Görüşme{yon_badge} &nbsp;•&nbsp; <span style="color:{PURPLE_LIGHT};">{b_item['durum']}</span>
                </div>
                <div style="color:{TEXT_LIGHT}; line-height:1.55; margin-bottom:0.45rem;">{b_item.get('genel_ozet_metin', '')}</div>
                <div style="font-size:0.88rem; color:{TEXT_MUTED};"><b>Olumlu Özet:</b> {" / ".join(b_item['olumlu'])}</div>
                <div style="font-size:0.88rem; color:{TEXT_MUTED};"><b>Gelişim / Olumsuz Özet:</b> {" / ".join(b_item['olumsuz'])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("🖨️ Tüm Birimlere Ait Verileri Ayrı Ayrı Detaylı Metin Halinde Görüntüle / Kopyala"):
        detay_bloklar = []
        for b_item in gosterilecek_birimler:
            detay_bloklar.append(
                f"=== {b_item['birim']} ({b_item['gorusme_sayisi']} Görüşme) ===\n"
                f"Genel Değerlendirme:\n{b_item.get('genel_ozet_metin', '')}\n"
                f"Öne Çıkan Olumlu Noktalar: {' ; '.join(b_item['olumlu'])}\n"
                f"Öne Çıkan Olumsuz Noktalar: {' ; '.join(b_item['olumsuz'])}\n"
            )
        full_text_report = "\n".join(detay_bloklar)
        st.text_area("Birim Bazlı Detaylı Metin Çıktısı", value=full_text_report, height=320)
        st.download_button(
            label="📄 Detaylı Birim Metin Raporunu İndir (.txt)",
            data=full_text_report.encode("utf-8"),
            file_name=f"Birim_Detayli_Metin_Raporu_{sel_donem_f.replace(' ', '_')}.txt",
            mime="text/plain",
        )


if "user" not in st.session_state:
    st.session_state.user = None

if st.session_state.user is None:
    st.markdown(
        """
        <div class="panel-header">
            <div class="panel-brand">İzlenim Paylaşım Görüşmesi</div>
            <div style="color:#94A3B8; margin-top:0.25rem;">İK İş Ortakları için 1-1 Görüşme Veri Havuzu, Standart Dashboard ve Sunum Üretici</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_l, col_m, col_r = st.columns([1, 1.3, 1])
    with col_m:
        with st.form("login_form"):
            st.subheader("İK Sorumlusu Girişi")
            username = st.text_input("Kullanıcı Adı", placeholder="örnek: seyma veya admin")
            password = st.text_input("Şifre", type="password")
            submitted = st.form_submit_button("Giriş Yap", use_container_width=True)

            if submitted:
                u_key = username.strip().lower()
                u_info = DEFAULT_USERS.get(u_key)
                if u_info and u_info["sifre"] == password:
                    st.session_state.user = {"username": u_key, **u_info}
                    st.rerun()
                else:
                    st.error("Kullanıcı adı veya şifre hatalı.")

        st.caption("Test hesapları: `seyma` / `panel2026` veya `admin` / `admin2026`")
    st.stop()


user = st.session_state.user
if user.get("username") == "seyma" and user.get("ad_soyad") != "Şeyma Oruç":
    user["ad_soyad"] = "Şeyma Oruç"
    st.session_state.user = user

df_all, data_source_label = load_records()
emp_directory = get_employee_directory(df_all)

with st.sidebar:
    st.markdown(f"### 👤 {user['ad_soyad']}")
    st.caption(data_source_label)
    if st.button("Çıkış Yap", use_container_width=True):
        st.session_state.user = None
        st.rerun()

    st.divider()
    sayfa = st.radio(
        "Modüller",
        ["1-1 Görüşme Kayıt", "İzlenim Paylaşım Görüşmeleri", "Ham Veri Havuzu"],
    )

assigned_gmy_map = user["gmy_birimler"]
gmy_options = list(assigned_gmy_map.keys())

MODULE_HEADER_TITLES = {
    "1-1 Görüşme Kayıt": "1-1 Görüşme Kayıt",
    "İzlenim Paylaşım Görüşmeleri": "İzlenim Paylaşım Görüşmeleri | Yönetici Sunumu",
    "Ham Veri Havuzu": "Ham Veri Havuzu",
}

hdr_left, hdr_right = st.columns([1.85, 2.15], vertical_alignment="top")
with hdr_left:
    header_title = MODULE_HEADER_TITLES.get(sayfa, sayfa)
    st.markdown(
        f"""
        <div class="panel-header">
            <div class="panel-brand">{header_title}</div>
            <div style="color:#94A3B8; margin-top:0.25rem; font-size:0.95rem;">{data_source_label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with hdr_right:
    mood_slot = st.empty()
    mood_slot.markdown(render_corner_mood_widget(df_all, highlight_positive=True), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if sayfa == "1-1 Görüşme Kayıt":
    st.markdown("#### 🔍 Kayıt Seçimi veya Yeni Görüşme Başlatma")
    mode_col1, mode_col2 = st.columns([1.2, 2])
    with mode_col1:
        kayit_modu = st.radio(
            "İşlem Türü",
            ["Yeni Görüşme Ekle", "Mevcut Kaydı Düzenle / Güncelle"],
            horizontal=True,
        )

    edit_record_row = None
    edit_kayit_id = None
    if kayit_modu == "Mevcut Kaydı Düzenle / Güncelle" and not df_all.empty:
        with mode_col2:
            secenekler = []
            id_map = {}
            for _, r in df_all.iterrows():
                kid = str(r.get("kayit_id", ""))
                sic = str(r.get("sicil_no", ""))
                ad = str(r.get("ad_soyad", "") or r.get("calisan_rumuzu", ""))
                don = str(r.get("donem", ""))
                gt = _format_tr_date_input(r.get("gorusme_tarihi", ""))
                lbl = f"{sic} - {ad} | {don} ({gt}) [ID: {kid}]"
                secenekler.append(lbl)
                id_map[lbl] = kid
            secilen_lbl = st.selectbox("Güncellenecek Görüşme Kaydını Seçin", secenekler)
            edit_kayit_id = id_map.get(secilen_lbl)
            matched_df = df_all[df_all["kayit_id"] == edit_kayit_id]
            if not matched_df.empty:
                edit_record_row = matched_df.iloc[0].to_dict()

    def _pick_idx(opts: list[str], val: str, fallback: int = 0) -> int:
        if val and val in opts:
            return opts.index(val)
        return fallback

    donem_opts = [
        "2026 - 1. Dönem",
        "2026 - 2. Dönem",
        "2027 - 1. Dönem",
        "2027 - 2. Dönem",
    ]

    s_col1, s_col2, s_col3, s_col4 = st.columns(4)
    default_sicil = str(edit_record_row.get("sicil_no", "")) if edit_record_row else ""
    with s_col1:
        sicil_input = st.text_input(
            "Sicil (Örn: 1001, 1002)",
            value=default_sicil,
            placeholder="Sicil girin...",
            key="inp_sicil_no",
        ).strip()

    emp_info_pre = emp_directory.get(sicil_input) if sicil_input else None
    pre_source = dict(edit_record_row) if edit_record_row else (dict(emp_info_pre) if emp_info_pre else {})

    with s_col4:
        sec_donem = st.selectbox(
            "Görüşme Dönemi",
            donem_opts,
            index=_pick_idx(donem_opts, str(pre_source.get("donem", "2026 - 2. Dönem")), 1),
            key="inp_donem",
        )

    has_sicil_and_donem = bool(sicil_input) and bool(str(sec_donem).strip())

    existing_sicil_donem_set = set()
    if not df_all.empty and "sicil_no" in df_all.columns:
        for _, r_row in df_all.iterrows():
            s_val = str(r_row.get("sicil_no", "") or "").strip()
            d_val = str(r_row.get("donem", "") or "").strip()
            k_val = str(r_row.get("kayit_id", "") or "").strip()
            if not s_val or not d_val:
                continue
            if edit_kayit_id and k_val == str(edit_kayit_id).strip():
                continue
            existing_sicil_donem_set.add((s_val, d_val))

    is_duplicate_sicil_donem = has_sicil_and_donem and ((sicil_input, str(sec_donem).strip()) in existing_sicil_donem_set)

    context_sig = f"{kayit_modu}|{edit_kayit_id or ''}|{sicil_input}|{sec_donem}"
    if st.session_state.get("_last_form_context_sig") != context_sig:
        st.session_state["_last_form_context_sig"] = context_sig
        is_edit_ctx = edit_record_row is not None
        st.session_state["inp_duygu"] = (
            str(pre_source.get("duygu_durumu", "") or "").strip() if is_edit_ctx else "Seçiniz"
        )
        raw_aks_ctx = str(pre_source.get("aksiyon_ihtiyaci", "") or "").strip()
        st.session_state["inp_aksiyon_ihtiyac"] = (
            raw_aks_ctx if (is_edit_ctx and raw_aks_ctx in ("Evet", "Hayır")) else "Seçiniz"
        )
        st.session_state["inp_calisma_modeli"] = (
            str(pre_source.get("calisma_modeli", "") or "").strip() if is_edit_ctx else "Seçiniz"
        )

    if not has_sicil_and_donem:
        st.info("Form alanlarını görüntülemek için lütfen **Sicil** ve **Görüşme Dönemi** seçin.")
    elif is_duplicate_sicil_donem:
        st.error(
            f"**{sicil_input}** sicili için **'{sec_donem}'** döneminde zaten kayıt mevcut! "
            "Aynı dönemde 2. kayıt yapılamaz. Farklı bir dönem seçtiğinizde veya farklı bir sicil girdiğinizde diğer alanlar açılacaktır."
        )
    else:
        if emp_info_pre:
            st.success(
                f"**{sicil_input}** ({emp_info_pre['ad_soyad']}) için **'{sec_donem}'** dönemi uygun; bilgileri girebilir veya düzenleyebilirsiniz."
            )
        else:
            st.info(f"**{sicil_input}** sicili ve **'{sec_donem}'** dönemi için yeni kayıt girişi yapabilirsiniz.")

    if has_sicil_and_donem:
        prefill = dict(pre_source) if not is_duplicate_sicil_donem else {}
        if not is_duplicate_sicil_donem:
            prefill["lokasyon"] = _normalize_location_value(prefill.get("lokasyon", ""), sicil_no=sicil_input, fallback_idx=0)

        with s_col2:
            ad_soyad_input = st.text_input(
                "Ad Soyad",
                value=str(prefill.get("ad_soyad", "")),
                placeholder="Ad Soyad",
                disabled=is_duplicate_sicil_donem,
                key="inp_ad_soyad",
            )
        with s_col3:
            default_gt_date = _parse_date_safe(prefill.get("gorusme_tarihi"), fallback=now_local().date())
            gorusme_tarihi_dt = st.date_input(
                "Görüşme Tarihi",
                value=default_gt_date,
                format="DD.MM.YYYY",
                disabled=is_duplicate_sicil_donem,
                key="inp_gorusme_tarihi",
            )

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            sec_gmy = st.selectbox(
                "Genel Müdür Yardımcılığı",
                gmy_options,
                index=_pick_idx(gmy_options, str(prefill.get("gmy", "")), 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_gmy",
            )
        with c2:
            sec_departman = st.selectbox(
                "Departman",
                DEPARTMAN_OPTIONS,
                index=_pick_idx(DEPARTMAN_OPTIONS, str(prefill.get("departman", "")), 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_departman",
            )
        with c3:
            birim_options = assigned_gmy_map.get(sec_gmy, GMY_BIRIM_MAP.get(sec_gmy, []))
            sec_birim = st.selectbox(
                "Birim",
                birim_options,
                index=_pick_idx(birim_options, str(prefill.get("birim", "")), 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_birim",
            )
        with c4:
            sec_bolum = st.selectbox(
                "Bölüm",
                BOLUM_OPTIONS,
                index=_pick_idx(BOLUM_OPTIONS, str(prefill.get("bolum", "")), 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_bolum",
            )

        c5, c6, c7, c8 = st.columns(4)
        with c5:
            yon_val_init = str(prefill.get("yonetici_bilgisi", "") or "")
            yon_extra = [yon_val_init] if (yon_val_init and yon_val_init not in YONETICI_OPTIONS) else []
            yon_opts = YONETICI_OPTIONS + yon_extra
            sec_yonetici = st.selectbox(
                "Yönetici Bilgisi",
                yon_opts,
                index=_pick_idx(yon_opts, yon_val_init, 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_yonetici",
            )
        with c6:
            ise_giris_str = st.text_input(
                "İşe Giriş Tarihi",
                value=_format_tr_date_input(prefill.get("ise_giris_tarihi", "")),
                placeholder="GG.AA.YYYY",
                disabled=is_duplicate_sicil_donem,
                key="inp_ise_giris",
            )
        with c7:
            sec_ik_ortagi = st.selectbox(
                "İK İş Ortağı",
                IK_IS_ORTAGI_OPTIONS,
                index=_pick_idx(IK_IS_ORTAGI_OPTIONS, str(prefill.get("ik_is_ortagi", "")), 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_ik_ortagi",
            )
        with c8:
            loc_val_clean = _normalize_location_value(prefill.get("lokasyon", ""), sicil_no=sicil_input, fallback_idx=0)
            sec_lokasyon = st.selectbox(
                "Lokasyon",
                LOKASYON_OPTIONS,
                index=_pick_idx(LOKASYON_OPTIONS, loc_val_clean, 0),
                disabled=is_duplicate_sicil_donem,
                key="inp_lokasyon",
            )

        st.markdown("#### 🎯 Duygu Durumu & Aksiyon Planı")
        pdf_aks_detay = AKSIYON_PLANI_MAP.get(sicil_input) if (sicil_input and not is_duplicate_sicil_donem) else None
        d_col1, d_col2, d_col3 = st.columns([1.0, 1.0, 1.2])
        with d_col1:
            duygu_opts = ["Seçiniz", "Olumlu", "Nötr", "Olumsuz"]
            duygu_durumu = st.selectbox(
                "Duygu Durumu (Olumlu, Nötr, Olumsuz)",
                duygu_opts,
                disabled=is_duplicate_sicil_donem,
                key="inp_duygu",
            )
        with d_col2:
            aksiyon_opts = ["Seçiniz", "Hayır", "Evet"]
            aksiyon_ihtiyaci = st.selectbox(
                "Aksiyon planı gerekli mi?",
                aksiyon_opts,
                disabled=is_duplicate_sicil_donem,
                key="inp_aksiyon_ihtiyac",
            )
        with d_col3:
            model_opts = ["Seçiniz", "Hibrit", "Uzaktan", "Ofis / Saha"]
            calisma_modeli = st.selectbox(
                "Çalışma Modeli",
                model_opts,
                disabled=is_duplicate_sicil_donem,
                key="inp_calisma_modeli",
            )

        st.markdown("##### 📝 Aksiyon Planı / Detay")
        default_aks_note = str(prefill.get("aksiyon_notu", "") or "").strip() or (pdf_aks_detay or "")
        aksiyon_notu = st.text_area(
            "Aksiyon Planı / Detay",
            value=default_aks_note,
            height=82,
            placeholder="Örn: Ekip değişikliği talebi bulunmakta, toplantı",
            disabled=is_duplicate_sicil_donem,
            key="inp_aksiyon_plani",
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "💬 15 Görüşme Sorusu",
            "🏢 İnsan Kaynakları & Kapanış",
            "⭐ 1-5 Puanlama (Bar + Manuel Giriş)",
            "📥 Excel'den Veri İçe Aktar",
        ])

        q_answers = {}
        ik_answers = {}
        p_scores = {}

        with tab1:
            for q_key, q_text in GORUSME_SORULARI:
                q_answers[q_key] = st.text_area(
                    q_text,
                    value=str(prefill.get(q_key, "") or ""),
                    height=70,
                    disabled=is_duplicate_sicil_donem,
                    key=f"inp_{q_key}",
                )

        with tab2:
            for ik_key, ik_label in IK_BASLIKLARI:
                ik_answers[ik_key] = st.text_area(
                    f"İK Başlığı — {ik_label}",
                    value=str(prefill.get(ik_key, "") or ""),
                    height=70,
                    disabled=is_duplicate_sicil_donem,
                    key=f"inp_{ik_key}",
                )
            q_kapanis = st.text_area(
                "Paylaşmak istediğin veya eklemek istediğin farklı konu var mı?",
                value=str(prefill.get("q_kapanis", "") or ""),
                height=80,
                disabled=is_duplicate_sicil_donem,
                key="inp_q_kapanis",
            )

        with tab3:
            st.caption(
                "Puanı aşağıdaki bardan seçebilir veya sağdaki alana manuel (1.0 - 5.0 arası, 0.5 adımlı) girebilirsiniz; "
                "manuel puan girildiğinde bar otomatik olarak güncellenir."
            )
            for p_key, p_label in PUAN_BASLIKLARI:
                default_p = _snap_half_step(prefill.get(p_key, None))
                p_scores[p_key] = render_synced_score_row(p_key, p_label, default_p, "main")

        with tab4:
            st.markdown("Excel'de tutulan mevcut görüşme kayıtlarını tek seferde havuza aktarabilirsiniz.")
            tmpl_bytes = build_excel_template()
            st.download_button(
                label="📄 Örnek Excel İçe Aktarma Şablonunu İndir (.xlsx)",
                data=tmpl_bytes,
                file_name="Izlenim_Paylasim_Import_Sablonu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            uploaded_xlsx = st.file_uploader("Doldurulmuş Excel Dosyasını Seçin (.xlsx)", type=["xlsx"])
            if uploaded_xlsx is not None:
                if st.button("Excel Verilerini Havuza Aktar", type="primary"):
                    cnt_imp, msg_imp = import_excel_records(uploaded_xlsx)
                    if cnt_imp > 0:
                        st.success(msg_imp)
                        st.rerun()
                    else:
                        st.error(msg_imp)

        st.markdown("<br>", unsafe_allow_html=True)
        btn_label = "💾 Değişiklikleri Güncelle" if edit_kayit_id else "💾 Görüşme Kaydını Kaydet"
        if st.button(btn_label, type="primary", use_container_width=True, disabled=is_duplicate_sicil_donem):
            clean_sicil = sicil_input.strip()
            clean_donem = str(sec_donem).strip()
            clean_ad_girilen = ad_soyad_input.strip()
            clean_note = aksiyon_notu.strip()
            has_note = _has_meaningful_action_note(clean_note)
            clean_kapanis = q_kapanis.strip()
            clean_q_vals = [v.strip() for v in q_answers.values() if v and v.strip()]
            clean_ik_vals = [v.strip() for v in ik_answers.values() if v and v.strip()]
            has_any_score = any(v is not None for v in p_scores.values())
            has_any_content = bool(
                clean_sicil
                or clean_ad_girilen
                or has_note
                or clean_kapanis
                or clean_q_vals
                or clean_ik_vals
                or has_any_score
            )

            if clean_sicil and (clean_sicil, clean_donem) in existing_sicil_donem_set:
                st.error(f"**{clean_sicil}** sicili için **'{clean_donem}'** döneminde zaten kayıt mevcut; aynı dönemde 2. kayıt yapılamaz.")
            elif (
                not duygu_durumu
                or duygu_durumu == "Seçiniz"
                or not aksiyon_ihtiyaci
                or aksiyon_ihtiyaci == "Seçiniz"
                or not calisma_modeli
                or calisma_modeli == "Seçiniz"
            ):
                st.error("Lütfen **Duygu Durumu**, **Aksiyon planı gerekli mi?** ve **Çalışma Modeli** alanlarının tamamını seçin.")
            elif not edit_kayit_id and not has_any_content:
                st.error("Hiçbir görüşme verisi girilmedi. Kayıt oluşturmak için en az Sicil/Ad Soyad, görüşme yanıtı, not veya puan girin.")
            else:
                now_tr = now_local()
                final_kayit_id = edit_kayit_id if edit_kayit_id else now_tr.strftime("%Y%m%d-%H%M%S")
                emp_matched = emp_directory.get(clean_sicil) if clean_sicil else None
                final_ad = clean_ad_girilen or (emp_matched["ad_soyad"] if emp_matched else "Anonim")
                g_tarih_formatted = gorusme_tarihi_dt.strftime("%d.%m.%Y")
                kayit_zamani_formatted = (
                    str(edit_record_row.get("kayit_zamani", "") or "").strip()
                    if edit_kayit_id and edit_record_row
                    else now_tr.strftime("%d.%m.%Y %H:%M")
                )
                ise_giris_formatted = _format_tr_date_input(ise_giris_str)
                final_aksiyon_ihtiyaci = "Evet" if (aksiyon_ihtiyaci == "Evet" or has_note) else "Hayır"

                existing_done = _is_completed_flag(prefill.get("aksiyon_tamamlandi", False)) if edit_kayit_id else False
                existing_done_note = str(prefill.get("aksiyon_tamamlanma_notu", "") or "").strip() if edit_kayit_id else ""
                existing_done_date = str(prefill.get("aksiyon_tamamlanma_tarihi", "") or "").strip() if edit_kayit_id else ""
                keep_done = (final_aksiyon_ihtiyaci == "Evet") and existing_done and _has_completion_note(existing_done_note)

                new_record = {
                    "kayit_id": final_kayit_id,
                    "sicil_no": clean_sicil or final_kayit_id,
                    "ad_soyad": final_ad,
                    "gorusme_tarihi": g_tarih_formatted,
                    "kayit_zamani": kayit_zamani_formatted,
                    "donem": clean_donem,
                    "hrbp_kullanici": user["username"],
                    "ik_is_ortagi": sec_ik_ortagi,
                    "gmy": sec_gmy,
                    "departman": sec_departman,
                    "birim": sec_birim,
                    "bolum": sec_bolum,
                    "yonetici_bilgisi": sec_yonetici,
                    "lokasyon": sec_lokasyon,
                    "ise_giris_tarihi": ise_giris_formatted,
                    "katilimci_turu": "Katılımcı",
                    "calisan_rumuzu": f"{final_ad} ({sec_birim.split('-', 1)[-1]})",
                    "calisma_modeli": calisma_modeli,
                    "duygu_durumu": duygu_durumu,
                    "ana_tema": "Güven ve Kültür",
                    "aksiyon_ihtiyaci": final_aksiyon_ihtiyaci,
                    "aksiyon_seviyesi": "Birim / YM" if final_aksiyon_ihtiyaci == "Evet" else "Yok",
                    "aksiyon_notu": clean_note if (final_aksiyon_ihtiyaci == "Evet" and has_note) else "",
                    "aksiyon_tamamlandi": keep_done,
                    "aksiyon_tamamlanma_tarihi": existing_done_date if keep_done else "",
                    "aksiyon_tamamlanma_notu": existing_done_note if keep_done else "",
                    "q_kapanis": clean_kapanis,
                    **{k: v.strip() for k, v in q_answers.items()},
                    **{k: v.strip() for k, v in ik_answers.items()},
                    **{k: (None if v is None else float(v)) for k, v in p_scores.items()},
                }
                ok, msg = save_or_update_record(new_record, edit_kayit_id=edit_kayit_id)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

elif sayfa == "İzlenim Paylaşım Görüşmeleri":
    render_izlenim_paylasim_module(df_all, assigned_gmy_map, gmy_options, mood_slot)

else:
    st.subheader("Kayıt Havuzu")
    if not df_all.empty:
        arama_metni = st.text_input(
            "Havuzda Ara (Sicil, Ad Soyad, Birim, Departman veya Yorum)",
            placeholder="örnek: 1001, uzaktan, onay, terfi, Mobil...",
        )
        if arama_metni and arama_metni.strip():
            q_low = tr_lower(arama_metni.strip())
            mask = df_all.apply(lambda s: s.astype(str).apply(tr_lower).str.contains(q_low, na=False)).any(axis=1)
            df_all = df_all[mask].copy()
            st.caption(f"Filtrelenen kayıt sayısı: {len(df_all)}")
            mood_slot.markdown(render_corner_mood_widget(df_all, highlight_positive=True), unsafe_allow_html=True)

    if df_all.empty:
        st.info("Henüz kayıt bulunmuyor.")
    else:
        df_pool_view = _sync_action_plan_to_df(df_all)
        if "aksiyon_tamamlandi" in df_pool_view.columns:
            df_pool_view["aksiyon_tamamlandi"] = df_pool_view["aksiyon_tamamlandi"].apply(
                lambda x: "Evet" if _is_completed_flag(x) else "Hayır"
            )
        df_display = df_pool_view.rename(columns=COLUMN_LABELS)
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        excel_data = build_excel_export(df_all)
        st.download_button(
            label="📥 Havuzu Excel Olarak İndir (.xlsx)",
            data=excel_data,
            file_name=f"Izlenim_Paylasim_Havuz_{now_local().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
