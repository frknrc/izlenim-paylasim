from collections import Counter
from datetime import date, datetime
import io
import json
import os
import re
import zlib
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

st.set_page_config(
    page_title="İzlenim Paylaşım Görüşmeleri | Yönetici Sunumu",
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
            height: 128px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .panel-brand {{
            font-size: 1.35rem;
            font-weight: 800;
            letter-spacing: -0.3px;
            color: {TEXT_LIGHT};
        }}
        .mood-corner-card {{
            background: linear-gradient(135deg, #0D2A29 0%, #131E33 55%, #1A1738 100%);
            border: 1.5px solid rgba(16, 185, 129, 0.55);
            box-shadow: 0 8px 24px rgba(16, 185, 129, 0.14);
            border-radius: 12px;
            padding: 0.85rem 1.1rem;
            height: 128px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }}
        .mood-top-row {{
            display: flex;
            align-items: center;
            gap: 0.7rem;
        }}
        .mood-svg-wrap {{
            width: 44px;
            height: 44px;
            flex-shrink: 0;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: rgba(16, 185, 129, 0.14);
            border: 1px solid rgba(16, 185, 129, 0.35);
        }}
        .mood-title {{
            font-size: 1.02rem;
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
            margin-top: 0.5rem;
            font-size: 0.74rem;
            font-weight: 700;
        }}
        .mood-bar-track {{
            width: 100%;
            height: 7px;
            border-radius: 999px;
            overflow: hidden;
            display: flex;
            background: #0F172A;
            margin-top: 0.3rem;
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
        .action-section-banner {{
            background: linear-gradient(90deg, rgba(139, 92, 246, 0.18) 0%, rgba(34, 211, 238, 0.12) 100%);
            border: 1px solid rgba(167, 139, 250, 0.45);
            border-left: 5px solid {PURPLE_LIGHT};
            border-radius: 10px;
            padding: 0.75rem 1rem;
            margin-bottom: 0.75rem;
        }}
        .action-note-box {{
            background: linear-gradient(120deg, #171D33 0%, #1D1838 100%);
            border: 1px solid rgba(167, 139, 250, 0.45);
            border-left: 5px solid {TEAL};
            box-shadow: 0 4px 16px rgba(139, 92, 246, 0.12);
            border-radius: 10px;
            padding: 0.9rem 1.15rem;
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


def render_corner_mood_widget(df_source: pd.DataFrame) -> str:
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
        <div class="mood-corner-card">
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
        durum_renk = "#34D399"
        vurgu_yuzde = pos_pct
        svg_face = """
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14.5" fill="rgba(52, 211, 153, 0.22)" stroke="#34D399" stroke-width="2.4"/>
            <circle cx="13" cy="14.5" r="2.1" fill="#34D399"/>
            <circle cx="23" cy="14.5" r="2.1" fill="#34D399"/>
            <path d="M11.5 20.2C13.2 23.6 15.5 25.2 18 25.2C20.5 25.2 22.8 23.6 24.5 20.2" stroke="#34D399" stroke-width="2.5" stroke-linecap="round"/>
        </svg>
        """
    elif neg_cnt > pos_cnt and neg_cnt >= neu_cnt:
        durum_baslik = "Genel Olumsuz"
        durum_renk = NEG_COLOR
        vurgu_yuzde = neg_pct
        svg_face = f"""
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14.5" fill="rgba(248, 113, 113, 0.2)" stroke="{NEG_COLOR}" stroke-width="2.4"/>
            <circle cx="13" cy="15" r="2.0" fill="{NEG_COLOR}"/>
            <circle cx="23" cy="15" r="2.0" fill="{NEG_COLOR}"/>
            <path d="M12 24.2C13.5 21.4 15.6 20 18 20C20.4 20 22.5 21.4 24 24.2" stroke="{NEG_COLOR}" stroke-width="2.4" stroke-linecap="round"/>
        </svg>
        """
    else:
        durum_baslik = "Genel Nötr"
        durum_renk = TEAL
        vurgu_yuzde = neu_pct
        svg_face = f"""
        <svg width="32" height="32" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg">
            <circle cx="18" cy="18" r="14.5" fill="rgba(34, 211, 238, 0.18)" stroke="{TEAL}" stroke-width="2.4"/>
            <circle cx="13" cy="15" r="2.0" fill="{TEAL}"/>
            <circle cx="23" cy="15" r="2.0" fill="{TEAL}"/>
            <path d="M12.2 22.2H23.8" stroke="{TEAL}" stroke-width="2.4" stroke-linecap="round"/>
        </svg>
        """

    return f"""
    <div class="mood-corner-card">
        <div class="mood-top-row">
            <div class="mood-svg-wrap">{svg_face}</div>
            <div>
                <div class="mood-title" style="color:{durum_renk};">{durum_baslik} <span style="color:#FFFFFF; font-size:0.96rem;">(%{vurgu_yuzde:.1f})</span></div>
                <div class="mood-sub">{total} görüşme • Öne çıkan genel izlenim</div>
            </div>
        </div>
        <div class="mood-pct-row">
            <span style="color:#34D399;">Olumlu %{pos_pct:.1f}</span>
            <span style="color:{NEU_COLOR};">Nötr %{neu_pct:.1f}</span>
            <span style="color:{NEG_COLOR};">Olumsuz %{neg_pct:.1f}</span>
        </div>
        <div class="mood-bar-track">
            <div style="width:{pos_pct}%; background:#34D399;"></div>
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

LOKASYON_OPTIONS = [
    "İstanbul - Ataşehir",
    "İstanbul - Teknopark (Pendik)",
    "İstanbul - Levent",
    "Ankara - ODTÜ Teknokent",
    "İzmir - Ofis",
    "Uzaktan / Hibrit",
]

IK_IS_ORTAGI_OPTIONS = [
    "Şeyma Oruç",
    "Örnek İK İş Ortağı",
    "İK İş Ortağı - Teknoloji",
    "İK İş Ortağı - Bankacılık",
]

DEFAULT_USERS = {
    "seyma": {
        "ad_soyad": "Şeyma (İK İş Ortağı)",
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

HALF_STEP_SCORE_OPTIONS = ["Boş", 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0]


def _parse_date_safe(val: object, fallback: date | None = None) -> date:
    default_d = fallback if fallback is not None else date.today()
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
    if isinstance(raw_val, str) and raw_val.strip().lower() in ("boş", "bos", "none", "nan"):
        return None
    try:
        if isinstance(raw_val, float) and np.isnan(raw_val):
            return None
        f_val = float(str(raw_val).replace(",", "."))
        if np.isnan(f_val):
            return None
        clamped = max(1.0, min(5.0, f_val))
        snapped = round(clamped * 2.0) / 2.0
        return float(max(1.0, min(5.0, snapped)))
    except Exception:
        return None


def _has_meaningful_action_note(note_text: object) -> bool:
    if note_text is None:
        return False
    s = str(note_text).strip()
    if not s:
        return False
    return s.lower() not in {"yok", "-", "hayır", "hayir", "none", "null", "boş", "bos"}


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
    "aksiyon_ihtiyaci": "Aksiyon İhtiyacı",
    "aksiyon_seviyesi": "Aksiyon Seviyesi",
    "aksiyon_notu": "Aksiyon Planı",
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
    "destekliyor", "katkı sağlıyor", "düzenli", "memnunum",
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


def get_employee_directory(df_all: pd.DataFrame) -> dict[str, dict]:
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
                "lokasyon": str(row.get("lokasyon", "") or LOKASYON_OPTIONS[idx % len(LOKASYON_OPTIONS)]),
                "ise_giris_tarihi": _format_tr_date_input(row.get("ise_giris_tarihi", "")) or f"15.0{(idx % 8) + 1}.202{(idx % 5) + 1}",
                "ik_is_ortagi": str(row.get("ik_is_ortagi", "") or IK_IS_ORTAGI_OPTIONS[idx % len(IK_IS_ORTAGI_OPTIONS)]),
            }

    sample_extra = {
        "1001": {
            "sicil_no": "1001",
            "ad_soyad": "Ahmet Yılmaz",
            "gmy": "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri",
            "departman": "Yazılım Geliştirme Departmanı",
            "birim": "10000474-Kurumsal Uygulamalar",
            "bolum": "Ölçme ve Raporlama",
            "yonetici_bilgisi": "Birim Yöneticisi - 1",
            "lokasyon": "İstanbul - Ataşehir",
            "ise_giris_tarihi": "14.03.2022",
            "ik_is_ortagi": "Şeyma Oruç",
        },
        "1002": {
            "sicil_no": "1002",
            "ad_soyad": "Zeynep Kaya",
            "gmy": "10000028-Teknoloji, Yapay Zeka ve Sürekli İyileştirme",
            "departman": "Veri ve Yapay Zeka Departmanı",
            "birim": "10000182-Veri ve Yapay Zeka Uygulamaları",
            "bolum": "Dönüşüm Atölyesi",
            "yonetici_bilgisi": "Yetkinlik Merkezi Lideri",
            "lokasyon": "İstanbul - Teknopark (Pendik)",
            "ise_giris_tarihi": "01.09.2023",
            "ik_is_ortagi": "Şeyma Oruç",
        },
    }
    for k, v in sample_extra.items():
        if k not in directory:
            directory[k] = v

    return directory


def normalize_turkish_text(text: str) -> str:
    low = text.lower()
    low = (
        low.replace("İ", "i")
        .replace("I", "ı")
        .replace("’", "'")
        .replace("“", '"')
        .replace("”", '"')
    )
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
    low = text.lower().replace("İ", "i").replace("I", "ı")
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


def classify_clause(clause: str, row_duygu: str = "") -> str:
    low = clause.lower()
    pos_hits = sum(1 for w in POS_SIGNALS if w in low)
    neg_hits = sum(1 for w in NEG_SIGNALS if w in low)
    if neg_hits > pos_hits:
        return "olumsuz"
    if pos_hits > neg_hits:
        return "olumlu"
    if row_duygu == "Olumsuz":
        return "olumsuz"
    if row_duygu == "Olumlu":
        return "olumlu"
    return "olumlu"


def cluster_similar_statements(
    records: list[tuple[int, str]],
    top_n: int = 3,
    sim_threshold: float = 0.34,
    max_items: int = 400,
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
        norm_docs.append(norm if norm.strip() else clean_txt.lower())

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
        return "empty"
    try:
        payload = pd.util.hash_pandas_object(df, index=True).values.tobytes()
        return str(zlib.crc32(payload))
    except Exception:
        return f"{len(df)}_{list(df.columns)}"


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
    if secili_gmy != "Tümü" and "gmy" in df_scoped.columns:
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
                for c_item in split_clauses(str(val)):
                    if classify_clause(c_item, row_duygu) == "olumsuz":
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
                        for c_item in split_clauses(str(val)):
                            if classify_clause(c_item, row_duygu) == "olumsuz":
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
                if classify_clause(txt) == "olumlu":
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

    def _score(v):
        if v is None or v == "" or (isinstance(v, float) and np.isnan(v)):
            return None
        try:
            val = float(str(v).replace(",", "."))
            return round(val, 2) if 1.0 <= val <= 5.0 else None
        except Exception:
            return None

    rows_out = []
    for idx, r in enumerate(records):
        g_mapped, b_mapped = all_pairs[idx % len(all_pairs)]
        g_val = _clean(r.get("gmy") or r.get("GMY") or r.get("Genel Müdür Yardımcılığı") or r.get("GMY / Domain")) or g_mapped
        b_val = _clean(r.get("birim") or r.get("Birim") or r.get("Birim / Ekip")) or b_mapped
        dept_val = _clean(r.get("departman") or r.get("Departman")) or DEPARTMAN_OPTIONS[idx % len(DEPARTMAN_OPTIONS)]
        bolum_val = _clean(r.get("bolum") or r.get("Bölüm")) or BOLUM_OPTIONS[idx % len(BOLUM_OPTIONS)]
        yon_val = _clean(r.get("yonetici_bilgisi") or r.get("Yönetici Bilgisi")) or YONETICI_OPTIONS[idx % len(YONETICI_OPTIONS)]
        lok_val = _clean(r.get("lokasyon") or r.get("Lokasyon")) or LOKASYON_OPTIONS[idx % len(LOKASYON_OPTIONS)]
        ise_giris_val = _format_tr_date_input(r.get("ise_giris_tarihi") or r.get("İşe Giriş Tarihi")) or f"15.0{(idx % 8) + 1}.202{(idx % 5) + 1}"
        ik_ortak_val = _clean(r.get("ik_is_ortagi") or r.get("İK İş Ortağı")) or IK_IS_ORTAGI_OPTIONS[idx % len(IK_IS_ORTAGI_OPTIONS)]
        sicil_val = _clean(r.get("sicil_no") or r.get("Sicil") or r.get("Sicil No")) or str(1001 + idx)
        ad_soyad_val = _clean(r.get("ad_soyad") or r.get("Ad Soyad")) or f"Kayıt {idx + 1}"
        g_tarih = _clean(r.get("gorusme_tarihi") or r.get("Görüşme Tarihi") or r.get("Görüşme Tarihi (Ay/Gün/Yıl)")) or "2026-06-04"

        mem = _score(r.get("p_genel_memnuniyet") if "p_genel_memnuniyet" in r else r.get("st Memnuniyeti"))
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
        aksiyon_not_raw = _clean(r.get("aksiyon_notu") or r.get("Aksiyon Planı") or r.get("Aksiyon Notu"))
        aksiyon_var_raw = _clean(r.get("aksiyon_ihtiyaci") or r.get("Aksiyon İhtiyacı"))
        if aksiyon_var_raw in ["Evet", "Hayır"]:
            aksiyon_var = aksiyon_var_raw
        else:
            aksiyon_var = (
                "Evet"
                if (
                    _has_meaningful_action_note(aksiyon_not_raw)
                    or (
                        q10_txt
                        and len(q10_txt) > 12
                        and "aklıma" not in q10_txt.lower()
                        and "yorum yapamıcam" not in q10_txt.lower()
                    )
                )
                else "Hayır"
            )

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
                "kayit_zamani": f"{g_tarih} 10:00",
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
                    if any(w in q10_txt.lower() for w in ["uzaktan", "remote", "ofis", "çalışma"])
                    else "Güven ve Kültür"
                ),
                "aksiyon_ihtiyaci": aksiyon_var,
                "aksiyon_seviyesi": _clean(r.get("aksiyon_seviyesi")) or ("Birim / YM" if aksiyon_var == "Evet" else "Yok"),
                "aksiyon_notu": aksiyon_not_raw if _has_meaningful_action_note(aksiyon_not_raw) else (q10_txt if aksiyon_var == "Evet" else ""),
                "q01_calisma_hayati": _clean(r.get("q01_calisma_hayati") or r.get("st'te çalışma hayatın nasıl gidiyor?")),
                "q02_guclu_yonler": _clean(r.get("q02_guclu_yonler") or r.get("Sana göre st’teki olumlu/güçlü yönler neler?")),
                "q03_ekip_rol_tatmin": _clean(r.get("q03_ekip_rol_tatmin") or r.get("Görev aldığın ekip, rolün ve projelerin seni tatmin ediyor, sana katkı sağlıyor mu?")),
                "q04_ekip_iyilestirme": _clean(r.get("q04_ekip_iyilestirme") or r.get("Ekibinde sana göre daha iyi olması gereken şeyler neler?")),
                "q05_motivasyon_arttiran": _clean(r.get("q05_motivasyon_arttiran") or r.get("Kendi motivasyonunu daha çok arttıracak şeyler neler olabilir? (İş, proje, kariyer vb.)")),
                "q06_yonetici_geri_bildirim": _clean(r.get("q06_yonetici_geri_bildirim") or r.get("Yöneticin gelişimin için zaman ayırıp düzenli geri bildirim verebiliyor mu? Görüşmenin içeriği beklentini karşılıyor mu? (1.Yönetici, varsa YHL, Direktör)")),
                "q07_ym_verimlilik": _clean(r.get("q07_ym_verimlilik") or r.get("Yetkinlik Merkezi yapısının verimliliğini nasıl değerlendirirsin? (Eğitim, iş süreçlerine katkı, matris yapı vb.)")),
                "q08_gmy_iletisimi": _clean(r.get("q08_gmy_iletisimi") or r.get("GMY iletişimi ile ilgili değerlendirmelerin nasıl olur?")),
                "q09_vizyon_strateji": _clean(r.get("q09_vizyon_strateji") or r.get("Şirketin sunduğu vizyon ve stratejiyi büyük resimde nasıl değerlendiriyorsun?")),
                "q10_degistirmek_istenen": q10_txt,
                "q11_mevcut_calisma_modeli": _clean(r.get("q11_mevcut_calisma_modeli") or r.get("Mevcut çalışma modelinle ilgili ne düşünüyorsun?")),
                "q12_ilerisi_calisma_modeli": _clean(r.get("q12_ilerisi_calisma_modeli") or r.get("İlerisi için nasıl bir çalışma modeli olursa daha verimli çalışacağına inanıyorsun?")),
                "q13_harmony": _clean(r.get("q13_harmony") or r.get("Harmony ile ilgili görüşlerin nasıl, çalışma sürecine olan etkisini nasıl değerlendirirsin?")),
                "q14_toplandigimiza_degsin": _clean(r.get("q14_toplandigimiza_degsin") or r.get('"Toplandığımıza Değsin" temasının ne kadar işlendiğini düşünüyorsun?')),
                "q15_karar_alma_surecleri": _clean(r.get("q15_karar_alma_surecleri") or r.get("Karar alma süreçlerini hem ekip içi işleyişler hem de üst yönetim seviyesinde nasıl değerlendirirsin?")),
                "ik_kariyer": _clean(r.get("ik_kariyer") or r.get("İşe Alım ve Kariyer Yönetimi")),
                "ik_deneyim_marka": _clean(r.get("ik_deneyim_marka") or r.get("Çalışan Deneyimi ve İşveren Markası Tanıtımı")),
                "ik_ucret_merkezi": _clean(r.get("ik_ucret_merkezi") or r.get("Ücretlendirme ve Çalışan Merkezi")),
                "ik_akademi": _clean(r.get("ik_akademi") or r.get("st Akademi")),
                "q_kapanis": kapanis_txt,
                "p_genel_memnuniyet": mem,
                "p_ekip_motivasyonu": _score(r.get("p_ekip_motivasyonu") if "p_ekip_motivasyonu" in r else r.get("Ekip Motivasyonu")),
                "p_kendi_motivasyonu": _score(r.get("p_kendi_motivasyonu") if "p_kendi_motivasyonu" in r else r.get("Kendi Motivasyonu")),
                "p_is_yasam_dengesi": _score(r.get("p_is_yasam_dengesi") if "p_is_yasam_dengesi" in r else r.get("İş Yaşam Dengesi")),
                "p_1_yonetici_iletisimi": _score(r.get("p_1_yonetici_iletisimi") if "p_1_yonetici_iletisimi" in r else r.get("1.Yönetici İletişimi")),
                "p_ym_degerlendirme": _score(r.get("p_ym_degerlendirme") if "p_ym_degerlendirme" in r else r.get("Yetkinlik Merkezi Değerlendirmesi")),
                "p_direktor_iletisimi": _score(r.get("p_direktor_iletisimi") if "p_direktor_iletisimi" in r else r.get("Direktör İletişimi")),
                "p_gmy_iletisimi": _score(r.get("p_gmy_iletisimi") if "p_gmy_iletisimi" in r else r.get("GMY İletişimi")),
                "p_ucret_yan_hak": _score(r.get("p_ucret_yan_hak") if "p_ucret_yan_hak" in r else r.get("Ücret Yan Haklar Motivasyonu")),
                "p_ik_iletisimi": _score(r.get("p_ik_iletisimi") if "p_ik_iletisimi" in r else r.get("İK İletişimi/Çalışmaları")),
                "p_yz_kullanim": _score(r.get("p_yz_kullanim") if "p_yz_kullanim" in r else r.get("Yapay Zeka Kullanım Sıklığı(CodeMantis, Amigo, Asistan vb.)")),
            }
        )
    return pd.DataFrame(rows_out)


def get_sample_data() -> pd.DataFrame:
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
                return _map_records_to_df(loaded)

    return pd.DataFrame([])


def load_records() -> tuple[pd.DataFrame, str]:
    if "local_df" not in st.session_state:
        st.session_state.local_df = get_sample_data()

    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            ws = spreadsheet.worksheet("gorusmeler")
            data = ws.get_all_records()
            if data:
                df = pd.DataFrame(data)
                for col, _ in PUAN_BASLIKLARI:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce")
                return df, f"Google Sheets ({len(df)} Kayıt)"
            return pd.DataFrame(columns=st.session_state.local_df.columns), "Google Sheets (0 Kayıt)"
        except Exception as e:
            return st.session_state.local_df.copy(), f"Yerel Havuz ({len(st.session_state.local_df)} Kayıt) • {e}"

    return st.session_state.local_df.copy(), f"Yerel Havuz ({len(st.session_state.local_df)} Kayıt)"


def save_or_update_record(record: dict, edit_kayit_id: str | None = None) -> tuple[bool, str]:
    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            try:
                ws = spreadsheet.worksheet("gorusmeler")
            except Exception:
                ws = spreadsheet.add_worksheet(title="gorusmeler", rows=1000, cols=70)
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
        st.session_state.local_df = df_curr
        return True, f"Kayıt ({edit_kayit_id}) başarıyla güncellendi."

    new_row = pd.DataFrame([record])
    st.session_state.local_df = pd.concat([df_curr, new_row], ignore_index=True)
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
        "Aksiyon İhtiyacı",
        "Aksiyon Planı",
    ] + [q_lbl for _, q_lbl in TUM_METIN_SORULARI] + [p_lbl for _, p_lbl in PUAN_BASLIKLARI]

    sample_row = {
        "Sicil": "1001",
        "Ad Soyad": "Ahmet Yılmaz",
        "Görüşme Tarihi": date.today().strftime("%d.%m.%Y"),
        "Dönem": "2026 - 2. Dönem",
        "Genel Müdür Yardımcılığı": "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri",
        "Departman": "Yazılım Geliştirme Departmanı",
        "Birim": "10000474-Kurumsal Uygulamalar",
        "Bölüm": "Ölçme ve Raporlama",
        "Yönetici Bilgisi": "Birim Yöneticisi - 1",
        "Lokasyon": "İstanbul - Ataşehir",
        "İşe Giriş Tarihi": "14.03.2022",
        "İK İş Ortağı": "Şeyma Oruç",
        "Çalışma Modeli": "Hibrit",
        "Duygu Durumu": "Olumlu",
        "Aksiyon İhtiyacı": "Evet",
        "Aksiyon Planı": "Ekip değişikliği talebi bulunmakta, yönetici ile görüşme organize edilecek.",
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

        df_mapped = _map_records_to_df(records_raw)
        if df_mapped.empty:
            return 0, "Geçerli satır dönüştürülemedi."

        st.session_state.local_df = pd.concat([st.session_state.local_df, df_mapped], ignore_index=True)
        return len(df_mapped), f"{len(df_mapped)} görüşme kaydı Excel'den havuza başarıyla aktarıldı."
    except Exception as e:
        return 0, f"Excel okuma hatası: {e}"


def build_excel_export(df_raw: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    df_clean = df_raw.copy()

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
                }
            )
    df_birim_ozet = pd.DataFrame(ozet_rows) if ozet_rows else pd.DataFrame(
        columns=["Genel Müdür Yardımcılığı", "Birim", "Görüşme Sayısı", "Genel Puan Ortalaması (1-5)", "Olumlu Görüşme", "Olumsuz Görüşme", "Aksiyon İhtiyacı Olan"]
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
        fp.text = f"İzlenim Paylaşım Görüşmeleri | Yönetici Sunumu  |  Slayt {slide_num}/{total_slides}"
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
        not_metni = str(r.get("aksiyon_notu", "") or r.get("q10_degistirmek_istenen", "") or "").strip()
        bir = str(r.get("birim", ""))
        kisi = str(r.get("ad_soyad", "") or r.get("calisan_rumuzu", ""))
        aksiyon_satirlari.append(f"[{bir} | {kisi}] {not_metni}")

    add_card(s_last, 0.6, 1.35, 12.1, 5.2, "Takip Edilecek Aksiyon Maddeleri", aksiyon_satirlari or ["Seçilen filtrede aksiyon planı bulunmuyor."], c_teal)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()


if "user" not in st.session_state:
    st.session_state.user = None

if st.session_state.user is None:
    st.markdown(
        """
        <div class="panel-header">
            <div class="panel-brand">İzlenim Paylaşım Görüşmeleri | Yönetici Sunumu</div>
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
df_all, data_source_label = load_records()
emp_directory = get_employee_directory(df_all)

with st.sidebar:
    st.markdown(f"### 👤 {user['ad_soyad']}")
    st.caption(f"Bağlantı: {data_source_label}")
    if st.button("Çıkış Yap", use_container_width=True):
        st.session_state.user = None
        st.rerun()

    st.divider()
    sayfa = st.radio(
        "Modüller",
        ["1-1 Görüşme Kayıt", "Standart Dashboard & Sunum", "Ham Veri Havuzu"],
    )

assigned_gmy_map = user["gmy_birimler"]
gmy_options = list(assigned_gmy_map.keys())

hdr_left, hdr_right = st.columns([2.1, 1.9], vertical_alignment="top")
with hdr_left:
    st.markdown(
        f"""
        <div class="panel-header">
            <div class="panel-brand">İzlenim Paylaşım Görüşmeleri | Yönetici Sunumu</div>
            <div style="color:#94A3B8; margin-top:0.25rem; font-size:0.95rem;">{data_source_label}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with hdr_right:
    mood_slot = st.empty()
    mood_slot.markdown(render_corner_mood_widget(df_all), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if sayfa == "1-1 Görüşme Kayıt":
    st.markdown("#### 🔍 Kayıt Seçimi veya Yeni Görüşme Başlatma")
    mode_col1, mode_col2 = st.columns([1.2, 2])
    with mode_col1:
        kayit_modu = st.radio("İşlem Türü", ["Yeni Görüşme Ekle", "Mevcut Kaydı Düzenle / Güncelle"], horizontal=True)

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
                gt = _format_tr_date_input(r.get("gorusme_tarihi", ""))
                lbl = f"{sic} - {ad} ({gt}) [ID: {kid}]"
                secenekler.append(lbl)
                id_map[lbl] = kid
            secilen_lbl = st.selectbox("Güncellenecek Görüşme Kaydını Seçin", secenekler)
            edit_kayit_id = id_map.get(secilen_lbl)
            matched_df = df_all[df_all["kayit_id"] == edit_kayit_id]
            if not matched_df.empty:
                edit_record_row = matched_df.iloc[0].to_dict()

    s_col1, s_col2, s_col3, s_col4 = st.columns(4)
    default_sicil = str(edit_record_row.get("sicil_no", "")) if edit_record_row else ""
    with s_col1:
        sicil_input = st.text_input(
            "Sicil (Örn: 1001, 1002)",
            value=default_sicil,
            placeholder="Sicil girin...",
        ).strip()

    emp_info = emp_directory.get(sicil_input) if sicil_input else None
    if sicil_input and emp_info:
        st.success(f"Sicil bulundu: **{emp_info['ad_soyad']}** — Bilgiler otomatik getirildi; dilediğiniz alanı manuel düzenleyebilirsiniz.")
    elif sicil_input and not emp_info:
        st.info("Bu sicil kayıt/katalogda bulunamadı; bilgileri aşağıdan manuel girebilir veya seçebilirsiniz.")

    def _pick_idx(opts: list[str], val: str, fallback: int = 0) -> int:
        if val and val in opts:
            return opts.index(val)
        return fallback

    prefill = edit_record_row if edit_record_row else (emp_info if emp_info else {})
    form_scope = f"{edit_kayit_id or 'new'}_{sicil_input or 'empty'}"

    with s_col2:
        ad_soyad_input = st.text_input(
            "Ad Soyad",
            value=str(prefill.get("ad_soyad", "")),
            placeholder="Ad Soyad",
            key=f"inp_ad_soyad_{form_scope}",
        )
    with s_col3:
        default_gt_date = _parse_date_safe(prefill.get("gorusme_tarihi"), fallback=date.today())
        gorusme_tarihi_dt = st.date_input(
            "Görüşme Tarihi",
            value=default_gt_date,
            format="DD.MM.YYYY",
            key=f"inp_gorusme_tarihi_{form_scope}",
        )
    with s_col4:
        donem_opts = ["2026 - 1. Dönem", "2026 - 2. Dönem", "2026 - 1. Çeyrek", "2026 - 2. Çeyrek", "2026 - 3. Çeyrek", "2026 - 4. Çeyrek"]
        sec_donem = st.selectbox(
            "Görüşme Dönemi",
            donem_opts,
            index=_pick_idx(donem_opts, str(prefill.get("donem", "2026 - 2. Dönem")), 1),
            key=f"inp_donem_{form_scope}",
        )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sec_gmy = st.selectbox(
            "Genel Müdür Yardımcılığı",
            gmy_options,
            index=_pick_idx(gmy_options, str(prefill.get("gmy", "")), 0),
            key=f"inp_gmy_{form_scope}",
        )
    with c2:
        sec_departman = st.selectbox(
            "Departman",
            DEPARTMAN_OPTIONS,
            index=_pick_idx(DEPARTMAN_OPTIONS, str(prefill.get("departman", "")), 0),
            key=f"inp_departman_{form_scope}",
        )
    with c3:
        birim_options = assigned_gmy_map.get(sec_gmy, GMY_BIRIM_MAP.get(sec_gmy, []))
        sec_birim = st.selectbox(
            "Birim",
            birim_options,
            index=_pick_idx(birim_options, str(prefill.get("birim", "")), 0),
            key=f"inp_birim_{form_scope}",
        )
    with c4:
        sec_bolum = st.selectbox(
            "Bölüm",
            BOLUM_OPTIONS,
            index=_pick_idx(BOLUM_OPTIONS, str(prefill.get("bolum", "")), 0),
            key=f"inp_bolum_{form_scope}",
        )

    c5, c6, c7, c8 = st.columns(4)
    with c5:
