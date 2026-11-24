import base64
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
import plotly.graph_objects as go
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import streamlit as st

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
        .stApp {{
            background-color: {BG_DARK};
            color: {TEXT_LIGHT};
        }}
        [data-testid="stSidebar"] {{
            background-color: {CARD_DARK};
            border-right: 1px solid {CARD_BORDER};
        }}
        .stTextInput > div > div > input,
        .stTextArea > div > div > textarea,
        .stSelectbox > div > div > div,
        .stMultiSelect > div > div > div {{
            background-color: #0F172A !important;
            color: {TEXT_LIGHT} !important;
            border-color: {CARD_BORDER} !important;
        }}
        .stButton > button, .stDownloadButton > button {{
            background: linear-gradient(90deg, {TEAL_SOFT} 0%, {PURPLE} 100%);
            color: #FFFFFF;
            font-weight: 600;
            border: none;
            border-radius: 8px;
            padding: 0.5rem 1rem;
        }}
        .stButton > button:hover, .stDownloadButton > button:hover {{
            opacity: 0.92;
            color: #FFFFFF;
        }}
        .panel-header {{
            background: linear-gradient(120deg, {CARD_DARK} 0%, #1B1638 60%, #0E293B 100%);
            border: 1px solid {CARD_BORDER};
            border-left: 5px solid {TEAL};
            border-radius: 12px;
            padding: 1.2rem 1.5rem;
            margin-bottom: 1rem;
        }}
        .panel-brand {{
            font-size: 1.5rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: {TEXT_LIGHT};
        }}
        .panel-accent {{
            color: {TEAL};
        }}
        .panel-purple {{
            color: {PURPLE_LIGHT};
        }}
        .metric-card {{
            background-color: {CARD_DARK};
            border: 1px solid {CARD_BORDER};
            border-radius: 10px;
            padding: 1rem;
            text-align: left;
        }}
        .metric-label {{
            color: {TEXT_MUTED};
            font-size: 0.85rem;
            margin-bottom: 0.25rem;
        }}
        .metric-value {{
            color: {TEXT_LIGHT};
            font-size: 1.6rem;
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
    </style>
    """,
    unsafe_allow_html=True,
)

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

DEFAULT_USERS = {
    "seyma": {
        "ad_soyad": "Şeyma (İK İş Ortağı)",
        "sifre": "panel2026",
        "rol": "hrbp",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
    "hrbp_ornek": {
        "ad_soyad": "Örnek HRBP",
        "sifre": "panel2026",
        "rol": "hrbp",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
    "admin": {
        "ad_soyad": "İK Raporlama Admin",
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

ANA_TEMALAR = [
    "Güven ve Kültür",
    "Ulaşılabilir İK",
    "İş Tatmini",
    "Yetkinlik Merkezi Yapısı",
    "Harmony ve Karne Etkisi",
    "Karar Süreçleri ve Bilgi Akışı",
    "Prim, Ücret Artışları ve Terfi",
    "Yaşanan Yetki Sorunları",
    "YZ ve BT Kaynaklı Kısıtlar",
    "Üst Yönetim / Direktör İletişimi",
    "Çalışma Modeli / Hibrit / İş-Yaşam",
    "Diğer",
]

COLUMN_LABELS = {
    "kayit_id": "Kayıt ID",
    "kayit_zamani": "Kayıt Zamanı",
    "donem": "Dönem",
    "hrbp_kullanici": "İK Sorumlusu",
    "gmy": "GMY / Domain",
    "birim": "Birim / Ekip",
    "katilimci_turu": "Katılımcı Tipi",
    "calisan_rumuzu": "Katılımcı Kodu / Adı",
    "calisma_modeli": "Çalışma Modeli",
    "duygu_durumu": "Duygu Durumu",
    "ana_tema": "Öne Çıkan Ana Tema",
    "aksiyon_ihtiyaci": "Aksiyon İhtiyacı",
    "aksiyon_seviyesi": "Aksiyon Seviyesi",
    "aksiyon_notu": "Aksiyon Notu",
    "q_kapanis": "Kapanış / Ek Not",
    **dict(GORUSME_SORULARI),
    **dict(IK_BASLIKLARI),
    **dict(PUAN_BASLIKLARI),
}

NEG_SIGNALS = [
    "yavaş",
    "geç",
    "sorun",
    "sıkıntı",
    "eksik",
    "yük",
    "yoğun",
    "zor",
    "karmaşık",
    "belirsiz",
    "yetersiz",
    "azalt",
    "kısıt",
    "olumsuz",
    "düşük",
    "problem",
    "zaman alıyor",
    "bekleniyor",
    "iyileşmeye açık",
    "sadeleştir",
    "şeffaf olabilir",
    "karne yükü",
    "yetki",
    "onay",
    "uzun sürüyor",
    "kopuk",
    "yetersiz kalıyor",
    "artırılmalı",
    "iyileştirilmeli",
    "net değil",
]

POS_SIGNALS = [
    "güçlü",
    "iyi",
    "olumlu",
    "memnun",
    "destekleyici",
    "başarılı",
    "verimli",
    "net",
    "açık",
    "kolay",
    "hızlı",
    "değerli",
    "faydalı",
    "yüksek",
    "yapıcı",
    "güven",
    "ulaşılabilir",
    "tatmin",
    "sorun yok",
    "akıcı",
    "uyumlu",
    "destekliyor",
    "katkı sağlıyor",
    "düzenli",
    "memnunum",
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
    """
    records: [(person_idx, raw_clause), ...]
    Farklı kelimelerle kurulsa bile aynı kavrama/anlama işaret eden ifadeleri
    kosinüs benzerliği + kavram eşleştirme ile kümeler ve kaç farklı kişide
    geçtiğini hesaplar.
    """
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


@st.cache_data(show_spinner=False)
def _cached_question_summaries(df_key: str, records_json: str) -> list[dict]:
    df_source = pd.read_json(io.StringIO(records_json), orient="records")
    return _analyze_questions_impl(df_source)


@st.cache_data(show_spinner=False)
def _cached_department_analysis(df_key: str, records_json: str, secili_gmy: str) -> dict:
    df_scoped = pd.read_json(io.StringIO(records_json), orient="records")
    return _analyze_department_impl(df_scoped, secili_gmy)


@st.cache_data(show_spinner=False)
def _cached_pptx_bytes(df_key: str, records_json: str, secili_gmy: str, secili_donem: str) -> bytes:
    df_filtered = pd.read_json(io.StringIO(records_json), orient="records")
    return _build_pptx_impl(df_filtered, secili_gmy, secili_donem)


def _df_to_json_safe(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return "[]"
    return df.to_json(orient="records", force_ascii=False)


def _analyze_questions_impl(df_source: pd.DataFrame) -> list[dict]:
    summaries = []
    for col_key, soru_metni in TUM_METIN_SORULARI:
        if col_key not in df_source.columns or df_source.empty:
            continue
        pos_records = []
        neg_records = []
        yanit_sayisi = 0

        for person_idx, (_, row) in enumerate(df_source.iterrows()):
            raw_val = row.get(col_key, "")
            if pd.isna(raw_val) or not str(raw_val).strip():
                continue
            yanit_sayisi += 1
            row_duygu = str(row.get("duygu_durumu", "") or "")
            for c_item in split_clauses(str(raw_val)):
                yon = classify_clause(c_item, row_duygu)
                if yon == "olumsuz":
                    neg_records.append((person_idx, c_item))
                else:
                    pos_records.append((person_idx, c_item))

        if yanit_sayisi > 0:
            summaries.append(
                {
                    "soru_kodu": col_key,
                    "soru": soru_metni,
                    "yanit_sayisi": yanit_sayisi,
                    "olumlu_noktalar": cluster_similar_statements(pos_records, top_n=3),
                    "olumsuz_noktalar": cluster_similar_statements(neg_records, top_n=3),
                }
            )
    return summaries


def analyze_questions_across_pool(df_source: pd.DataFrame) -> list[dict]:
    if df_source is None or df_source.empty:
        return []
    return _cached_question_summaries(_df_fingerprint(df_source), _df_to_json_safe(df_source))


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

            if df_b.empty:
                durum_etiketi = "Kayıt Yok"
            elif top_b_neg or olumsuz_kisi > 0 or aksiyon_kisi > 0:
                durum_etiketi = "Dikkat / İyileştirme Alanı Var"
            else:
                durum_etiketi = "Genel Olumlu / Sorun Görünmüyor"

            birim_listesi.append(
                {
                    "gmy": g_ad,
                    "birim": b_ad,
                    "gorusme_sayisi": len(df_b),
                    "durum": durum_etiketi,
                    "olumlu": top_b_pos or (["Kayıt bulunmuyor"] if df_b.empty else ["Belirgin olumlu ifade yok"]),
                    "olumsuz": top_b_neg or (["Kayıt bulunmuyor"] if df_b.empty else ["Öne çıkan olumsuz ifade yok"]),
                }
            )

    return {
        "gorusme_sayisi": len(df_dept),
        "olumlu": cluster_similar_statements(dept_pos_all, top_n=4) or ["Yeterli yorum bulunmuyor"],
        "olumsuz": cluster_similar_statements(dept_neg_all, top_n=4) or ["Öne çıkan olumsuz ifade bulunmuyor"],
        "birimler": birim_listesi,
    }


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
        if v is None:
            return ""
        return " ".join(str(v).split()).strip()

    def _score(v):
        if v is None or v == "":
            return None
        try:
            val = float(v)
            return round(val, 2) if 1.0 <= val <= 5.0 else None
        except Exception:
            return None

    rows_out = []
    for idx, r in enumerate(records):
        g_mapped, b_mapped = all_pairs[idx % len(all_pairs)]
        mem = _score(r.get("st Memnuniyeti"))
        if mem is None:
            duygu = "Nötr"
        elif mem >= 4.0:
            duygu = "Olumlu"
        elif mem < 3.0:
            duygu = "Olumsuz"
        else:
            duygu = "Nötr"

        q10_txt = _clean(r.get("st'te neyi değiştirmek isterdin?"))
        aksiyon_var = (
            "Evet"
            if (
                q10_txt
                and len(q10_txt) > 12
                and "aklıma" not in q10_txt.lower()
                and "yorum yapamıcam" not in q10_txt.lower()
            )
            else "Hayır"
        )
        yz_yorum = _clean(r.get("Yapay Zeka Kullanım Sıklığı Yorum(CodeMantis, Amigo, Asistan vb.)"))
        kapanis_txt = _clean(r.get("Paylaşmak istediğin veya eklemek istediğin farklı konu var mı?"))
        if yz_yorum:
            kapanis_txt = f"{kapanis_txt} | YZ Yorum: {yz_yorum}".strip(" |")

        rows_out.append(
            {
                "kayit_id": f"2026-{r.get('Sicil', 1001 + idx)}",
                "kayit_zamani": f"{_clean(r.get('Görüşme Tarihi (Ay/Gün/Yıl)')) or '2026-06-04'} 10:00",
                "donem": "2026 - 2. Dönem",
                "hrbp_kullanici": "seyma",
                "gmy": g_mapped,
                "birim": b_mapped,
                "katilimci_turu": "Çalışan",
                "calisan_rumuzu": f"{_clean(r.get('Ad Soyad')) or f'Çalışan {idx+1}'} ({_clean(r.get('Birim')) or b_mapped.split('-', 1)[-1]})",
                "calisma_modeli": "Hibrit",
                "duygu_durumu": duygu,
                "ana_tema": "Çalışma Modeli / Hibrit / İş-Yaşam"
                if any(w in q10_txt.lower() for w in ["uzaktan", "remote", "ofis", "çalışma"])
                else "Güven ve Kültür",
                "aksiyon_ihtiyaci": aksiyon_var,
                "aksiyon_seviyesi": "Birim / YM" if aksiyon_var == "Evet" else "Yok",
                "aksiyon_notu": q10_txt if aksiyon_var == "Evet" else "",
                "q01_calisma_hayati": _clean(r.get("st'te çalışma hayatın nasıl gidiyor?")),
                "q02_guclu_yonler": _clean(r.get("Sana göre st’teki olumlu/güçlü yönler neler?")),
                "q03_ekip_rol_tatmin": _clean(r.get("Görev aldığın ekip, rolün ve projelerin seni tatmin ediyor, sana katkı sağlıyor mu?")),
                "q04_ekip_iyilestirme": _clean(r.get("Ekibinde sana göre daha iyi olması gereken şeyler neler?")),
                "q05_motivasyon_arttiran": _clean(r.get("Kendi motivasyonunu daha çok arttıracak şeyler neler olabilir? (İş, proje, kariyer vb.)")),
                "q06_yonetici_geri_bildirim": _clean(r.get("Yöneticin gelişimin için zaman ayırıp düzenli geri bildirim verebiliyor mu? Görüşmenin içeriği beklentini karşılıyor mu? (1.Yönetici, varsa YHL, Direktör)")),
                "q07_ym_verimlilik": _clean(r.get("Yetkinlik Merkezi yapısının verimliliğini nasıl değerlendirirsin? (Eğitim, iş süreçlerine katkı, matris yapı vb.)")),
                "q08_gmy_iletisimi": _clean(r.get("GMY iletişimi ile ilgili değerlendirmelerin nasıl olur?")),
                "q09_vizyon_strateji": _clean(r.get("Şirketin sunduğu vizyon ve stratejiyi büyük resimde nasıl değerlendiriyorsun?")),
                "q10_degistirmek_istenen": q10_txt,
                "q11_mevcut_calisma_modeli": _clean(r.get("Mevcut çalışma modelinle ilgili ne düşünüyorsun?")),
                "q12_ilerisi_calisma_modeli": _clean(r.get("İlerisi için nasıl bir çalışma modeli olursa daha verimli çalışacağına inanıyorsun?")),
                "q13_harmony": _clean(r.get("Harmony ile ilgili görüşlerin nasıl, çalışma sürecine olan etkisini nasıl değerlendirirsin?")),
                "q14_toplandigimiza_degsin": _clean(r.get('"Toplandığımıza Değsin" temasının ne kadar işlendiğini düşünüyorsun?')),
                "q15_karar_alma_surecleri": _clean(r.get("Karar alma süreçlerini hem ekip içi işleyişler hem de üst yönetim seviyesinde nasıl değerlendirirsin?")),
                "ik_kariyer": _clean(r.get("İşe Alım ve Kariyer Yönetimi")),
                "ik_deneyim_marka": _clean(r.get("Çalışan Deneyimi ve İşveren Markası Tanıtımı")),
                "ik_ucret_merkezi": _clean(r.get("Ücretlendirme ve Çalışan Merkezi")),
                "ik_akademi": _clean(r.get("st Akademi")),
                "q_kapanis": kapanis_txt,
                "p_genel_memnuniyet": _score(r.get("st Memnuniyeti")),
                "p_ekip_motivasyonu": _score(r.get("Ekip Motivasyonu")),
                "p_kendi_motivasyonu": _score(r.get("Kendi Motivasyonu")),
                "p_is_yasam_dengesi": _score(r.get("İş Yaşam Dengesi")),
                "p_1_yonetici_iletisimi": _score(r.get("1.Yönetici İletişimi")),
                "p_ym_degerlendirme": _score(r.get("Yetkinlik Merkezi Değerlendirmesi")),
                "p_direktor_iletisimi": _score(r.get("Direktör İletişimi")),
                "p_gmy_iletisimi": _score(r.get("GMY İletişimi")),
                "p_ucret_yan_hak": _score(r.get("Ücret Yan Haklar Motivasyonu")),
                "p_ik_iletisimi": _score(r.get("İK İletişimi/Çalışmaları")),
                "p_yz_kullanim": _score(r.get("Yapay Zeka Kullanım Sıklığı(CodeMantis, Amigo, Asistan vb.)")),
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
                return df, "Google Sheets (Canlı)"
            return pd.DataFrame(columns=st.session_state.local_df.columns), "Google Sheets (Boş)"
        except Exception as e:
            return st.session_state.local_df.copy(), f"Yerel/Demo Mod (Sheets hatası: {e})"

    return st.session_state.local_df.copy(), f"Yerel Havuz ({len(st.session_state.local_df)} Kayıt)"


def save_record(record: dict) -> tuple[bool, str]:
    _, spreadsheet = get_gsheet_client()
    if spreadsheet is not None:
        try:
            try:
                ws = spreadsheet.worksheet("gorusmeler")
            except Exception:
                ws = spreadsheet.add_worksheet(title="gorusmeler", rows=1000, cols=60)
            headers = ws.row_values(1)
            if not headers:
                headers = list(record.keys())
                ws.append_row(headers)
            row_values = [record.get(h, "") for h in headers]
            ws.append_row(row_values)
            return True, "Kayıt Google Sheets'e eklendi."
        except Exception as e:
            return False, f"Google Sheets yazma hatası: {e}"

    new_row = pd.DataFrame([record])
    st.session_state.local_df = pd.concat([st.session_state.local_df, new_row], ignore_index=True)
    return True, "Kayıt oturum verisine eklendi."


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
                    "GMY / Domain": gmy_ad,
                    "Birim / Ekip": birim_ad,
                    "Görüşme Sayısı": len(grp),
                    "Genel Puan Ortalaması (1-5)": ort_skor,
                    "Olumlu Görüşme": olumlu_sayisi,
                    "Olumsuz Görüşme": olumsuz_sayisi,
                    "Aksiyon İhtiyacı Olan": aksiyon_sayisi,
                }
            )
    df_birim_ozet = (
        pd.DataFrame(ozet_rows)
        if ozet_rows
        else pd.DataFrame(
            columns=[
                "GMY / Domain",
                "Birim / Ekip",
                "Görüşme Sayısı",
                "Genel Puan Ortalaması (1-5)",
                "Olumlu Görüşme",
                "Olumsuz Görüşme",
                "Aksiyon İhtiyacı Olan",
            ]
        )
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
                if ws_name == "Görüşme Kayıtları":
                    if len(header_text) > 35:
                        ws.column_dimensions[col_letter].width = 42
                    elif "GMY" in header_text or "Birim" in header_text:
                        ws.column_dimensions[col_letter].width = 34
                    elif any(
                        k in header_text
                        for k in ["Memnuniyet", "Motivasyon", "İletişimi", "Dengesi", "Sıklığı", "Değerlendirmesi"]
                    ):
                        ws.column_dimensions[col_letter].width = 18
                    else:
                        ws.column_dimensions[col_letter].width = max(16, min(28, len(header_text) + 5))
                else:
                    max_len = len(header_text)
                    for r_i in range(2, ws.max_row + 1):
                        v_len = len(str(ws.cell(row=r_i, column=col_idx).value or ""))
                        if v_len > max_len:
                            max_len = v_len
                    ws.column_dimensions[col_letter].width = min(48, max(18, max_len + 4))

    bio.seek(0)
    return bio.getvalue()


def analyze_department_and_units(df_scoped: pd.DataFrame, secili_gmy: str) -> dict:
    if df_scoped is None:
        df_scoped = pd.DataFrame()
    return _cached_department_analysis(_df_fingerprint(df_scoped), _df_to_json_safe(df_scoped), secili_gmy)


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

    soru_ozetleri = analyze_questions_across_pool(df_filtered)
    soru_slayt_gruplari = [soru_ozetleri[i : i + 2] for i in range(0, len(soru_ozetleri), 2)]
    dept_analiz = analyze_department_and_units(df_filtered, secili_gmy)
    birim_gruplari = [dept_analiz["birimler"][i : i + 4] for i in range(0, len(dept_analiz["birimler"]), 4)]

    total_slides = 6 + max(1, len(soru_slayt_gruplari)) + max(1, len(birim_gruplari))
    slide_counter = 1

    n_kisi = len(df_filtered)
    birim_sayisi = df_filtered["birim"].nunique() if not df_filtered.empty and "birim" in df_filtered.columns else 0
    puan_ort = {}
    for col, label in PUAN_BASLIKLARI:
        if col in df_filtered.columns:
            s_num = pd.to_numeric(df_filtered[col], errors="coerce").dropna()
            if not s_num.empty:
                puan_ort[label] = round(float(s_num.mean()), 2)

    en_yuksek = sorted(puan_ort.items(), key=lambda x: x[1], reverse=True)[:3]
    en_dusuk = sorted(puan_ort.items(), key=lambda x: x[1])[:3]

    s1 = prs.slides.add_slide(blank_layout)
    add_header(s1, slide_counter, total_slides, "Executive Summary", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    add_card(
        s1,
        0.6,
        1.35,
        5.9,
        5.2,
        "Öne Çıkan Yüksek Başlıklar (1-5)",
        [f"{k}: {v}/5.00" for k, v in en_yuksek] or ["Henüz puan verisi yok"],
        c_teal,
    )
    add_card(
        s1,
        6.8,
        1.35,
        5.9,
        5.2,
        "Öncelikli Gelişim / Takip Başlıkları",
        [f"{k}: {v}/5.00" for k, v in en_dusuk] or ["Henüz puan verisi yok"],
        c_purple,
    )

    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, slide_counter, total_slides, "Görüşme Kapsamı", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    modeller = (
        df_filtered["calisma_modeli"].value_counts().to_dict()
        if "calisma_modeli" in df_filtered.columns and not df_filtered.empty
        else {}
    )
    add_card(
        s2,
        0.6,
        1.35,
        5.9,
        5.2,
        "Kapsam Metrikleri",
        [
            f"Toplam Görüşülen Kişi: {n_kisi}",
            f"Kapsanan Birim Sayısı: {birim_sayisi}",
            f"Aksiyon İhtiyacı Olan Kayıt: {int((df_filtered['aksiyon_ihtiyaci'] == 'Evet').sum()) if 'aksiyon_ihtiyaci' in df_filtered.columns and not df_filtered.empty else 0}",
        ],
        c_teal,
    )
    add_card(
        s2,
        6.8,
        1.35,
        5.9,
        5.2,
        "Çalışma Modeli Dağılımı",
        [f"{k}: {v} kişi" for k, v in modeller.items()] or ["Veri bulunmuyor"],
        c_purple,
    )

    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, slide_counter, total_slides, "Genel Değerlendirme (1-5 Puan Ortalamaları)", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    tum_puanlar = [f"{k}: {v}" for k, v in sorted(puan_ort.items(), key=lambda x: x[1], reverse=True)]
    yarim = (len(tum_puanlar) + 1) // 2
    add_card(s3, 0.6, 1.35, 5.9, 5.2, "Kriter Ortalamaları (1/2)", tum_puanlar[:yarim] or ["Veri yok"], c_teal)
    add_card(s3, 6.8, 1.35, 5.9, 5.2, "Kriter Ortalamaları (2/2)", tum_puanlar[yarim:] or ["Veri yok"], c_purple)

    if soru_slayt_gruplari:
        for grp_idx, grup in enumerate(soru_slayt_gruplari, start=1):
            s_q = prs.slides.add_slide(blank_layout)
            add_header(
                s_q,
                slide_counter,
                total_slides,
                f"Soru Bazlı Benzer / Ortak Noktalar ({grp_idx}/{len(soru_slayt_gruplari)})",
                f"{secili_gmy}  |  Dönem: {secili_donem}",
            )
            slide_counter += 1
            for slot_i, s_item in enumerate(grup):
                top_y = 1.32 + slot_i * 2.75
                pos_str = " ; ".join(s_item["olumlu_noktalar"]) if s_item["olumlu_noktalar"] else "Öne çıkan olumlu ifade yok"
                neg_str = " ; ".join(s_item["olumsuz_noktalar"]) if s_item["olumsuz_noktalar"] else "Öne çıkan olumsuz ifade yok"
                add_card(
                    s_q,
                    0.6,
                    top_y,
                    12.1,
                    2.55,
                    f"{s_item['soru']}  ({s_item['yanit_sayisi']} yanıt)",
                    [
                        f"Öne Çıkan Olumlu Noktalar: {pos_str}",
                        f"Öne Çıkan Olumsuz Noktalar: {neg_str}",
                    ],
                    c_teal if slot_i % 2 == 0 else c_purple,
                    font_pt=10.0,
                )
    else:
        s_q = prs.slides.add_slide(blank_layout)
        add_header(s_q, slide_counter, total_slides, "Soru Bazlı Olumlu & Olumsuz Noktalar", f"{secili_gmy}  |  Dönem: {secili_donem}")
        slide_counter += 1
        add_card(s_q, 0.6, 1.35, 12.1, 5.2, "Soru Açıklamaları", ["Seçili kapsamda yanıtlanmış soru açıklaması bulunmuyor."], c_teal)

    s_dept = prs.slides.add_slide(blank_layout)
    add_header(
        s_dept,
        slide_counter,
        total_slides,
        "Departman (GMY) Geneli — Olumlu & Olumsuz Değerlendirme",
        f"{secili_gmy}  |  Görüşme Sayısı: {dept_analiz['gorusme_sayisi']}  |  Dönem: {secili_donem}",
    )
    slide_counter += 1
    add_card(s_dept, 0.6, 1.35, 5.9, 5.2, "Departman Geneli Öne Çıkan Olumlu Yönler", dept_analiz["olumlu"], c_teal)
    add_card(s_dept, 6.8, 1.35, 5.9, 5.2, "Departman Geneli Öne Çıkan Olumsuz Yönler", dept_analiz["olumsuz"], c_purple)

    if birim_gruplari:
        for b_idx, b_grp in enumerate(birim_gruplari, start=1):
            s_b = prs.slides.add_slide(blank_layout)
            add_header(
                s_b,
                slide_counter,
                total_slides,
                f"Birim Bazlı Olumlu & Olumsuz Kırılım ({b_idx}/{len(birim_gruplari)})",
                f"{secili_gmy}  |  Dönem: {secili_donem}",
            )
            slide_counter += 1
            for idx_in_page, b_item in enumerate(b_grp):
                col_x = 0.6 if idx_in_page % 2 == 0 else 6.8
                row_y = 1.32 + (idx_in_page // 2) * 2.72
                b_pos_txt = " / ".join(b_item["olumlu"])
                b_neg_txt = " / ".join(b_item["olumsuz"])
                add_card(
                    s_b,
                    col_x,
                    row_y,
                    5.9,
                    2.52,
                    f"{b_item['birim']} ({b_item['gorusme_sayisi']} görüşme — {b_item['durum']})",
                    [
                        f"Olumlu: {b_pos_txt}",
                        f"Olumsuz: {b_neg_txt}",
                    ],
                    c_purple if "Dikkat" in b_item["durum"] else c_teal,
                    font_pt=9.5,
                )
    else:
        s_b = prs.slides.add_slide(blank_layout)
        add_header(s_b, slide_counter, total_slides, "Birim Bazlı Olumlu & Olumsuz Kırılım", f"{secili_gmy}  |  Dönem: {secili_donem}")
        slide_counter += 1
        add_card(s_b, 0.6, 1.35, 12.1, 5.2, "Birim Kırılımı", ["Seçili kapsamda birim kaydı bulunmuyor."], c_teal)

    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, slide_counter, total_slides, "Öne Çıkan Temalar ve Duygu Dağılımı", f"{secili_gmy}  |  Dönem: {secili_donem}")
    slide_counter += 1
    temalar = (
        [f"{k}: {v}" for k, v in df_filtered["ana_tema"].value_counts().head(6).items()]
        if "ana_tema" in df_filtered.columns and not df_filtered.empty
        else ["Veri yok"]
    )
    duygular = (
        [f"{k}: {v}" for k, v in df_filtered["duygu_durumu"].value_counts().items()]
        if "duygu_durumu" in df_filtered.columns and not df_filtered.empty
        else ["Veri yok"]
    )
    add_card(s6, 0.6, 1.35, 5.9, 5.2, "Öne Çıkan Ana Temalar", temalar, c_teal)
    add_card(s6, 6.8, 1.35, 5.9, 5.2, "Genel Duygu Dağılımı", duygular, c_purple)

    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, slide_counter, total_slides, "İK Önerileri ve Aksiyon Planı", f"{secili_gmy}  |  Dönem: {secili_donem}")
    aksiyonlar = (
        df_filtered[df_filtered["aksiyon_ihtiyaci"] == "Evet"]
        if "aksiyon_ihtiyaci" in df_filtered.columns and not df_filtered.empty
        else pd.DataFrame()
    )
    aksiyon_satirlari = []
    for _, r in aksiyonlar.head(6).iterrows():
        not_metni = str(r.get("aksiyon_notu", "") or r.get("q10_degistirmek_istenen", "") or "").strip()
        sev = str(r.get("aksiyon_seviyesi", "Birim"))
        bir = str(r.get("birim", ""))
        aksiyon_satirlari.append(f"[{bir} | {sev}] {not_metni}")

    add_card(
        s7,
        0.6,
        1.35,
        12.1,
        5.2,
        "Öncelikli Aksiyon Maddeleri",
        aksiyon_satirlari or ["Seçilen filtrede işaretlenmiş aksiyon notu bulunmuyor."],
        c_teal,
    )

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
            <div class="panel-brand"><span class="panel-accent">&lt;</span> izlenim<span class="panel-accent">.</span> <span class="panel-purple">|</span> İzlenim Paylaşım Paneli</div>
            <div style="color:#94A3B8; margin-top:0.3rem;">İK İş Ortakları için 1-1 Görüşme Veri Havuzu, Standart Dashboard ve Sunum Üretici</div>
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

st.markdown(
    f"""
    <div class="panel-header">
        <div class="panel-brand"><span class="panel-accent">&lt;</span> izlenim<span class="panel-accent">.</span> <span class="panel-purple">|</span> {sayfa}</div>
        <div style="color:#94A3B8; margin-top:0.25rem;">Aktif Kullanıcı: <b>{user['ad_soyad']}</b> &nbsp;•&nbsp; {data_source_label}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

assigned_gmy_map = user["gmy_birimler"]
gmy_options = list(assigned_gmy_map.keys())

if sayfa == "1-1 Görüşme Kayıt":
    c1, c2, c3 = st.columns(3)
    with c1:
        sec_gmy = st.selectbox("GMY / Domain", gmy_options)
    with c2:
        birim_options = assigned_gmy_map.get(sec_gmy, GMY_BIRIM_MAP.get(sec_gmy, []))
        sec_birim = st.selectbox("Birim / Ekip", birim_options)
    with c3:
        sec_donem = st.selectbox(
            "Görüşme Dönemi",
            ["2026 - 1. Dönem", "2026 - 2. Dönem", "2026 - Q1", "2026 - Q2", "2026 - Q3", "2026 - Q4"],
            index=1,
        )

    c4, c5, c6, c7 = st.columns(4)
    with c4:
        gorusme_tarihi = st.date_input("Görüşme Tarihi", value=date.today())
    with c5:
        katilimci_turu = st.selectbox("Katılımcı Tipi", ["Çalışan", "Direktör"])
    with c6:
        calisan_rumuzu = st.text_input("Çalışan / Katılımcı Kodu veya Adı", placeholder="örnek: Çalışan #05 veya Anonim")
    with c7:
        calisma_modeli = st.selectbox("Çalışma Modeli", ["Hibrit", "Remote", "Ofis / Saha"])

    c8, c9, c10, c11 = st.columns(4)
    with c8:
        duygu_durumu = st.selectbox("Genel Duygu Sınıflandırması", ["Olumlu", "Nötr", "Olumsuz"])
    with c9:
        ana_tema = st.selectbox("Öne Çıkan Ana Tema", ANA_TEMALAR)
    with c10:
        aksiyon_ihtiyaci = st.selectbox("Aksiyon İhtiyacı", ["Hayır", "Evet"])
    with c11:
        aksiyon_seviyesi = st.selectbox("Aksiyon Seviyesi", ["Yok", "Birim / YM", "GMY / Üst Yönetim", "İK / Kurumsal"])

    aksiyon_notu = st.text_input(
        "Aksiyon Notu (Varsa)",
        placeholder="Örn: Harmony metrikleri veya karar alma süreci için önerilen adım...",
    )

    tab1, tab2, tab3 = st.tabs([
        "💬 15 Görüşme Sorusu",
        "🏢 İnsan Kaynakları & Kapanış",
        "⭐ 1-5 Sohbet İçi Puanlama",
    ])

    q_answers = {}
    ik_answers = {}
    p_scores = {}

    with tab1:
        for q_key, q_text in GORUSME_SORULARI:
            q_answers[q_key] = st.text_area(q_text, height=70, key=f"inp_{q_key}")

    with tab2:
        for ik_key, ik_label in IK_BASLIKLARI:
            ik_answers[ik_key] = st.text_area(f"İK Başlığı — {ik_label}", height=70, key=f"inp_{ik_key}")
        q_kapanis = st.text_area(
            "Paylaşmak istediğin veya eklemek istediğin farklı konu var mı?",
            height=80,
            key="inp_q_kapanis",
        )

    with tab3:
        st.caption("Görüşme akışında puanlanmayan konuları boş bırakabilirsin; boş bırakılanlar ortalamalara dahil edilmez.")
        for p_key, p_label in PUAN_BASLIKLARI:
            p_scores[p_key] = st.select_slider(
                p_label,
                options=["Boş", 1, 2, 3, 4, 5],
                value="Boş",
                key=f"inp_{p_key}",
            )

    if st.button("Görüşme Kaydını Havuza Aktar", type="primary", use_container_width=True):
        new_record = {
            "kayit_id": datetime.now().strftime("%Y%m%d-%H%M%S"),
            "kayit_zamani": f"{gorusme_tarihi.strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M')}",
            "donem": sec_donem,
            "hrbp_kullanici": user["username"],
            "gmy": sec_gmy,
            "birim": sec_birim,
            "katilimci_turu": katilimci_turu,
            "calisan_rumuzu": calisan_rumuzu.strip() or "Anonim",
            "calisma_modeli": calisma_modeli,
            "duygu_durumu": duygu_durumu,
            "ana_tema": ana_tema,
            "aksiyon_ihtiyaci": aksiyon_ihtiyaci,
            "aksiyon_seviyesi": aksiyon_seviyesi,
            "aksiyon_notu": aksiyon_notu.strip(),
            "q_kapanis": q_kapanis.strip(),
            **{k: v.strip() for k, v in q_answers.items()},
            **{k: v.strip() for k, v in ik_answers.items()},
            **{k: (None if v == "Boş" else float(v)) for k, v in p_scores.items()},
        }
        ok, msg = save_record(new_record)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

elif sayfa == "Standart Dashboard & Sunum":
    allowed_birimler = [b for b_list in assigned_gmy_map.values() for b in b_list]
    if user["rol"] != "admin" and not df_all.empty and "birim" in df_all.columns:
        df_scoped = df_all[df_all["birim"].isin(allowed_birimler)].copy()
    else:
        df_scoped = df_all.copy()

    f1, f2, f3 = st.columns(3)
    with f1:
        gmy_filter_opts = ["Tümü"] + gmy_options
        sel_gmy_f = st.selectbox("GMY / Domain Filtresi", gmy_filter_opts)
    with f2:
        if sel_gmy_f == "Tümü":
            birim_filter_opts = ["Tümü"] + allowed_birimler
        else:
            birim_filter_opts = ["Tümü"] + assigned_gmy_map.get(sel_gmy_f, [])
        sel_birim_f = st.selectbox("Birim / Ekip Filtresi", birim_filter_opts)
    with f3:
        donem_vals = (
            sorted(df_scoped["donem"].dropna().unique().tolist())
            if "donem" in df_scoped.columns and not df_scoped.empty
            else ["2026 - 2. Dönem"]
        )
        sel_donem_f = st.selectbox(
            "Görüşme Dönemi",
            ["Tümü"] + donem_vals,
            index=len(donem_vals) if donem_vals else 0,
        )

    df_view = df_scoped.copy()
    if not df_view.empty:
        if sel_gmy_f != "Tümü" and "gmy" in df_view.columns:
            df_view = df_view[df_view["gmy"] == sel_gmy_f]
        if sel_birim_f != "Tümü" and "birim" in df_view.columns:
            df_view = df_view[df_view["birim"] == sel_birim_f]
        if sel_donem_f != "Tümü" and "donem" in df_view.columns:
            df_view = df_view[df_view["donem"] == sel_donem_f]

    k1, k2, k3, k4 = st.columns(4)
    puan_cols = [c for c, _ in PUAN_BASLIKLARI if c in df_view.columns]
    genel_ort = (
        round(float(df_view[puan_cols].apply(pd.to_numeric, errors="coerce").stack().mean()), 2)
        if not df_view.empty
        and puan_cols
        and not df_view[puan_cols].apply(pd.to_numeric, errors="coerce").stack().empty
        else 0.0
    )
    aksiyon_adet = (
        int((df_view["aksiyon_ihtiyaci"] == "Evet").sum())
        if "aksiyon_ihtiyaci" in df_view.columns and not df_view.empty
        else 0
    )

    for col_box, lbl, val, sub in [
        (k1, "Görüşülen Kişi Sayısı", str(len(df_view)), "Seçili kapsam"),
        (
            k2,
            "Kapsanan Birim Sayısı",
            str(df_view["birim"].nunique() if not df_view.empty and "birim" in df_view.columns else 0),
            "Aktif ekipler",
        ),
        (k3, "Genel Puan Ortalaması", f"{genel_ort} / 5.00", "Boş puanlar hariç"),
        (k4, "Aksiyon Gerektiren Kayıt", str(aksiyon_adet), "Takip bekleyen"),
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

    pptx_state_key = f"pptx_{_df_fingerprint(df_view)}_{sel_gmy_f}_{sel_donem_f}"
    if st.button("📊 Standart Direktör Sunumu Hazırla (.pptx)", use_container_width=True):
        with st.spinner("Sunum oluşturuluyor..."):
            st.session_state[pptx_state_key] = _cached_pptx_bytes(
                _df_fingerprint(df_view),
                _df_to_json_safe(df_view),
                sel_gmy_f,
                sel_donem_f,
            )
    if pptx_state_key in st.session_state:
        st.download_button(
            label="📥 Hazırlanan Sunumu İndir (.pptx)",
            data=st.session_state[pptx_state_key],
            file_name=f"Izlenim_Paylasim_{sel_donem_f.replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    if df_view.empty:
        st.info("Seçili filtrelere uygun görüşme kaydı bulunamadı.")
    else:
        g_left, g_right = st.columns([1.3, 1])

        with g_left:
            st.subheader("11 Başlık İçin Ortalama Puanlar (1 - 5)")
            ort_rows = []
            for col_k, lbl_k in PUAN_BASLIKLARI:
                if col_k in df_view.columns:
                    s_num = pd.to_numeric(df_view[col_k], errors="coerce").dropna()
                    if not s_num.empty:
                        ort_rows.append({"Başlık": lbl_k, "Ortalama": round(float(s_num.mean()), 2)})
            if ort_rows:
                df_ort = pd.DataFrame(ort_rows).sort_values("Ortalama", ascending=True)
                fig_bar = px.bar(
                    df_ort,
                    x="Ortalama",
                    y="Başlık",
                    orientation="h",
                    range_x=[0, 5],
                    text="Ortalama",
                    color="Ortalama",
                    color_continuous_scale=[NEG_COLOR, PURPLE, TEAL],
                )
                fig_bar.update_layout(
                    paper_bgcolor=CARD_DARK,
                    plot_bgcolor=CARD_DARK,
                    font_color=TEXT_LIGHT,
                    coloraxis_showscale=False,
                    margin=dict(l=10, r=20, t=20, b=20),
                    height=430,
                )
                st.plotly_chart(fig_bar, use_container_width=True)

        with g_right:
            st.subheader("Çalışma Modeli & Duygu Dağılımı")
            if "calisma_modeli" in df_view.columns:
                df_mod = df_view["calisma_modeli"].value_counts().reset_index()
                df_mod.columns = ["Model", "Adet"]
                fig_donut = go.Figure(
                    data=[
                        go.Pie(
                            labels=df_mod["Model"],
                            values=df_mod["Adet"],
                            hole=0.6,
                            marker=dict(colors=[TEAL, PURPLE, TEAL_SOFT]),
                        )
                    ]
                )
                fig_donut.update_layout(
                    paper_bgcolor=CARD_DARK,
                    plot_bgcolor=CARD_DARK,
                    font_color=TEXT_LIGHT,
                    margin=dict(l=10, r=10, t=20, b=20),
                    height=205,
                )
                st.plotly_chart(fig_donut, use_container_width=True)

            if "duygu_durumu" in df_view.columns:
                df_duy = df_view["duygu_durumu"].value_counts().reset_index()
                df_duy.columns = ["Duygu", "Adet"]
                color_map = {"Olumlu": POS_COLOR, "Nötr": NEU_COLOR, "Olumsuz": NEG_COLOR}
                fig_duy = px.bar(
                    df_duy,
                    x="Duygu",
                    y="Adet",
                    color="Duygu",
                    color_discrete_map=color_map,
                )
                fig_duy.update_layout(
                    paper_bgcolor=CARD_DARK,
                    plot_bgcolor=CARD_DARK,
                    font_color=TEXT_LIGHT,
                    showlegend=False,
                    margin=dict(l=10, r=10, t=20, b=20),
                    height=205,
                )
                st.plotly_chart(fig_duy, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Departman Geneli & Birim Bazlı Olumlu / Olumsuz Kırılım")
        dept_cache_key = f"dept_{_df_fingerprint(df_view)}_{sel_gmy_f}"
        if dept_cache_key not in st.session_state:
            with st.spinner("Departman ve birim kırılımı hazırlanıyor..."):
                st.session_state[dept_cache_key] = _analyze_department_impl(df_view, sel_gmy_f)
        dept_res = st.session_state[dept_cache_key]

        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.markdown("**Departman Geneli — Öne Çıkan Olumlu Yönler**")
            olumlu_list = dept_res.get("olumlu") or ["Kayıtlı olumlu ifade bulunmuyor"]
            for item in olumlu_list:
                st.markdown(f"- {item}")
        with d_col2:
            st.markdown("**Departman Geneli — Öne Çıkan Olumsuz Yönler**")
            olumsuz_list = dept_res.get("olumsuz") or ["Kayıtlı olumsuz ifade bulunmuyor"]
            for item in olumsuz_list:
                st.markdown(f"- {item}")

        st.markdown("<br>", unsafe_allow_html=True)
        birimler_all = dept_res.get("birimler") or []
        if not birimler_all:
            st.info("Seçili kapsamda birim kırılımı bulunmuyor.")
        else:
            tum_birim_goster = st.checkbox(
                f"Tüm birimleri göster ({len(birimler_all)} birim)",
                value=False,
                key=f"chk_all_units_{sel_gmy_f}",
            )
            gosterilecek_birimler = birimler_all if tum_birim_goster else birimler_all[:10]
            if not tum_birim_goster and len(birimler_all) > 10:
                st.caption(f"Performans için ilk 10 birim gösteriliyor (toplam {len(birimler_all)} birim). Tamamını görmek için yukarıdaki kutuyu işaretle.")
            for b_item in gosterilecek_birimler:
                b_pos_str = " / ".join(b_item.get("olumlu") or ["Kayıt yok"])
                b_neg_str = " / ".join(b_item.get("olumsuz") or ["Kayıt yok"])
                st.markdown(
                    f"""
                    <div class="detail-box">
                        <div style="font-weight:700; color:{TEAL}; margin-bottom:0.35rem;">
                            {b_item['birim']} &nbsp;•&nbsp; {b_item['gorusme_sayisi']} Görüşme &nbsp;•&nbsp; <span style="color:{PURPLE_LIGHT};">{b_item['durum']}</span>
                        </div>
                        <div><b>Olumlu:</b> {b_pos_str}</div>
                        <div><b>Olumsuz:</b> {b_neg_str}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Her Soru İçin Benzer / Ortak Noktalar (Anlamsal Gruplama)")
        q_cache_key = f"qsum_{_df_fingerprint(df_view)}"
        if q_cache_key not in st.session_state:
            st.session_state[q_cache_key] = _analyze_questions_impl(df_view)
        soru_ozet_list = st.session_state[q_cache_key]
        for s_oz in soru_ozet_list:
            with st.expander(f"{s_oz['soru']} ({s_oz['yanit_sayisi']} yanıt)"):
                st.markdown("**Öne Çıkan Olumlu Noktalar:**")
                for p_txt in s_oz["olumlu_noktalar"] or ["Öne çıkan olumlu ifade yok"]:
                    st.markdown(f"- {p_txt}")
                st.markdown("**Öne Çıkan Olumsuz Noktalar:**")
                for n_txt in s_oz["olumsuz_noktalar"] or ["Öne çıkan olumsuz ifade yok"]:
                    st.markdown(f"- {n_txt}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"Çalışan Yorumları — Bireysel Görüş Kartları ({len(df_view)} Kayıt)")
        tum_kart_goster = st.checkbox(
            f"Tüm çalışan kartlarını göster ({len(df_view)} kayıt)",
            value=False,
            key="chk_all_cards",
        )
        kart_df = df_view if tum_kart_goster else df_view.head(25)
        if not tum_kart_goster and len(df_view) > 25:
            st.caption(f"Sayfanın kasmaması için ilk 25 kayıt gösteriliyor (toplam {len(df_view)} kayıt).")
        for _, r_item in kart_df.iterrows():
            kisi_adi = r_item.get("calisan_rumuzu", "Katılımcı")
            birim_adi = r_item.get("birim", "")
            duygu_adi = r_item.get("duygu_durumu", "Olumlu")
            mem_puan = r_item.get("p_genel_memnuniyet", "-")
            q01_txt = r_item.get("q01_calisma_hayati", "")
            q02_txt = r_item.get("q02_guclu_yonler", "")
            q04_txt = r_item.get("q04_ekip_iyilestirme", "")
            q10_txt = r_item.get("q10_degistirmek_istenen", "")
            st.markdown(
                f"""
                <div class="detail-box">
                    <div style="font-weight:700; color:{TEAL}; margin-bottom:0.35rem;">
                        {kisi_adi} &nbsp;•&nbsp; {birim_adi} &nbsp;•&nbsp; <span style="color:{PURPLE_LIGHT};">{duygu_adi}</span> &nbsp;•&nbsp; Memnuniyet: <b>{mem_puan}</b>
                    </div>
                    <div><b>Çalışma Hayatı:</b> {q01_txt}</div>
                    <div><b>Güçlü Yönler:</b> {q02_txt}</div>
                    <div><b>İyileştirme / Ekip:</b> {q04_txt}</div>
                    <div><b>Değiştirmek İstenen:</b> {q10_txt}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

else:
    st.subheader("Kayıt Havuzu")
    if not df_all.empty:
        arama_metni = st.text_input(
            "Havuzda Ara (Katılımcı, Birim, Yorum veya Tema)",
            placeholder="örnek: uzaktan, onay, terfi, Mobil...",
        )
        if arama_metni and arama_metni.strip():
            q_low = arama_metni.strip().lower()
            mask = df_all.apply(lambda s: s.astype(str).str.lower().str.contains(q_low, na=False)).any(axis=1)
            df_all = df_all[mask].copy()
            st.caption(f"Filtrelenen kayıt sayısı: {len(df_all)}")

    if df_all.empty:
        st.info("Henüz kayıt bulunmuyor.")
    else:
        df_display = df_all.rename(columns=COLUMN_LABELS)
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        excel_data = build_excel_export(df_all)
        st.download_button(
            label="📥 Havuzu Excel Olarak İndir (.xlsx)",
            data=excel_data,
            file_name=f"Izlenim_Paylasim_Havuz_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
