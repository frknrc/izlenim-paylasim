import io
from datetime import date, datetime
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
import streamlit as st

# ==============================================================================
# SAYFA AYARLARI VE SOFTTECH KURUMSAL TEMA
# ==============================================================================
st.set_page_config(
    page_title="Softtech İzlenim Paylaşım Paneli",
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
        .softtech-header {{
            background: linear-gradient(120deg, {CARD_DARK} 0%, #1B1638 60%, #0E293B 100%);
            border: 1px solid {CARD_BORDER};
            border-left: 5px solid {TEAL};
            border-radius: 12px;
            padding: 1.2rem 1.5rem;
            margin-bottom: 1rem;
        }}
        .softtech-brand {{
            font-size: 1.5rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            color: {TEXT_LIGHT};
        }}
        .softtech-accent {{
            color: {TEAL};
        }}
        .softtech-purple {{
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
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# GMY / BİRİM HİYerarşisi
# ==============================================================================
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
        "sifre": "softtech2026",
        "rol": "hrbp",
        "gmy_birimler": {
            "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri": [
                "10000474-Kurumsal Uygulamalar",
                "10000475-Mobil Uygulamalar",
                "10000127-Çok Kanallı Bankacılık",
            ],
            "10000028-Teknoloji, Yapay Zeka ve Sürekli İyileştirme": [
                "10000464-İş Analizi Yetkinlik Merkezi",
                "10000483-Yapay Zeka Yetkinlik Merkezi",
                "10000486-Mimari Yetkinlik Merkezi",
            ],
        },
    },
    "hrbp_ornek": {
        "ad_soyad": "Örnek HRBP",
        "sifre": "softtech2026",
        "rol": "hrbp",
        "gmy_birimler": {
            "10000023-Mevduat, Kredi, Süreç ve Ödemeler": [
                "10000134-Ödemeler",
                "10000160-Ticari Krediler",
            ],
        },
    },
    "admin": {
        "ad_soyad": "İK Raporlama Admin",
        "sifre": "admin2026",
        "rol": "admin",
        "gmy_birimler": GMY_BIRIM_MAP,
    },
}

GORUSME_SORULARI = [
    ("q01_calisma_hayati", "Softtech'te çalışma hayatın nasıl gidiyor?"),
    ("q02_guclu_yonler", "Sana göre Softtech'teki olumlu/güçlü yönler neler?"),
    ("q03_ekip_rol_tatmin", "Görev aldığın ekip, rolün ve projelerin seni tatmin ediyor, sana katkı sağlıyor mu?"),
    ("q04_ekip_iyilestirme", "Ekibinde sana göre daha iyi olması gereken şeyler neler?"),
    ("q05_motivasyon_arttiran", "Kendi motivasyonunu daha çok arttıracak şeyler neler olabilir? (İş, proje, kariyer vb.)"),
    ("q06_yonetici_geri_bildirim", "Yöneticin gelişimin için zaman ayırıp düzenli geri bildirim verebiliyor mu? Görüşme sıklığı nasıl?"),
    ("q07_ym_verimlilik", "Yetkinlik Merkezi yapısının verimliliğini nasıl değerlendirirsin? (Eğitim, iş süreçleri vb.)"),
    ("q08_gmy_iletisimi", "GMY iletişimi ile ilgili değerlendirmelerin nasıl olur?"),
    ("q09_vizyon_strateji", "Şirketin sunduğu vizyon ve stratejiyi büyük resimde nasıl değerlendiriyorsun?"),
    ("q10_degistirmek_istenen", "Softtech'te neyi değiştirmek isterdin?"),
    ("q11_mevcut_calisma_modeli", "Mevcut çalışma modelinle ilgili ne düşünüyorsun?"),
    ("q12_ilerisi_calisma_modeli", "İlerisi için nasıl bir çalışma modeli olursa daha verimli çalışacağına inanıyorsun?"),
    ("q13_harmony", "Harmony ile ilgili görüşlerin nasıl, çalışma sürecine olan etkisini nasıl değerlendirirsin?"),
    ("q14_toplandigimiza_degsin", "\"Toplandığımıza Değsin\" temasının ne kadar işlendiğini düşünüyorsun?"),
    ("q15_karar_alma_surecleri", "Karar alma süreçlerini hem ekip içi işleyişler hem de üst yönetim seviyesinde nasıl değerlendirirsin?"),
]

IK_BASLIKLARI = [
    ("ik_kariyer", "İşe Alım ve Kariyer Yönetimi"),
    ("ik_deneyim_marka", "Çalışan Deneyimi ve İşveren Markası Tanıtımı"),
    ("ik_ucret_merkezi", "Ücretlendirme ve Çalışan Merkezi"),
    ("ik_akademi", "Softtech Akademi"),
]

PUAN_BASLIKLARI = [
    ("p_softtech_memnuniyet", "Softtech Memnuniyeti"),
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


def get_sample_data() -> pd.DataFrame:
    rows = [
        {
            "kayit_id": "2026-001",
            "kayit_zamani": "2026-10-12 10:30",
            "donem": "2026 - 2. Dönem",
            "hrbp_kullanici": "seyma",
            "gmy": "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri",
            "birim": "10000474-Kurumsal Uygulamalar",
            "katilimci_turu": "Çalışan",
            "calisan_rumuzu": "Çalışan #01",
            "calisma_modeli": "Hibrit",
            "duygu_durumu": "Olumlu",
            "ana_tema": "Güven ve Kültür",
            "aksiyon_ihtiyaci": "Hayır",
            "aksiyon_seviyesi": "Yok",
            "aksiyon_notu": "",
            "q01_calisma_hayati": "Ekip içi iletişim çok güçlü, günlük iş akışı düzenli ilerliyor.",
            "q02_guclu_yonler": "Ulaşılabilir İK, güvenli çalışma ortamı ve destekleyici ekip kültürü.",
            "q03_ekip_rol_tatmin": "Sorumluluklarım net, yeni teknolojiler deneme şansı buluyorum.",
            "q04_ekip_iyilestirme": "Yoğun sprint dönemlerinde toplantı sürelerinin daha kısa tutulması.",
            "q05_motivasyon_arttiran": "Net kariyer patikası ve eğitim olanakları.",
            "q06_yonetici_geri_bildirim": "Yöneticim düzenli geri bildirim veriyor ve gelişimimi destekliyor.",
            "q07_ym_verimlilik": "Mimari ve proje yönetim YM desteği faydalı bulunuyor.",
            "q08_gmy_iletisimi": "GMY bilgilendirmeleri açık ve erişilebilir.",
            "q09_vizyon_strateji": "Dijitalleşme ve yapay zeka odağı büyük resimde anlamlı.",
            "q10_degistirmek_istenen": "Onay ve karar süreçlerinin bazı adımlarının hızlanması.",
            "q11_mevcut_calisma_modeli": "Mevcut hibrit düzenden memnunum.",
            "q12_ilerisi_calisma_modeli": "Esnek hibrit yapının korunması verimliliği artırır.",
            "q13_harmony": "Kriterlerin netleşmesiyle daha sağlıklı ölçülebilir.",
            "q14_toplandigimiza_degsin": "Gündemi net toplantılarda verim yüksek.",
            "q15_karar_alma_surecleri": "Birim içi hızlı, üst düzeyde zaman zaman daha yavaş.",
            "ik_kariyer": "Kariyer gelişim görüşmeleri yapıcı ilerliyor.",
            "ik_deneyim_marka": "Şirket içi etkinlikler ve kültür olumlu.",
            "ik_ucret_merkezi": "Ücret ve yan hakların piyasa ile uyumlu kalması önemli.",
            "ik_akademi": "Akademi eğitim içerikleri güncel ve pratik.",
            "q_kapanis": "Eklemek istediğim farklı bir konu yok.",
            "p_softtech_memnuniyet": 4.4,
            "p_ekip_motivasyonu": 4.2,
            "p_kendi_motivasyonu": 4.3,
            "p_is_yasam_dengesi": 3.8,
            "p_1_yonetici_iletisimi": 4.5,
            "p_ym_degerlendirme": 3.6,
            "p_direktor_iletisimi": 4.2,
            "p_gmy_iletisimi": 4.0,
            "p_ucret_yan_hak": 3.4,
            "p_ik_iletisimi": 4.4,
            "p_yz_kullanim": 4.1,
        },
        {
            "kayit_id": "2026-002",
            "kayit_zamani": "2026-10-15 14:15",
            "donem": "2026 - 2. Dönem",
            "hrbp_kullanici": "seyma",
            "gmy": "10000028-Teknoloji, Yapay Zeka ve Sürekli İyileştirme",
            "birim": "10000464-İş Analizi Yetkinlik Merkezi",
            "katilimci_turu": "Çalışan",
            "calisan_rumuzu": "Çalışan #02",
            "calisma_modeli": "Hibrit",
            "duygu_durumu": "Olumsuz",
            "ana_tema": "Harmony ve Karne Etkisi",
            "aksiyon_ihtiyaci": "Evet",
            "aksiyon_seviyesi": "Birim / YM",
            "aksiyon_notu": "Harmony metriklerinin saha iş yüküyle uyumu gözden geçirilmeli.",
            "q01_calisma_hayati": "İş yoğunluğu dönem dönem artıyor, önceliklendirme kritik.",
            "q02_guclu_yonler": "Ekip arkadaşlığı ve teknik bilgi paylaşımı.",
            "q03_ekip_rol_tatmin": "Farklı domainlerle çalışmak perspektif kazandırıyor.",
            "q04_ekip_iyilestirme": "İş birimleriyle erken aşamada kapsam netleştirme.",
            "q05_motivasyon_arttiran": "Emek karşılığının net görünürlüğü ve eğitim bütçesi.",
            "q06_yonetici_geri_bildirim": "Düzenli görüşmeler yapılıyor.",
            "q07_ym_verimlilik": "YM süreçlerinde rol ve sorumlulukların netleşmesi bekleniyor.",
            "q08_gmy_iletisimi": "Düzenli paylaşımlar yapılıyor.",
            "q09_vizyon_strateji": "YZ dönüşümü doğru yönde ilerliyor.",
            "q10_degistirmek_istenen": "Harmony ölçüm ve karne yükünün azaltılması.",
            "q11_mevcut_calisma_modeli": "Hibrit model uygun ancak yoğun dönemlerde esneklik önemli.",
            "q12_ilerisi_calisma_modeli": "Çıktı odaklı hibrit model.",
            "q13_harmony": "Operasyonel yük yarattığı noktalar var.",
            "q14_toplandigimiza_degsin": "Bazı toplantılar aksiyon odaklı değil.",
            "q15_karar_alma_surecleri": "Kararların netleşmesi zaman alabiliyor.",
            "ik_kariyer": "Kariyer adımları daha şeffaf olabilir.",
            "ik_deneyim_marka": "İletişim kanalları aktif.",
            "ik_ucret_merkezi": "Prim ve ücret artış süreçlerinde netlik beklentisi var.",
            "ik_akademi": "Teknik eğitimlerin artırılması faydalı olur.",
            "q_kapanis": "Geri bildirimlerin takip edilmesi değerli.",
            "p_softtech_memnuniyet": 3.7,
            "p_ekip_motivasyonu": 3.6,
            "p_kendi_motivasyonu": 3.5,
            "p_is_yasam_dengesi": 3.1,
            "p_1_yonetici_iletisimi": 4.0,
            "p_ym_degerlendirme": 3.0,
            "p_direktor_iletisimi": 3.8,
            "p_gmy_iletisimi": 3.7,
            "p_ucret_yan_hak": 2.9,
            "p_ik_iletisimi": 4.1,
            "p_yz_kullanim": 3.8,
        },
        {
            "kayit_id": "2026-003",
            "kayit_zamani": "2026-10-18 11:00",
            "donem": "2026 - 2. Dönem",
            "hrbp_kullanici": "seyma",
            "gmy": "10000027-Dijital Bankacılık ve Uluslararası Bankacılık Çözümleri",
            "birim": "10000475-Mobil Uygulamalar",
            "katilimci_turu": "Direktör",
            "calisan_rumuzu": "Direktör #01",
            "calisma_modeli": "Hibrit",
            "duygu_durumu": "Olumlu",
            "ana_tema": "Karar Süreçleri ve Bilgi Akışı",
            "aksiyon_ihtiyaci": "Evet",
            "aksiyon_seviyesi": "GMY / Üst Yönetim",
            "aksiyon_notu": "Süreç değişikliklerinde önce direktörlere ön bilgilendirme yapılması.",
            "q01_calisma_hayati": "Ekip performansı ve iş birliği güçlü.",
            "q02_guclu_yonler": "Güven kültürü, ulaşılabilir İK ve güçlü teknik kadro.",
            "q03_ekip_rol_tatmin": "Rolüm ve ekibimin etkisi yüksek.",
            "q04_ekip_iyilestirme": "Platformlar arası koordinasyonun daha akıcı olması.",
            "q05_motivasyon_arttiran": "Ekip başarısı ve YZ araçlarının etkin kullanımı.",
            "q06_yonetici_geri_bildirim": "Açık ve çift yönlü geri bildirim kültürü var.",
            "q07_ym_verimlilik": "YZ ve mimari alanlarında beklentiler yüksek.",
            "q08_gmy_iletisimi": "Doğrudan ve yapıcı iletişim kurulabiliyor.",
            "q09_vizyon_strateji": "Stratejik hedefler net.",
            "q10_degistirmek_istenen": "Değişikliklerde bilgilendirme sırasının netleşmesi.",
            "q11_mevcut_calisma_modeli": "Hibrit çalışma verimli.",
            "q12_ilerisi_calisma_modeli": "Hibrit modelin sürdürülmesi.",
            "q13_harmony": "Metriklerin sadeleştirilmesi önerilir.",
            "q14_toplandigimiza_degsin": "Odaklı toplantılar değer yaratıyor.",
            "q15_karar_alma_surecleri": "Birim içi hızlı, üst ölçekte iyileşmeye açık.",
            "ik_kariyer": "İK iş ortaklığı desteği güçlü.",
            "ik_deneyim_marka": "Kurumsal marka bilinirliği yüksek.",
            "ik_ucret_merkezi": "Bütçe ve kota netliği önemli.",
            "ik_akademi": "Yurt dışı ve uzmanlık eğitimleri değerli.",
            "q_kapanis": "Çalışan sesinin düzenli toplanması çok kıymetli.",
            "p_softtech_memnuniyet": 4.3,
            "p_ekip_motivasyonu": 4.4,
            "p_kendi_motivasyonu": 4.2,
            "p_is_yasam_dengesi": 3.7,
            "p_1_yonetici_iletisimi": 4.4,
            "p_ym_degerlendirme": 3.5,
            "p_direktor_iletisimi": 4.5,
            "p_gmy_iletisimi": 4.1,
            "p_ucret_yan_hak": 3.3,
            "p_ik_iletisimi": 4.5,
            "p_yz_kullanim": 4.3,
        },
    ]
    return pd.DataFrame(rows)


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

    return st.session_state.local_df.copy(), "Demo / Yerel Mod (Google Sheets bağlı değil)"


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
    return True, "Kayıt demo/oturum verisine eklendi (Google Sheets bağlandığında doğrudan Sheets'e yazılacak)."


def build_pptx_report(df_filtered: pd.DataFrame, secili_gmy: str, secili_donem: str) -> bytes:
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

    def add_header(slide, slide_num: int, title: str, subtitle: str = ""):
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
        fp.text = f"softtech.com.tr  |  İzlenim Paylaşım Değerlendirmesi  |  Slayt {slide_num}/7"
        fp.font.size = Pt(9.5)
        fp.font.color.rgb = c_muted

    def add_card(slide, left, top, width, height, heading, body_lines, accent=c_teal):
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
        p0.font.size = Pt(14)
        p0.font.bold = True
        p0.font.color.rgb = accent
        p0.space_after = Pt(6)

        for line in body_lines:
            p = tf.add_paragraph()
            p.text = f"• {line}" if not line.startswith("•") else line
            p.font.size = Pt(11.5)
            p.font.color.rgb = c_white
            p.space_after = Pt(4)

    n_kisi = len(df_filtered)
    birim_sayisi = df_filtered["birim"].nunique() if not df_filtered.empty else 0

    puan_ort = {}
    for col, label in PUAN_BASLIKLARI:
        if col in df_filtered.columns and not df_filtered[col].dropna().empty:
            puan_ort[label] = round(float(df_filtered[col].dropna().mean()), 2)

    en_yuksek = sorted(puan_ort.items(), key=lambda x: x[1], reverse=True)[:3]
    en_dusuk = sorted(puan_ort.items(), key=lambda x: x[1])[:3]

    s1 = prs.slides.add_slide(blank_layout)
    add_header(s1, 1, "Executive Summary", f"{secili_gmy}  |  Dönem: {secili_donem}")
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
    add_header(s2, 2, "Görüşme Kapsamı", f"{secili_gmy}  |  Dönem: {secili_donem}")
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
            f"Aksiyon İhtiyacı Olan Kayıt: {int((df_filtered['aksiyon_ihtiyaci'] == 'Evet').sum()) if 'aksiyon_ihtiyaci' in df_filtered.columns else 0}",
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
    add_header(s3, 3, "Genel Değerlendirme (1-5 Puan Ortalamaları)", f"{secili_gmy}  |  Dönem: {secili_donem}")
    tum_puanlar = [f"{k}: {v}" for k, v in sorted(puan_ort.items(), key=lambda x: x[1], reverse=True)]
    yarim = (len(tum_puanlar) + 1) // 2
    add_card(s3, 0.6, 1.35, 5.9, 5.2, "Kriter Ortalamaları (1/2)", tum_puanlar[:yarim] or ["Veri yok"], c_teal)
    add_card(s3, 6.8, 1.35, 5.9, 5.2, "Kriter Ortalamaları (2/2)", tum_puanlar[yarim:] or ["Veri yok"], c_purple)

    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, 4, "Birim Bazlı Karşılaştırma", f"{secili_gmy}  |  Dönem: {secili_donem}")
    birim_ozet = []
    if not df_filtered.empty:
        num_cols = [c for c, _ in PUAN_BASLIKLARI if c in df_filtered.columns]
        for b_adi, g in df_filtered.groupby("birim"):
            ort_val = round(float(g[num_cols].stack().mean()), 2) if num_cols and not g[num_cols].stack().empty else 0.0
            birim_ozet.append(f"{b_adi}: {len(g)} görüşme | Ort. Puan: {ort_val}")
    add_card(s4, 0.6, 1.35, 12.1, 5.2, "Birim Bazlı Görüşme ve Ortalama Skor Özeti", birim_ozet or ["Kayıt bulunmuyor"], c_teal)

    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, 5, "Çalışan Sesi (Voice of Employee)", f"{secili_gmy}  |  Dönem: {secili_donem}")
    guclu_yorumlar = (
        df_filtered["q02_guclu_yonler"].dropna().astype(str).str.strip().loc[lambda s: s != ""].head(4).tolist()
        if "q02_guclu_yonler" in df_filtered.columns
        else []
    )
    degisim_yorumlar = (
        df_filtered["q10_degistirmek_istenen"].dropna().astype(str).str.strip().loc[lambda s: s != ""].head(4).tolist()
        if "q10_degistirmek_istenen" in df_filtered.columns
        else []
    )
    add_card(s5, 0.6, 1.35, 5.9, 5.2, "Güçlü Yönlere Dair İfadeler", guclu_yorumlar or ["Yorum bulunmuyor"], c_teal)
    add_card(s5, 6.8, 1.35, 5.9, 5.2, "Değişim / Gelişim İfadeleri", degisim_yorumlar or ["Yorum bulunmuyor"], c_purple)

    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, 6, "Öne Çıkan Temalar ve Duygu Dağılımı", f"{secili_gmy}  |  Dönem: {secili_donem}")
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
    add_header(s7, 7, "İK Önerileri ve Aksiyon Planı", f"{secili_gmy}  |  Dönem: {secili_donem}")
    aksiyonlar = (
        df_filtered[df_filtered["aksiyon_ihtiyaci"] == "Evet"]
        if "aksiyon_ihtiyaci" in df_filtered.columns
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
        <div class="softtech-header">
            <div class="softtech-brand"><span class="softtech-accent">&lt;</span> softtech<span class="softtech-accent">.</span> <span class="softtech-purple">|</span> İzlenim Paylaşım Paneli</div>
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

        st.caption("Test hesapları: `seyma` / `softtech2026` veya `admin` / `admin2026`")
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
    <div class="softtech-header">
        <div class="softtech-brand"><span class="softtech-accent">&lt;</span> softtech<span class="softtech-accent">.</span> <span class="softtech-purple">|</span> {sayfa}</div>
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
            "kayit_zamani": datetime.now().strftime("%Y-%m-%d %H:%M"),
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
    with c2:
        pass
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

    pptx_bytes = build_pptx_report(df_view, sel_gmy_f, sel_donem_f)
    st.download_button(
        label="📊 Standart Direktör Sunumunu İndir (.pptx)",
        data=pptx_bytes,
        file_name=f"Softtech_Izlenim_Paylasim_{sel_donem_f.replace(' ', '_')}.pptx",
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

else:
    st.subheader("Kayıt Havuzu")
    if df_all.empty:
        st.info("Henüz kayıt bulunmuyor.")
    else:
        st.dataframe(df_all, use_container_width=True, hide_index=True)
        csv_data = df_all.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="Havuzu CSV Olarak İndir",
            data=csv_data,
            file_name="softtech_izlenim_paylasim_havuz.csv",
            mime="text/csv",
        )
