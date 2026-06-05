
# ============================================================
# AYÇA Insight V5.0 Executive Demo
# Eczanenin Dijital Yönetim Paneli
# Revizyon: Demo giriş/kayıt ekranı eklendi; Kritik Merkez kartları aktif bölüm değiştirir; eski yatay bölüm menüsü korunur
# ------------------------------------------------------------
# Bu app.py, kullanıcının verdiği AYÇA Excel formatına göre yazılmıştır.
#
# Beklenen Excel sayfası:
# - AYCA_V2_Data
#
# Beklenen ana kolonlar:
# - Tarih
# - Fiş/Reçete No
# - Kaynak
# - Barkod
# - Ürün Adı
# - Kategori
# - Alt Kategori
# - Adet
# - Alış Birim TL
# - Satış Birim TL
# - Ciro TL
# - Maliyet TL
# - Brüt Kar TL
# - Brüt Kar %
# - SGK/Tahsilat Tipi
# - Mevcut Stok
# - Miad Tarihi
# - Tedarikçi
# - Raf Lokasyonu
# - Son 60 Gün Çıkış
# - Stok Ay Karşılığı
# - Sipariş Önerisi
#
# Çalıştırma:
# pip install streamlit pandas numpy openpyxl plotly
# streamlit run app.py
# ============================================================

from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
from html import escape
from typing import Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# STREAMLIT AYARI
# ============================================================
st.set_page_config(
    page_title="AYÇA Insight V5.0",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# SOFT TEMA CSS
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --bg: #F8FAFC;
        --panel: #FFFFFF;
        --panel2: #F9FBFF;
        --border: #E2E8F0;
        --text: #0F172A;
        --muted: #64748B;

        --blue: #2563EB;
        --blue-soft: #DBEAFE;

        --green: #10B981;
        --green-soft: #DCFCE7;

        --orange: #F59E0B;
        --orange-soft: #FEF3C7;

        --red: #EF4444;
        --red-soft: #FEE2E2;

        --purple: #8B5CF6;
        --purple-soft: #EDE9FE;

        --cyan: #06B6D4;
        --cyan-soft: #CFFAFE;
    }

    html, body, [data-testid="stAppViewContainer"] {
        background: var(--bg);
        color: var(--text);
    }

    [data-testid="stHeader"] {
        background: rgba(248, 250, 252, 0.88);
        backdrop-filter: blur(10px);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FFFFFF 0%, #F1F5F9 100%);
        border-right: 1px solid var(--border);
    }

    [data-testid="stSidebar"] * {
        color: var(--text);
    }

    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        max-width: 1520px;
    }

    .ayca-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        margin-bottom: 16px;
    }

    .ayca-title h1 {
        margin: 0;
        color: var(--text);
        font-size: 32px;
        letter-spacing: -0.7px;
        font-weight: 900;
    }

    .ayca-title p {
        margin: 6px 0 0 0;
        color: var(--muted);
        font-size: 14px;
    }

    .header-pill {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 12px 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
        color: var(--text);
        font-weight: 800;
        font-size: 13px;
        white-space: nowrap;
    }

    .metric-card {
        min-height: 142px;
        background: linear-gradient(180deg, #FFFFFF 0%, #F8FBFF 100%);
        border: 1px solid var(--border);
        border-radius: 20px;
        padding: 18px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
        overflow: hidden;
        position: relative;
    }

    .metric-card::after {
        content: "";
        position: absolute;
        right: -42px;
        bottom: -42px;
        width: 120px;
        height: 120px;
        border-radius: 999px;
        background: radial-gradient(circle, rgba(37, 99, 235, 0.14), rgba(37, 99, 235, 0));
    }

    .metric-label {
        color: var(--muted);
        font-size: 12px;
        font-weight: 900;
        letter-spacing: .4px;
        text-transform: uppercase;
        margin-bottom: 10px;
    }

    .metric-value {
        color: var(--text);
        font-size: 30px;
        font-weight: 950;
        margin-bottom: 8px;
        letter-spacing: -0.5px;
    }

    .metric-sub {
        color: var(--muted);
        font-size: 13px;
    }

    .metric-up {
        color: var(--green);
        font-size: 13px;
        font-weight: 900;
    }

    .metric-down {
        color: var(--red);
        font-size: 13px;
        font-weight: 900;
    }

    .mini-card {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 18px;
        padding: 16px;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05);
        min-height: 104px;
    }

    .mini-title {
        color: var(--text);
        font-size: 14px;
        font-weight: 900;
        margin-bottom: 8px;
    }

    .mini-value {
        color: var(--text);
        font-size: 24px;
        font-weight: 950;
        margin-bottom: 4px;
    }

    .mini-note {
        color: var(--muted);
        font-size: 13px;
    }

    .alert-red {
        background: linear-gradient(135deg, #FFFFFF 0%, var(--red-soft) 100%);
        border-color: #FECACA;
    }

    .alert-orange {
        background: linear-gradient(135deg, #FFFFFF 0%, var(--orange-soft) 100%);
        border-color: #FDE68A;
    }

    .alert-green {
        background: linear-gradient(135deg, #FFFFFF 0%, var(--green-soft) 100%);
        border-color: #BBF7D0;
    }

    .alert-blue {
        background: linear-gradient(135deg, #FFFFFF 0%, var(--blue-soft) 100%);
        border-color: #BFDBFE;
    }

    .alert-purple {
        background: linear-gradient(135deg, #FFFFFF 0%, var(--purple-soft) 100%);
        border-color: #DDD6FE;
    }

    .ai-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%);
        border: 1px solid #BFDBFE;
        border-radius: 22px;
        padding: 20px;
        box-shadow: 0 12px 30px rgba(37, 99, 235, 0.08);
        margin: 14px 0 18px 0;
    }

    .ai-title {
        color: var(--blue);
        font-size: 18px;
        font-weight: 950;
        margin-bottom: 8px;
    }

    .ai-text {
        color: var(--text);
        font-size: 14px;
        line-height: 1.55;
    }

    .section-title {
        color: var(--text);
        font-size: 21px;
        font-weight: 950;
        margin: 20px 0 12px 0;
        letter-spacing: -0.3px;
    }

    .soft-panel {
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 20px;
        padding: 16px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.05);
    }

    .chip {
        display: inline-block;
        border-radius: 999px;
        padding: 5px 10px;
        font-size: 12px;
        font-weight: 900;
    }

    .chip-red {
        background: var(--red-soft);
        color: #B91C1C;
    }

    .chip-orange {
        background: var(--orange-soft);
        color: #B45309;
    }

    .chip-green {
        background: var(--green-soft);
        color: #047857;
    }

    .chip-blue {
        background: var(--blue-soft);
        color: #1D4ED8;
    }

    .chip-purple {
        background: var(--purple-soft);
        color: #6D28D9;
    }

    .chip-gray {
        background: #E2E8F0;
        color: #334155;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        background: #FFFFFF;
        border: 1px solid var(--border);
        border-radius: 13px;
        padding: 10px 16px;
        color: var(--text);
        font-weight: 800;
    }

    .stTabs [aria-selected="true"] {
        background: var(--blue-soft);
        color: var(--blue);
        border-color: #BFDBFE;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid var(--border);
    }

    .stButton > button,
    .stDownloadButton > button {
        border-radius: 13px;
        border: 1px solid #BFDBFE;
        background: #FFFFFF;
        color: var(--blue);
        font-weight: 900;
    }

     .stButton > button:hover,
    .stDownloadButton > button:hover {
        border-color: var(--blue);
        background: var(--blue-soft);
        color: var(--blue);
    }

    /* Kritik Merkez kartları artık gerçek Streamlit butonudur. */
    div[data-testid="stHorizontalBlock"] .stButton > button {
        min-height: 132px;
        border-radius: 20px;
        padding: 18px;
        text-align: left;
        white-space: pre-line;
        font-size: 15px;
        line-height: 1.45;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05);
    }



    .click-card-link {
        display: block;
        text-decoration: none !important;
        color: inherit !important;
    }

    .click-card-link .mini-card {
        cursor: pointer;
        transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease;
    }

    .click-card-link:hover .mini-card {
        transform: translateY(-4px);
        box-shadow: 0 16px 34px rgba(15, 23, 42, 0.12);
        border-color: #2563EB;
    }

    .click-card-hint {
        margin-top: 8px;
        color: #2563EB;
        font-size: 12px;
        font-weight: 900;
    }

    .small-muted {
        color: var(--muted);
        font-size: 13px;
    }

    /* Yatay bölüm menüsü eski sekme hissini korusun. */
    div[role="radiogroup"] {
        gap: 10px;
    }

    div[role="radiogroup"] label {
        background: #FFFFFF;
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 10px 14px;
        box-shadow: 0 4px 12px rgba(15,23,42,0.04);
    }

    div[role="radiogroup"] label:has(input:checked) {
        background: var(--blue-soft);
        border-color: #BFDBFE;
        color: var(--blue);
    }


    /* V5 Executive Dashboard */
    .exec-grid { display: grid; grid-template-columns: 1.25fr .75fr; gap: 16px; margin: 14px 0 18px 0; }
    .exec-card { background: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%); border: 1px solid #BFDBFE; border-radius: 24px; padding: 20px; box-shadow: 0 14px 34px rgba(37, 99, 235, 0.08); }
    .exec-title { color: #0F172A; font-size: 22px; font-weight: 950; letter-spacing: -0.4px; margin-bottom: 8px; }
    .exec-sub { color: #64748B; font-size: 14px; line-height: 1.55; margin-bottom: 14px; }
    .exec-list-item { background: rgba(255,255,255,.78); border: 1px solid rgba(226,232,240,.95); border-radius: 16px; padding: 12px 13px; margin: 9px 0; font-size: 14px; color: #0F172A; line-height: 1.45; font-weight: 700; }
    .score-big { font-size: 54px; font-weight: 950; letter-spacing: -2px; color: #2563EB; line-height: 1; margin: 4px 0 8px 0; }
    .score-label { color: #64748B; font-size: 13px; font-weight: 850; text-transform: uppercase; letter-spacing: .35px; }
    .health-row { margin: 12px 0; }
    .health-head { display: flex; justify-content: space-between; color: #334155; font-weight: 900; font-size: 13px; margin-bottom: 6px; }
    .health-bar-bg { width: 100%; height: 10px; background: #E2E8F0; border-radius: 999px; overflow: hidden; }
    .health-bar-fill { height: 10px; border-radius: 999px; background: linear-gradient(90deg, #2563EB, #10B981); }
    .radar-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin: 12px 0 18px 0; }
    .radar-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 18px; padding: 15px; box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05); min-height: 110px; }
    .radar-title { color: #64748B; font-size: 12px; font-weight: 900; text-transform: uppercase; letter-spacing: .35px; margin-bottom: 9px; }
    .radar-value { color: #0F172A; font-size: 24px; font-weight: 950; margin-bottom: 5px; }
    .radar-note { color: #64748B; font-size: 13px; line-height: 1.35; }
    .task-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 14px 0 18px 0; }
    .task-card, .lost-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 22px; padding: 18px; box-shadow: 0 12px 30px rgba(15, 23, 42, 0.05); }
    .task-item { border-bottom: 1px solid #E2E8F0; padding: 10px 0; color: #0F172A; font-size: 14px; font-weight: 750; line-height: 1.45; }
    .task-item:last-child { border-bottom: 0; }
    .lost-number { font-size: 34px; font-weight: 950; color: #B91C1C; letter-spacing: -1px; margin: 4px 0; }
    @media (max-width: 1000px) { .exec-grid, .task-grid { grid-template-columns: 1fr; } .radar-grid { grid-template-columns: 1fr 1fr; } }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DEMO GİRİŞ / KAYIT EKRANI
# ============================================================
DEMO_USERS = {
    "basic": {
        "password": "basic2026",
        "name": "Basic Demo Kullanıcı",
        "pharmacy": "İdil Eczanesi",
        "membership": "Basic",
    },
    "premium": {
        "password": "premium2026",
        "name": "Premium Demo Kullanıcı",
        "pharmacy": "İdil Eczanesi",
        "membership": "Premium",
    },
}


def safe_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def get_membership() -> str:
    return st.session_state.get("membership", "Basic")


def is_premium_user() -> bool:
    return get_membership().lower() == "premium"


def show_basic_info(message: str = "Basic demo kullanıcısı bu bölümde yalnızca kısa önizleme görür. Detaylı tablo, grafik ve raporlar Premium üyelikte açılır."):
    st.markdown(
        f''' 
        <div class="mini-card alert-orange" style="margin: 10px 0 16px 0; min-height: auto;">
            <div class="mini-title">🔒 Basic Üyelik Önizlemesi</div>
            <div class="mini-note">{message}</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def show_demo_auth_screen():
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {display: none;}
        .login-shell {
            max-width: 980px;
            margin: 32px auto 0 auto;
            display: grid;
            grid-template-columns: 1.1fr .9fr;
            gap: 22px;
            align-items: stretch;
        }
        .login-hero {
            background: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 55%, #DCFCE7 100%);
            border: 1px solid #BFDBFE;
            border-radius: 28px;
            padding: 34px;
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.08);
            min-height: 430px;
        }
        .login-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 28px;
            padding: 28px;
            box-shadow: 0 18px 45px rgba(15, 23, 42, 0.08);
        }
        .login-logo {
            width: 64px;
            height: 64px;
            border-radius: 20px;
            background: #DBEAFE;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 34px;
            margin-bottom: 18px;
        }
        .login-title {
            font-size: 38px;
            line-height: 1.08;
            font-weight: 950;
            letter-spacing: -1px;
            color: #0F172A;
            margin-bottom: 12px;
        }
        .login-sub {
            color: #64748B;
            font-size: 15px;
            line-height: 1.6;
            margin-bottom: 24px;
        }
        .feature-row {
            display: flex;
            gap: 10px;
            align-items: center;
            background: rgba(255,255,255,.75);
            border: 1px solid rgba(226,232,240,.8);
            border-radius: 16px;
            padding: 11px 13px;
            margin: 10px 0;
            color: #0F172A;
            font-weight: 750;
            font-size: 14px;
        }
        .demo-badge {
            display: inline-block;
            background: #EDE9FE;
            color: #6D28D9;
            border-radius: 999px;
            padding: 7px 12px;
            font-size: 12px;
            font-weight: 900;
            margin-bottom: 14px;
        }
        .credential-box {
            background: #F8FAFC;
            border: 1px dashed #CBD5E1;
            border-radius: 16px;
            padding: 12px 14px;
            color: #334155;
            font-size: 13px;
            margin-top: 12px;
        }
        @media (max-width: 900px) {
            .login-shell {grid-template-columns: 1fr;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.12, .88])
    with left:
        st.markdown(
            """
            <div class="login-hero">
                <div class="login-logo">💊</div>
                <div class="demo-badge">AYÇA Insight Demo</div>
                <div class="login-title">Eczanenizin dijital yönetim paneli</div>
                <div class="login-sub">
                    Stok, miad, kârlılık, sipariş ve ölü stok risklerini tek ekranda yorumlayan demo panel.
                    Bu ekran sunum ve ürün tanıtımı için hazırlanmıştır.
                </div>
                <div class="feature-row">📦 Kritik stok ve bitiş tahmini</div>
                <div class="feature-row">⏳ Miad yaklaşan ürün uyarıları</div>
                <div class="feature-row">🛒 Sipariş önerisi ve maliyet analizi</div>
                <div class="feature-row">💰 Kârlılık ve kategori performansı</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        st.markdown("### Giriş Yap")
        mode = st.radio("İşlem", ["Giriş", "Demo Kayıt"], horizontal=True, label_visibility="collapsed")

        if mode == "Giriş":
            username = st.text_input("Kullanıcı adı", value="premium")
            password = st.text_input("Şifre", value="", type="password")
            login_clicked = st.button("🚀 Dashboard'a Giriş Yap", use_container_width=True)

            st.markdown(
                """
                <div class="credential-box">
                    <b>Premium demo</b><br>
                    Kullanıcı: <b>premium</b> · Şifre: <b>premium2026</b><br><br>
                    <b>Basic demo</b><br>
                    Kullanıcı: <b>basic</b> · Şifre: <b>basic2026</b>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if login_clicked:
                user_key = username.strip().lower()
                user_record = DEMO_USERS.get(user_key)
                if user_record and password == user_record["password"]:
                    st.session_state["authenticated"] = True
                    st.session_state["auth_user"] = user_record["name"]
                    st.session_state["auth_pharmacy"] = user_record["pharmacy"]
                    st.session_state["membership"] = user_record["membership"]
                    safe_rerun()
                else:
                    st.error("Kullanıcı adı veya şifre hatalı. Premium: premium / premium2026 · Basic: basic / basic2026")

        else:
            pharmacy = st.text_input("Eczane adı")
            name = st.text_input("Yetkili adı")
            phone = st.text_input("Telefon / e-posta")
            st.text_input("Şifre", type="password")
            st.text_input("Şifre tekrar", type="password")
            if st.button("📝 Demo Talebi Oluştur", use_container_width=True):
                st.success("Demo talebi oluşturuldu. Bu demo sürümde gerçek kayıt yapılmaz.")
                if pharmacy or name or phone:
                    st.info("Sonraki gerçek sistemde bu alanlar kullanıcı veritabanına bağlanacak.")

        st.markdown("</div>", unsafe_allow_html=True)


if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    show_demo_auth_screen()
    st.stop()


# ============================================================
# BASIC / PREMIUM DEMO ERİŞİM KONTROLÜ
# ============================================================
_original_dataframe = st.dataframe
_original_plotly_chart = st.plotly_chart
_original_download_button = st.download_button


def premium_locked_chart(*args, **kwargs):
    if is_premium_user():
        return _original_plotly_chart(*args, **kwargs)
    show_basic_info("Basic üyelikte grafikler kapalıdır. Bu bölümde sadece kısa özet ve en fazla 2 satır önizleme gösterilir.")
    return None


def limited_dataframe(data=None, *args, **kwargs):
    if is_premium_user():
        return _original_dataframe(data, *args, **kwargs)

    try:
        preview = data.head(2).copy() if hasattr(data, "head") else data
        show_basic_info("Basic üyelikte tablo önizlemesi en fazla 2 satırdır. Tüm ürün listesi ve detaylı analiz Premium üyelikte görünür.")
        return _original_dataframe(preview, *args, **kwargs)
    except Exception:
        show_basic_info()
        return _original_dataframe(data, *args, **kwargs)


def premium_download_button(*args, **kwargs):
    if is_premium_user():
        return _original_download_button(*args, **kwargs)
    show_basic_info("Basic üyelikte Excel raporu indirme kapalıdır. Rapor indirme Premium üyelikte açılır.")
    return False


st.plotly_chart = premium_locked_chart
st.dataframe = limited_dataframe
st.download_button = premium_download_button


# ============================================================
# GENEL YARDIMCI FONKSİYONLAR
# ============================================================
def normalize_col_name(name: str) -> str:
    name = str(name).strip().lower()
    tr_map = str.maketrans("çğıöşüİı", "cgiosuii")
    name = name.translate(tr_map)
    name = re.sub(r"[^a-z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name


def money_fmt(x) -> str:
    try:
        if pd.isna(x) or np.isinf(float(x)):
            return "₺0"
        return f"₺{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "₺0"


def num_fmt(x, digits: int = 1) -> str:
    try:
        if pd.isna(x) or np.isinf(float(x)):
            return "0"
        return f"{float(x):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"


def pct_fmt(x) -> str:
    try:
        if pd.isna(x) or np.isinf(float(x)):
            return "%0,0"
        return "%" + num_fmt(float(x) * 100, 1)
    except Exception:
        return "%0,0"


def rate_fmt(current, previous):
    try:
        current = float(current)
        previous = float(previous)
        if previous == 0 and current > 0:
            return "▲ yeni", "metric-up"
        if previous == 0:
            return "▲ %0,0", "metric-up"
        rate = (current - previous) / previous
        if rate >= 0:
            return "▲ " + pct_fmt(rate), "metric-up"
        return "▼ " + pct_fmt(abs(rate)), "metric-down"
    except Exception:
        return "▲ %0,0", "metric-up"


def find_col(columns, candidates) -> Optional[str]:
    normalized = {c: normalize_col_name(c) for c in columns}
    for cand in candidates:
        cand_norm = normalize_col_name(cand)
        for original, norm in normalized.items():
            if cand_norm == norm:
                return original
    for cand in candidates:
        cand_norm = normalize_col_name(cand)
        for original, norm in normalized.items():
            if cand_norm in norm:
                return original
    return None


def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
    """
    Excel seri tarihleri ve normal tarihleri birlikte dönüştürür.
    Örn: 46082 -> 2026-03-01 gibi.
    """
    s = series.copy()

    if pd.api.types.is_numeric_dtype(s):
        return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")

    dt = pd.to_datetime(s, errors="coerce", dayfirst=True)

    mask = dt.isna()
    if mask.any():
        numeric = pd.to_numeric(s[mask], errors="coerce")
        converted = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
        dt.loc[mask] = converted

    return dt


def safe_numeric(series_or_value, default=0):
    try:
        return pd.to_numeric(series_or_value, errors="coerce").fillna(default)
    except Exception:
        return default


def read_excel_smart(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)

    preferred_sheet = None
    for sheet in xls.sheet_names:
        if normalize_col_name(sheet) == "ayca_v2_data":
            preferred_sheet = sheet
            break

    if preferred_sheet is None:
        for sheet in xls.sheet_names:
            norm = normalize_col_name(sheet)
            if any(key in norm for key in ["data", "veri", "satis", "stok"]):
                preferred_sheet = sheet
                break

    if preferred_sheet is None:
        preferred_sheet = xls.sheet_names[0]

    df = pd.read_excel(uploaded_file, sheet_name=preferred_sheet)
    return df, preferred_sheet, xls.sheet_names


def make_metric_card(label, value, sub="", trend_text_value=None, trend_class="metric-up"):
    trend_html = ""
    if trend_text_value:
        trend_html = f"<span class='{trend_class}'>{trend_text_value}</span>"
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub} {trend_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_mini_card(title, value, note, css_class=""):
    st.markdown(
        f"""
        <div class="mini-card {css_class}">
            <div class="mini-title">{title}</div>
            <div class="mini-value">{value}</div>
            <div class="mini-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_clickable_mini_card(title, value, note, css_class="", page_name=None, hint="Detayı göster", key=None):
    """
    Kritik Merkez kartını gerçek Streamlit butonu gibi çalıştırır.
    URL değiştirmez, başka sekmeye atmaz; sadece session_state içindeki aktif sayfayı değiştirir.
    """
    label = f"{title}\n\n{value}\n{note}"
    clicked = st.button(label, key=key or f"card_{normalize_col_name(title)}", use_container_width=True)
    if clicked and page_name is not None:
        st.session_state["aktif_sayfa"] = page_name
        st.rerun()


def sparkline(values, color="#2563EB"):
    values = pd.Series(values).fillna(0).astype(float).tolist()
    if len(values) < 2:
        values = [0, values[0] if values else 0]
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            y=values,
            mode="lines",
            line=dict(color=color, width=3),
            fill="tozeroy",
            fillcolor="rgba(37, 99, 235, 0.10)",
            hoverinfo="skip",
        )
    )
    fig.update_layout(
        height=58,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
    )
    return fig


def apply_plot_theme(fig, height=360):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#0F172A", family="Arial"),
        margin=dict(l=10, r=10, t=48, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=False, color="#64748B")
    fig.update_yaxes(gridcolor="#E2E8F0", color="#64748B")
    return fig


# ============================================================
# VERİ STANDARDİZASYONU
# ============================================================
def standardize_ayca_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw_df.columns)

    mapping = {
        "tarih": find_col(cols, ["Tarih"]),
        "fis": find_col(cols, ["Fiş/Reçete No", "Fis Recete No", "Fiş No", "Reçete No"]),
        "kaynak": find_col(cols, ["Kaynak"]),
        "barkod": find_col(cols, ["Barkod"]),
        "urun": find_col(cols, ["Ürün Adı", "Urun Adi", "İlaç Adı", "Malzeme Adı"]),
        "kategori": find_col(cols, ["Kategori"]),
        "alt_kategori": find_col(cols, ["Alt Kategori"]),
        "adet": find_col(cols, ["Adet", "Miktar", "Satış Adet"]),
        "alis_birim": find_col(cols, ["Alış Birim TL", "Alis Birim TL", "Alış Fiyatı", "Maliyet Birim"]),
        "satis_birim": find_col(cols, ["Satış Birim TL", "Satis Birim TL", "Satış Fiyatı"]),
        "ciro": find_col(cols, ["Ciro TL", "Ciro", "Satış Tutarı"]),
        "maliyet": find_col(cols, ["Maliyet TL", "Maliyet"]),
        "brut_kar": find_col(cols, ["Brüt Kar TL", "Brut Kar TL", "Kar TL"]),
        "brut_kar_oran": find_col(cols, ["Brüt Kar %", "Brut Kar %", "Kar %"]),
        "tahsilat": find_col(cols, ["SGK/Tahsilat Tipi", "Tahsilat Tipi", "Ödeme Tipi"]),
        "stok": find_col(cols, ["Mevcut Stok", "Stok", "Kalan Stok"]),
        "miad": find_col(cols, ["Miad Tarihi", "Miat Tarihi", "SKT", "Son Kullanma Tarihi"]),
        "tedarikci": find_col(cols, ["Tedarikçi", "Tedarikci"]),
        "raf": find_col(cols, ["Raf Lokasyonu", "Raf", "Lokasyon"]),
        "son_60": find_col(cols, ["Son 60 Gün Çıkış", "Son 60 Gun Cikis"]),
        "stok_ay": find_col(cols, ["Stok Ay Karşılığı", "Stok Ay Karsiligi"]),
        "siparis": find_col(cols, ["Sipariş Önerisi", "Siparis Onerisi"]),
    }

    required = ["tarih", "urun", "kategori", "adet", "ciro", "brut_kar", "stok", "miad", "son_60"]
    missing = [k for k in required if mapping.get(k) is None]
    if missing:
        readable = ", ".join(missing)
        raise ValueError(f"Eksik zorunlu kolonlar: {readable}. Lütfen Excel formatını kontrol edin.")

    df = pd.DataFrame()
    df["tarih"] = excel_serial_to_datetime(raw_df[mapping["tarih"]])
    df["fis"] = raw_df[mapping["fis"]].astype(str) if mapping["fis"] else ""
    df["kaynak"] = raw_df[mapping["kaynak"]].astype(str) if mapping["kaynak"] else "Bilinmiyor"
    df["barkod"] = raw_df[mapping["barkod"]].astype(str) if mapping["barkod"] else ""
    df["urun"] = raw_df[mapping["urun"]].astype(str)
    df["kategori"] = raw_df[mapping["kategori"]].astype(str)
    df["alt_kategori"] = raw_df[mapping["alt_kategori"]].astype(str) if mapping["alt_kategori"] else "Genel"

    df["adet"] = pd.to_numeric(raw_df[mapping["adet"]], errors="coerce").fillna(0)
    df["alis_birim"] = pd.to_numeric(raw_df[mapping["alis_birim"]], errors="coerce").fillna(0) if mapping["alis_birim"] else 0
    df["satis_birim"] = pd.to_numeric(raw_df[mapping["satis_birim"]], errors="coerce").fillna(0) if mapping["satis_birim"] else 0
    df["ciro"] = pd.to_numeric(raw_df[mapping["ciro"]], errors="coerce").fillna(0)
    df["maliyet"] = pd.to_numeric(raw_df[mapping["maliyet"]], errors="coerce").fillna(0) if mapping["maliyet"] else df["adet"] * df["alis_birim"]
    df["brut_kar"] = pd.to_numeric(raw_df[mapping["brut_kar"]], errors="coerce").fillna(0)
    df["brut_kar_oran"] = pd.to_numeric(raw_df[mapping["brut_kar_oran"]], errors="coerce").fillna(np.nan) if mapping["brut_kar_oran"] else np.nan

    df["tahsilat"] = raw_df[mapping["tahsilat"]].astype(str) if mapping["tahsilat"] else "Bilinmiyor"
    df["stok"] = pd.to_numeric(raw_df[mapping["stok"]], errors="coerce").fillna(0)
    df["miad"] = excel_serial_to_datetime(raw_df[mapping["miad"]])
    df["tedarikci"] = raw_df[mapping["tedarikci"]].astype(str) if mapping["tedarikci"] else "Bilinmiyor"
    df["raf"] = raw_df[mapping["raf"]].astype(str) if mapping["raf"] else "Bilinmiyor"
    df["son_60"] = pd.to_numeric(raw_df[mapping["son_60"]], errors="coerce").fillna(0)
    df["stok_ay"] = pd.to_numeric(raw_df[mapping["stok_ay"]], errors="coerce").fillna(np.nan) if mapping["stok_ay"] else np.nan
    df["siparis_onerisi"] = raw_df[mapping["siparis"]].astype(str) if mapping["siparis"] else ""

    df = df.dropna(subset=["tarih"])
    return df


def build_product_table(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """
    Satış satırlarını ürün bazında tekilleştirir.
    Aynı ürün birden fazla satırda olduğu için stok/miad gibi alanlarda son kayıt kullanılır.
    """
    df_sorted = df.sort_values("tarih")

    agg = df_sorted.groupby(["barkod", "urun"], dropna=False).agg(
        kategori=("kategori", lambda x: x.dropna().iloc[-1] if len(x.dropna()) else "Bilinmiyor"),
        alt_kategori=("alt_kategori", lambda x: x.dropna().iloc[-1] if len(x.dropna()) else "Genel"),
        toplam_adet=("adet", "sum"),
        toplam_ciro=("ciro", "sum"),
        toplam_maliyet=("maliyet", "sum"),
        toplam_brut_kar=("brut_kar", "sum"),
        ort_satis_birim=("satis_birim", "mean"),
        ort_alis_birim=("alis_birim", "mean"),
        mevcut_stok=("stok", "last"),
        miad=("miad", "last"),
        tedarikci=("tedarikci", "last"),
        raf=("raf", "last"),
        son_60=("son_60", "last"),
        stok_ay=("stok_ay", "last"),
        siparis_onerisi=("siparis_onerisi", "last"),
        son_satis_tarihi=("tarih", "max"),
        islem_sayisi=("fis", pd.Series.nunique),
    ).reset_index()

    agg["kar_marji"] = np.where(agg["toplam_ciro"] > 0, agg["toplam_brut_kar"] / agg["toplam_ciro"], 0)
    agg["gunluk_tuketim_60"] = agg["son_60"] / 60
    agg["tahmini_bitis_gunu"] = np.where(
        agg["gunluk_tuketim_60"] > 0,
        agg["mevcut_stok"] / agg["gunluk_tuketim_60"],
        np.inf,
    )
    agg["stok_degeri"] = agg["mevcut_stok"] * agg["ort_alis_birim"]
    agg["miad_kalan_gun"] = (agg["miad"] - today).dt.days
    agg["son_satis_kac_gun"] = (today - agg["son_satis_tarihi"]).dt.days
    agg["stok_ay_hesap"] = np.where(agg["son_60"] > 0, agg["mevcut_stok"] / (agg["son_60"] / 2), np.inf)

    return agg


def classify_inventory(product_df: pd.DataFrame, critical_days, warning_days, miad_days, dead_days) -> pd.DataFrame:
    p = product_df.copy()

    p["stok_durumu"] = np.select(
        [
            (p["mevcut_stok"] > 0) & (p["son_60"] <= 0),
            (p["gunluk_tuketim_60"] > 0) & (p["tahmini_bitis_gunu"] <= critical_days),
            (p["gunluk_tuketim_60"] > 0) & (p["tahmini_bitis_gunu"] <= warning_days),
            (p["gunluk_tuketim_60"] > 0) & (p["tahmini_bitis_gunu"] > warning_days),
        ],
        ["Ölü Stok", "Kritik", "Dikkat", "Güvenli"],
        default="Veri Yok",
    )

    p["miad_durumu"] = np.select(
        [
            p["miad"].isna(),
            p["miad_kalan_gun"] < 0,
            p["miad_kalan_gun"] <= 30,
            p["miad_kalan_gun"] <= miad_days,
        ],
        ["Miad Yok", "Miad Geçmiş", "Çok Yakın", "Yaklaşıyor"],
        default="Güvenli",
    )

    p["olu_stok_mu"] = (p["mevcut_stok"] > 0) & (
        (p["son_60"] <= 0) | (p["son_satis_kac_gun"] >= dead_days)
    )

    p["onerilen_siparis_adedi"] = np.where(
        p["gunluk_tuketim_60"] > 0,
        np.ceil((warning_days * 2 * p["gunluk_tuketim_60"]) - p["mevcut_stok"]),
        0,
    )
    p["onerilen_siparis_adedi"] = p["onerilen_siparis_adedi"].clip(lower=0)
    p["onerilen_siparis_maliyeti"] = p["onerilen_siparis_adedi"] * p["ort_alis_birim"]

    return p


def ayca_score(p: pd.DataFrame, critical_days, miad_days) -> int:
    if p.empty:
        return 0

    total = len(p)
    critical_ratio = len(p[(p["gunluk_tuketim_60"] > 0) & (p["tahmini_bitis_gunu"] <= critical_days)]) / total
    miad_ratio = len(p[(p["miad_kalan_gun"].notna()) & (p["miad_kalan_gun"] <= miad_days)]) / total
    dead_ratio = len(p[p["olu_stok_mu"]]) / total
    low_margin_ratio = len(p[(p["toplam_ciro"] > 0) & (p["kar_marji"] < 0.12)]) / total

    score = 100
    score -= critical_ratio * 28
    score -= miad_ratio * 24
    score -= dead_ratio * 24
    score -= low_margin_ratio * 14

    return int(max(0, min(100, round(score))))


def create_ai_comment(
    df: pd.DataFrame,
    p: pd.DataFrame,
    critical_df: pd.DataFrame,
    miad_df: pd.DataFrame,
    dead_df: pd.DataFrame,
    margin_df: pd.DataFrame,
    current_ciro,
    previous_ciro,
):
    messages = []

    if previous_ciro > 0:
        rate = (current_ciro - previous_ciro) / previous_ciro
        if rate >= 0:
            messages.append(f"Güncel dönem ciro performansı önceki döneme göre {pct_fmt(rate)} artış gösteriyor.")
        else:
            messages.append(f"Güncel dönem ciro performansı önceki döneme göre {pct_fmt(abs(rate))} düşüş gösteriyor.")

    if not critical_df.empty:
        first = critical_df.sort_values("tahmini_bitis_gunu").iloc[0]
        messages.append(
            f"{first['urun']} mevcut tüketim hızına göre yaklaşık {num_fmt(first['tahmini_bitis_gunu'], 0)} gün içinde kritik seviyeye düşebilir."
        )

    if not miad_df.empty:
        first = miad_df.sort_values("miad_kalan_gun").iloc[0]
        messages.append(
            f"{first['urun']} için miada {num_fmt(first['miad_kalan_gun'], 0)} gün kaldı; raf önceliği verilmelidir."
        )

    if not dead_df.empty:
        messages.append(
            f"Toplam {money_fmt(dead_df['stok_degeri'].sum())} değerinde ölü/hareketsiz stok tespit edildi."
        )

    if not margin_df.empty:
        first = margin_df.sort_values("kar_marji").iloc[0]
        messages.append(
            f"{first['urun']} ürününde kâr marjı düşük görünüyor; fiyat ve maliyet kontrolü önerilir."
        )

    if not p.empty:
        top_cat = p.groupby("kategori")["toplam_ciro"].sum().sort_values(ascending=False)
        if not top_cat.empty:
            messages.append(f"En yüksek ciro katkısı {top_cat.index[0]} kategorisinden geliyor.")

    if not messages:
        messages.append("Genel tablo dengeli görünüyor. Kritik stok, miad ve ölü stok baskısı düşük.")

    return " ".join(messages)


def create_excel_report(df, p, critical_df, miad_df, dead_df, order_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ham_Satis_Verisi", index=False)
        p.replace([np.inf, -np.inf], np.nan).to_excel(writer, sheet_name="Urun_Analizi", index=False)
        critical_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, sheet_name="Kritik_Stok", index=False)
        miad_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, sheet_name="Miad_Takibi", index=False)
        dead_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, sheet_name="Olu_Stok", index=False)
        order_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, sheet_name="Siparis_Onerisi", index=False)

    return output.getvalue()




# ============================================================
# V5 EXECUTIVE DASHBOARD YARDIMCI FONKSİYONLAR
# ============================================================
def clamp_score(value) -> int:
    try:
        return int(max(0, min(100, round(float(value)))))
    except Exception:
        return 0


def score_status(score_value: int) -> str:
    if score_value >= 80:
        return "Güçlü"
    if score_value >= 60:
        return "Takip edilmeli"
    return "Riskli"


def calculate_health_scores(product_df, period_df, previous_period_df, current_margin, previous_margin, critical_days, miad_warning_days):
    total_products = max(1, len(product_df))
    critical_ratio = len(product_df[(product_df["gunluk_tuketim_60"] > 0) & (product_df["tahmini_bitis_gunu"] <= critical_days)]) / total_products
    miad_ratio = len(product_df[(product_df["miad_kalan_gun"].notna()) & (product_df["miad_kalan_gun"] <= miad_warning_days)]) / total_products
    dead_ratio = len(product_df[product_df["olu_stok_mu"]]) / total_products
    low_margin_ratio = len(product_df[(product_df["toplam_ciro"] > 0) & (product_df["kar_marji"] < 0.12)]) / total_products
    current_ciro_local = period_df["ciro"].sum() if not period_df.empty else 0
    previous_ciro_local = previous_period_df["ciro"].sum() if not previous_period_df.empty else 0
    if previous_ciro_local > 0:
        sales_trend = (current_ciro_local - previous_ciro_local) / previous_ciro_local
        sales_score = 70 + (sales_trend * 45)
    else:
        sales_score = 72 if current_ciro_local > 0 else 50
    margin_delta = current_margin - previous_margin if previous_margin > 0 else current_margin
    profitability_score = 70 + (current_margin * 80) + (margin_delta * 40) - (low_margin_ratio * 30)
    return {
        "Karlılık": clamp_score(profitability_score),
        "Stok Yönetimi": clamp_score(100 - (critical_ratio * 75) - (dead_ratio * 30)),
        "Miad Yönetimi": clamp_score(100 - (miad_ratio * 90)),
        "Satış Performansı": clamp_score(sales_score),
        "Nakit Verimliliği": clamp_score(100 - (dead_ratio * 70) - (critical_ratio * 20)),
    }


def best_growth_category(period_df, prev_df):
    if period_df.empty:
        return None, 0
    current = period_df.groupby("kategori")["ciro"].sum().sort_values(ascending=False)
    if current.empty:
        return None, 0
    prev = prev_df.groupby("kategori")["ciro"].sum() if not prev_df.empty else pd.Series(dtype=float)
    best_cat = current.index[0]
    best_rate = -999
    for cat, val in current.items():
        prev_val = float(prev.get(cat, 0)) if len(prev) else 0
        rate = (float(val) - prev_val) / prev_val if prev_val > 0 else 0.0
        if rate > best_rate:
            best_cat = cat
            best_rate = rate
    return best_cat, best_rate


def missed_profit_analysis(product_df, warning_days):
    p = product_df.copy()
    risk = p[(p["gunluk_tuketim_60"] > 0) & (p["tahmini_bitis_gunu"].replace(np.inf, np.nan).notna())].copy()
    risk = risk[risk["tahmini_bitis_gunu"] <= warning_days]
    if risk.empty:
        return 0, 0, 0
    risk["riskli_gun"] = (warning_days - risk["tahmini_bitis_gunu"]).clip(lower=0)
    risk["tahmini_kacirilan_ciro"] = risk["riskli_gun"] * risk["gunluk_tuketim_60"] * risk["ort_satis_birim"].fillna(0)
    risk["tahmini_kacirilan_kar"] = risk["tahmini_kacirilan_ciro"] * risk["kar_marji"].clip(lower=0, upper=0.75).fillna(0)
    return int(len(risk)), float(risk["tahmini_kacirilan_ciro"].sum()), float(risk["tahmini_kacirilan_kar"].sum())


def build_today_tasks(critical_df, miad_df, dead_df, low_margin_df, order_df, period_df, prev_df):
    tasks = []
    if not critical_df.empty:
        row = critical_df.sort_values("tahmini_bitis_gunu").iloc[0]
        tasks.append(f"☐ {escape(str(row['urun']))} için stok kontrolü yap; yaklaşık {num_fmt(row['tahmini_bitis_gunu'], 0)} gün içinde bitebilir.")
    if not miad_df.empty:
        row = miad_df.sort_values("miad_kalan_gun").iloc[0]
        tasks.append(f"☐ {escape(str(row['urun']))} ürününü ön rafa al; miada {num_fmt(row['miad_kalan_gun'], 0)} gün kaldı.")
    if not dead_df.empty:
        row = dead_df.sort_values("stok_degeri", ascending=False).iloc[0]
        tasks.append(f"☐ {escape(str(row['urun']))} için ölü stok aksiyonu planla; bağlı para {money_fmt(row['stok_degeri'])}.")
    if not order_df.empty:
        tasks.append(f"☐ Sipariş asistanındaki {len(order_df)} ürün için tedarikçi kontrolü yap.")
    if not low_margin_df.empty:
        row = low_margin_df.sort_values("kar_marji").iloc[0]
        tasks.append(f"☐ {escape(str(row['urun']))} ürününde fiyat/maliyet kontrolü yap; marj düşük görünüyor.")
    top_cat, rate = best_growth_category(period_df, prev_df)
    if top_cat:
        if rate > 0:
            tasks.append(f"☐ {escape(str(top_cat))} kategorisini takip et; önceki döneme göre {pct_fmt(rate)} büyüme sinyali var.")
        else:
            tasks.append(f"☐ {escape(str(top_cat))} kategorisini takip et; güncel dönemin en yüksek ciro katkısı burada.")
    fallback = [
        "☐ Kritik stok listesini kontrol et.",
        "☐ Miadı yaklaşan ürünleri raf önceliğine al.",
        "☐ Ölü stok değerini azaltacak aksiyon belirle.",
        "☐ Düşük marjlı ürünlerde fiyat/maliyet kontrolü yap.",
        "☐ Gün sonunda satış ve stok dengesini tekrar incele.",
    ]
    for item in fallback:
        if len(tasks) >= 5:
            break
        tasks.append(item)
    return tasks[:5]


def render_executive_dashboard(kullanici_adi, score, product_df, period_df, prev_df, critical_df, miad_df, dead_df, low_margin_df, order_df, current_margin, previous_margin, critical_days, miad_warning_days, warning_days):
    health = calculate_health_scores(product_df, period_df, prev_df, current_margin, previous_margin, critical_days, miad_warning_days)
    lost_count, lost_revenue, lost_profit = missed_profit_analysis(product_df, warning_days)
    tasks = build_today_tasks(critical_df, miad_df, dead_df, low_margin_df, order_df, period_df, prev_df)
    top_cat, cat_rate = best_growth_category(period_df, prev_df)

    if not critical_df.empty:
        first_critical = critical_df.sort_values("tahmini_bitis_gunu").iloc[0]
        critical_line = f"🔴 <b>{escape(str(first_critical['urun']))}</b> yaklaşık {num_fmt(first_critical['tahmini_bitis_gunu'], 0)} gün içinde kritik seviyeye düşebilir."
    else:
        critical_line = "🟢 Kritik stok baskısı düşük görünüyor."
    if not miad_df.empty:
        first_miad = miad_df.sort_values("miad_kalan_gun").iloc[0]
        miad_line = f"🟠 <b>{escape(str(first_miad['urun']))}</b> için miada {num_fmt(first_miad['miad_kalan_gun'], 0)} gün kaldı."
    else:
        miad_line = "🟢 Seçilen eşiğe göre yakın miad riski düşük."
    dead_line = f"💀 Ölü/hareketsiz stokta <b>{money_fmt(dead_df['stok_degeri'].sum())}</b> bağlı para var." if not dead_df.empty else "🟢 Ölü stok baskısı düşük görünüyor."
    order_line = f"🛒 Sipariş asistanı <b>{len(order_df)} ürün</b> için aksiyon öneriyor." if not order_df.empty else "🟢 Sipariş eşiğine takılan ürün görünmüyor."
    category_line = f"📈 <b>{escape(str(top_cat))}</b> kategorisi güncel dönemde öne çıkıyor." if top_cat else "📈 Kategori hareketi için yeterli veri bekleniyor."

    health_html = "".join([f'''<div class="health-row"><div class="health-head"><span>{escape(name)}</span><span>{val}/100</span></div><div class="health-bar-bg"><div class="health-bar-fill" style="width:{val}%;"></div></div></div>''' for name, val in health.items()])
    tasks_html = "".join([f'<div class="task-item">{task}</div>' for task in tasks])

    st.markdown(f'''
        <div class="exec-grid"><div class="exec-card"><div class="exec-title">🤖 Günaydın {escape(str(kullanici_adi))}</div><div class="exec-sub">Bugün dikkat etmeniz gereken 5 ana konu aşağıda özetlendi. AYÇA bu alanı stok, miad, satış ve kârlılık verilerinden otomatik yorumlar.</div><div class="exec-list-item">{critical_line}</div><div class="exec-list-item">{miad_line}</div><div class="exec-list-item">{dead_line}</div><div class="exec-list-item">{order_line}</div><div class="exec-list-item">{category_line}</div></div><div class="exec-card"><div class="score-label">Eczane Sağlık Skoru</div><div class="score-big">{score}</div><div class="exec-sub">Durum: <b>{score_status(score)}</b>. Skor; stok, miad, ölü stok ve kârlılık risklerine göre hesaplanır.</div>{health_html}</div></div>
    ''', unsafe_allow_html=True)

    st.markdown('<div class="section-title">AYÇA Radar Merkezi</div>', unsafe_allow_html=True)
    st.markdown(f'''
        <div class="radar-grid"><div class="radar-card"><div class="radar-title">Kritik Stok</div><div class="radar-value">🔴 {len(critical_df)}</div><div class="radar-note">{critical_days} gün ve altı</div></div><div class="radar-card"><div class="radar-title">Miad Riski</div><div class="radar-value">🟠 {len(miad_df)}</div><div class="radar-note">{miad_warning_days} gün içinde</div></div><div class="radar-card"><div class="radar-title">Ölü Stok</div><div class="radar-value">💀 {len(dead_df)}</div><div class="radar-note">{money_fmt(dead_df['stok_degeri'].sum())}</div></div><div class="radar-card"><div class="radar-title">Karlılık</div><div class="radar-value">💰 {pct_fmt(current_margin)}</div><div class="radar-note">Dönem brüt marjı</div></div><div class="radar-card"><div class="radar-title">Sipariş</div><div class="radar-value">🛒 {len(order_df)}</div><div class="radar-note">Öneri bekleyen ürün</div></div></div>
    ''', unsafe_allow_html=True)

    st.markdown(f'''
        <div class="task-grid"><div class="task-card"><div class="exec-title" style="font-size:20px;">✅ Bugün Ne Yapmalıyım?</div><div class="exec-sub">Eczacı için günlük aksiyon listesi.</div>{tasks_html}</div><div class="lost-card"><div class="exec-title" style="font-size:20px;">💰 Kaçırılan Kâr Analizi</div><div class="exec-sub">Bu demo model, stok bitiş riski olan ürünlerde yaklaşık ciro/kâr kaybı ihtimalini hesaplar.</div><div class="score-label">Riskli ürün sayısı</div><div class="lost-number">{lost_count}</div><div class="exec-list-item">Tahmini kaçırılan ciro: <b>{money_fmt(lost_revenue)}</b></div><div class="exec-list-item">Tahmini kaçırılan kâr: <b>{money_fmt(lost_profit)}</b></div></div></div>
    ''', unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.success(f"Giriş: {st.session_state.get('auth_user', 'Demo Kullanıcı')} · {get_membership()}")
if st.sidebar.button("Çıkış Yap", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state.pop("auth_user", None)
    st.session_state.pop("auth_pharmacy", None)
    st.session_state.pop("membership", None)
    safe_rerun()

st.sidebar.title("💊 AYÇA Insight")
st.sidebar.caption("V5.0 Executive Demo")

eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")

uploaded_file = st.sidebar.file_uploader(
    "AYÇA Excel dosyasını yükle",
    type=["xlsx", "xls"],
)

st.sidebar.markdown("---")
st.sidebar.subheader("Analiz Eşikleri")
critical_days = st.sidebar.slider("Kritik stok günü", 1, 15, 5)
warning_days = st.sidebar.slider("Dikkat stok günü", 6, 45, 15)
miad_warning_days = st.sidebar.slider("Miad uyarı günü", 30, 180, 90)
dead_stock_days = st.sidebar.slider("Ölü stok günü", 30, 180, 90)

st.sidebar.markdown("---")
selected_period = st.sidebar.selectbox(
    "Dashboard Dönemi",
    ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Tüm veri"],
    index=2,
)

st.sidebar.caption("Not: Hasta TC, reçete detayları ve kişisel sağlık verisi analiz dışı tutulmalıdır.")


# ============================================================
# DOSYA KONTROL
# ============================================================
if uploaded_file is None:
    st.markdown(
        """
        <div class="ayca-header">
            <div class="ayca-title">
                <h1>AYÇA Insight V5.0</h1>
                <p>Soft dashboard · Excel yükleyerek dinamik analiz alın.</p>
            </div>
            <div class="header-pill">Dosya bekleniyor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info("Sol menüden AYÇA Excel dosyanı yükle. İlk sayfa olarak AYCA_V2_Data okunacaktır.")
    st.stop()


try:
    raw_df, active_sheet, sheet_names = read_excel_smart(uploaded_file)
    df = standardize_ayca_data(raw_df)
except Exception as exc:
    st.error(f"Dosya okunurken hata oluştu: {exc}")
    st.stop()


today = pd.Timestamp.today().normalize()
product_df = build_product_table(df, today)
product_df = classify_inventory(product_df, critical_days, warning_days, miad_warning_days, dead_stock_days)


# ============================================================
# DÖNEM FİLTRESİ
# ============================================================
max_date = df["tarih"].max()
if selected_period == "Son 7 gün":
    start_date = max_date - pd.Timedelta(days=7)
elif selected_period == "Son 14 gün":
    start_date = max_date - pd.Timedelta(days=14)
elif selected_period == "Son 30 gün":
    start_date = max_date - pd.Timedelta(days=30)
else:
    start_date = df["tarih"].min()

period_df = df[df["tarih"] >= start_date].copy()

period_days = max(1, (max_date - start_date).days)
prev_start = start_date - pd.Timedelta(days=period_days)
prev_end = start_date
prev_df = df[(df["tarih"] >= prev_start) & (df["tarih"] < prev_end)].copy()


# ============================================================
# ANA METRİKLER
# ============================================================
current_ciro = period_df["ciro"].sum()
previous_ciro = prev_df["ciro"].sum()

current_profit = period_df["brut_kar"].sum()
previous_profit = prev_df["brut_kar"].sum()

current_cost = period_df["maliyet"].sum()
current_margin = current_profit / current_ciro if current_ciro > 0 else 0
previous_margin = previous_profit / previous_ciro if previous_ciro > 0 else 0

current_transactions = period_df["fis"].nunique()
previous_transactions = prev_df["fis"].nunique()

avg_basket = current_ciro / current_transactions if current_transactions > 0 else 0
prev_avg_basket = previous_ciro / previous_transactions if previous_transactions > 0 else 0

total_units = period_df["adet"].sum()
total_stock_value = product_df["stok_degeri"].sum()

critical_df = product_df[
    (product_df["gunluk_tuketim_60"] > 0)
    & (product_df["tahmini_bitis_gunu"] <= critical_days)
].copy()

warning_df = product_df[
    (product_df["gunluk_tuketim_60"] > 0)
    & (product_df["tahmini_bitis_gunu"] <= warning_days)
].copy()

miad_df = product_df[
    product_df["miad_kalan_gun"].notna()
    & (product_df["miad_kalan_gun"] <= miad_warning_days)
].copy()

dead_df = product_df[product_df["olu_stok_mu"]].copy()

low_margin_df = product_df[
    (product_df["toplam_ciro"] > 0)
    & (product_df["kar_marji"] < 0.12)
].copy()

order_df = product_df[
    (product_df["gunluk_tuketim_60"] > 0)
    & (product_df["tahmini_bitis_gunu"] <= warning_days)
    & (product_df["onerilen_siparis_adedi"] > 0)
].copy()

score = ayca_score(product_df, critical_days, miad_warning_days)


# ============================================================
# SAYFA NAVİGASYONU DURUMU
# ============================================================
PAGE_GENERAL = "🏠 Genel Performans"
PAGE_STOCK = "📦 Stok Bitiş Tahmini"
PAGE_ORDER = "🛒 Sipariş Asistanı"
PAGE_MIAD = "⏳ Miad Takibi"
PAGE_DEAD = "💀 Ölü Stok Analizi"
PAGE_PROFIT = "💰 Kârlılık"
PAGE_CATEGORY = "📊 Kategori Analizi"
PAGE_REPORT = "📥 Rapor"

PAGE_BY_KEY = {
    "genel": PAGE_GENERAL,
    "stok": PAGE_STOCK,
    "siparis": PAGE_ORDER,
    "miad": PAGE_MIAD,
    "olu_stok": PAGE_DEAD,
    "karlilik": PAGE_PROFIT,
    "kategori": PAGE_CATEGORY,
    "rapor": PAGE_REPORT,
}

if "aktif_sayfa" not in st.session_state:
    st.session_state["aktif_sayfa"] = PAGE_GENERAL

def go_page(page_name: str):
    st.session_state["aktif_sayfa"] = page_name


# ============================================================
# HEADER
# ============================================================
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-header">
        <div class="ayca-title">
            <h1>AYÇA Insight V5.0</h1>
            <p>{eczane_adi} · {selected_period} · Sayfa: {active_sheet} · {today_str} · {get_membership()} Demo</p>
        </div>
        <div class="header-pill">AYÇA Skoru: {score}/100</div>
    </div>
    """,
    unsafe_allow_html=True,
)


if not is_premium_user():
    st.markdown(
        """
        <div class="ai-card" style="background: linear-gradient(135deg, #FFFFFF 0%, #FEF3C7 100%); border-color: #FDE68A;">
            <div class="ai-title" style="color:#B45309;">🔐 Basic Demo Modu</div>
            <div class="ai-text">Bu kullanıcı kısa özet, KPI kartları ve 1-2 satırlık önizleme görür. Tüm grafikler, detaylı tablolar ve Excel raporu Premium kullanıcıda açılır.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# KPI KARTLARI
# ============================================================
ciro_trend, ciro_class = rate_fmt(current_ciro, previous_ciro)
profit_trend, profit_class = rate_fmt(current_profit, previous_profit)
margin_trend, margin_class = rate_fmt(current_margin, previous_margin)
basket_trend, basket_class = rate_fmt(avg_basket, prev_avg_basket)

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    make_metric_card("Güncel Ciro", money_fmt(current_ciro), selected_period, ciro_trend, ciro_class)
with k2:
    make_metric_card("Brüt Kâr", money_fmt(current_profit), "Dönem kârı", profit_trend, profit_class)
with k3:
    make_metric_card("Kâr Marjı", pct_fmt(current_margin), "Brüt kâr / ciro", margin_trend, margin_class)
with k4:
    make_metric_card("İşlem Sayısı", f"{current_transactions}", "Fiş/Reçete tekil", None)
with k5:
    make_metric_card("Ortalama Sepet", money_fmt(avg_basket), "Ciro / işlem", basket_trend, basket_class)




# ============================================================
# V5 EXECUTIVE DASHBOARD
# ============================================================
if st.session_state.get("aktif_sayfa", PAGE_GENERAL) == PAGE_GENERAL:
    render_executive_dashboard(
        kullanici_adi=kullanici_adi,
        score=score,
        product_df=product_df,
        period_df=period_df,
        prev_df=prev_df,
        critical_df=critical_df,
        miad_df=miad_df,
        dead_df=dead_df,
        low_margin_df=low_margin_df,
        order_df=order_df,
        current_margin=current_margin,
        previous_margin=previous_margin,
        critical_days=critical_days,
        miad_warning_days=miad_warning_days,
        warning_days=warning_days,
    )


# ============================================================
# RİSK KARTLARI
# ============================================================
st.markdown('<div class="section-title">Kritik Merkez</div>', unsafe_allow_html=True)

r1, r2, r3, r4, r5 = st.columns(5)
with r1:
    make_clickable_mini_card("Kritik Stok", f"{len(critical_df)} ürün", f"{critical_days} gün ve altı", "alert-red", PAGE_STOCK, "Kritik stokları göster", "card_kritik_stok")
with r2:
    make_clickable_mini_card("Miad Uyarısı", f"{len(miad_df)} ürün", f"{miad_warning_days} gün içinde", "alert-orange", PAGE_MIAD, "Miad listesini aç", "card_miad")
with r3:
    make_clickable_mini_card("Ölü Stok", money_fmt(dead_df["stok_degeri"].sum()), f"{dead_stock_days}+ gün / çıkış yok", "alert-purple", PAGE_DEAD, "Ölü stokları incele", "card_olu_stok")
with r4:
    make_clickable_mini_card("Stok Değeri", money_fmt(total_stock_value), "Mevcut stok maliyeti", "alert-blue", PAGE_STOCK, "Stok analizine git", "card_stok_degeri")
with r5:
    make_clickable_mini_card("Sipariş Önerisi", f"{len(order_df)} ürün", money_fmt(order_df["onerilen_siparis_maliyeti"].sum()), "alert-green", PAGE_ORDER, "Sipariş listesini aç", "card_siparis")


# ============================================================
# ESKİ YATAY BÖLÜM MENÜSÜ
# ============================================================
# Not:
# Streamlit'in native st.tabs bileşeni dışarıdan programatik olarak
# seçtirilemediği için burada aynı görünüm/mantığa yakın yatay radio kullanılır.
# Kritik Merkez kartlarına basıldığında st.session_state["aktif_sayfa"] değişir;
# yani alttaki bölüm sanki o sekmeye basılmış gibi açılır.

st.markdown('<div style="height: 8px;"></div>', unsafe_allow_html=True)
st.radio(
    "Bölüm seçimi",
    [PAGE_GENERAL, PAGE_STOCK, PAGE_ORDER, PAGE_MIAD, PAGE_DEAD, PAGE_PROFIT, PAGE_CATEGORY, PAGE_REPORT],
    key="aktif_sayfa",
    horizontal=True,
    label_visibility="collapsed",
)

st.markdown('<div style="height: 12px;"></div>', unsafe_allow_html=True)


if st.session_state["aktif_sayfa"] == PAGE_GENERAL:
    # ============================================================
    # AYÇA AI YORUMU
    # ============================================================
    ai_text = create_ai_comment(
        df=period_df,
        p=product_df,
        critical_df=critical_df,
        miad_df=miad_df,
        dead_df=dead_df,
        margin_df=low_margin_df,
        current_ciro=current_ciro,
        previous_ciro=previous_ciro,
    )

    st.markdown(
        f"""
        <div class="ai-card">
            <div class="ai-title">🤖 AYÇA AI Yorumu</div>
            <div class="ai-text">
                Günaydın {kullanici_adi}. {ai_text}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


    # ============================================================
    # ANA GRAFİKLER
    # ============================================================
    st.markdown('<div class="section-title">Genel Performans</div>', unsafe_allow_html=True)

    daily = df.groupby("tarih", as_index=False).agg(
        ciro=("ciro", "sum"),
        brut_kar=("brut_kar", "sum"),
        adet=("adet", "sum"),
        islem=("fis", pd.Series.nunique),
    )

    monthly = df.copy()
    monthly["ay"] = monthly["tarih"].dt.to_period("M").astype(str)
    monthly_summary = monthly.groupby("ay", as_index=False).agg(
        ciro=("ciro", "sum"),
        brut_kar=("brut_kar", "sum"),
        adet=("adet", "sum"),
        islem=("fis", pd.Series.nunique),
    )
    monthly_summary["kar_marji"] = np.where(monthly_summary["ciro"] > 0, monthly_summary["brut_kar"] / monthly_summary["ciro"], 0)

    cat_summary = period_df.groupby("kategori", as_index=False).agg(
        ciro=("ciro", "sum"),
        brut_kar=("brut_kar", "sum"),
        adet=("adet", "sum"),
        islem=("fis", pd.Series.nunique),
    )
    cat_summary["kar_marji"] = np.where(cat_summary["ciro"] > 0, cat_summary["brut_kar"] / cat_summary["ciro"], 0)
    cat_summary = cat_summary.sort_values("ciro", ascending=False)

    g1, g2 = st.columns([1.35, 1])
    with g1:
        fig = px.area(
            monthly_summary,
            x="ay",
            y="ciro",
            title="Aylık Ciro Trendi",
            markers=True,
            labels={"ay": "Ay", "ciro": "Ciro TL"},
        )
        fig.update_traces(line_color="#2563EB", fillcolor="rgba(37,99,235,0.12)")
        st.plotly_chart(apply_plot_theme(fig, 360), use_container_width=True)

    with g2:
        if not cat_summary.empty:
            fig = px.pie(
                cat_summary,
                values="ciro",
                names="kategori",
                hole=0.55,
                title="Kategori Ciro Dağılımı",
                color_discrete_sequence=px.colors.qualitative.Set3,
            )
            st.plotly_chart(apply_plot_theme(fig, 360), use_container_width=True)


    g3, g4 = st.columns(2)
    with g3:
        top_products = period_df.groupby("urun", as_index=False).agg(
            ciro=("ciro", "sum"),
            adet=("adet", "sum"),
            brut_kar=("brut_kar", "sum"),
        ).sort_values("ciro", ascending=False).head(12)

        fig = px.bar(
            top_products.sort_values("ciro"),
            x="ciro",
            y="urun",
            orientation="h",
            title="En Çok Ciro Getiren Ürünler",
            labels={"ciro": "Ciro TL", "urun": "Ürün"},
            color="ciro",
            color_continuous_scale=["#DBEAFE", "#2563EB"],
        )
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)

    with g4:
        profit_products = period_df.groupby("urun", as_index=False).agg(
            brut_kar=("brut_kar", "sum"),
            ciro=("ciro", "sum"),
            adet=("adet", "sum"),
        ).sort_values("brut_kar", ascending=False).head(12)

        fig = px.bar(
            profit_products.sort_values("brut_kar"),
            x="brut_kar",
            y="urun",
            orientation="h",
            title="En Karlı Ürünler",
            labels={"brut_kar": "Brüt Kâr TL", "urun": "Ürün"},
            color="brut_kar",
            color_continuous_scale=["#DCFCE7", "#10B981"],
        )
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)


# ============================================================
# SAYFA İÇERİKLERİ
# ============================================================

# ------------------------------------------------------------
# TAB 1 - STOK BİTİŞ TAHMİNİ
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_STOCK:
    st.markdown('<div class="section-title">📦 Stok Bitiş Tahmini</div>', unsafe_allow_html=True)

    view = product_df.copy()
    view["Durum"] = view["stok_durumu"].map(
        {
            "Kritik": "🔴 Kritik",
            "Dikkat": "🟠 Dikkat",
            "Güvenli": "🟢 Güvenli",
            "Ölü Stok": "⚫ Ölü Stok",
            "Veri Yok": "Veri Yok",
        }
    )
    view["Tahmini Bitiş Günü"] = view["tahmini_bitis_gunu"].replace(np.inf, np.nan).round(1)
    view["Günlük Tüketim"] = view["gunluk_tuketim_60"].round(2)
    view["Stok Değeri"] = view["stok_degeri"].round(2)
    view["Kâr Marjı"] = (view["kar_marji"] * 100).round(1)

    c_a, c_b, c_c = st.columns(3)
    with c_a:
        st.metric("Kritik Ürün", len(critical_df))
    with c_b:
        st.metric("Dikkat Gerektiren", len(warning_df))
    with c_c:
        st.metric("Ortalama Stok Değeri", money_fmt(product_df["stok_degeri"].mean()))

    st.dataframe(
        view[
            [
                "Durum",
                "urun",
                "kategori",
                "alt_kategori",
                "mevcut_stok",
                "son_60",
                "Günlük Tüketim",
                "Tahmini Bitiş Günü",
                "Stok Değeri",
                "Kâr Marjı",
                "siparis_onerisi",
            ]
        ].sort_values(["Tahmini Bitiş Günü", "mevcut_stok"], ascending=[True, True]),
        use_container_width=True,
        hide_index=True,
    )

    closest = view[view["gunluk_tuketim_60"] > 0].sort_values("tahmini_bitis_gunu").head(15)
    if not closest.empty:
        fig = px.bar(
            closest.sort_values("tahmini_bitis_gunu", ascending=True),
            x="urun",
            y="tahmini_bitis_gunu",
            title="En Yakın Bitecek Ürünler",
            labels={"urun": "Ürün", "tahmini_bitis_gunu": "Tahmini Bitiş Günü"},
            color="tahmini_bitis_gunu",
            color_continuous_scale=["#FEE2E2", "#FEF3C7", "#DBEAFE"],
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(apply_plot_theme(fig, 420), use_container_width=True)


# ------------------------------------------------------------
# TAB 2 - SİPARİŞ ASİSTANI
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_ORDER:
    st.markdown('<div class="section-title">🛒 Sipariş Asistanı</div>', unsafe_allow_html=True)

    st.info("Sipariş önerisi; son 60 gün çıkış hızına ve mevcut stok bitiş süresine göre hesaplanır.")

    order_view = order_df.copy()
    if order_view.empty:
        st.success("Şu anda eşiklere göre sipariş önerisi gerektiren ürün görünmüyor.")
    else:
        o1, o2, o3 = st.columns(3)
        with o1:
            st.metric("Sipariş Önerilen Ürün", len(order_view))
        with o2:
            st.metric("Tahmini Sipariş Maliyeti", money_fmt(order_view["onerilen_siparis_maliyeti"].sum()))
        with o3:
            st.metric("Ortalama Bitiş Günü", num_fmt(order_view["tahmini_bitis_gunu"].mean(), 1))

        order_view["Tahmini Bitiş"] = order_view["tahmini_bitis_gunu"].round(1)
        order_view["Günlük Tüketim"] = order_view["gunluk_tuketim_60"].round(2)
        order_view["Önerilen Sipariş"] = order_view["onerilen_siparis_adedi"].astype(int)
        order_view["Tahmini Maliyet"] = order_view["onerilen_siparis_maliyeti"].round(2)

        st.dataframe(
            order_view[
                [
                    "stok_durumu",
                    "urun",
                    "kategori",
                    "mevcut_stok",
                    "son_60",
                    "Günlük Tüketim",
                    "Tahmini Bitiş",
                    "Önerilen Sipariş",
                    "Tahmini Maliyet",
                    "tedarikci",
                    "raf",
                ]
            ].sort_values("Tahmini Bitiş"),
            use_container_width=True,
            hide_index=True,
        )

        fig = px.bar(
            order_view.sort_values("onerilen_siparis_maliyeti", ascending=False).head(15),
            x="urun",
            y="onerilen_siparis_maliyeti",
            title="Sipariş Maliyeti En Yüksek Öneriler",
            labels={"urun": "Ürün", "onerilen_siparis_maliyeti": "Tahmini Maliyet"},
            color="onerilen_siparis_maliyeti",
            color_continuous_scale=["#DCFCE7", "#10B981"],
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)


# ------------------------------------------------------------
# TAB 3 - MİAD TAKİBİ
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_MIAD:
    st.markdown('<div class="section-title">⏳ Miad Takibi</div>', unsafe_allow_html=True)

    miad_filter = st.radio(
        "Miad filtresi",
        ["30 gün", "60 gün", "90 gün", "180 gün", "Tümü"],
        horizontal=True,
        index=2,
    )

    if miad_filter == "30 gün":
        miad_limit = 30
    elif miad_filter == "60 gün":
        miad_limit = 60
    elif miad_filter == "90 gün":
        miad_limit = 90
    elif miad_filter == "180 gün":
        miad_limit = 180
    else:
        miad_limit = 99999

    miad_active = product_df[
        product_df["miad_kalan_gun"].notna()
        & (product_df["miad_kalan_gun"] <= miad_limit)
    ].copy()

    expired = product_df[
        product_df["miad_kalan_gun"].notna()
        & (product_df["miad_kalan_gun"] < 0)
    ].copy()

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Miad Uyarısı", len(miad_active))
    with m2:
        st.metric("Miadı Geçmiş", len(expired))
    with m3:
        st.metric("Riskli Stok Değeri", money_fmt(miad_active["stok_degeri"].sum()))
    with m4:
        if not miad_active.empty:
            st.metric("En Yakın Miad", f"{int(miad_active['miad_kalan_gun'].min())} gün")
        else:
            st.metric("En Yakın Miad", "-")

    if miad_active.empty:
        st.success("Seçilen filtreye göre miadı yaklaşan ürün bulunmuyor.")
    else:
        miad_active["Miad Tarihi"] = miad_active["miad"].dt.strftime("%d.%m.%Y")
        miad_active["Kalan Gün"] = miad_active["miad_kalan_gun"].astype(int)
        miad_active["Stok Değeri"] = miad_active["stok_degeri"].round(2)

        st.dataframe(
            miad_active[
                [
                    "miad_durumu",
                    "urun",
                    "kategori",
                    "mevcut_stok",
                    "Miad Tarihi",
                    "Kalan Gün",
                    "Stok Değeri",
                    "raf",
                    "tedarikci",
                ]
            ].sort_values("Kalan Gün"),
            use_container_width=True,
            hide_index=True,
        )

        fig = px.bar(
            miad_active.sort_values("miad_kalan_gun").head(20),
            x="urun",
            y="miad_kalan_gun",
            title="Miadı En Yakın Ürünler",
            labels={"urun": "Ürün", "miad_kalan_gun": "Kalan Gün"},
            color="miad_kalan_gun",
            color_continuous_scale=["#FEE2E2", "#FEF3C7", "#DBEAFE"],
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)

        fig = px.pie(
            miad_active.groupby("kategori", as_index=False)["stok_degeri"].sum(),
            values="stok_degeri",
            names="kategori",
            hole=0.50,
            title="Miad Riski Kategori Dağılımı",
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        st.plotly_chart(apply_plot_theme(fig, 360), use_container_width=True)


# ------------------------------------------------------------
# TAB 4 - ÖLÜ STOK ANALİZİ
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_DEAD:
    st.markdown('<div class="section-title">💀 Ölü Stok Analizi</div>', unsafe_allow_html=True)

    dead_mode = st.radio(
        "Ölü stok kriteri",
        ["Son 60 gün çıkışı olmayanlar", f"{dead_stock_days}+ gündür satılmayanlar", "İkisi birlikte"],
        horizontal=True,
    )

    if dead_mode == "Son 60 gün çıkışı olmayanlar":
        dead_active = product_df[(product_df["mevcut_stok"] > 0) & (product_df["son_60"] <= 0)].copy()
    elif dead_mode == f"{dead_stock_days}+ gündür satılmayanlar":
        dead_active = product_df[(product_df["mevcut_stok"] > 0) & (product_df["son_satis_kac_gun"] >= dead_stock_days)].copy()
    else:
        dead_active = dead_df.copy()

    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.metric("Ölü Stok Ürün", len(dead_active))
    with d2:
        st.metric("Bağlı Para", money_fmt(dead_active["stok_degeri"].sum()))
    with d3:
        st.metric("Toplam Stok Adedi", num_fmt(dead_active["mevcut_stok"].sum(), 0))
    with d4:
        if len(product_df) > 0:
            st.metric("Ürün Oranı", pct_fmt(len(dead_active) / len(product_df)))
        else:
            st.metric("Ürün Oranı", "%0,0")

    if dead_active.empty:
        st.success("Seçilen kritere göre ölü stok görünmüyor.")
    else:
        dead_active["Son Satış"] = dead_active["son_satis_tarihi"].dt.strftime("%d.%m.%Y")
        dead_active["Satışsız Gün"] = dead_active["son_satis_kac_gun"].astype(int)
        dead_active["Stok Değeri"] = dead_active["stok_degeri"].round(2)

        st.dataframe(
            dead_active[
                [
                    "urun",
                    "kategori",
                    "mevcut_stok",
                    "son_60",
                    "Son Satış",
                    "Satışsız Gün",
                    "Stok Değeri",
                    "raf",
                    "tedarikci",
                ]
            ].sort_values("Stok Değeri", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

        col_dead1, col_dead2 = st.columns(2)
        with col_dead1:
            fig = px.bar(
                dead_active.sort_values("stok_degeri", ascending=False).head(15),
                x="urun",
                y="stok_degeri",
                title="En Yüksek Ölü Stok Değeri",
                labels={"urun": "Ürün", "stok_degeri": "Stok Değeri"},
                color="stok_degeri",
                color_continuous_scale=["#EDE9FE", "#8B5CF6"],
            )
            fig.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)

        with col_dead2:
            dead_cat = dead_active.groupby("kategori", as_index=False)["stok_degeri"].sum().sort_values("stok_degeri", ascending=False)
            fig = px.pie(
                dead_cat,
                values="stok_degeri",
                names="kategori",
                hole=0.52,
                title="Ölü Stok Kategori Dağılımı",
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)


# ------------------------------------------------------------
# TAB 5 - KÂRLILIK
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_PROFIT:
    st.markdown('<div class="section-title">💰 Kârlılık Analizi</div>', unsafe_allow_html=True)

    p_profit = product_df.copy()
    p_profit["Kâr Marjı %"] = (p_profit["kar_marji"] * 100).round(1)
    p_profit["Toplam Brüt Kâr"] = p_profit["toplam_brut_kar"].round(2)
    p_profit["Toplam Ciro"] = p_profit["toplam_ciro"].round(2)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Dönem Brüt Kâr", money_fmt(current_profit))
    with c2:
        st.metric("Dönem Kâr Marjı", pct_fmt(current_margin))
    with c3:
        st.metric("Düşük Marjlı Ürün", len(low_margin_df))

    st.dataframe(
        p_profit[
            [
                "urun",
                "kategori",
                "toplam_adet",
                "Toplam Ciro",
                "Toplam Brüt Kâr",
                "Kâr Marjı %",
                "ort_alis_birim",
                "ort_satis_birim",
            ]
        ].sort_values("Toplam Brüt Kâr", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

    colp1, colp2 = st.columns(2)
    with colp1:
        fig = px.scatter(
            p_profit,
            x="toplam_ciro",
            y="kar_marji",
            size="toplam_adet",
            color="kategori",
            hover_name="urun",
            title="Ciro / Kâr Marjı Haritası",
            labels={"toplam_ciro": "Toplam Ciro", "kar_marji": "Kâr Marjı"},
        )
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)

    with colp2:
        low = p_profit[p_profit["toplam_ciro"] > 0].sort_values("kar_marji").head(15)
        fig = px.bar(
            low,
            x="urun",
            y="kar_marji",
            title="Kâr Marjı En Düşük Ürünler",
            labels={"urun": "Ürün", "kar_marji": "Kâr Marjı"},
            color="kar_marji",
            color_continuous_scale=["#FEE2E2", "#FEF3C7", "#DCFCE7"],
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)


# ------------------------------------------------------------
# TAB 6 - KATEGORİ ANALİZİ
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_CATEGORY:
    st.markdown('<div class="section-title">📊 Kategori Analizi</div>', unsafe_allow_html=True)

    cat_all = df.groupby("kategori", as_index=False).agg(
        ciro=("ciro", "sum"),
        brut_kar=("brut_kar", "sum"),
        adet=("adet", "sum"),
        islem=("fis", pd.Series.nunique),
    )
    cat_all["kar_marji"] = np.where(cat_all["ciro"] > 0, cat_all["brut_kar"] / cat_all["ciro"], 0)
    cat_all["Ortalama Sepet"] = np.where(cat_all["islem"] > 0, cat_all["ciro"] / cat_all["islem"], 0)

    st.dataframe(
        cat_all.sort_values("ciro", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

    cat_col1, cat_col2 = st.columns(2)
    with cat_col1:
        fig = px.bar(
            cat_all.sort_values("ciro", ascending=False),
            x="kategori",
            y="ciro",
            title="Kategori Bazlı Ciro",
            labels={"kategori": "Kategori", "ciro": "Ciro"},
            color="ciro",
            color_continuous_scale=["#DBEAFE", "#2563EB"],
        )
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)

    with cat_col2:
        fig = px.bar(
            cat_all.sort_values("brut_kar", ascending=False),
            x="kategori",
            y="brut_kar",
            title="Kategori Bazlı Brüt Kâr",
            labels={"kategori": "Kategori", "brut_kar": "Brüt Kâr"},
            color="brut_kar",
            color_continuous_scale=["#DCFCE7", "#10B981"],
        )
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(apply_plot_theme(fig, 430), use_container_width=True)


# ------------------------------------------------------------
# TAB 7 - RAPOR
# ------------------------------------------------------------
if st.session_state["aktif_sayfa"] == PAGE_REPORT:
    st.markdown('<div class="section-title">📥 Rapor ve Veri Kontrol</div>', unsafe_allow_html=True)

    st.write("Analiz sonucunu Excel olarak indirebilirsiniz.")

    report_bytes = create_excel_report(df, product_df, critical_df, miad_df, dead_df, order_df)

    st.download_button(
        label="📥 AYÇA V4.2 Analiz Raporu İndir",
        data=report_bytes,
        file_name=f"AYCA_Insight_V4_2_Rapor_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("### Okunan Excel Bilgisi")
    st.write(f"Aktif sayfa: **{active_sheet}**")
    st.write(f"Dosyadaki sayfalar: {', '.join(sheet_names)}")
    st.write(f"Satır sayısı: **{len(df)}**")
    st.write(f"Ürün sayısı: **{len(product_df)}**")

    st.markdown("### Ham Veri Önizleme")
    st.dataframe(df.head(100), use_container_width=True, hide_index=True)

    st.markdown("### Ürün Analizi Önizleme")
    st.dataframe(product_df.head(100).replace([np.inf, -np.inf], np.nan), use_container_width=True, hide_index=True)


# ============================================================
# ALT BİLGİ
# ============================================================
st.markdown("---")
st.caption(
    "AYÇA Insight V4.2 Soft · Raporları okumaz, eczanenizi yorumlar. "
    "Bu uygulama karar destek amaçlıdır; nihai ticari ve mesleki karar kullanıcıya aittir."
)
