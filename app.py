# ============================================================
# AYÇA Insight V7.1 - Ürün, Doktor Intelligence ve Hasta Sadakat Sürümü
# Satış Excel + Envanter Excel + Ürün Bazında Toplamlar ayrı ayrı yüklenir.
# ------------------------------------------------------------
# Bu sürüm ürün satış kalem raporu olmadan güvenilir analiz üretir:
# - Ciro, brüt kâr, marj, toplam stok değeri, tahsilat açığı
# - Önceki eş dönem kıyaslama
# - Gün / saat / hafta günü satış ritmi
# - Kurum, doktor, tahsilat, satış tipi analizi
# - Tahsilat açığı ve sonlanmamış satış uyarıları
# - Stok sermayesi, kritik stok, stok yok, yüksek stok
# - Pareto stok sermayesi analizi
# - Günlük aksiyon listesi ve yönetici özeti
#
# Not: Ürün bazlı satış hızı, sipariş tavsiyesi, ölü stok ve miad motoru
# ürün satış raporu gelene kadar bilinçli olarak kapalıdır.
# ============================================================

from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
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
    page_title="AYÇA Insight V7.1",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CSS
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --bg: #F8FAFC;
        --panel: #FFFFFF;
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
    }
    html, body, [data-testid="stAppViewContainer"] {background: var(--bg); color: var(--text);} 
    [data-testid="stHeader"] {background: rgba(248,250,252,.88); backdrop-filter: blur(10px);} 
    [data-testid="stSidebar"] {background: linear-gradient(180deg,#FFFFFF 0%,#F1F5F9 100%); border-right: 1px solid var(--border);} 
    .block-container {padding-top: 1.2rem; max-width: 1540px;}
    .ayca-header {display:flex; justify-content:space-between; align-items:center; gap:16px; margin-bottom:16px;}
    .ayca-title h1 {margin:0; color:var(--text); font-size:32px; letter-spacing:-.7px; font-weight:950;}
    .ayca-title p {margin:6px 0 0 0; color:var(--muted); font-size:14px;}
    .header-pill {background:#fff; border:1px solid var(--border); border-radius:16px; padding:12px 16px; box-shadow:0 8px 24px rgba(15,23,42,.05); color:var(--text); font-weight:900; font-size:13px; white-space:nowrap;}
    .metric-card {min-height:132px; background:linear-gradient(180deg,#fff 0%,#F8FBFF 100%); border:1px solid var(--border); border-radius:20px; padding:18px; box-shadow:0 12px 30px rgba(15,23,42,.06); overflow:hidden; position:relative;}
    .metric-label {color:var(--muted); font-size:12px; font-weight:900; letter-spacing:.4px; text-transform:uppercase; margin-bottom:10px;}
    .metric-value {color:var(--text); font-size:28px; font-weight:950; margin-bottom:8px; letter-spacing:-.5px;}
    .metric-sub {color:var(--muted); font-size:13px;}
    .metric-up {color:var(--green); font-size:13px; font-weight:900;}
    .metric-down {color:var(--red); font-size:13px; font-weight:900;}
    .mini-card {background:var(--panel); border:1px solid var(--border); border-radius:18px; padding:16px; box-shadow:0 10px 26px rgba(15,23,42,.05); min-height:104px;}
    .mini-title {color:var(--text); font-size:14px; font-weight:900; margin-bottom:8px;}
    .mini-value {color:var(--text); font-size:24px; font-weight:950; margin-bottom:4px;}
    .mini-note {color:var(--muted); font-size:13px; line-height:1.45;}
    .alert-red {background:linear-gradient(135deg,#fff 0%,var(--red-soft) 100%); border-color:#FECACA;}
    .alert-orange {background:linear-gradient(135deg,#fff 0%,var(--orange-soft) 100%); border-color:#FDE68A;}
    .alert-green {background:linear-gradient(135deg,#fff 0%,var(--green-soft) 100%); border-color:#BBF7D0;}
    .alert-blue {background:linear-gradient(135deg,#fff 0%,var(--blue-soft) 100%); border-color:#BFDBFE;}
    .alert-purple {background:linear-gradient(135deg,#fff 0%,var(--purple-soft) 100%); border-color:#DDD6FE;}
    .ai-card {background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%); border:1px solid #BFDBFE; border-radius:22px; padding:20px; box-shadow:0 12px 30px rgba(37,99,235,.08); margin:14px 0 18px 0;}
    .ai-title {color:var(--blue); font-size:18px; font-weight:950; margin-bottom:8px;}
    .ai-text {color:var(--text); font-size:14px; line-height:1.55;}
    .section-title {color:var(--text); font-size:21px; font-weight:950; margin:20px 0 12px 0; letter-spacing:-.3px;}
    .exec-grid {display:grid; grid-template-columns:1.25fr .75fr; gap:16px; margin:14px 0 18px 0;}
    .exec-card {background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%); border:1px solid #BFDBFE; border-radius:24px; padding:20px; box-shadow:0 14px 34px rgba(37,99,235,.08);}
    .exec-title {color:#0F172A; font-size:22px; font-weight:950; letter-spacing:-.4px; margin-bottom:8px;}
    .exec-sub {color:#64748B; font-size:14px; line-height:1.55; margin-bottom:14px;}
    .exec-list-item {background:rgba(255,255,255,.78); border:1px solid rgba(226,232,240,.95); border-radius:16px; padding:12px 13px; margin:9px 0; font-size:14px; color:#0F172A; line-height:1.45; font-weight:700;}
    .score-big {font-size:54px; font-weight:950; letter-spacing:-2px; color:#2563EB; line-height:1; margin:4px 0 8px 0;}
    .health-row {margin:12px 0;}
    .health-head {display:flex; justify-content:space-between; color:#334155; font-weight:900; font-size:13px; margin-bottom:6px;}
    .health-bar-bg {width:100%; height:10px; background:#E2E8F0; border-radius:999px; overflow:hidden;}
    .health-bar-fill {height:10px; border-radius:999px; background:linear-gradient(90deg,#2563EB,#10B981);}
    @media (max-width:1000px){.exec-grid{grid-template-columns:1fr;}.ayca-header{display:block}.header-pill{display:inline-block;margin-top:10px}}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DEMO GİRİŞ
# ============================================================
DEMO_USERS = {
    "basic": {"password": "basic2026", "name": "Basic Demo Kullanıcı", "pharmacy": "İdil Eczanesi", "membership": "Basic"},
    "premium": {"password": "premium2026", "name": "Premium Demo Kullanıcı", "pharmacy": "İdil Eczanesi", "membership": "Premium"},
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


def show_basic_info(message: str):
    st.markdown(
        f"""
        <div class="mini-card alert-orange" style="margin: 10px 0 16px 0; min-height: auto;">
            <div class="mini-title">🔒 Basic Üyelik Önizlemesi</div>
            <div class="mini-note">{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_demo_auth_screen():
    st.markdown("### 💊 AYÇA Insight Demo Giriş")
    c1, c2 = st.columns([1.1, .9])
    with c1:
        st.markdown(
            """
            <div class="ai-card">
                <div class="ai-title">AYÇA Insight V7.1</div>
                <div class="ai-text">
                Bu sürüm iki ayrı TEBEOS Excel çıktısını okur: <b>Satış</b> ve <b>Envanter</b>.
                Ürün satış raporu olmadan ciro, kârlılık, tahsilat, doktor/kurum ve stok sermayesi analizini güçlendirir.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        username = st.text_input("Kullanıcı adı", value="premium")
        password = st.text_input("Şifre", value="premium2026", type="password")
        if st.button("🚀 Giriş Yap", use_container_width=True):
            record = DEMO_USERS.get(username.strip().lower())
            if record and password == record["password"]:
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = record["name"]
                st.session_state["auth_pharmacy"] = record["pharmacy"]
                st.session_state["membership"] = record["membership"]
                safe_rerun()
            else:
                st.error("Kullanıcı adı veya şifre hatalı. Premium: premium / premium2026")


if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    show_demo_auth_screen()
    st.stop()


# ============================================================
# BASIC/PREMIUM KİLİDİ
# ============================================================
_original_dataframe = st.dataframe
_original_plotly_chart = st.plotly_chart
_original_download_button = st.download_button


def premium_locked_chart(*args, **kwargs):
    if is_premium_user():
        return _original_plotly_chart(*args, **kwargs)
    show_basic_info("Basic üyelikte grafikler kapalıdır. Premium kullanıcıda açılır.")
    return None


def limited_dataframe(data=None, *args, **kwargs):
    if is_premium_user():
        return _original_dataframe(data, *args, **kwargs)
    try:
        preview = data.head(2).copy() if hasattr(data, "head") else data
        show_basic_info("Basic üyelikte tablo önizlemesi en fazla 2 satırdır.")
        return _original_dataframe(preview, *args, **kwargs)
    except Exception:
        return _original_dataframe(data, *args, **kwargs)


def premium_download_button(*args, **kwargs):
    if is_premium_user():
        return _original_download_button(*args, **kwargs)
    show_basic_info("Basic üyelikte Excel raporu indirme kapalıdır.")
    return False


st.plotly_chart = premium_locked_chart
st.dataframe = limited_dataframe
st.download_button = premium_download_button


# ============================================================
# GENEL YARDIMCILAR
# ============================================================
def normalize_col_name(name: str) -> str:
    name = str(name).strip().lower()
    tr_map = str.maketrans("çğıöşüİı", "cgiosuii")
    name = name.translate(tr_map)
    name = re.sub(r"[^a-z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name


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
            if cand_norm and cand_norm in norm:
                return original
    return None


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
        return ("▲ " + pct_fmt(rate), "metric-up") if rate >= 0 else ("▼ " + pct_fmt(abs(rate)), "metric-down")
    except Exception:
        return "▲ %0,0", "metric-up"


def excel_serial_to_datetime(series: pd.Series) -> pd.Series:
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


def read_excel_first_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=sheet)
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    return df, sheet, xls.sheet_names


def make_metric_card(label, value, sub="", trend_text_value=None, trend_class="metric-up"):
    trend_html = f" <span class='{trend_class}'>{trend_text_value}</span>" if trend_text_value else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-sub">{sub}{trend_html}</div>
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


def safe_div(a, b):
    return float(a) / float(b) if float(b or 0) != 0 else 0.0


# ============================================================
# SATIŞ VE ENVANTER STANDARDİZASYONU
# ============================================================
def standardize_sales(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw_df.columns)
    mapping = {
        "satis_no": find_col(cols, ["Satış No", "Satis No", "Fiş No", "Fis No"]),
        "satis_tipi": find_col(cols, ["Satış Tipi", "Satis Tipi"]),
        "tahsilat": find_col(cols, ["Tahsilat", "Ödeme Tipi", "Odeme Tipi"]),
        "hasta": find_col(cols, ["Hasta Adı Soyadı", "Hasta Adi Soyadi"]),
        "recete_no": find_col(cols, ["Reç. No", "Rec. No", "Reçete No", "Recete No"]),
        "erecete_no": find_col(cols, ["EReç. No", "ERec. No", "E Reçete No"]),
        "doktor": find_col(cols, ["Doktor Adı", "Doktor Adi", "Doktor"]),
        "recete_tarihi": find_col(cols, ["Reç. Tar", "Rec. Tar", "Reçete Tarihi"]),
        "kurum": find_col(cols, ["Kurum Adı", "Kurum Adi"]),
        "grup": find_col(cols, ["Grubu", "Grup"]),
        "alım_tarihi": find_col(cols, ["Alım Tarih", "Alim Tarih"]),
        "odenen": find_col(cols, ["Ödenen Tutar", "Odenen Tutar"]),
        "toplam": find_col(cols, ["Toplam Tutar", "Ciro TL", "Ciro"]),
        "iskonto": find_col(cols, ["İskonto", "Iskonto"]),
        "elde_toplam": find_col(cols, ["Eld. Top. Tut", "Elden Toplam Tutar"]),
        "kar": find_col(cols, ["Kar Tutarı", "Kar Tutari", "Brüt Kar TL", "Brut Kar TL"]),
        "maliyet": find_col(cols, ["Maliyet Tutarı", "Maliyet Tutari", "Maliyet TL"]),
        "fiyat_farki": find_col(cols, ["Fiy. Farkı", "Fiyat Farkı", "Fiy Farki"]),
        "sonlandi": find_col(cols, ["Sonlandı", "Sonlandi"]),
        "islem_tarihi": find_col(cols, ["İşlem Tarihi", "Islem Tarihi", "Tarih"]),
        "kullanici": find_col(cols, ["Kullanıcı", "Kullanici"]),
    }
    required = ["satis_no", "toplam", "kar", "maliyet", "islem_tarihi"]
    missing = [k for k in required if mapping.get(k) is None]
    if missing:
        raise ValueError("Satış Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["satis_no"] = raw_df[mapping["satis_no"]].astype(str)
    df["satis_tipi"] = raw_df[mapping["satis_tipi"]].astype(str) if mapping["satis_tipi"] else "Bilinmiyor"
    df["tahsilat"] = raw_df[mapping["tahsilat"]].astype(str) if mapping["tahsilat"] else "Bilinmiyor"
    df["hasta"] = raw_df[mapping["hasta"]].astype(str) if mapping["hasta"] else ""
    df["recete_no"] = raw_df[mapping["recete_no"]].astype(str) if mapping["recete_no"] else ""
    df["erecete_no"] = raw_df[mapping["erecete_no"]].astype(str) if mapping["erecete_no"] else ""
    df["doktor"] = raw_df[mapping["doktor"]].astype(str) if mapping["doktor"] else "Bilinmiyor"
    df["kurum"] = raw_df[mapping["kurum"]].astype(str) if mapping["kurum"] else "Bilinmiyor"
    df["grup"] = raw_df[mapping["grup"]].astype(str) if mapping["grup"] else "Bilinmiyor"
    df["ciro"] = pd.to_numeric(raw_df[mapping["toplam"]], errors="coerce").fillna(0)
    df["odenen"] = pd.to_numeric(raw_df[mapping["odenen"]], errors="coerce").fillna(0) if mapping["odenen"] else df["ciro"]
    df["iskonto"] = pd.to_numeric(raw_df[mapping["iskonto"]], errors="coerce").fillna(0) if mapping["iskonto"] else 0
    df["elden_toplam"] = pd.to_numeric(raw_df[mapping["elde_toplam"]], errors="coerce").fillna(0) if mapping["elde_toplam"] else 0
    df["brut_kar"] = pd.to_numeric(raw_df[mapping["kar"]], errors="coerce").fillna(0)
    df["maliyet"] = pd.to_numeric(raw_df[mapping["maliyet"]], errors="coerce").fillna(0)
    df["fiyat_farki"] = pd.to_numeric(raw_df[mapping["fiyat_farki"]], errors="coerce").fillna(0) if mapping["fiyat_farki"] else 0
    if mapping["sonlandi"]:
        df["sonlandi"] = raw_df[mapping["sonlandi"]].astype(str).str.lower().isin(["true", "1", "evet", "yes"])
    else:
        df["sonlandi"] = True
    df["tarih"] = excel_serial_to_datetime(raw_df[mapping["islem_tarihi"]])
    df["recete_tarihi"] = excel_serial_to_datetime(raw_df[mapping["recete_tarihi"]]) if mapping["recete_tarihi"] else pd.NaT
    df["alım_tarihi"] = excel_serial_to_datetime(raw_df[mapping["alım_tarihi"]]) if mapping["alım_tarihi"] else pd.NaT
    df["kullanici"] = raw_df[mapping["kullanici"]].astype(str) if mapping["kullanici"] else ""
    df = df.dropna(subset=["tarih"])
    df["gun"] = df["tarih"].dt.date
    df["ay"] = df["tarih"].dt.to_period("M").astype(str)
    df["saat"] = df["tarih"].dt.hour
    day_map = {0: "Pazartesi", 1: "Salı", 2: "Çarşamba", 3: "Perşembe", 4: "Cuma", 5: "Cumartesi", 6: "Pazar"}
    df["hafta_gunu"] = df["tarih"].dt.dayofweek.map(day_map)
    df["hafta_gunu_no"] = df["tarih"].dt.dayofweek
    df["kar_marji"] = np.where(df["ciro"] > 0, df["brut_kar"] / df["ciro"], 0)
    df["tahsilat_acigi"] = (df["ciro"] - df["odenen"]).clip(lower=0)
    df["recete_gecikme_gun"] = (df["tarih"].dt.normalize() - df["recete_tarihi"].dt.normalize()).dt.days
    return df


def standardize_inventory(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw_df.columns)
    mapping = {
        "barkod": find_col(cols, ["Barkod"]),
        "urun": find_col(cols, ["Ürün Adı", "Urun Adi", "İlaç Adı", "Ilac Adi", "Malzeme Adı"]),
        "kademe": find_col(cols, ["Kademe"]),
        "psf": find_col(cols, ["Psf", "PSF", "Satış Fiyatı"]),
        "kamu": find_col(cols, ["Kamu"]),
        "stok": find_col(cols, ["Stok", "Mevcut Stok", "Kalan Stok"]),
        "kritik_stok": find_col(cols, ["Kritik Stok", "Kritik"]),
        "raf": find_col(cols, ["Raf", "Raf Lokasyonu"]),
        "kdv": find_col(cols, ["Kdv", "KDV"]),
        "psf_toplam": find_col(cols, ["Psf Toplam", "PSF Toplam"]),
        "kamu_toplam": find_col(cols, ["Kamu Toplam"]),
        "mal_haric": find_col(cols, ["Mal Top(Kdv Hariç)", "Mal Top Kdv Haric", "Mal Top(KDV Hariç)"]),
        "mal_dahil": find_col(cols, ["Mal Top(Kdv Dahil)", "Mal Top Kdv Dahil", "Mal Top(KDV Dahil)"]),
    }
    required = ["barkod", "urun", "stok"]
    missing = [k for k in required if mapping.get(k) is None]
    if missing:
        raise ValueError("Envanter Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["barkod"] = raw_df[mapping["barkod"]].astype(str)
    df["urun"] = raw_df[mapping["urun"]].astype(str)
    df["kademe"] = pd.to_numeric(raw_df[mapping["kademe"]], errors="coerce").fillna(0) if mapping["kademe"] else 0
    df["psf"] = pd.to_numeric(raw_df[mapping["psf"]], errors="coerce").fillna(0) if mapping["psf"] else 0
    df["kamu"] = pd.to_numeric(raw_df[mapping["kamu"]], errors="coerce").fillna(0) if mapping["kamu"] else 0
    df["stok"] = pd.to_numeric(raw_df[mapping["stok"]], errors="coerce").fillna(0)
    df["kritik_stok"] = pd.to_numeric(raw_df[mapping["kritik_stok"]], errors="coerce").fillna(0) if mapping["kritik_stok"] else 0
    df["raf"] = raw_df[mapping["raf"]].astype(str) if mapping["raf"] else "Bilinmiyor"
    df["kdv"] = pd.to_numeric(raw_df[mapping["kdv"]], errors="coerce").fillna(0) if mapping["kdv"] else 0
    df["psf_toplam"] = pd.to_numeric(raw_df[mapping["psf_toplam"]], errors="coerce").fillna(0) if mapping["psf_toplam"] else df["psf"] * df["stok"]
    df["kamu_toplam"] = pd.to_numeric(raw_df[mapping["kamu_toplam"]], errors="coerce").fillna(0) if mapping["kamu_toplam"] else df["kamu"] * df["stok"]
    df["mal_haric"] = pd.to_numeric(raw_df[mapping["mal_haric"]], errors="coerce").fillna(0) if mapping["mal_haric"] else 0
    df["mal_dahil"] = pd.to_numeric(raw_df[mapping["mal_dahil"]], errors="coerce").fillna(0) if mapping["mal_dahil"] else df["mal_haric"]
    df["stok_degeri"] = np.where(df["mal_dahil"] > 0, df["mal_dahil"], df["psf_toplam"])
    df["kritik_mi"] = np.where(df["kritik_stok"] > 0, df["stok"] <= df["kritik_stok"], df["stok"] <= 1)
    return df




def standardize_product_totals(raw_df: pd.DataFrame) -> pd.DataFrame:
    """TEBEOS Ürün Bazında Toplamlar raporunu standart kolonlara çevirir."""
    cols = list(raw_df.columns)
    mapping = {
        "barkod": find_col(cols, ["Barkod"]),
        "urun": find_col(cols, ["Ürün Adı", "Urun Adi", "Ürün Adı (İçinde Geçen İsim Şeklinde Arama Yapılabilir)"]),
        "stok": find_col(cols, ["Stok", "Mevcut Stok"]),
        "psf": find_col(cols, ["Psf", "PSF", "Satış Fiyatı"]),
        "alis_adet": find_col(cols, ["Alış Adet", "Alis Adet"]),
        "alis_maliyet": find_col(cols, ["Alış Maliyet Topl", "Alis Maliyet Topl", "Alış Maliyet Toplam"]),
        "satilan_adet": find_col(cols, ["Satılan Adet", "Satilan Adet", "Satış Adet", "Satis Adet"]),
        "satis_tutari": find_col(cols, ["Satış Tutarı", "Satis Tutari", "Ciro"]),
        "iade_adet": find_col(cols, ["İade Adet", "Iade Adet"]),
        "iade_tutari": find_col(cols, ["İade Tutarı", "Iade Tutari"]),
        "kar_tutari": find_col(cols, ["Kar Tutarı", "Kar Tutari", "Brüt Kar", "Brut Kar"]),
        "fark_toplami": find_col(cols, ["Fark Toplamı", "Fark Toplami"]),
        "urun_grubu": find_col(cols, ["İlaç Dışı Ürün Grubu", "Ilac Disi Urun Grubu", "Ürün Grubu", "Urun Grubu"]),
    }
    required = ["barkod", "urun", "satilan_adet", "satis_tutari", "kar_tutari"]
    missing = [k for k in required if mapping.get(k) is None]
    if missing:
        raise ValueError("Ürün Bazında Toplamlar Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["barkod"] = raw_df[mapping["barkod"]].astype(str) if mapping["barkod"] else ""
    df["urun"] = raw_df[mapping["urun"]].astype(str)
    df["stok"] = pd.to_numeric(raw_df[mapping["stok"]], errors="coerce").fillna(0) if mapping["stok"] else 0
    df["psf"] = pd.to_numeric(raw_df[mapping["psf"]], errors="coerce").fillna(0) if mapping["psf"] else 0
    df["alis_adet"] = pd.to_numeric(raw_df[mapping["alis_adet"]], errors="coerce").fillna(0) if mapping["alis_adet"] else 0
    df["alis_maliyet"] = pd.to_numeric(raw_df[mapping["alis_maliyet"]], errors="coerce").fillna(0) if mapping["alis_maliyet"] else 0
    df["satilan_adet"] = pd.to_numeric(raw_df[mapping["satilan_adet"]], errors="coerce").fillna(0)
    df["satis_tutari"] = pd.to_numeric(raw_df[mapping["satis_tutari"]], errors="coerce").fillna(0)
    df["iade_adet"] = pd.to_numeric(raw_df[mapping["iade_adet"]], errors="coerce").fillna(0) if mapping["iade_adet"] else 0
    df["iade_tutari"] = pd.to_numeric(raw_df[mapping["iade_tutari"]], errors="coerce").fillna(0) if mapping["iade_tutari"] else 0
    df["kar_tutari"] = pd.to_numeric(raw_df[mapping["kar_tutari"]], errors="coerce").fillna(0)
    df["fark_toplami"] = pd.to_numeric(raw_df[mapping["fark_toplami"]], errors="coerce").fillna(0) if mapping["fark_toplami"] else 0
    df["urun_grubu"] = raw_df[mapping["urun_grubu"]].astype(str) if mapping["urun_grubu"] else "Bilinmiyor"
    df["kar_marji"] = np.where(df["satis_tutari"] > 0, df["kar_tutari"] / df["satis_tutari"], 0)
    df["gunluk_satis_hizi"] = 0.0
    df["tahmini_bitis_gun"] = np.inf
    return df


def enrich_product_totals(product_df: pd.DataFrame, selected_days: int) -> pd.DataFrame:
    """Ürün bazında satış hızı, stok bitiş günü ve temel risk etiketleri üretir."""
    df = product_df.copy()
    days = max(1, int(selected_days or 1))
    df["gunluk_satis_hizi"] = df["satilan_adet"] / days
    df["tahmini_bitis_gun"] = np.where(df["gunluk_satis_hizi"] > 0, df["stok"] / df["gunluk_satis_hizi"], np.inf)
    df["stokta_yok_ama_satan"] = (df["stok"] <= 0) & (df["satilan_adet"] > 0)
    df["hizli_tukeniyor"] = (df["satilan_adet"] >= 5) & (df["stok"] > 0) & (df["tahmini_bitis_gun"] <= 7)
    df["olu_stok_aday"] = (df["stok"] > 0) & (df["satilan_adet"] <= 0)
    df["dusuk_karli_cok_satan"] = (df["satilan_adet"] >= df["satilan_adet"].quantile(0.75)) & (df["kar_marji"] < 0.08)
    df["urun_saglik_skoru"] = (
        np.clip(df["kar_marji"] / 0.20, 0, 1) * 35
        + np.clip(df["satilan_adet"] / max(1, df["satilan_adet"].quantile(0.90)), 0, 1) * 35
        + np.where(df["stokta_yok_ama_satan"], 0, np.where(df["hizli_tukeniyor"], 12, 30))
    ).round().astype(int)
    return df


def build_patient_loyalty(period_df: pd.DataFrame, all_sales_df: pd.DataFrame) -> dict:
    """Hasta sadakat analizleri. KVKK nedeniyle hasta adı varsayılan olarak maskelenebilir."""
    df = all_sales_df.copy()
    df["hasta_clean"] = df["hasta"].astype(str).str.strip()
    df = df[(df["hasta_clean"] != "") & (df["hasta_clean"].str.lower() != "nan")]
    if df.empty:
        empty = pd.DataFrame()
        return {"summary": {}, "frequency": empty, "basket": empty, "lost": empty, "vip": empty}

    max_dt = df["tarih"].max()
    last_12_start = max_dt - pd.Timedelta(days=365)
    last_90_start = max_dt - pd.Timedelta(days=90)
    last_12 = df[df["tarih"] >= last_12_start].copy()
    before_90 = last_12[last_12["tarih"] < last_90_start].copy()

    patient = last_12.groupby("hasta_clean", as_index=False).agg(
        ziyaret=("satis_no", "nunique"),
        ciro=("ciro", "sum"),
        kar=("brut_kar", "sum"),
        son_gelis=("tarih", "max"),
        ilk_gelis=("tarih", "min"),
    )
    patient["ortalama_sepet"] = np.where(patient["ziyaret"] > 0, patient["ciro"] / patient["ziyaret"], 0)
    patient["kar_marji"] = np.where(patient["ciro"] > 0, patient["kar"] / patient["ciro"], 0)
    patient["gelmeyen_gun"] = (max_dt - patient["son_gelis"]).dt.days

    previous_active = before_90.groupby("hasta_clean", as_index=False).agg(
        onceki_ziyaret=("satis_no", "nunique"),
        onceki_ciro=("ciro", "sum"),
        son_gelis_onceki=("tarih", "max"),
    )
    lost = previous_active.merge(patient[["hasta_clean", "son_gelis"]], on="hasta_clean", how="left")
    lost = lost[(lost["onceki_ziyaret"] >= 2) & ((lost["son_gelis"].isna()) | (lost["son_gelis"] < last_90_start))]
    lost["kayip_risk_notu"] = "Son 90 gündür gelmeyen düzenli müşteri"

    vip = patient[(patient["ziyaret"] >= 3) | (patient["ciro"] >= patient["ciro"].quantile(0.90))].copy()
    summary = {
        "aktif_hasta": int(patient["hasta_clean"].nunique()),
        "vip_hasta": int(len(vip)),
        "kayip_riski": int(len(lost)),
        "toplam_hasta_ciro": float(patient["ciro"].sum()),
        "toplam_hasta_kar": float(patient["kar"].sum()),
    }
    return {
        "summary": summary,
        "frequency": patient.sort_values(["ziyaret", "ciro"], ascending=False).head(50),
        "basket": patient.sort_values("ortalama_sepet", ascending=False).head(50),
        "lost": lost.sort_values(["onceki_ciro", "onceki_ziyaret"], ascending=False).head(100),
        "vip": vip.sort_values("ciro", ascending=False).head(100),
    }


def mask_patient_names(df: pd.DataFrame, name_col: str = "hasta_clean") -> pd.DataFrame:
    """Hasta adlarını güvenli ekranda Hasta-0001 formatında maskeler."""
    if df is None or df.empty or name_col not in df.columns:
        return df
    out = df.copy()
    unique_names = {name: f"Hasta-{i+1:04d}" for i, name in enumerate(out[name_col].dropna().astype(str).unique())}
    out[name_col] = out[name_col].map(unique_names).fillna("Hasta")
    return out


def build_doctor_intelligence(period_df: pd.DataFrame, all_sales_df: pd.DataFrame) -> dict:
    """Doktor bazlı ciro/kârlılık/trend analizi."""
    df = all_sales_df.copy()
    df["doktor_clean"] = df["doktor"].astype(str).str.strip()
    df = df[(df["doktor_clean"] != "") & (df["doktor_clean"].str.lower() != "nan") & (df["doktor_clean"].str.lower() != "bilinmiyor")]
    if df.empty:
        empty = pd.DataFrame()
        return {"doctor_kpi": empty, "doctor_month": empty, "doctor_growth": empty, "doctor_institution": empty}

    doctor_kpi = period_df.copy()
    doctor_kpi["doktor_clean"] = doctor_kpi["doktor"].astype(str).str.strip()
    doctor_kpi = doctor_kpi[(doctor_kpi["doktor_clean"] != "") & (doctor_kpi["doktor_clean"].str.lower() != "nan") & (doctor_kpi["doktor_clean"].str.lower() != "bilinmiyor")]
    doctor_kpi = doctor_kpi.groupby("doktor_clean", as_index=False).agg(
        ciro=("ciro", "sum"),
        kar=("brut_kar", "sum"),
        islem=("satis_no", "nunique"),
        hasta_sayisi=("hasta", pd.Series.nunique),
        tahsilat_acigi=("tahsilat_acigi", "sum"),
    ).sort_values("ciro", ascending=False)
    doctor_kpi["marj"] = np.where(doctor_kpi["ciro"] > 0, doctor_kpi["kar"] / doctor_kpi["ciro"], 0)
    doctor_kpi["ortalama_recete"] = np.where(doctor_kpi["islem"] > 0, doctor_kpi["ciro"] / doctor_kpi["islem"], 0)

    df["ay"] = df["tarih"].dt.to_period("M").astype(str)
    doctor_month = df.groupby(["doktor_clean", "ay"], as_index=False).agg(
        ciro=("ciro", "sum"),
        kar=("brut_kar", "sum"),
        islem=("satis_no", "nunique"),
    ).sort_values(["doktor_clean", "ay"])

    last_months = sorted(doctor_month["ay"].unique())[-3:]
    growth_base = doctor_month[doctor_month["ay"].isin(last_months)].copy()
    if len(last_months) >= 2:
        first_m, last_m = last_months[0], last_months[-1]
        first = growth_base[growth_base["ay"] == first_m][["doktor_clean", "ciro"]].rename(columns={"ciro": "ilk_ay_ciro"})
        last = growth_base[growth_base["ay"] == last_m][["doktor_clean", "ciro"]].rename(columns={"ciro": "son_ay_ciro"})
        doctor_growth = first.merge(last, on="doktor_clean", how="outer").fillna(0)
        doctor_growth["uc_ay_degisim"] = np.where(doctor_growth["ilk_ay_ciro"] > 0, (doctor_growth["son_ay_ciro"] - doctor_growth["ilk_ay_ciro"]) / doctor_growth["ilk_ay_ciro"], np.nan)
        doctor_growth = doctor_growth.sort_values("son_ay_ciro", ascending=False)
    else:
        doctor_growth = pd.DataFrame()

    doctor_institution = period_df.groupby(["doktor", "kurum"], as_index=False).agg(
        ciro=("ciro", "sum"),
        islem=("satis_no", "nunique"),
    ).sort_values(["doktor", "ciro"], ascending=[True, False])

    return {
        "doctor_kpi": doctor_kpi,
        "doctor_month": doctor_month,
        "doctor_growth": doctor_growth,
        "doctor_institution": doctor_institution,
    }

# ============================================================
# ANALİZ FONKSİYONLARI
# ============================================================
def summarize_sales(df: pd.DataFrame) -> dict:
    ciro = df["ciro"].sum()
    kar = df["brut_kar"].sum()
    maliyet = df["maliyet"].sum()
    islem = df["satis_no"].nunique()
    return {
        "ciro": ciro,
        "kar": kar,
        "maliyet": maliyet,
        "marj": safe_div(kar, ciro),
        "islem": islem,
        "ortalama_sepet": safe_div(ciro, islem),
        "recete_basi_kar": safe_div(kar, islem),
        "tahsilat_acigi": df["tahsilat_acigi"].sum(),
        "tahsilat_orani": safe_div(df["odenen"].sum(), ciro),
        "sonlanmamis": int((~df["sonlandi"]).sum()),
    }


def compare_text(cur, prev, positive_label="arttı", negative_label="azaldı"):
    if prev == 0 and cur > 0:
        return "yeni veri oluştu"
    if prev == 0:
        return "değişim hesaplanamadı"
    rate = (cur - prev) / prev
    return f"{pct_fmt(abs(rate))} {positive_label if rate >= 0 else negative_label}"


def score_from_threshold(value: float, good: float, warning: float, bad: float, higher_is_better: bool = True) -> int:
    """0-100 arası okunabilir skor üretir."""
    try:
        value = float(value)
    except Exception:
        return 0
    if higher_is_better:
        if value >= good:
            return 100
        if value >= warning:
            return 75
        if value >= bad:
            return 45
        return 15
    if value <= good:
        return 100
    if value <= warning:
        return 75
    if value <= bad:
        return 45
    return 15


def health_score_breakdown(sales_stats: dict, previous_stats: dict, inventory_df: pd.DataFrame, critical_df: pd.DataFrame, no_stock_df: pd.DataFrame, high_stock_df: pd.DataFrame) -> dict:
    inv_total = max(1, len(inventory_df))
    stock_value = float(inventory_df["stok_degeri"].sum())
    high_stock_value = float(high_stock_df["stok_degeri"].sum())
    ciro = float(sales_stats.get("ciro", 0))
    margin = float(sales_stats.get("marj", 0))
    collection_gap = float(sales_stats.get("tahsilat_acigi", 0))
    collection_gap_ratio = safe_div(collection_gap, ciro)
    critical_ratio = len(critical_df) / inv_total
    no_stock_ratio = len(no_stock_df) / inv_total
    high_stock_value_ratio = safe_div(high_stock_value, stock_value)
    stock_to_sales_ratio = safe_div(stock_value, ciro)
    growth_rate = safe_div(sales_stats.get("ciro", 0) - previous_stats.get("ciro", 0), previous_stats.get("ciro", 0)) if previous_stats.get("ciro", 0) else 0

    scores = {
        "Kârlılık": score_from_threshold(margin, good=0.20, warning=0.15, bad=0.10, higher_is_better=True),
        "Tahsilat": score_from_threshold(collection_gap_ratio, good=0.02, warning=0.05, bad=0.10, higher_is_better=False),
        "Stok Yönetimi": int(max(0, min(100, 100 - critical_ratio * 55 - no_stock_ratio * 90))),
        "Sermaye Verimliliği": int(max(0, min(100, 100 - high_stock_value_ratio * 85 - max(0, stock_to_sales_ratio - 2.0) * 12))),
        "Büyüme": score_from_threshold(growth_rate, good=0.08, warning=0.00, bad=-0.08, higher_is_better=True),
    }
    weights = {"Kârlılık": 20, "Tahsilat": 25, "Stok Yönetimi": 20, "Sermaye Verimliliği": 20, "Büyüme": 15}
    total = sum(scores[k] * weights[k] for k in scores) / sum(weights.values())
    return {
        "total": int(round(max(0, min(100, total)))),
        "scores": scores,
        "weights": weights,
        "collection_gap_ratio": collection_gap_ratio,
        "critical_ratio": critical_ratio,
        "no_stock_ratio": no_stock_ratio,
        "high_stock_value_ratio": high_stock_value_ratio,
        "stock_to_sales_ratio": stock_to_sales_ratio,
        "growth_rate": growth_rate,
    }


def ayca_score(sales_stats: dict, previous_stats: dict, inventory_df: pd.DataFrame, critical_df: pd.DataFrame, no_stock_df: pd.DataFrame, high_stock_df: pd.DataFrame) -> int:
    return health_score_breakdown(sales_stats, previous_stats, inventory_df, critical_df, no_stock_df, high_stock_df)["total"]


def score_status(score_value: int) -> str:
    if score_value >= 82:
        return "Güçlü"
    if score_value >= 65:
        return "Kontrollü"
    if score_value >= 50:
        return "Takip edilmeli"
    return "Riskli"


def make_pareto_inventory(inventory_df: pd.DataFrame) -> pd.DataFrame:
    pareto = inventory_df.sort_values("stok_degeri", ascending=False).copy()
    total = max(1, pareto["stok_degeri"].sum())
    pareto["kumulatif_deger"] = pareto["stok_degeri"].cumsum()
    pareto["kumulatif_oran"] = pareto["kumulatif_deger"] / total
    pareto["pareto_sinif"] = np.where(pareto["kumulatif_oran"] <= 0.80, "A - Sermaye Yoğun", "B/C - Diğer")
    return pareto


def detect_daily_anomalies(daily_df: pd.DataFrame) -> pd.DataFrame:
    if len(daily_df) < 5:
        daily_df["anomali"] = "Normal"
        return daily_df
    mean = daily_df["ciro"].mean()
    std = daily_df["ciro"].std(ddof=0)
    if std == 0 or pd.isna(std):
        daily_df["anomali"] = "Normal"
        return daily_df
    daily_df["z_skor"] = (daily_df["ciro"] - mean) / std
    daily_df["anomali"] = np.select(
        [daily_df["z_skor"] >= 1.5, daily_df["z_skor"] <= -1.5],
        ["Yüksek gün", "Düşük gün"],
        default="Normal",
    )
    return daily_df


def create_ai_comment(current_stats, previous_stats, critical_df, no_stock_df, high_stock_df, best_day, worst_day, top_kurum, top_doctor):
    messages = []
    messages.append(f"Ciro önceki eş döneme göre {compare_text(current_stats['ciro'], previous_stats['ciro'])}.")
    messages.append(f"Brüt kâr {compare_text(current_stats['kar'], previous_stats['kar'])}; güncel marj {pct_fmt(current_stats['marj'])}.")
    messages.append(f"Tahsilat oranı {pct_fmt(current_stats['tahsilat_orani'])}; tahsilat açığı {money_fmt(current_stats['tahsilat_acigi'])}.")
    if top_kurum:
        messages.append(f"En yüksek ciro oluşturan kurum: {top_kurum}.")
    if top_doctor:
        messages.append(f"En çok ciro getiren doktor kaydı: {top_doctor}.")
    if best_day and worst_day:
        messages.append(f"Haftanın en güçlü günü {best_day}; en zayıf günü {worst_day} görünüyor.")
    if not critical_df.empty:
        messages.append(f"Kritik stokta {len(critical_df)} ürün var.")
    if not no_stock_df.empty:
        messages.append(f"Stokta olmayan {len(no_stock_df)} ürün satış kaçırma riski oluşturabilir.")
    if not high_stock_df.empty:
        messages.append(f"Yüksek stok grubunda {money_fmt(high_stock_df['stok_degeri'].sum())} sermaye bağlı.")
    return " ".join(messages)


def create_action_items(current_stats, previous_stats, critical_df, no_stock_df, high_stock_df, pareto_df, daily_anomaly_df):
    actions = []
    if current_stats["marj"] < 0.12:
        actions.append("💰 Brüt kâr marjı düşük. İskonto, maliyet ve fiyat farkı kayıtlarını kontrol et.")
    if current_stats["tahsilat_acigi"] > 0:
        actions.append(f"🧾 Tahsilat açığı {money_fmt(current_stats['tahsilat_acigi'])}. Ödenen tutar ve toplam tutar farklarını incele.")
    if current_stats["sonlanmamis"] > 0:
        actions.append(f"⏳ Sonlanmamış {current_stats['sonlanmamis']} satış kaydı var. Gün sonu kontrolü yap.")
    if not no_stock_df.empty:
        actions.append(f"⛔ Stokta olmayan {len(no_stock_df)} ürün var. En sık sorulan ürünleri manuel önceliklendir.")
    if not critical_df.empty:
        row = critical_df.sort_values(["stok", "stok_degeri"], ascending=[True, False]).iloc[0]
        actions.append(f"🔴 Kritik stok listesinin başında {row['urun']} var. Mevcut stok: {num_fmt(row['stok'], 0)}.")
    if not high_stock_df.empty:
        row = high_stock_df.sort_values("stok_degeri", ascending=False).iloc[0]
        actions.append(f"📦 En çok sermaye bağlayan yüksek stok ürünü: {row['urun']} ({money_fmt(row['stok_degeri'])}).")
    pareto_a_count = int((pareto_df["pareto_sinif"] == "A - Sermaye Yoğun").sum()) if not pareto_df.empty else 0
    if pareto_a_count:
        actions.append(f"🎯 Stok sermayesinin yaklaşık %80'i {pareto_a_count} üründe toplanıyor. Bu ürünleri ayrı takip et.")
    low_days = daily_anomaly_df[daily_anomaly_df.get("anomali", "") == "Düşük gün"]
    if not low_days.empty:
        d = low_days.sort_values("ciro").iloc[0]
        actions.append(f"📉 {d['gun']} düşük ciro günü olarak ayrışıyor. Nöbet, hava, kampanya veya işlem kapanış etkisi olabilir.")
    if not actions:
        actions.append("✅ Kritik aksiyon görünmüyor. Günlük ciro, tahsilat ve stok sermayesini takip et.")
    return actions[:7]


def create_excel_report(sales_df, inv_df, product_df, period_df, critical_df, no_stock_df, high_stock_df, pareto_df, kurum_df, doktor_df, weekday_df, hourly_df, doctor_kpi=None, patient_frequency=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        period_df.to_excel(writer, sheet_name="Secili_Donem_Satis", index=False)
        sales_df.to_excel(writer, sheet_name="Satis_Standart", index=False)
        inv_df.to_excel(writer, sheet_name="Envanter_Standart", index=False)
        product_df.to_excel(writer, sheet_name="Urun_Bazinda_Toplamlar", index=False)
        critical_df.to_excel(writer, sheet_name="Kritik_Stok", index=False)
        no_stock_df.to_excel(writer, sheet_name="Stok_Yok", index=False)
        high_stock_df.to_excel(writer, sheet_name="Yuksek_Stok", index=False)
        pareto_df.to_excel(writer, sheet_name="Pareto_Stok", index=False)
        kurum_df.to_excel(writer, sheet_name="Kurum_Analizi", index=False)
        doktor_df.to_excel(writer, sheet_name="Doktor_Analizi", index=False)
        weekday_df.to_excel(writer, sheet_name="Hafta_Gunu", index=False)
        hourly_df.to_excel(writer, sheet_name="Saatlik", index=False)
        if doctor_kpi is not None and not doctor_kpi.empty:
            doctor_kpi.to_excel(writer, sheet_name="Doktor_Intelligence", index=False)
        if patient_frequency is not None and not patient_frequency.empty:
            patient_frequency.to_excel(writer, sheet_name="Hasta_Sadakat", index=False)
    return output.getvalue()


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.success(f"Giriş: {st.session_state.get('auth_user', 'Demo Kullanıcı')} · {get_membership()}")
if st.sidebar.button("Çıkış Yap", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state.pop("auth_user", None)
    st.session_state.pop("membership", None)
    safe_rerun()

st.sidebar.title("💊 AYÇA Insight")
st.sidebar.caption("V7.1 · Ciro + Stok + Ürün + Doktor + Hasta")
eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")

sales_file = st.sidebar.file_uploader("1) Satış Excel dosyasını yükle", type=["xlsx", "xls"], key="sales_file")
inventory_file = st.sidebar.file_uploader("2) Envanter Excel dosyasını yükle", type=["xlsx", "xls"], key="inventory_file")
product_file = st.sidebar.file_uploader("3) Ürün Bazında Toplamlar Excelini yükle", type=["xlsx", "xls"], key="product_file")

st.sidebar.markdown("---")
selected_period = st.sidebar.selectbox("Satış Dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Tüm veri"], index=2)
high_stock_limit = st.sidebar.slider("Yüksek stok adedi eşiği", 20, 500, 100)
min_stock_value_filter = st.sidebar.number_input("Yüksek stok için minimum stok değeri", min_value=0, value=0, step=1000)
show_patient_columns = st.sidebar.checkbox("Hasta isim kolonunu göster", value=False)
mask_patient_display = st.sidebar.checkbox("Hasta adlarını maskele", value=True)
st.sidebar.caption("Hasta TC ve kişisel sağlık verisi analiz dışı tutulmalıdır. Hasta adı varsayılan olarak gizlidir.")


# ============================================================
# DOSYA KONTROL
# ============================================================
if sales_file is None or inventory_file is None or product_file is None:
    st.markdown(
        f"""
        <div class="ayca-header">
            <div class="ayca-title">
                <h1>AYÇA Insight V7.1</h1>
                <p>{eczane_adi} · Satış, envanter ve ürün bazında toplamlar dosyalarını ayrı ayrı yükle.</p>
            </div>
            <div class="header-pill">Dosya bekleniyor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        make_mini_card("1. Satış Exceli", "Zorunlu", "Ciro, kâr, tahsilat, kurum, doktor, işlem tarihi", "alert-blue")
    with c2:
        make_mini_card("2. Envanter Exceli", "Zorunlu", "Stok, kritik stok, raf, stok değeri", "alert-green")
    with c3:
        make_mini_card("3. Ürün Bazında Toplamlar", "Zorunlu", "Satılan adet, ürün kârı, stok bitiş ve ürün sağlık analizi", "alert-orange")
    st.info("Sol menüden üç dosyayı da yüklediğinde dashboard otomatik çalışır.")
    st.stop()

try:
    raw_sales, sales_sheet, sales_sheets = read_excel_first_sheet(sales_file)
    raw_inventory, inv_sheet, inv_sheets = read_excel_first_sheet(inventory_file)
    raw_product, product_sheet, product_sheets = read_excel_first_sheet(product_file)
    sales_df = standardize_sales(raw_sales)
    inventory_df = standardize_inventory(raw_inventory)
    product_df = standardize_product_totals(raw_product)
except Exception as exc:
    st.error(f"Dosya okunurken hata oluştu: {exc}")
    st.stop()

inventory_df["fazla_stok_mu"] = (inventory_df["stok"] >= high_stock_limit) & (inventory_df["stok_degeri"] >= min_stock_value_filter)
inventory_df["stok_durumu"] = np.select(
    [inventory_df["stok"] <= 0, inventory_df["kritik_mi"], inventory_df["fazla_stok_mu"]],
    ["Stok Yok", "Kritik", "Yüksek Stok"],
    default="Normal",
)


# ============================================================
# DÖNEM HESABI
# ============================================================
max_date = sales_df["tarih"].max()
if selected_period == "Son 7 gün":
    start_date = max_date - pd.Timedelta(days=7)
elif selected_period == "Son 14 gün":
    start_date = max_date - pd.Timedelta(days=14)
elif selected_period == "Son 30 gün":
    start_date = max_date - pd.Timedelta(days=30)
else:
    start_date = sales_df["tarih"].min()

period_df = sales_df[sales_df["tarih"] >= start_date].copy()
period_days = max(1, (max_date - start_date).days)
product_df = enrich_product_totals(product_df, period_days)
prev_start = start_date - pd.Timedelta(days=period_days)
prev_df = sales_df[(sales_df["tarih"] >= prev_start) & (sales_df["tarih"] < start_date)].copy()

current_stats = summarize_sales(period_df)
previous_stats = summarize_sales(prev_df)

critical_df = inventory_df[inventory_df["kritik_mi"] & (inventory_df["stok"] > 0)].copy()
no_stock_df = inventory_df[inventory_df["stok"] <= 0].copy()
high_stock_df = inventory_df[inventory_df["fazla_stok_mu"]].copy()
pareto_df = make_pareto_inventory(inventory_df)
health = health_score_breakdown(current_stats, previous_stats, inventory_df, critical_df, no_stock_df, high_stock_df)
score = health["total"]

# Ana özet tablolar
kurum_df = period_df.groupby("kurum", as_index=False).agg(
    ciro=("ciro", "sum"),
    kar=("brut_kar", "sum"),
    maliyet=("maliyet", "sum"),
    islem=("satis_no", "nunique"),
    tahsilat_acigi=("tahsilat_acigi", "sum"),
).sort_values("ciro", ascending=False)
kurum_df["marj"] = np.where(kurum_df["ciro"] > 0, kurum_df["kar"] / kurum_df["ciro"], 0)
kurum_df["ortalama_sepet"] = np.where(kurum_df["islem"] > 0, kurum_df["ciro"] / kurum_df["islem"], 0)

doktor_df = period_df.groupby("doktor", as_index=False).agg(
    ciro=("ciro", "sum"),
    kar=("brut_kar", "sum"),
    islem=("satis_no", "nunique"),
).sort_values("ciro", ascending=False)
doktor_df["marj"] = np.where(doktor_df["ciro"] > 0, doktor_df["kar"] / doktor_df["ciro"], 0)
doktor_df["ortalama_sepet"] = np.where(doktor_df["islem"] > 0, doktor_df["ciro"] / doktor_df["islem"], 0)

weekday_df = period_df.groupby(["hafta_gunu_no", "hafta_gunu"], as_index=False).agg(
    ciro=("ciro", "sum"),
    kar=("brut_kar", "sum"),
    islem=("satis_no", "nunique"),
    gun_sayisi=("gun", "nunique"),
).sort_values("hafta_gunu_no")
weekday_df["gunluk_ortalama_ciro"] = np.where(weekday_df["gun_sayisi"] > 0, weekday_df["ciro"] / weekday_df["gun_sayisi"], 0)
weekday_df["marj"] = np.where(weekday_df["ciro"] > 0, weekday_df["kar"] / weekday_df["ciro"], 0)

hourly_df = period_df.groupby("saat", as_index=False).agg(
    ciro=("ciro", "sum"),
    kar=("brut_kar", "sum"),
    islem=("satis_no", "nunique"),
).sort_values("saat")

daily_df = period_df.groupby("gun", as_index=False).agg(
    ciro=("ciro", "sum"),
    kar=("brut_kar", "sum"),
    islem=("satis_no", "nunique"),
    tahsilat_acigi=("tahsilat_acigi", "sum"),
)
daily_df["marj"] = np.where(daily_df["ciro"] > 0, daily_df["kar"] / daily_df["ciro"], 0)
daily_df = detect_daily_anomalies(daily_df)

best_day = None
worst_day = None
if not weekday_df.empty:
    best_day = weekday_df.sort_values("gunluk_ortalama_ciro", ascending=False).iloc[0]["hafta_gunu"]
    worst_day = weekday_df.sort_values("gunluk_ortalama_ciro", ascending=True).iloc[0]["hafta_gunu"]

top_kurum = kurum_df.iloc[0]["kurum"] if not kurum_df.empty else ""
top_doctor = doktor_df.iloc[0]["doktor"] if not doktor_df.empty else ""
ai_comment = create_ai_comment(current_stats, previous_stats, critical_df, no_stock_df, high_stock_df, best_day, worst_day, top_kurum, top_doctor)
actions = create_action_items(current_stats, previous_stats, critical_df, no_stock_df, high_stock_df, pareto_df, daily_df)
patient_loyalty = build_patient_loyalty(period_df, sales_df)
doctor_intel = build_doctor_intelligence(period_df, sales_df)
fast_depleting_df = product_df[product_df["hizli_tukeniyor"]].sort_values(["tahmini_bitis_gun", "satilan_adet"], ascending=[True, False])
out_of_stock_selling_df = product_df[product_df["stokta_yok_ama_satan"]].sort_values("satis_tutari", ascending=False)
dead_stock_df = product_df[product_df["olu_stok_aday"]].sort_values("stok", ascending=False)
most_profitable_products_df = product_df.sort_values("kar_tutari", ascending=False).head(100)


# ============================================================
# HEADER + KPI
# ============================================================
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-header">
        <div class="ayca-title">
            <h1>AYÇA Insight V7.1</h1>
            <p>{eczane_adi} · {selected_period} · Satış: {sales_sheet} · Envanter: {inv_sheet} · Ürün: {product_sheet} · {today_str}</p>
        </div>
        <div class="header-pill">AYÇA Sağlık Puanı: {score}/100 · {score_status(score)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

ciro_trend, ciro_class = rate_fmt(current_stats["ciro"], previous_stats["ciro"])
profit_trend, profit_class = rate_fmt(current_stats["kar"], previous_stats["kar"])
margin_trend, margin_class = rate_fmt(current_stats["marj"], previous_stats["marj"])
collection_trend, collection_class = rate_fmt(current_stats["tahsilat_acigi"], previous_stats["tahsilat_acigi"])
stock_value_total = inventory_df["stok_degeri"].sum()

k1, k2, k3, k4, k5 = st.columns(5)
with k1: make_metric_card("Güncel Ciro", money_fmt(current_stats["ciro"]), selected_period, ciro_trend, ciro_class)
with k2: make_metric_card("Brüt Kâr", money_fmt(current_stats["kar"]), "Satış Exceli", profit_trend, profit_class)
with k3: make_metric_card("Kâr Marjı", pct_fmt(current_stats["marj"]), "Brüt kâr / ciro", margin_trend, margin_class)
with k4: make_metric_card("Toplam Stok Değeri", money_fmt(stock_value_total), "İçerideki satılabilir mal değeri")
with k5: make_metric_card("Tahsilat Açığı", money_fmt(current_stats["tahsilat_acigi"]), f"Cironun {pct_fmt(health['collection_gap_ratio'])}", collection_trend, "metric-down" if current_stats["tahsilat_acigi"] > previous_stats["tahsilat_acigi"] else "metric-up")


# ============================================================
# EXECUTIVE ÖZET
# ============================================================
stock_value = inventory_df["stok_degeri"].sum()
high_stock_value = high_stock_df["stok_degeri"].sum()
pareto_a_count = int((pareto_df["pareto_sinif"] == "A - Sermaye Yoğun").sum()) if not pareto_df.empty else 0
health_items = health["scores"]
health_weights = health["weights"]
health_html = "".join([
    f'<div class="health-row"><div class="health-head"><span>{k} <small>({health_weights[k]}%)</small></span><span>{v}/100</span></div><div class="health-bar-bg"><div class="health-bar-fill" style="width:{v}%;"></div></div></div>'
    for k, v in health_items.items()
])
health_explain_html = f"""
<div class="exec-list-item">🧾 Tahsilat açığı oranı: <b>{pct_fmt(health['collection_gap_ratio'])}</b> · Tahsilat skoru bu orana göre hesaplanır.</div>
<div class="exec-list-item">📦 Toplam stok değeri: <b>{money_fmt(stock_value)}</b> · Stok/ciro oranı: <b>{num_fmt(health['stock_to_sales_ratio'], 2)}x</b></div>
<div class="exec-list-item">💸 Yüksek stokta bağlı sermaye: <b>{money_fmt(high_stock_value)}</b> · Toplam stokun <b>{pct_fmt(health['high_stock_value_ratio'])}</b></div>
"""

action_html = "".join([f"<div class='exec-list-item'>{item}</div>" for item in actions])

st.markdown(
    f"""
    <div class="exec-grid">
        <div class="exec-card">
            <div class="exec-title">🤖 Günaydın {kullanici_adi}</div>
            <div class="exec-sub">AYÇA, ürün satış raporu olmadan mevcut satış ve envanter dosyalarından güvenilir yönetim özeti çıkardı.</div>
            <div class="exec-list-item">{ai_comment}</div>
            {action_html}
        </div>
        <div class="exec-card">
            <div class="exec-sub">Eczane Sağlık Puanı</div>
            <div class="score-big">{score}</div>
            <div class="exec-sub">Durum: <b>{score_status(score)}</b></div>
            {health_html}
            {health_explain_html}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

r1, r2, r3, r4, r5 = st.columns(5)
with r1: make_mini_card("Toplam Stok Değeri", money_fmt(stock_value), "İçerideki satılabilir mal", "alert-blue")
with r2: make_mini_card("Tahsilat Açığı", money_fmt(current_stats["tahsilat_acigi"]), f"Cironun {pct_fmt(health['collection_gap_ratio'])}", "alert-red" if current_stats["tahsilat_acigi"] > 0 else "alert-green")
with r3: make_mini_card("Kritik / Yok", f"{len(critical_df)} / {len(no_stock_df)}", "Kritik stok ve stok yok", "alert-red" if len(critical_df) + len(no_stock_df) else "alert-green")
with r4: make_mini_card("Yüksek Stok", str(len(high_stock_df)), f"{money_fmt(high_stock_value)} bağlı para", "alert-orange" if len(high_stock_df) else "alert-green")
with r5: make_mini_card("Stok / Ciro", f"{num_fmt(health['stock_to_sales_ratio'], 2)}x", "Stok değeri / seçili dönem ciro", "alert-purple")


# ============================================================
# SAYFALAR
# ============================================================
pages = ["🏠 Sabah Ekranı", "📈 Ciro Analizi", "👨‍⚕️ Doktor Intelligence", "🧑‍🤝‍🧑 Hasta Sadakat", "🏥 Kurum & Doktor", "💰 Kârlılık", "📦 Stok Sermayesi", "🔴 Risk Listeleri", "📥 Rapor"]
page = st.radio("Bölüm", pages, horizontal=True, label_visibility="collapsed")

if page == "🏠 Sabah Ekranı":
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Günlük Ciro ve Kâr</div>', unsafe_allow_html=True)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=daily_df["gun"], y=daily_df["ciro"], mode="lines+markers", name="Ciro"))
        fig.add_trace(go.Scatter(x=daily_df["gun"], y=daily_df["kar"], mode="lines+markers", name="Brüt Kâr"))
        fig.update_layout(title="Günlük Ciro / Kâr")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        st.markdown('<div class="section-title">Tahsilat Dağılımı</div>', unsafe_allow_html=True)
        tah = period_df.groupby("tahsilat", as_index=False)["ciro"].sum().sort_values("ciro", ascending=False).head(10)
        fig = px.pie(tah, names="tahsilat", values="ciro", title="Tahsilata Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    st.markdown('<div class="section-title">Bugünün Aksiyon Listesi</div>', unsafe_allow_html=True)
    for item in actions:
        st.markdown(f"<div class='exec-list-item'>{item}</div>", unsafe_allow_html=True)

    st.markdown(
        """
        <div class="ai-card">
            <div class="ai-title">Bilinçli Veri Sınırı</div>
            <div class="ai-text">
            Bu sürümde ürün satış raporu olmadığı için ürün bazlı satış hızı, otomatik sipariş önerisi, gerçek ölü stok ve miad riski hesaplanmaz.
            Bu analizleri gerçekmiş gibi üretmek yerine, mevcut satış özeti ve envanter verisinden güvenilir finans/stok sermayesi yorumu yapılır.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

elif page == "📈 Ciro Analizi":
    st.markdown('<div class="section-title">Satış Ritmi</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(weekday_df, x="hafta_gunu", y="gunluk_ortalama_ciro", title="Hafta Gününe Göre Ortalama Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        fig = px.bar(hourly_df, x="saat", y="ciro", title="Saatlere Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        sales_type = period_df.groupby("satis_tipi", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique")).sort_values("ciro", ascending=False)
        fig = px.bar(sales_type, x="satis_tipi", y="ciro", title="Satış Tipine Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c4:
        fig = px.scatter(daily_df, x="ciro", y="kar", size="islem", hover_name="gun", title="Gün Bazlı Ciro / Kâr İlişkisi")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    st.markdown('<div class="section-title">Günlük Detay</div>', unsafe_allow_html=True)
    st.dataframe(daily_df.sort_values("gun", ascending=False), use_container_width=True, hide_index=True)


elif page == "👨‍⚕️ Doktor Intelligence":
    st.markdown('<div class="section-title">👨‍⚕️ Doktor Intelligence</div>', unsafe_allow_html=True)
    doctor_kpi = doctor_intel["doctor_kpi"]
    doctor_month = doctor_intel["doctor_month"]
    doctor_growth = doctor_intel["doctor_growth"]
    doctor_institution = doctor_intel["doctor_institution"]

    if doctor_kpi.empty:
        st.info("Doktor analizi için satış dosyasında doktor bilgisi bulunamadı.")
    else:
        d1, d2, d3, d4 = st.columns(4)
        with d1: make_mini_card("Doktor Sayısı", str(doctor_kpi["doktor_clean"].nunique()), "Seçili dönem", "alert-blue")
        with d2: make_mini_card("İlk 10 Doktor Ciro", money_fmt(doctor_kpi.head(10)["ciro"].sum()), "Ciro liderleri", "alert-green")
        with d3: make_mini_card("En Karlı Doktor", doctor_kpi.sort_values("kar", ascending=False).iloc[0]["doktor_clean"], money_fmt(doctor_kpi.sort_values("kar", ascending=False).iloc[0]["kar"]), "alert-purple")
        with d4: make_mini_card("En Yüksek Ortalama Reçete", doctor_kpi.sort_values("ortalama_recete", ascending=False).iloc[0]["doktor_clean"], money_fmt(doctor_kpi.sort_values("ortalama_recete", ascending=False).iloc[0]["ortalama_recete"]), "alert-orange")

        c1, c2 = st.columns(2)
        with c1:
            top_doc = doctor_kpi.head(15).copy()
            fig = px.bar(top_doc, x="ciro", y="doktor_clean", orientation="h", title="Doktor Bazlı Ciro")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)
        with c2:
            top_profit_doc = doctor_kpi.sort_values("kar", ascending=False).head(15).copy()
            fig = px.bar(top_profit_doc, x="kar", y="doktor_clean", orientation="h", title="Doktor Bazlı Brüt Kâr")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)

        st.markdown('<div class="section-title">Son 3 Ay Doktor Trendleri</div>', unsafe_allow_html=True)
        if not doctor_growth.empty:
            st.dataframe(doctor_growth.head(50), use_container_width=True, hide_index=True)
        else:
            st.info("Trend için en az iki farklı ay verisi gerekir.")

        st.markdown('<div class="section-title">Doktor Detay Kartı</div>', unsafe_allow_html=True)
        selected_doctor = st.selectbox("Doktor seç", doctor_kpi["doktor_clean"].tolist())
        selected_doc_month = doctor_month[doctor_month["doktor_clean"] == selected_doctor].copy()
        selected_doc_inst = doctor_institution[doctor_institution["doktor"] == selected_doctor].head(10).copy()
        c3, c4 = st.columns(2)
        with c3:
            if not selected_doc_month.empty:
                fig = px.line(selected_doc_month, x="ay", y="ciro", markers=True, title=f"{selected_doctor} - Aylık Ciro")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
        with c4:
            if not selected_doc_inst.empty:
                fig = px.bar(selected_doc_inst, x="ciro", y="kurum", orientation="h", title=f"{selected_doctor} - Kurum Dağılımı")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

        st.markdown(
            """
            <div class="ai-card">
                <div class="ai-title">İlaç Eğilim Analizi İçin Veri Notu</div>
                <div class="ai-text">
                Mevcut ürün bazında toplamlar raporu ürünleri genel toplam olarak verir; doktor veya reçete numarası içermez.
                Bu nedenle <b>hangi doktor hangi ilacı yazıyor</b> analizi tam güvenilir şekilde ancak reçete kalemleri / ürün satış detay raporu doktor veya reçete no ile geldiğinde açılmalıdır.
                Şimdilik doktor bazlı ciro, kârlılık, hasta sayısı, kurum dağılımı ve trend analizi güvenilir şekilde çalışır.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.dataframe(doctor_kpi, use_container_width=True, hide_index=True)

elif page == "🧑‍🤝‍🧑 Hasta Sadakat":
    st.markdown('<div class="section-title">🧑‍🤝‍🧑 Hasta Sadakat Merkezi</div>', unsafe_allow_html=True)
    summary = patient_loyalty["summary"]
    if not summary:
        st.info("Hasta sadakat analizi için satış dosyasında hasta adı bilgisi bulunamadı.")
    else:
        h1, h2, h3, h4 = st.columns(4)
        with h1: make_mini_card("Aktif Hasta", str(summary["aktif_hasta"]), "Son 12 ay verisi", "alert-blue")
        with h2: make_mini_card("VIP Segment", str(summary["vip_hasta"]), "3+ ziyaret veya yüksek ciro", "alert-green")
        with h3: make_mini_card("Kaybedilen Müşteri Riski", str(summary["kayip_riski"]), "Son 90 gündür gelmeyenler", "alert-red" if summary["kayip_riski"] else "alert-green")
        with h4: make_mini_card("Hasta Kaynaklı Ciro", money_fmt(summary["toplam_hasta_ciro"]), "Eşleşen hasta kayıtları", "alert-purple")

        freq_df = patient_loyalty["frequency"]
        basket_df = patient_loyalty["basket"]
        lost_df = patient_loyalty["lost"]
        vip_df = patient_loyalty["vip"]
        if mask_patient_display:
            freq_df = mask_patient_names(freq_df)
            basket_df = mask_patient_names(basket_df)
            lost_df = mask_patient_names(lost_df)
            vip_df = mask_patient_names(vip_df)

        t1, t2, t3, t4 = st.tabs(["En Sık Gelen 50", "En Yüksek Sepet", "Kaybedilen Müşteri", "VIP Segment"])
        with t1:
            st.dataframe(freq_df, use_container_width=True, hide_index=True)
        with t2:
            st.dataframe(basket_df, use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(lost_df, use_container_width=True, hide_index=True)
        with t4:
            st.dataframe(vip_df, use_container_width=True, hide_index=True)

        st.markdown(
            """
            <div class="ai-card">
                <div class="ai-title">KVKK Güvenli Kullanım Notu</div>
                <div class="ai-text">
                Bu ekran ticari karar desteği için tasarlanmıştır. Hasta isimleri varsayılan olarak maskelenir.
                Hasta TC, açık sağlık tanısı ve hassas sağlık verisi analiz ekranına alınmamalıdır.
                Hedef, kişiyi ifşa etmek değil; sadakat, sepet ve kayıp müşteri riskini eczane seviyesinde anlamaktır.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

elif page == "🏥 Kurum & Doktor":
    st.markdown('<div class="section-title">Kurum ve Doktor Performansı</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        top_kurum_df = kurum_df.head(12).copy()
        fig = px.bar(top_kurum_df, x="ciro", y="kurum", orientation="h", title="İlk 12 Kurum - Ciro")
        st.plotly_chart(apply_plot_theme(fig, height=460), use_container_width=True)
    with c2:
        top_doc_df = doktor_df[doktor_df["doktor"].str.lower() != "nan"].head(12).copy()
        fig = px.bar(top_doc_df, x="ciro", y="doktor", orientation="h", title="İlk 12 Doktor - Ciro")
        st.plotly_chart(apply_plot_theme(fig, height=460), use_container_width=True)

    t1, t2 = st.tabs(["Kurum Detayı", "Doktor Detayı"])
    with t1:
        st.dataframe(kurum_df, use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(doktor_df, use_container_width=True, hide_index=True)

elif page == "💰 Kârlılık":
    st.markdown('<div class="section-title">Kârlılık ve Tahsilat Kontrolü</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(daily_df, x="gun", y="marj", markers=True, title="Günlük Brüt Marj")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        profit_by_type = period_df.groupby("satis_tipi", as_index=False).agg(kar=("brut_kar", "sum"), ciro=("ciro", "sum"), maliyet=("maliyet", "sum"))
        profit_by_type["marj"] = np.where(profit_by_type["ciro"] > 0, profit_by_type["kar"] / profit_by_type["ciro"], 0)
        fig = px.bar(profit_by_type, x="satis_tipi", y="kar", title="Satış Tipine Göre Brüt Kâr")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    c3, c4, c5 = st.columns(3)
    with c3: make_mini_card("Tahsilat Oranı", pct_fmt(current_stats["tahsilat_orani"]), "Ödenen / toplam ciro", "alert-green" if current_stats["tahsilat_orani"] > 0.95 else "alert-orange")
    with c4: make_mini_card("İskonto Toplamı", money_fmt(period_df["iskonto"].sum()), "Seçili dönem", "alert-blue")
    with c5: make_mini_card("Fiyat Farkı", money_fmt(period_df["fiyat_farki"].sum()), "Seçili dönem", "alert-purple")

    st.dataframe(profit_by_type.sort_values("kar", ascending=False), use_container_width=True, hide_index=True)

elif page == "📦 Stok Sermayesi":
    st.markdown('<div class="section-title">Stok Sermayesi ve Pareto Analizi</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        status = inventory_df.groupby("stok_durumu", as_index=False).agg(urun_sayisi=("urun", "count"), stok_degeri=("stok_degeri", "sum"))
        fig = px.bar(status, x="stok_durumu", y="stok_degeri", title="Stok Durumuna Göre Sermaye")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        pareto_plot = pareto_df.head(80).copy()
        fig = px.line(pareto_plot, x=np.arange(1, len(pareto_plot) + 1), y="kumulatif_oran", title="Stok Sermayesi Pareto Eğrisi")
        fig.update_xaxes(title="Ürün sırası")
        fig.update_yaxes(title="Kümülatif oran", tickformat=".0%")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        top_stock = inventory_df.sort_values("stok_degeri", ascending=False).head(15)
        fig = px.bar(top_stock, x="stok_degeri", y="urun", orientation="h", title="Stok Değeri En Yüksek 15 Ürün")
        st.plotly_chart(apply_plot_theme(fig, height=500), use_container_width=True)
    with c4:
        kdv_df = inventory_df.groupby("kdv", as_index=False).agg(stok_degeri=("stok_degeri", "sum"), urun_sayisi=("urun", "count"))
        fig = px.pie(kdv_df, names="kdv", values="stok_degeri", title="KDV Grubuna Göre Stok Sermayesi")
        st.plotly_chart(apply_plot_theme(fig, height=500), use_container_width=True)

    st.dataframe(pareto_df[["barkod", "urun", "stok", "psf", "raf", "stok_degeri", "kumulatif_oran", "pareto_sinif"]].head(300), use_container_width=True, hide_index=True)

elif page == "🔴 Risk Listeleri":
    st.markdown('<div class="section-title">Risk Listeleri</div>', unsafe_allow_html=True)
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["Kritik Stok", "Stok Yok", "Yüksek Stok", "Satış/Tahsilat Riski", "Hızlı Tükenen", "Stokta Yok Ama Satan", "Ölü Stok Adayı"])
    with t1:
        st.dataframe(critical_df[["barkod", "urun", "stok", "kritik_stok", "psf", "raf", "stok_degeri"]].sort_values("stok"), use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(no_stock_df[["barkod", "urun", "stok", "kritik_stok", "psf", "raf"]].sort_values("urun"), use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(high_stock_df[["barkod", "urun", "stok", "psf", "raf", "stok_degeri"]].sort_values("stok_degeri", ascending=False), use_container_width=True, hide_index=True)
    with t4:
        risk_cols = ["tarih", "satis_no", "satis_tipi", "tahsilat", "kurum", "doktor", "ciro", "odenen", "tahsilat_acigi", "brut_kar", "kar_marji", "sonlandi"]
        risk_df = period_df[(period_df["tahsilat_acigi"] > 0) | (~period_df["sonlandi"])].copy()
        if show_patient_columns:
            risk_cols.insert(4, "hasta")
        st.dataframe(risk_df[risk_cols].sort_values("tarih", ascending=False), use_container_width=True, hide_index=True)
    with t5:
        st.dataframe(fast_depleting_df[["barkod", "urun", "stok", "satilan_adet", "gunluk_satis_hizi", "tahmini_bitis_gun", "satis_tutari", "kar_tutari", "kar_marji"]].head(200), use_container_width=True, hide_index=True)
    with t6:
        st.dataframe(out_of_stock_selling_df[["barkod", "urun", "stok", "satilan_adet", "satis_tutari", "kar_tutari"]].head(200), use_container_width=True, hide_index=True)
    with t7:
        st.dataframe(dead_stock_df[["barkod", "urun", "stok", "psf", "satilan_adet", "satis_tutari", "kar_tutari", "urun_grubu"]].head(300), use_container_width=True, hide_index=True)

elif page == "📥 Rapor":
    st.markdown('<div class="section-title">Excel Raporu</div>', unsafe_allow_html=True)
    report = create_excel_report(sales_df, inventory_df, product_df, period_df, critical_df, no_stock_df, high_stock_df, pareto_df, kurum_df, doktor_df, weekday_df, hourly_df, doctor_intel.get("doctor_kpi"), patient_loyalty.get("frequency"))
    st.download_button(
        "📥 AYÇA Insight V7.1 Raporunu İndir",
        data=report,
        file_name=f"ayca_insight_v71_rapor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown(
        """
        <div class="ai-card">
            <div class="ai-title">Ürün satış raporu gelince açılacak motorlar</div>
            <div class="ai-text">
            Ürün satış raporu geldiğinde bu yapıya üçüncü dosya olarak eklenecek. Barkod üzerinden satış ve envanter birleşince
            satış hızı, stok bitiş günü, sipariş önerisi, gerçek ölü stok, ABC/XYZ ve miad riski modülleri aktif edilecek.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
