# ============================================================
# AYÇA Insight V7.1 - 3 Excel Ürün Zekası Sürümü
# ------------------------------------------------------------
# Zorunlu / Önerilen dosyalar:
# 1) Envanter Exceli
# 2) Ürün Bazında Toplamlar Exceli
# 3) Satış Hareketleri Exceli
#
# Bu sürüm üç dosyayı birlikte kullanır:
# - Envanter: stok, raf, kritik stok, stok değeri
# - Ürün bazında toplamlar: satılan adet, satış tutarı, kar, ürün grubu
# - Satış hareketleri: tarih, saat, kurum, doktor, tahsilat, ciro, kar
#
# Açılan ana motorlar:
# - Ürün satış hızı
# - Stok bitiş günü
# - Sipariş tavsiyesi
# - Ölü stok / yavaş stok / hızlı dönen ürün
# - Çok satan ama stokta az kalan ürünler
# - Çok karlı ürünler
# - Sermaye bağlayan ürünler
# - ABC ürün sınıflaması
# - Kurum / doktor / tahsilat / saat ritmi
# - Yönetici sabah ekranı
# - Excel rapor çıktısı
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
        --bg:#F8FAFC; --panel:#FFFFFF; --border:#E2E8F0; --text:#0F172A; --muted:#64748B;
        --blue:#2563EB; --green:#10B981; --orange:#F59E0B; --red:#EF4444; --purple:#8B5CF6;
        --blue-soft:#DBEAFE; --green-soft:#DCFCE7; --orange-soft:#FEF3C7; --red-soft:#FEE2E2; --purple-soft:#EDE9FE;
    }
    html, body, [data-testid="stAppViewContainer"] {background:var(--bg); color:var(--text);} 
    [data-testid="stHeader"] {background:rgba(248,250,252,.88); backdrop-filter:blur(10px);} 
    [data-testid="stSidebar"] {background:linear-gradient(180deg,#FFFFFF 0%,#F1F5F9 100%); border-right:1px solid var(--border);} 
    .block-container {padding-top:1.1rem; max-width:1580px;}
    .ayca-header {display:flex; justify-content:space-between; align-items:center; gap:16px; margin-bottom:14px;}
    .ayca-title h1 {margin:0; color:var(--text); font-size:32px; letter-spacing:-.7px; font-weight:950;}
    .ayca-title p {margin:6px 0 0 0; color:var(--muted); font-size:14px;}
    .header-pill {background:#fff; border:1px solid var(--border); border-radius:16px; padding:12px 16px; box-shadow:0 8px 24px rgba(15,23,42,.05); color:var(--text); font-weight:900; font-size:13px; white-space:nowrap;}
    .metric-card {min-height:128px; background:linear-gradient(180deg,#fff 0%,#F8FBFF 100%); border:1px solid var(--border); border-radius:20px; padding:17px; box-shadow:0 12px 30px rgba(15,23,42,.06); overflow:hidden; position:relative;}
    .metric-label {color:var(--muted); font-size:12px; font-weight:900; letter-spacing:.4px; text-transform:uppercase; margin-bottom:10px;}
    .metric-value {color:var(--text); font-size:27px; font-weight:950; margin-bottom:8px; letter-spacing:-.5px;}
    .metric-sub {color:var(--muted); font-size:13px; line-height:1.35;}
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
    .exec-grid {display:grid; grid-template-columns:1.25fr .75fr; gap:16px; margin:14px 0 18px 0;}
    .exec-card {background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%); border:1px solid #BFDBFE; border-radius:24px; padding:20px; box-shadow:0 14px 34px rgba(37,99,235,.08);}
    .exec-title {color:#0F172A; font-size:22px; font-weight:950; letter-spacing:-.4px; margin-bottom:8px;}
    .exec-sub {color:#64748B; font-size:14px; line-height:1.55; margin-bottom:12px;}
    .exec-list-item {background:rgba(255,255,255,.82); border:1px solid rgba(226,232,240,.95); border-radius:16px; padding:12px 13px; margin:9px 0; font-size:14px; color:#0F172A; line-height:1.45; font-weight:700;}
    .score-big {font-size:54px; font-weight:950; letter-spacing:-2px; color:#2563EB; line-height:1; margin:4px 0 8px 0;}
    .section-title {color:var(--text); font-size:21px; font-weight:950; margin:20px 0 12px 0; letter-spacing:-.3px;}
    .ai-card {background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%); border:1px solid #BFDBFE; border-radius:22px; padding:18px; box-shadow:0 12px 30px rgba(37,99,235,.08); margin:14px 0 18px 0;}
    .ai-title {color:var(--blue); font-size:18px; font-weight:950; margin-bottom:8px;}
    .ai-text {color:var(--text); font-size:14px; line-height:1.55;}
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
        <div class="mini-card alert-orange" style="margin:10px 0 16px 0; min-height:auto;">
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
                Bu sürüm üç TEBEOS Excel çıktısını birlikte okur: <b>Envanter</b>, <b>Ürün Bazında Toplamlar</b> ve <b>Satış Hareketleri</b>.
                Böylece ürün bazlı satış hızı, stok bitiş günü, sipariş tavsiyesi, ölü stok ve kârlılık motoru aktif olur.
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
# BASIC / PREMIUM KİLİDİ
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


def safe_div(a, b):
    try:
        return float(a) / float(b) if float(b or 0) != 0 else 0.0
    except Exception:
        return 0.0


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


def clean_text_series(s):
    return s.astype(str).replace({"nan": "", "None": ""})


# ============================================================
# STANDARDİZASYON FONKSİYONLARI
# ============================================================
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
    missing = [k for k in ["barkod", "urun", "stok"] if mapping.get(k) is None]
    if missing:
        raise ValueError("Envanter Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["barkod"] = clean_text_series(raw_df[mapping["barkod"]]).str.replace(r"\.0$", "", regex=True).str.strip()
    df["urun"] = clean_text_series(raw_df[mapping["urun"]]).str.strip()
    df["kademe"] = pd.to_numeric(raw_df[mapping["kademe"]], errors="coerce").fillna(0) if mapping["kademe"] else 0
    df["psf"] = pd.to_numeric(raw_df[mapping["psf"]], errors="coerce").fillna(0) if mapping["psf"] else 0
    df["kamu"] = pd.to_numeric(raw_df[mapping["kamu"]], errors="coerce").fillna(0) if mapping["kamu"] else 0
    df["stok_envanter"] = pd.to_numeric(raw_df[mapping["stok"]], errors="coerce").fillna(0)
    df["kritik_stok"] = pd.to_numeric(raw_df[mapping["kritik_stok"]], errors="coerce").fillna(0) if mapping["kritik_stok"] else 0
    df["raf"] = clean_text_series(raw_df[mapping["raf"]]).str.strip() if mapping["raf"] else "Bilinmiyor"
    df["kdv"] = pd.to_numeric(raw_df[mapping["kdv"]], errors="coerce").fillna(0) if mapping["kdv"] else 0
    df["psf_toplam"] = pd.to_numeric(raw_df[mapping["psf_toplam"]], errors="coerce").fillna(0) if mapping["psf_toplam"] else df["psf"] * df["stok_envanter"]
    df["kamu_toplam"] = pd.to_numeric(raw_df[mapping["kamu_toplam"]], errors="coerce").fillna(0) if mapping["kamu_toplam"] else df["kamu"] * df["stok_envanter"]
    df["mal_haric"] = pd.to_numeric(raw_df[mapping["mal_haric"]], errors="coerce").fillna(0) if mapping["mal_haric"] else 0
    df["mal_dahil"] = pd.to_numeric(raw_df[mapping["mal_dahil"]], errors="coerce").fillna(0) if mapping["mal_dahil"] else df["mal_haric"]
    df["stok_degeri"] = np.where(df["mal_dahil"] > 0, df["mal_dahil"], df["psf_toplam"])
    df = df[df["barkod"].ne("")].drop_duplicates("barkod", keep="first")
    return df


def standardize_product_sales(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw_df.columns)
    mapping = {
        "barkod": find_col(cols, ["Barkod"]),
        "urun": find_col(cols, ["Ürün Adı", "Urun Adi", "Ürün Adı (İçinde Geçen İsim Şeklinde Arama Yapılabilir)"]),
        "stok_rapor": find_col(cols, ["Stok"]),
        "psf": find_col(cols, ["Psf", "PSF"]),
        "alis_adet": find_col(cols, ["Alış Adet", "Alis Adet"]),
        "alis_maliyet": find_col(cols, ["Alış Maliyet Topl", "Alis Maliyet Topl"]),
        "satilan_adet": find_col(cols, ["Satılan Adet", "Satilan Adet"]),
        "satis_tutari": find_col(cols, ["Satış Tutarı", "Satis Tutari"]),
        "iade_adet": find_col(cols, ["İade Adet", "Iade Adet"]),
        "iade_tutari": find_col(cols, ["İade Tutarı", "Iade Tutari"]),
        "kar_tutari": find_col(cols, ["Kar Tutarı", "Kar Tutari", "Brüt Kar"]),
        "fark_toplami": find_col(cols, ["Fark Toplamı", "Fark Toplami"]),
        "urun_grubu": find_col(cols, ["İlaç Dışı Ürün Grubu", "Ilac Disi Urun Grubu", "Ürün Grubu"]),
    }
    missing = [k for k in ["barkod", "urun", "satilan_adet", "satis_tutari", "kar_tutari"] if mapping.get(k) is None]
    if missing:
        raise ValueError("Ürün bazında toplamlar Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["barkod"] = clean_text_series(raw_df[mapping["barkod"]]).str.replace(r"\.0$", "", regex=True).str.strip()
    df["urun_satis"] = clean_text_series(raw_df[mapping["urun"]]).str.strip()
    df["stok_urun_raporu"] = pd.to_numeric(raw_df[mapping["stok_rapor"]], errors="coerce").fillna(0) if mapping["stok_rapor"] else 0
    df["psf_urun_raporu"] = pd.to_numeric(raw_df[mapping["psf"]], errors="coerce").fillna(0) if mapping["psf"] else 0
    df["alis_adet"] = pd.to_numeric(raw_df[mapping["alis_adet"]], errors="coerce").fillna(0) if mapping["alis_adet"] else 0
    df["alis_maliyet_toplam"] = pd.to_numeric(raw_df[mapping["alis_maliyet"]], errors="coerce").fillna(0) if mapping["alis_maliyet"] else 0
    df["satilan_adet"] = pd.to_numeric(raw_df[mapping["satilan_adet"]], errors="coerce").fillna(0)
    df["satis_tutari"] = pd.to_numeric(raw_df[mapping["satis_tutari"]], errors="coerce").fillna(0)
    df["iade_adet"] = pd.to_numeric(raw_df[mapping["iade_adet"]], errors="coerce").fillna(0) if mapping["iade_adet"] else 0
    df["iade_tutari"] = pd.to_numeric(raw_df[mapping["iade_tutari"]], errors="coerce").fillna(0) if mapping["iade_tutari"] else 0
    df["kar_tutari"] = pd.to_numeric(raw_df[mapping["kar_tutari"]], errors="coerce").fillna(0)
    df["fark_toplami"] = pd.to_numeric(raw_df[mapping["fark_toplami"]], errors="coerce").fillna(0) if mapping["fark_toplami"] else 0
    df["urun_grubu"] = clean_text_series(raw_df[mapping["urun_grubu"]]).str.strip() if mapping["urun_grubu"] else "Bilinmiyor"
    df = df[df["barkod"].ne("")]

    # Aynı barkod raporda birden fazla gelirse tekilleştir.
    agg = df.groupby("barkod", as_index=False).agg(
        urun_satis=("urun_satis", "first"),
        stok_urun_raporu=("stok_urun_raporu", "max"),
        psf_urun_raporu=("psf_urun_raporu", "max"),
        alis_adet=("alis_adet", "sum"),
        alis_maliyet_toplam=("alis_maliyet_toplam", "sum"),
        satilan_adet=("satilan_adet", "sum"),
        satis_tutari=("satis_tutari", "sum"),
        iade_adet=("iade_adet", "sum"),
        iade_tutari=("iade_tutari", "sum"),
        kar_tutari=("kar_tutari", "sum"),
        fark_toplami=("fark_toplami", "sum"),
        urun_grubu=("urun_grubu", "first"),
    )
    return agg


def standardize_sales(raw_df: pd.DataFrame) -> pd.DataFrame:
    cols = list(raw_df.columns)
    mapping = {
        "satis_no": find_col(cols, ["Satış No", "Satis No", "Fiş No", "Fis No"]),
        "satis_tipi": find_col(cols, ["Satış Tipi", "Satis Tipi"]),
        "tahsilat": find_col(cols, ["Tahsilat", "Ödeme Tipi", "Odeme Tipi"]),
        "hasta": find_col(cols, ["Hasta Adı Soyadı", "Hasta Adi Soyadi"]),
        "recete_no": find_col(cols, ["Reç. No", "Rec. No", "Reçete No", "Recete No"]),
        "doktor": find_col(cols, ["Doktor Adı", "Doktor Adi", "Doktor"]),
        "kurum": find_col(cols, ["Kurum Adı", "Kurum Adi"]),
        "grup": find_col(cols, ["Grubu", "Grup"]),
        "recete_tarihi": find_col(cols, ["Reç. Tar", "Rec. Tar", "Reçete Tarihi"]),
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
    missing = [k for k in ["satis_no", "toplam", "kar", "maliyet", "islem_tarihi"] if mapping.get(k) is None]
    if missing:
        raise ValueError("Satış hareketleri Excelinde eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["satis_no"] = clean_text_series(raw_df[mapping["satis_no"]]).str.strip()
    df["satis_tipi"] = clean_text_series(raw_df[mapping["satis_tipi"]]).str.strip() if mapping["satis_tipi"] else "Bilinmiyor"
    df["tahsilat"] = clean_text_series(raw_df[mapping["tahsilat"]]).str.strip() if mapping["tahsilat"] else "Bilinmiyor"
    df["hasta"] = clean_text_series(raw_df[mapping["hasta"]]).str.strip() if mapping["hasta"] else ""
    df["recete_no"] = clean_text_series(raw_df[mapping["recete_no"]]).str.strip() if mapping["recete_no"] else ""
    df["doktor"] = clean_text_series(raw_df[mapping["doktor"]]).str.strip() if mapping["doktor"] else "Bilinmiyor"
    df["kurum"] = clean_text_series(raw_df[mapping["kurum"]]).str.strip() if mapping["kurum"] else "Bilinmiyor"
    df["grup"] = clean_text_series(raw_df[mapping["grup"]]).str.strip() if mapping["grup"] else "Bilinmiyor"
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
    df["kullanici"] = clean_text_series(raw_df[mapping["kullanici"]]).str.strip() if mapping["kullanici"] else ""
    df = df.dropna(subset=["tarih"])
    df["gun"] = df["tarih"].dt.date
    df["ay"] = df["tarih"].dt.to_period("M").astype(str)
    df["saat"] = df["tarih"].dt.hour
    day_map = {0:"Pazartesi",1:"Salı",2:"Çarşamba",3:"Perşembe",4:"Cuma",5:"Cumartesi",6:"Pazar"}
    df["hafta_gunu"] = df["tarih"].dt.dayofweek.map(day_map)
    df["hafta_gunu_no"] = df["tarih"].dt.dayofweek
    df["kar_marji"] = np.where(df["ciro"] > 0, df["brut_kar"] / df["ciro"], 0)
    df["tahsilat_acigi"] = (df["ciro"] - df["odenen"]).clip(lower=0)
    return df


# ============================================================
# ANALİZ MOTORU
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
        "tahsilat_acigi": df["tahsilat_acigi"].sum(),
        "tahsilat_orani": safe_div(df["odenen"].sum(), ciro),
        "sonlanmamis": int((~df["sonlandi"]).sum()),
    }


def make_product_master(inv_df: pd.DataFrame, prod_df: pd.DataFrame, analysis_days: int, target_days: int, safety_days: int) -> pd.DataFrame:
    master = prod_df.merge(inv_df, on="barkod", how="outer", suffixes=("", "_env"))
    master["urun"] = master["urun"].fillna(master["urun_satis"]).fillna(master["barkod"])
    master["urun_satis"] = master["urun_satis"].fillna(master["urun"])
    master["urun_grubu"] = master["urun_grubu"].fillna("Bilinmiyor")
    for col in [
        "satilan_adet", "satis_tutari", "kar_tutari", "iade_adet", "iade_tutari", "alis_adet",
        "alis_maliyet_toplam", "fark_toplami", "stok_envanter", "stok_urun_raporu", "psf", "psf_urun_raporu",
        "kritik_stok", "stok_degeri", "mal_dahil", "mal_haric", "kamu", "kamu_toplam"
    ]:
        if col not in master.columns:
            master[col] = 0
        master[col] = pd.to_numeric(master[col], errors="coerce").fillna(0)

    master["raf"] = master.get("raf", "Bilinmiyor")
    master["raf"] = master["raf"].fillna("Bilinmiyor")
    master["stok"] = np.where(master["stok_envanter"].notna(), master["stok_envanter"], master["stok_urun_raporu"])
    master["stok"] = pd.to_numeric(master["stok"], errors="coerce").fillna(0)
    master["psf_final"] = np.where(master["psf"] > 0, master["psf"], master["psf_urun_raporu"])
    master["stok_degeri"] = np.where(master["stok_degeri"] > 0, master["stok_degeri"], master["stok"] * master["psf_final"])
    master["maliyet_tahmini"] = np.where(master["satis_tutari"] - master["kar_tutari"] > 0, master["satis_tutari"] - master["kar_tutari"], 0)
    master["birim_kar"] = np.where(master["satilan_adet"] > 0, master["kar_tutari"] / master["satilan_adet"], 0)
    master["birim_satis"] = np.where(master["satilan_adet"] > 0, master["satis_tutari"] / master["satilan_adet"], master["psf_final"])
    master["kar_marji"] = np.where(master["satis_tutari"] > 0, master["kar_tutari"] / master["satis_tutari"], 0)
    master["gunluk_satis_hizi"] = master["satilan_adet"] / max(1, analysis_days)
    master["stok_bitis_gunu"] = np.where(master["gunluk_satis_hizi"] > 0, master["stok"] / master["gunluk_satis_hizi"], np.inf)
    master["stok_bitis_gunu_goster"] = master["stok_bitis_gunu"].replace(np.inf, np.nan)
    master["hedef_stok"] = np.ceil(master["gunluk_satis_hizi"] * (target_days + safety_days))
    master["siparis_onerisi"] = np.maximum(0, master["hedef_stok"] - master["stok"]).round(0)
    master["siparis_tahmini_tutar"] = master["siparis_onerisi"] * master["birim_satis"].fillna(0)
    master["kritik_mi"] = np.where(master["kritik_stok"] > 0, master["stok"] <= master["kritik_stok"], master["stok"] <= 1)
    master["stok_yok_mu"] = master["stok"] <= 0
    master["hizli_tukeniyor_mu"] = (master["stok_bitis_gunu"] <= 14) & (master["satilan_adet"] > 0)
    master["siparis_gerekli_mi"] = master["siparis_onerisi"] > 0
    master["olu_stok_mu"] = (master["satilan_adet"] <= 0) & (master["stok"] > 0)
    master["yavas_stok_mu"] = (master["satilan_adet"] > 0) & (master["stok_bitis_gunu"] > 90) & (master["stok"] > 0)
    master["sermaye_riski_mi"] = (master["stok_degeri"] > master["stok_degeri"].quantile(0.90)) & (master["stok_bitis_gunu"] > 60)
    master["stokta_yok_satmis_mi"] = (master["satilan_adet"] > 0) & (master["stok"] <= 0)

    master = master.sort_values(["satis_tutari", "satilan_adet"], ascending=False)
    total_sales = master["satis_tutari"].sum()
    master["ciro_payi"] = np.where(total_sales > 0, master["satis_tutari"] / total_sales, 0)
    master["kumulatif_ciro_payi"] = master["ciro_payi"].cumsum()
    master["abc_sinif"] = np.select(
        [master["kumulatif_ciro_payi"] <= 0.80, master["kumulatif_ciro_payi"] <= 0.95],
        ["A - Ciro Motoru", "B - Destek"],
        default="C - Uzun Kuyruk",
    )
    master["aksiyon"] = np.select(
        [
            master["stokta_yok_satmis_mi"],
            master["hizli_tukeniyor_mu"] & master["siparis_gerekli_mi"],
            master["olu_stok_mu"],
            master["yavas_stok_mu"],
            master["sermaye_riski_mi"],
        ],
        [
            "Acil sipariş / satış kaçırma riski",
            "Sipariş öner",
            "Ölü stok - aksiyon al",
            "Yavaş stok - kampanya / raf kontrolü",
            "Sermaye riski - stok azalt",
        ],
        default="Normal takip",
    )
    return master


def create_action_items(product_master, current_stats, previous_stats, daily_df):
    actions = []
    urgent = product_master[product_master["stokta_yok_satmis_mi"]].sort_values("satis_tutari", ascending=False)
    reorder = product_master[product_master["siparis_gerekli_mi"]].sort_values("siparis_tahmini_tutar", ascending=False)
    dead = product_master[product_master["olu_stok_mu"]].sort_values("stok_degeri", ascending=False)
    slow = product_master[product_master["yavas_stok_mu"]].sort_values("stok_degeri", ascending=False)
    profitable = product_master[product_master["satilan_adet"] > 0].sort_values("kar_tutari", ascending=False)

    if not urgent.empty:
        r = urgent.iloc[0]
        actions.append(f"⛔ {len(urgent)} ürün satılmış ama stok 0/eksi. İlk risk: {r['urun']} · satış {num_fmt(r['satilan_adet'],0)} adet.")
    if not reorder.empty:
        r = reorder.iloc[0]
        actions.append(f"🛒 Sipariş listesinde {len(reorder)} ürün var. En yüksek tutarlı öneri: {r['urun']} · {num_fmt(r['siparis_onerisi'],0)} adet.")
    if not dead.empty:
        r = dead.iloc[0]
        actions.append(f"🧊 Ölü stokta {len(dead)} ürün var. En çok sermaye bağlayan: {r['urun']} · {money_fmt(r['stok_degeri'])}.")
    if not slow.empty:
        r = slow.iloc[0]
        actions.append(f"🐢 Yavaş stokta {len(slow)} ürün var. Raf/kampanya kontrolü: {r['urun']}.")
    if not profitable.empty:
        r = profitable.iloc[0]
        actions.append(f"💎 En yüksek kâr getiren ürün: {r['urun']} · kâr {money_fmt(r['kar_tutari'])} · marj {pct_fmt(r['kar_marji'])}.")
    if current_stats["tahsilat_acigi"] > 0:
        actions.append(f"🧾 Tahsilat açığı {money_fmt(current_stats['tahsilat_acigi'])}. Satış hareketleri ekranında kapanmayan kayıtları incele.")
    if current_stats["marj"] < 0.12:
        actions.append("💰 Brüt kâr marjı düşük. İskonto, fiyat farkı ve maliyet kayıtlarını kontrol et.")
    if daily_df is not None and not daily_df.empty:
        min_day = daily_df.sort_values("ciro").iloc[0]
        max_day = daily_df.sort_values("ciro", ascending=False).iloc[0]
        actions.append(f"📈 En güçlü gün {max_day['gun']} ({money_fmt(max_day['ciro'])}); en zayıf gün {min_day['gun']} ({money_fmt(min_day['ciro'])}).")
    return actions[:8] if actions else ["✅ Kritik aksiyon görünmüyor. Günlük ciro, tahsilat ve stok hızını takip et."]


def score_from_threshold(value: float, good: float, warning: float, bad: float, higher_is_better: bool = True) -> int:
    try:
        value = float(value)
    except Exception:
        return 0
    if higher_is_better:
        if value >= good: return 100
        if value >= warning: return 75
        if value >= bad: return 45
        return 15
    if value <= good: return 100
    if value <= warning: return 75
    if value <= bad: return 45
    return 15


def health_score(product_master, current_stats, previous_stats):
    ciro = current_stats["ciro"]
    margin = current_stats["marj"]
    tahsilat_gap_ratio = safe_div(current_stats["tahsilat_acigi"], ciro)
    urgent_ratio = safe_div(product_master["stokta_yok_satmis_mi"].sum(), len(product_master))
    dead_value_ratio = safe_div(product_master.loc[product_master["olu_stok_mu"], "stok_degeri"].sum(), product_master["stok_degeri"].sum())
    reorder_ratio = safe_div(product_master["siparis_gerekli_mi"].sum(), len(product_master))
    growth = safe_div(current_stats["ciro"] - previous_stats["ciro"], previous_stats["ciro"]) if previous_stats["ciro"] else 0
    scores = {
        "Kârlılık": score_from_threshold(margin, 0.20, 0.15, 0.10, True),
        "Tahsilat": score_from_threshold(tahsilat_gap_ratio, 0.02, 0.05, 0.10, False),
        "Ürün Bulunurluğu": int(max(0, min(100, 100 - urgent_ratio * 250))),
        "Stok Verimliliği": int(max(0, min(100, 100 - dead_value_ratio * 130 - reorder_ratio * 20))),
        "Büyüme": score_from_threshold(growth, 0.08, 0.00, -0.08, True),
    }
    weights = {"Kârlılık": 20, "Tahsilat": 20, "Ürün Bulunurluğu": 25, "Stok Verimliliği": 25, "Büyüme": 10}
    total = sum(scores[k] * weights[k] for k in scores) / sum(weights.values())
    return int(round(max(0, min(100, total)))), scores, weights


def score_status(score_value: int) -> str:
    if score_value >= 82: return "Güçlü"
    if score_value >= 65: return "Kontrollü"
    if score_value >= 50: return "Takip edilmeli"
    return "Riskli"


def create_excel_report(product_master, sales_df, period_df, kurum_df, doktor_df, daily_df, weekday_df, hourly_df):
    output = BytesIO()
    export_cols = [
        "barkod", "urun", "urun_grubu", "raf", "stok", "kritik_stok", "psf_final", "stok_degeri",
        "satilan_adet", "satis_tutari", "kar_tutari", "kar_marji", "gunluk_satis_hizi", "stok_bitis_gunu_goster",
        "hedef_stok", "siparis_onerisi", "siparis_tahmini_tutar", "abc_sinif", "aksiyon"
    ]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        product_master[export_cols].to_excel(writer, sheet_name="Urun_Zekasi", index=False)
        product_master[product_master["siparis_gerekli_mi"]][export_cols].to_excel(writer, sheet_name="Siparis_Onerisi", index=False)
        product_master[product_master["olu_stok_mu"]][export_cols].to_excel(writer, sheet_name="Olu_Stok", index=False)
        product_master[product_master["yavas_stok_mu"]][export_cols].to_excel(writer, sheet_name="Yavas_Stok", index=False)
        product_master[product_master["stokta_yok_satmis_mi"]][export_cols].to_excel(writer, sheet_name="Stokta_Yok_Satmis", index=False)
        kurum_df.to_excel(writer, sheet_name="Kurum", index=False)
        doktor_df.to_excel(writer, sheet_name="Doktor", index=False)
        daily_df.to_excel(writer, sheet_name="Gunluk", index=False)
        weekday_df.to_excel(writer, sheet_name="Hafta_Gunu", index=False)
        hourly_df.to_excel(writer, sheet_name="Saatlik", index=False)
        period_df.to_excel(writer, sheet_name="Satis_Hareket_Secili", index=False)
        sales_df.to_excel(writer, sheet_name="Satis_Hareket_Tum", index=False)
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
st.sidebar.caption("V7.0 · 3 Excel Ürün Zekası")
eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")

inventory_file = st.sidebar.file_uploader("1/3) Envanter Exceli - ZORUNLU", type=["xlsx", "xls"], key="inventory_file")
product_file = st.sidebar.file_uploader("2/3) Ürün Bazında Toplamlar Exceli - ZORUNLU", type=["xlsx", "xls"], key="product_file")
sales_file = st.sidebar.file_uploader("3/3) Satış Hareketleri Exceli - ZORUNLU", type=["xlsx", "xls"], key="sales_file")

st.sidebar.markdown("---")
selected_period = st.sidebar.selectbox("Satış hareket dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Tüm veri"], index=2)
target_days = st.sidebar.slider("Sipariş hedef stok günü", 7, 90, 30)
safety_days = st.sidebar.slider("Güvenlik stok günü", 0, 30, 7)
manual_days = st.sidebar.number_input("Ürün raporu kaç günü kapsıyor?", min_value=1, max_value=365, value=30, step=1)
use_sales_date_span = st.sidebar.checkbox("Gün hesabında satış hareket tarih aralığını kullan", value=True)
show_patient_columns = st.sidebar.checkbox("Hasta isim kolonunu göster", value=False)
st.sidebar.caption("Hasta TC ve kişisel sağlık verisi analiz dışı tutulmalıdır. Hasta adı varsayılan olarak gizlidir.")


# ============================================================
# DOSYA BEKLEME EKRANI
# ============================================================
if inventory_file is None or product_file is None or sales_file is None:
    st.markdown(
        f"""
        <div class="ayca-header">
            <div class="ayca-title">
                <h1>AYÇA Insight V7.1</h1>
                <p>{eczane_adi} · Üç Excel dosyasını yükle: envanter, ürün bazında toplamlar, satış hareketleri.</p>
            </div>
            <div class="header-pill">Dosya bekleniyor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3 = st.columns(3)
    with c1: make_mini_card("1. Envanter", "Zorunlu", "Barkod, stok, raf, kritik stok, stok değeri", "alert-blue")
    with c2: make_mini_card("2. Ürün Bazında Toplamlar", "Zorunlu", "Satılan adet, satış tutarı, kâr, ürün grubu", "alert-green")
    with c3: make_mini_card("3. Satış Hareketleri", "Zorunlu", "Tarih, saat, kurum, doktor, tahsilat", "alert-purple")
    st.info("Sol menüden üç dosyayı da yüklediğinde ürün zekâsı otomatik çalışır.")
    st.stop()


# ============================================================
# DOSYALARI OKU
# ============================================================
try:
    raw_inventory, inv_sheet, _ = read_excel_first_sheet(inventory_file)
    raw_product, product_sheet, _ = read_excel_first_sheet(product_file)
    raw_sales, sales_sheet, _ = read_excel_first_sheet(sales_file)

    inventory_df = standardize_inventory(raw_inventory)
    product_df = standardize_product_sales(raw_product)
    sales_df = standardize_sales(raw_sales)
except Exception as exc:
    st.error(f"Dosya okunurken hata oluştu: {exc}")
    st.stop()


# ============================================================
# DÖNEM VE BİRLEŞİK MOTOR
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
prev_start = start_date - pd.Timedelta(days=period_days)
prev_df = sales_df[(sales_df["tarih"] >= prev_start) & (sales_df["tarih"] < start_date)].copy()

analysis_days = period_days if use_sales_date_span else int(manual_days)
analysis_days = max(1, int(analysis_days))

product_master = make_product_master(inventory_df, product_df, analysis_days, target_days, safety_days)
current_stats = summarize_sales(period_df)
previous_stats = summarize_sales(prev_df)
score, score_items, score_weights = health_score(product_master, current_stats, previous_stats)

# Özet tablolar
kurum_df = period_df.groupby("kurum", as_index=False).agg(
    ciro=("ciro", "sum"), kar=("brut_kar", "sum"), maliyet=("maliyet", "sum"), islem=("satis_no", "nunique"), tahsilat_acigi=("tahsilat_acigi", "sum")
).sort_values("ciro", ascending=False)
kurum_df["marj"] = np.where(kurum_df["ciro"] > 0, kurum_df["kar"] / kurum_df["ciro"], 0)

doktor_df = period_df.groupby("doktor", as_index=False).agg(
    ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique")
).sort_values("ciro", ascending=False)
doktor_df["marj"] = np.where(doktor_df["ciro"] > 0, doktor_df["kar"] / doktor_df["ciro"], 0)

weekday_df = period_df.groupby(["hafta_gunu_no", "hafta_gunu"], as_index=False).agg(
    ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique"), gun_sayisi=("gun", "nunique")
).sort_values("hafta_gunu_no")
weekday_df["gunluk_ortalama_ciro"] = np.where(weekday_df["gun_sayisi"] > 0, weekday_df["ciro"] / weekday_df["gun_sayisi"], 0)

hourly_df = period_df.groupby("saat", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique")).sort_values("saat")

daily_df = period_df.groupby("gun", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique"), tahsilat_acigi=("tahsilat_acigi", "sum"))
daily_df["marj"] = np.where(daily_df["ciro"] > 0, daily_df["kar"] / daily_df["ciro"], 0)

actions = create_action_items(product_master, current_stats, previous_stats, daily_df)

# Segmentler
reorder_df = product_master[product_master["siparis_gerekli_mi"]].sort_values("siparis_tahmini_tutar", ascending=False)
dead_df = product_master[product_master["olu_stok_mu"]].sort_values("stok_degeri", ascending=False)
slow_df = product_master[product_master["yavas_stok_mu"]].sort_values("stok_degeri", ascending=False)
urgent_df = product_master[product_master["stokta_yok_satmis_mi"]].sort_values("satis_tutari", ascending=False)
fast_df = product_master[product_master["hizli_tukeniyor_mu"]].sort_values("stok_bitis_gunu")
profit_df = product_master[product_master["satilan_adet"] > 0].sort_values("kar_tutari", ascending=False)
capital_df = product_master.sort_values("stok_degeri", ascending=False)
abc_df = product_master.groupby("abc_sinif", as_index=False).agg(urun_sayisi=("barkod", "count"), ciro=("satis_tutari", "sum"), kar=("kar_tutari", "sum"), stok_degeri=("stok_degeri", "sum"))
group_df = product_master.groupby("urun_grubu", as_index=False).agg(urun_sayisi=("barkod", "count"), satilan_adet=("satilan_adet", "sum"), ciro=("satis_tutari", "sum"), kar=("kar_tutari", "sum"), stok_degeri=("stok_degeri", "sum"), siparis_onerisi=("siparis_onerisi", "sum")).sort_values("ciro", ascending=False)
group_df["marj"] = np.where(group_df["ciro"] > 0, group_df["kar"] / group_df["ciro"], 0)

# Eşleşme kalitesi
product_barcodes = set(product_df["barkod"])
inv_barcodes = set(inventory_df["barkod"])
matched_count = len(product_barcodes & inv_barcodes)
match_ratio = safe_div(matched_count, len(product_barcodes))


# ============================================================
# HEADER + KPI
# ============================================================
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-header">
        <div class="ayca-title">
            <h1>AYÇA Insight V7.1</h1>
            <p>{eczane_adi} · {selected_period} · Gün hesabı: {analysis_days} gün · {today_str}</p>
        </div>
        <div class="header-pill">AYÇA Ürün Puanı: {score}/100 · {score_status(score)}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

ciro_trend, ciro_class = rate_fmt(current_stats["ciro"], previous_stats["ciro"])
profit_trend, profit_class = rate_fmt(current_stats["kar"], previous_stats["kar"])
margin_trend, margin_class = rate_fmt(current_stats["marj"], previous_stats["marj"])

k1, k2, k3, k4, k5 = st.columns(5)
with k1: make_metric_card("Ciro", money_fmt(current_stats["ciro"]), selected_period, ciro_trend, ciro_class)
with k2: make_metric_card("Brüt Kâr", money_fmt(current_stats["kar"]), "Satış hareketleri", profit_trend, profit_class)
with k3: make_metric_card("Ürün Satış Tutarı", money_fmt(product_master["satis_tutari"].sum()), f"{num_fmt(product_master['satilan_adet'].sum(),0)} adet")
with k4: make_metric_card("Stok Değeri", money_fmt(product_master["stok_degeri"].sum()), "Envanter + ürün raporu")
with k5: make_metric_card("Sipariş Önerisi", money_fmt(reorder_df["siparis_tahmini_tutar"].sum()), f"{len(reorder_df)} ürün")

r1, r2, r3, r4, r5 = st.columns(5)
with r1: make_mini_card("Barkod Eşleşme", pct_fmt(match_ratio), f"{matched_count} / {len(product_barcodes)} ürün", "alert-green" if match_ratio >= .70 else "alert-orange")
with r2: make_mini_card("Stokta Yok Satmış", str(len(urgent_df)), "Satılmış ama stok 0/eksi", "alert-red" if len(urgent_df) else "alert-green")
with r3: make_mini_card("Ölü Stok", str(len(dead_df)), money_fmt(dead_df["stok_degeri"].sum()), "alert-orange" if len(dead_df) else "alert-green")
with r4: make_mini_card("Yavaş Stok", str(len(slow_df)), money_fmt(slow_df["stok_degeri"].sum()), "alert-purple" if len(slow_df) else "alert-green")
with r5: make_mini_card("Tahsilat Açığı", money_fmt(current_stats["tahsilat_acigi"]), f"Oran {pct_fmt(safe_div(current_stats['tahsilat_acigi'], current_stats['ciro']))}", "alert-red" if current_stats["tahsilat_acigi"] > 0 else "alert-green")

health_html = "".join([
    f'<div class="health-row"><div class="health-head"><span>{k} <small>({score_weights[k]}%)</small></span><span>{v}/100</span></div><div class="health-bar-bg"><div class="health-bar-fill" style="width:{v}%;"></div></div></div>'
    for k, v in score_items.items()
])
action_html = "".join([f"<div class='exec-list-item'>{item}</div>" for item in actions])
st.markdown(
    f"""
    <div class="exec-grid">
        <div class="exec-card">
            <div class="exec-title">🤖 Günaydın {kullanici_adi}</div>
            <div class="exec-sub">AYÇA üç dosyayı birleştirerek ürün bazlı satış, stok, sipariş ve risk analizini çıkardı.</div>
            {action_html}
        </div>
        <div class="exec-card">
            <div class="exec-sub">Ürün Zekâsı Puanı</div>
            <div class="score-big">{score}</div>
            <div class="exec-sub">Durum: <b>{score_status(score)}</b></div>
            {health_html}
            <div class="exec-list-item">📦 Ürün raporu: <b>{len(product_df)}</b> barkod · Envanter: <b>{len(inventory_df)}</b> barkod · Eşleşme: <b>{pct_fmt(match_ratio)}</b></div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SAYFALAR
# ============================================================
pages = ["🏠 Sabah Ekranı", "🛒 Sipariş Motoru", "📦 Ürün Zekası", "💰 Kârlılık", "🧊 Ölü/Yavaş Stok", "📈 Ciro & Tahsilat", "🏥 Kurum & Doktor", "📥 Rapor"]
page = st.radio("Bölüm", pages, horizontal=True, label_visibility="collapsed")

product_cols = [
    "barkod", "urun", "urun_grubu", "raf", "stok", "kritik_stok", "psf_final", "stok_degeri",
    "satilan_adet", "satis_tutari", "kar_tutari", "kar_marji", "gunluk_satis_hizi", "stok_bitis_gunu_goster",
    "siparis_onerisi", "siparis_tahmini_tutar", "abc_sinif", "aksiyon"
]

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
        st.markdown('<div class="section-title">En Kritik Siparişler</div>', unsafe_allow_html=True)
        top_reorder = reorder_df.head(12).copy()
        fig = px.bar(top_reorder, x="siparis_tahmini_tutar", y="urun", orientation="h", title="Tutar Bazlı İlk Sipariş Önerileri")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    st.markdown('<div class="section-title">Bugünün Aksiyon Listesi</div>', unsafe_allow_html=True)
    for item in actions:
        st.markdown(f"<div class='exec-list-item'>{item}</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">En Acil 20 Ürün</div>', unsafe_allow_html=True)
    acil = pd.concat([urgent_df, fast_df, reorder_df], ignore_index=True).drop_duplicates("barkod").head(20)
    st.dataframe(acil[product_cols], use_container_width=True, hide_index=True)

elif page == "🛒 Sipariş Motoru":
    st.markdown('<div class="section-title">Sipariş Tavsiye Motoru</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: make_mini_card("Sipariş Ürünü", str(len(reorder_df)), "Önerilen ürün sayısı", "alert-blue")
    with c2: make_mini_card("Tahmini Tutar", money_fmt(reorder_df["siparis_tahmini_tutar"].sum()), "PSF/birim satışa göre", "alert-green")
    with c3: make_mini_card("Acil Stok Yok", str(len(urgent_df)), "Satılmış ama stok yok", "alert-red" if len(urgent_df) else "alert-green")
    with c4: make_mini_card("14 Gün Altı", str(len(fast_df)), "Hızlı tükenen stok", "alert-orange" if len(fast_df) else "alert-green")

    c5, c6 = st.columns(2)
    with c5:
        fig = px.bar(reorder_df.head(20), x="siparis_tahmini_tutar", y="urun", orientation="h", title="İlk 20 Sipariş Önerisi")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
    with c6:
        fig = px.scatter(product_master[product_master["satilan_adet"] > 0], x="stok_bitis_gunu_goster", y="kar_tutari", size="satilan_adet", hover_name="urun", title="Stok Bitiş Günü / Kâr")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)

    t1, t2, t3 = st.tabs(["Sipariş Listesi", "Stokta Yok Satmış", "Hızlı Tükenen"])
    with t1: st.dataframe(reorder_df[product_cols], use_container_width=True, hide_index=True)
    with t2: st.dataframe(urgent_df[product_cols], use_container_width=True, hide_index=True)
    with t3: st.dataframe(fast_df[product_cols], use_container_width=True, hide_index=True)

elif page == "📦 Ürün Zekası":
    st.markdown('<div class="section-title">Ürün Performans Merkezi</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(abc_df, x="abc_sinif", y="ciro", title="ABC Sınıfına Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        fig = px.bar(group_df.head(12), x="ciro", y="urun_grubu", orientation="h", title="Ürün Grubuna Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig = px.bar(product_master.head(20), x="satis_tutari", y="urun", orientation="h", title="En Çok Ciro Yapan 20 Ürün")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
    with c4:
        fig = px.bar(product_master.sort_values("satilan_adet", ascending=False).head(20), x="satilan_adet", y="urun", orientation="h", title="En Çok Adet Satan 20 Ürün")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)

    t1, t2 = st.tabs(["Ürün Ana Tablo", "Ürün Grubu"])
    with t1: st.dataframe(product_master[product_cols], use_container_width=True, hide_index=True)
    with t2: st.dataframe(group_df, use_container_width=True, hide_index=True)

elif page == "💰 Kârlılık":
    st.markdown('<div class="section-title">Ürün Kârlılığı</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: make_mini_card("Ürün Kârı", money_fmt(product_master["kar_tutari"].sum()), "Ürün bazında toplam", "alert-green")
    with c2: make_mini_card("Ürün Marjı", pct_fmt(safe_div(product_master["kar_tutari"].sum(), product_master["satis_tutari"].sum())), "Kâr / satış", "alert-blue")
    with c3: make_mini_card("En Karlı Ürün", profit_df.iloc[0]["urun"] if not profit_df.empty else "-", money_fmt(profit_df.iloc[0]["kar_tutari"]) if not profit_df.empty else "₺0", "alert-purple")

    c4, c5 = st.columns(2)
    with c4:
        fig = px.bar(profit_df.head(20), x="kar_tutari", y="urun", orientation="h", title="En Çok Kâr Getiren Ürünler")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
    with c5:
        high_margin = product_master[(product_master["satilan_adet"] >= 3) & (product_master["satis_tutari"] > 0)].sort_values("kar_marji", ascending=False).head(20)
        fig = px.bar(high_margin, x="kar_marji", y="urun", orientation="h", title="Yüksek Marjlı Ürünler")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)

    st.dataframe(profit_df[product_cols], use_container_width=True, hide_index=True)

elif page == "🧊 Ölü/Yavaş Stok":
    st.markdown('<div class="section-title">Ölü Stok, Yavaş Stok ve Sermaye Riski</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: make_mini_card("Ölü Stok", str(len(dead_df)), money_fmt(dead_df["stok_degeri"].sum()), "alert-red" if len(dead_df) else "alert-green")
    with c2: make_mini_card("Yavaş Stok", str(len(slow_df)), money_fmt(slow_df["stok_degeri"].sum()), "alert-orange" if len(slow_df) else "alert-green")
    with c3: make_mini_card("En Çok Sermaye", capital_df.iloc[0]["urun"] if not capital_df.empty else "-", money_fmt(capital_df.iloc[0]["stok_degeri"]) if not capital_df.empty else "₺0", "alert-purple")

    c4, c5 = st.columns(2)
    with c4:
        fig = px.bar(dead_df.head(20), x="stok_degeri", y="urun", orientation="h", title="Ölü Stok - Sermaye Bazlı")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
    with c5:
        fig = px.bar(slow_df.head(20), x="stok_degeri", y="urun", orientation="h", title="Yavaş Stok - Sermaye Bazlı")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)

    t1, t2, t3 = st.tabs(["Ölü Stok", "Yavaş Stok", "Sermaye Bağlayanlar"])
    with t1: st.dataframe(dead_df[product_cols], use_container_width=True, hide_index=True)
    with t2: st.dataframe(slow_df[product_cols], use_container_width=True, hide_index=True)
    with t3: st.dataframe(capital_df[product_cols], use_container_width=True, hide_index=True)

elif page == "📈 Ciro & Tahsilat":
    st.markdown('<div class="section-title">Ciro, Saat ve Tahsilat Analizi</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(weekday_df, x="hafta_gunu", y="gunluk_ortalama_ciro", title="Hafta Gününe Göre Ortalama Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        fig = px.bar(hourly_df, x="saat", y="ciro", title="Saatlere Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        tah = period_df.groupby("tahsilat", as_index=False)["ciro"].sum().sort_values("ciro", ascending=False).head(12)
        fig = px.pie(tah, names="tahsilat", values="ciro", title="Tahsilata Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c4:
        fig = px.line(daily_df, x="gun", y="tahsilat_acigi", markers=True, title="Günlük Tahsilat Açığı")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    risk_cols = ["tarih", "satis_no", "satis_tipi", "tahsilat", "kurum", "doktor", "ciro", "odenen", "tahsilat_acigi", "brut_kar", "kar_marji", "sonlandi"]
    if show_patient_columns:
        risk_cols.insert(4, "hasta")
    risk_df = period_df[(period_df["tahsilat_acigi"] > 0) | (~period_df["sonlandi"])].copy()
    st.dataframe(risk_df[risk_cols].sort_values("tarih", ascending=False), use_container_width=True, hide_index=True)

elif page == "🏥 Kurum & Doktor":
    st.markdown('<div class="section-title">Kurum ve Doktor Performansı</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(kurum_df.head(12), x="ciro", y="kurum", orientation="h", title="İlk 12 Kurum - Ciro")
        st.plotly_chart(apply_plot_theme(fig, height=460), use_container_width=True)
    with c2:
        top_doc_df = doktor_df[doktor_df["doktor"].str.lower() != "nan"].head(12).copy()
        fig = px.bar(top_doc_df, x="ciro", y="doktor", orientation="h", title="İlk 12 Doktor - Ciro")
        st.plotly_chart(apply_plot_theme(fig, height=460), use_container_width=True)

    t1, t2 = st.tabs(["Kurum Detayı", "Doktor Detayı"])
    with t1: st.dataframe(kurum_df, use_container_width=True, hide_index=True)
    with t2: st.dataframe(doktor_df, use_container_width=True, hide_index=True)

elif page == "📥 Rapor":
    st.markdown('<div class="section-title">Excel Raporu</div>', unsafe_allow_html=True)
    report = create_excel_report(product_master, sales_df, period_df, kurum_df, doktor_df, daily_df, weekday_df, hourly_df)
    st.download_button(
        "📥 AYÇA Insight V7.1 Raporunu İndir",
        data=report,
        file_name=f"ayca_insight_v7_rapor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown(
        f"""
        <div class="ai-card">
            <div class="ai-title">Veri Kalitesi Özeti</div>
            <div class="ai-text">
            Envanter dosyası: <b>{len(inventory_df)}</b> barkod · Ürün bazında satış dosyası: <b>{len(product_df)}</b> barkod ·
            Eşleşen barkod: <b>{matched_count}</b> · Eşleşme oranı: <b>{pct_fmt(match_ratio)}</b>.
            Satış hareket dosyası tarih aralığı: <b>{sales_df['tarih'].min().strftime('%d.%m.%Y')}</b> - <b>{sales_df['tarih'].max().strftime('%d.%m.%Y')}</b>.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
