# ============================================================
# AYÇA Insight V5.1 - İki Excel Uyumlu Demo
# Satış Excel + Envanter Excel ayrı ayrı yüklenir.
# ------------------------------------------------------------
# Beklenen dosyalar:
# 1) Satış Exceli: sats_all.xlsx benzeri TEBEOS satış özeti
#    Örnek kolonlar: Satış No, Satış Tipi, Tahsilat, Toplam Tutar,
#    Ödenen Tutar, Kar Tutarı, Maliyet Tutarı, İşlem Tarihi, Sonlandı
#
# 2) Envanter Exceli: Envanter_...xlsx benzeri stok listesi
#    Örnek kolonlar: Barkod, Ürün Adı, Psf, Kamu, Stok, Kritik Stok,
#    Raf, Mal Top(Kdv Hariç), Mal Top(Kdv Dahil)
#
# Çalıştırma:
# pip install streamlit pandas numpy openpyxl plotly
# streamlit run app.py
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
    page_title="AYÇA Insight V5.1",
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
    .block-container {padding-top: 1.4rem; max-width: 1520px;}
    .ayca-header {display:flex; justify-content:space-between; align-items:center; gap:16px; margin-bottom:16px;}
    .ayca-title h1 {margin:0; color:var(--text); font-size:32px; letter-spacing:-.7px; font-weight:950;}
    .ayca-title p {margin:6px 0 0 0; color:var(--muted); font-size:14px;}
    .header-pill {background:#fff; border:1px solid var(--border); border-radius:16px; padding:12px 16px; box-shadow:0 8px 24px rgba(15,23,42,.05); color:var(--text); font-weight:900; font-size:13px; white-space:nowrap;}
    .metric-card {min-height:138px; background:linear-gradient(180deg,#fff 0%,#F8FBFF 100%); border:1px solid var(--border); border-radius:20px; padding:18px; box-shadow:0 12px 30px rgba(15,23,42,.06); overflow:hidden; position:relative;}
    .metric-label {color:var(--muted); font-size:12px; font-weight:900; letter-spacing:.4px; text-transform:uppercase; margin-bottom:10px;}
    .metric-value {color:var(--text); font-size:29px; font-weight:950; margin-bottom:8px; letter-spacing:-.5px;}
    .metric-sub {color:var(--muted); font-size:13px;}
    .metric-up {color:var(--green); font-size:13px; font-weight:900;}
    .metric-down {color:var(--red); font-size:13px; font-weight:900;}
    .mini-card {background:var(--panel); border:1px solid var(--border); border-radius:18px; padding:16px; box-shadow:0 10px 26px rgba(15,23,42,.05); min-height:104px;}
    .mini-title {color:var(--text); font-size:14px; font-weight:900; margin-bottom:8px;}
    .mini-value {color:var(--text); font-size:24px; font-weight:950; margin-bottom:4px;}
    .mini-note {color:var(--muted); font-size:13px;}
    .alert-red {background:linear-gradient(135deg,#fff 0%,var(--red-soft) 100%); border-color:#FECACA;}
    .alert-orange {background:linear-gradient(135deg,#fff 0%,var(--orange-soft) 100%); border-color:#FDE68A;}
    .alert-green {background:linear-gradient(135deg,#fff 0%,var(--green-soft) 100%); border-color:#BBF7D0;}
    .alert-blue {background:linear-gradient(135deg,#fff 0%,var(--blue-soft) 100%); border-color:#BFDBFE;}
    .alert-purple {background:linear-gradient(135deg,#fff 0%,var(--purple-soft) 100%); border-color:#DDD6FE;}
    .ai-card {background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%); border:1px solid #BFDBFE; border-radius:22px; padding:20px; box-shadow:0 12px 30px rgba(37,99,235,.08); margin:14px 0 18px 0;}
    .ai-title {color:var(--blue); font-size:18px; font-weight:950; margin-bottom:8px;}
    .ai-text {color:var(--text); font-size:14px; line-height:1.55;}
    .section-title {color:var(--text); font-size:21px; font-weight:950; margin:20px 0 12px 0; letter-spacing:-.3px;}
    .soft-panel {background:var(--panel); border:1px solid var(--border); border-radius:20px; padding:16px; box-shadow:0 12px 30px rgba(15,23,42,.05);}
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
                <div class="ai-title">AYÇA Insight V5.1</div>
                <div class="ai-text">
                Bu sürüm iki ayrı TEBEOS Excel çıktısını okur: <b>Satış</b> ve <b>Envanter</b>.
                Satış dosyasından ciro/kâr/tahsilat; envanter dosyasından stok, kritik stok ve stok değeri analiz edilir.
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
    df["kurum"] = raw_df[mapping["kurum"]].astype(str) if mapping["kurum"] else "Bilinmiyor"
    df["grup"] = raw_df[mapping["grup"]].astype(str) if mapping["grup"] else "Bilinmiyor"
    df["ciro"] = pd.to_numeric(raw_df[mapping["toplam"]], errors="coerce").fillna(0)
    df["odenen"] = pd.to_numeric(raw_df[mapping["odenen"]], errors="coerce").fillna(0) if mapping["odenen"] else df["ciro"]
    df["iskonto"] = pd.to_numeric(raw_df[mapping["iskonto"]], errors="coerce").fillna(0) if mapping["iskonto"] else 0
    df["elden_toplam"] = pd.to_numeric(raw_df[mapping["elde_toplam"]], errors="coerce").fillna(0) if mapping["elde_toplam"] else 0
    df["brut_kar"] = pd.to_numeric(raw_df[mapping["kar"]], errors="coerce").fillna(0)
    df["maliyet"] = pd.to_numeric(raw_df[mapping["maliyet"]], errors="coerce").fillna(0)
    df["fiyat_farki"] = pd.to_numeric(raw_df[mapping["fiyat_farki"]], errors="coerce").fillna(0) if mapping["fiyat_farki"] else 0
    df["sonlandi"] = raw_df[mapping["sonlandi"]].astype(bool) if mapping["sonlandi"] else True
    df["tarih"] = excel_serial_to_datetime(raw_df[mapping["islem_tarihi"]])
    df["recete_tarihi"] = excel_serial_to_datetime(raw_df[mapping["recete_tarihi"]]) if mapping["recete_tarihi"] else pd.NaT
    df["alım_tarihi"] = excel_serial_to_datetime(raw_df[mapping["alım_tarihi"]]) if mapping["alım_tarihi"] else pd.NaT
    df["kullanici"] = raw_df[mapping["kullanici"]].astype(str) if mapping["kullanici"] else ""
    df = df.dropna(subset=["tarih"])
    df["gun"] = df["tarih"].dt.date
    df["ay"] = df["tarih"].dt.to_period("M").astype(str)
    df["kar_marji"] = np.where(df["ciro"] > 0, df["brut_kar"] / df["ciro"], 0)
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
    df["fazla_stok_mu"] = df["stok"] >= 100
    df["stok_durumu"] = np.select(
        [df["stok"] <= 0, df["kritik_mi"], df["fazla_stok_mu"]],
        ["Stok Yok", "Kritik", "Yüksek Stok"],
        default="Normal",
    )
    return df


# ============================================================
# SKOR VE YORUM
# ============================================================
def ayca_score(sales_df: pd.DataFrame, inventory_df: pd.DataFrame, current_margin: float) -> int:
    inv_total = max(1, len(inventory_df))
    critical_ratio = len(inventory_df[inventory_df["kritik_mi"]]) / inv_total
    no_stock_ratio = len(inventory_df[inventory_df["stok"] <= 0]) / inv_total
    high_stock_value = inventory_df.loc[inventory_df["fazla_stok_mu"], "stok_degeri"].sum()
    total_stock_value = max(1, inventory_df["stok_degeri"].sum())
    high_stock_ratio = high_stock_value / total_stock_value

    score = 100
    score -= critical_ratio * 30
    score -= no_stock_ratio * 20
    score -= high_stock_ratio * 18
    if current_margin < 0.10:
        score -= 18
    elif current_margin < 0.18:
        score -= 8
    return int(max(0, min(100, round(score))))


def score_status(score_value: int) -> str:
    if score_value >= 80:
        return "Güçlü"
    if score_value >= 60:
        return "Takip edilmeli"
    return "Riskli"


def create_ai_comment(sales_df, inv_df, critical_df, no_stock_df, high_stock_df, current_ciro, previous_ciro, current_margin):
    messages = []
    if previous_ciro > 0:
        rate = (current_ciro - previous_ciro) / previous_ciro
        messages.append(f"Seçilen dönemde ciro önceki eş döneme göre {pct_fmt(abs(rate))} {'artmış' if rate >= 0 else 'azalmış'} görünüyor.")
    if not critical_df.empty:
        row = critical_df.sort_values("stok").iloc[0]
        messages.append(f"{row['urun']} ürünü kritik stokta; mevcut stok {num_fmt(row['stok'],0)}.")
    if not no_stock_df.empty:
        messages.append(f"Stok sıfır görünen {len(no_stock_df)} ürün var; satış kaçırma riski oluşturabilir.")
    if not high_stock_df.empty:
        messages.append(f"Yüksek stok grubunda yaklaşık {money_fmt(high_stock_df['stok_degeri'].sum())} bağlı para var.")
    if current_margin > 0:
        messages.append(f"Güncel brüt kâr marjı {pct_fmt(current_margin)} seviyesinde.")
    if not messages:
        messages.append("Genel tablo dengeli görünüyor; kritik stok ve nakit bağlayan stoklar düzenli takip edilmeli.")
    return " ".join(messages)


def create_excel_report(sales_df, inv_df, critical_df, no_stock_df, high_stock_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sales_df.to_excel(writer, sheet_name="Satis_Standart", index=False)
        inv_df.to_excel(writer, sheet_name="Envanter_Standart", index=False)
        critical_df.to_excel(writer, sheet_name="Kritik_Stok", index=False)
        no_stock_df.to_excel(writer, sheet_name="Stok_Yok", index=False)
        high_stock_df.to_excel(writer, sheet_name="Yuksek_Stok", index=False)
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
st.sidebar.caption("V5.1 · Satış + Envanter")
eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")

sales_file = st.sidebar.file_uploader("1) Satış Excel dosyasını yükle", type=["xlsx", "xls"], key="sales_file")
inventory_file = st.sidebar.file_uploader("2) Envanter Excel dosyasını yükle", type=["xlsx", "xls"], key="inventory_file")

st.sidebar.markdown("---")
selected_period = st.sidebar.selectbox("Satış Dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Tüm veri"], index=2)
high_stock_limit = st.sidebar.slider("Yüksek stok adedi eşiği", 20, 500, 100)
show_patient_columns = st.sidebar.checkbox("Hasta isim kolonunu göster", value=False)
st.sidebar.caption("Not: Hasta TC ve kişisel sağlık verisi analiz dışı tutulmalıdır. Hasta adı varsayılan olarak gösterilmez.")


# ============================================================
# DOSYA KONTROL
# ============================================================
if sales_file is None or inventory_file is None:
    st.markdown(
        f"""
        <div class="ayca-header">
            <div class="ayca-title">
                <h1>AYÇA Insight V5.1</h1>
                <p>{eczane_adi} · Satış ve envanter dosyalarını ayrı ayrı yükle.</p>
            </div>
            <div class="header-pill">Dosya bekleniyor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    with c1:
        make_mini_card("1. Satış Exceli", "Zorunlu", "sats_all.xlsx formatı: ciro, kâr, tahsilat, işlem tarihi", "alert-blue")
    with c2:
        make_mini_card("2. Envanter Exceli", "Zorunlu", "stok, kritik stok, raf, stok değeri", "alert-green")
    st.info("Sol menüden iki dosyayı da yüklediğinde dashboard otomatik çalışır.")
    st.stop()

try:
    raw_sales, sales_sheet, sales_sheets = read_excel_first_sheet(sales_file)
    raw_inventory, inv_sheet, inv_sheets = read_excel_first_sheet(inventory_file)
    sales_df = standardize_sales(raw_sales)
    inventory_df = standardize_inventory(raw_inventory)
except Exception as exc:
    st.error(f"Dosya okunurken hata oluştu: {exc}")
    st.stop()

inventory_df["fazla_stok_mu"] = inventory_df["stok"] >= high_stock_limit
inventory_df["stok_durumu"] = np.select(
    [inventory_df["stok"] <= 0, inventory_df["kritik_mi"], inventory_df["fazla_stok_mu"]],
    ["Stok Yok", "Kritik", "Yüksek Stok"],
    default="Normal",
)


# ============================================================
# SATIŞ DÖNEMİ
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

current_ciro = period_df["ciro"].sum()
previous_ciro = prev_df["ciro"].sum()
current_profit = period_df["brut_kar"].sum()
previous_profit = prev_df["brut_kar"].sum()
current_margin = current_profit / current_ciro if current_ciro > 0 else 0
previous_margin = previous_profit / previous_ciro if previous_ciro > 0 else 0
current_transactions = period_df["satis_no"].nunique()
previous_transactions = prev_df["satis_no"].nunique()
avg_basket = current_ciro / current_transactions if current_transactions else 0
prev_avg_basket = previous_ciro / previous_transactions if previous_transactions else 0

critical_df = inventory_df[inventory_df["kritik_mi"] & (inventory_df["stok"] > 0)].copy()
no_stock_df = inventory_df[inventory_df["stok"] <= 0].copy()
high_stock_df = inventory_df[inventory_df["fazla_stok_mu"]].copy()
score = ayca_score(sales_df, inventory_df, current_margin)


# ============================================================
# HEADER + KPI
# ============================================================
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-header">
        <div class="ayca-title">
            <h1>AYÇA Insight V5.1</h1>
            <p>{eczane_adi} · {selected_period} · Satış: {sales_sheet} · Envanter: {inv_sheet} · {today_str}</p>
        </div>
        <div class="header-pill">AYÇA Skoru: {score}/100</div>
    </div>
    """,
    unsafe_allow_html=True,
)

ciro_trend, ciro_class = rate_fmt(current_ciro, previous_ciro)
profit_trend, profit_class = rate_fmt(current_profit, previous_profit)
margin_trend, margin_class = rate_fmt(current_margin, previous_margin)
basket_trend, basket_class = rate_fmt(avg_basket, prev_avg_basket)

k1, k2, k3, k4, k5 = st.columns(5)
with k1: make_metric_card("Güncel Ciro", money_fmt(current_ciro), selected_period, ciro_trend, ciro_class)
with k2: make_metric_card("Brüt Kâr", money_fmt(current_profit), "Satış Exceli", profit_trend, profit_class)
with k3: make_metric_card("Kâr Marjı", pct_fmt(current_margin), "Brüt kâr / ciro", margin_trend, margin_class)
with k4: make_metric_card("İşlem Sayısı", f"{current_transactions}", "Tekil satış no")
with k5: make_metric_card("Ortalama Sepet", money_fmt(avg_basket), "Ciro / işlem", basket_trend, basket_class)


# ============================================================
# EXECUTIVE ÖZET
# ============================================================
ai_comment = create_ai_comment(sales_df, inventory_df, critical_df, no_stock_df, high_stock_df, current_ciro, previous_ciro, current_margin)
health_items = {
    "Kârlılık": max(0, min(100, int(60 + current_margin * 160))),
    "Stok Yönetimi": max(0, min(100, int(100 - (len(critical_df) / max(1, len(inventory_df)) * 70) - (len(no_stock_df) / max(1, len(inventory_df)) * 40)))),
    "Nakit Verimliliği": max(0, min(100, int(100 - (high_stock_df["stok_degeri"].sum() / max(1, inventory_df["stok_degeri"].sum()) * 80)))),
    "Satış Performansı": max(0, min(100, int(70 + (((current_ciro - previous_ciro) / previous_ciro) * 40 if previous_ciro else 0))))
}
health_html = "".join([
    f'<div class="health-row"><div class="health-head"><span>{k}</span><span>{v}/100</span></div><div class="health-bar-bg"><div class="health-bar-fill" style="width:{v}%;"></div></div></div>'
    for k, v in health_items.items()
])

st.markdown(
    f"""
    <div class="exec-grid">
        <div class="exec-card">
            <div class="exec-title">🤖 Günaydın {kullanici_adi}</div>
            <div class="exec-sub">AYÇA iki Excel dosyasını beraber yorumladı.</div>
            <div class="exec-list-item">{ai_comment}</div>
            <div class="exec-list-item">📦 Toplam ürün: <b>{len(inventory_df)}</b> · Toplam stok değeri: <b>{money_fmt(inventory_df['stok_degeri'].sum())}</b></div>
            <div class="exec-list-item">🔴 Kritik stok: <b>{len(critical_df)}</b> · ⛔ Stok yok: <b>{len(no_stock_df)}</b> · 📦 Yüksek stok: <b>{len(high_stock_df)}</b></div>
        </div>
        <div class="exec-card">
            <div class="exec-sub">Eczane Sağlık Skoru</div>
            <div class="score-big">{score}</div>
            <div class="exec-sub">Durum: <b>{score_status(score)}</b></div>
            {health_html}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

r1, r2, r3, r4 = st.columns(4)
with r1: make_mini_card("Kritik Stok", str(len(critical_df)), "Stok kritik seviyede", "alert-red" if len(critical_df) else "alert-green")
with r2: make_mini_card("Stok Yok", str(len(no_stock_df)), "Satış kaçırma riski", "alert-red" if len(no_stock_df) else "alert-green")
with r3: make_mini_card("Yüksek Stok", str(len(high_stock_df)), f"{money_fmt(high_stock_df['stok_degeri'].sum())} bağlı para", "alert-orange" if len(high_stock_df) else "alert-green")
with r4: make_mini_card("Toplam Stok Değeri", money_fmt(inventory_df["stok_degeri"].sum()), "Mal Top KDV dahil / PSF toplam", "alert-blue")


# ============================================================
# SAYFALAR
# ============================================================
pages = ["🏠 Genel", "📈 Satış Analizi", "📦 Envanter", "🔴 Kritik Stok", "💰 Kârlılık", "📥 Rapor"]
page = st.radio("Bölüm", pages, horizontal=True, label_visibility="collapsed")

if page == "🏠 Genel":
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Günlük Ciro Trendi</div>', unsafe_allow_html=True)
        daily = period_df.groupby("gun", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique"))
        fig = px.line(daily, x="gun", y="ciro", markers=True, title="Günlük Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        st.markdown('<div class="section-title">Tahsilat Dağılımı</div>', unsafe_allow_html=True)
        tah = period_df.groupby("tahsilat", as_index=False)["ciro"].sum().sort_values("ciro", ascending=False).head(10)
        fig = px.pie(tah, names="tahsilat", values="ciro", title="Tahsilata Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    st.markdown('<div class="section-title">Bugünün Öncelikleri</div>', unsafe_allow_html=True)
    todo = []
    if not critical_df.empty:
        row = critical_df.sort_values("stok").iloc[0]
        todo.append(f"🔴 {row['urun']} kritik stokta. Stok: {num_fmt(row['stok'],0)}")
    if not no_stock_df.empty:
        todo.append(f"⛔ {len(no_stock_df)} ürün stokta yok. En çok ihtiyaç duyulan ürünler manuel kontrol edilmeli.")
    if not high_stock_df.empty:
        row = high_stock_df.sort_values("stok_degeri", ascending=False).iloc[0]
        todo.append(f"📦 {row['urun']} yüksek stok değerinde öne çıkıyor. Bağlı para: {money_fmt(row['stok_degeri'])}")
    if current_margin < 0.12:
        todo.append("💰 Kâr marjı düşük. İndirim, maliyet ve satış fiyatı kontrol edilmeli.")
    if not todo:
        todo.append("✅ Kritik aksiyon görünmüyor. Günlük satış ve stok dengesi takip edilmeli.")
    for t in todo[:5]:
        st.markdown(f"<div class='exec-list-item'>{t}</div>", unsafe_allow_html=True)

elif page == "📈 Satış Analizi":
    st.markdown('<div class="section-title">Satış Tipi ve Kurum Analizi</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        sales_type = period_df.groupby("satis_tipi", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), islem=("satis_no", "nunique")).sort_values("ciro", ascending=False)
        fig = px.bar(sales_type, x="satis_tipi", y="ciro", title="Satış Tipine Göre Ciro")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        kurum = period_df.groupby("kurum", as_index=False)["ciro"].sum().sort_values("ciro", ascending=False).head(10)
        fig = px.bar(kurum, x="ciro", y="kurum", orientation="h", title="İlk 10 Kurum")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    show_cols = ["tarih", "satis_no", "satis_tipi", "tahsilat", "kurum", "ciro", "brut_kar", "maliyet", "kar_marji", "sonlandi"]
    if show_patient_columns:
        show_cols.insert(4, "hasta")
    st.dataframe(period_df[show_cols].sort_values("tarih", ascending=False), use_container_width=True, hide_index=True)

elif page == "📦 Envanter":
    st.markdown('<div class="section-title">Envanter Genel Görünüm</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        status = inventory_df.groupby("stok_durumu", as_index=False).agg(urun_sayisi=("urun", "count"), stok_degeri=("stok_degeri", "sum"))
        fig = px.bar(status, x="stok_durumu", y="urun_sayisi", title="Stok Durumu Dağılımı")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        top_stock = inventory_df.sort_values("stok_degeri", ascending=False).head(15)
        fig = px.bar(top_stock, x="stok_degeri", y="urun", orientation="h", title="Stok Değeri En Yüksek 15 Ürün")
        st.plotly_chart(apply_plot_theme(fig, height=440), use_container_width=True)
    inv_cols = ["barkod", "urun", "stok", "kritik_stok", "stok_durumu", "psf", "kamu", "raf", "kdv", "stok_degeri"]
    st.dataframe(inventory_df[inv_cols].sort_values("stok_degeri", ascending=False), use_container_width=True, hide_index=True)

elif page == "🔴 Kritik Stok":
    st.markdown('<div class="section-title">Kritik Stok ve Stok Yok Listeleri</div>', unsafe_allow_html=True)
    t1, t2, t3 = st.tabs(["Kritik Stok", "Stok Yok", "Yüksek Stok"])
    with t1:
        st.dataframe(critical_df[["barkod", "urun", "stok", "kritik_stok", "psf", "raf", "stok_degeri"]].sort_values("stok"), use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(no_stock_df[["barkod", "urun", "stok", "kritik_stok", "psf", "raf"]].sort_values("urun"), use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(high_stock_df[["barkod", "urun", "stok", "psf", "raf", "stok_degeri"]].sort_values("stok_degeri", ascending=False), use_container_width=True, hide_index=True)

elif page == "💰 Kârlılık":
    st.markdown('<div class="section-title">Kârlılık ve Nakit Akışı</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        daily = period_df.groupby("gun", as_index=False).agg(ciro=("ciro", "sum"), kar=("brut_kar", "sum"), maliyet=("maliyet", "sum"))
        daily["marj"] = np.where(daily["ciro"] > 0, daily["kar"] / daily["ciro"], 0)
        fig = px.line(daily, x="gun", y="marj", markers=True, title="Günlük Brüt Marj")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    with c2:
        profit_by_type = period_df.groupby("satis_tipi", as_index=False).agg(kar=("brut_kar", "sum"), ciro=("ciro", "sum"))
        profit_by_type["marj"] = np.where(profit_by_type["ciro"] > 0, profit_by_type["kar"] / profit_by_type["ciro"], 0)
        fig = px.bar(profit_by_type, x="satis_tipi", y="kar", title="Satış Tipine Göre Brüt Kâr")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
    st.dataframe(profit_by_type.sort_values("kar", ascending=False), use_container_width=True, hide_index=True)

elif page == "📥 Rapor":
    st.markdown('<div class="section-title">Excel Raporu</div>', unsafe_allow_html=True)
    report = create_excel_report(sales_df, inventory_df, critical_df, no_stock_df, high_stock_df)
    st.download_button(
        "📥 AYÇA Insight Raporunu İndir",
        data=report,
        file_name=f"ayca_insight_iki_excel_rapor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown(
        """
        <div class="ai-card">
            <div class="ai-title">Bu sürümde bilinçli ayrım</div>
            <div class="ai-text">
            Satış Exceliniz ürün satır detayı içermediği için barkod/ürün bazlı satış hızı hesaplanmaz.
            Bu nedenle miad ve ürün bazlı son 60 gün çıkış analizi kapalıdır. Satış dosyasından ciro/kâr/tahsilat;
            envanter dosyasından stok/kritik stok/stok değeri analiz edilir.
            Ürün bazlı satış detayı olan ayrı bir satış kalemleri raporu eklendiğinde sistem barkod bazında stok bitiş tahmini de yapabilir.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
