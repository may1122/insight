# ============================================================
# AYÇA Insight V10.7 Copilot Slim Health Analysis - Kontrol Merkezi + Asistan + Ürün Fırsatları
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
import html
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
    page_title="AYÇA Insight V10.7 Copilot Slim Health Analysis",
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

    .saas-shell {background:linear-gradient(135deg,#0F172A 0%,#1E3A8A 48%,#2563EB 100%); border-radius:30px; padding:22px; box-shadow:0 22px 60px rgba(37,99,235,.20); margin:12px 0 18px 0; color:white; overflow:hidden; position:relative;}
    .saas-shell:after {content:""; position:absolute; width:360px; height:360px; right:-120px; top:-150px; background:radial-gradient(circle,rgba(255,255,255,.22),rgba(255,255,255,0)); border-radius:999px;}
    .saas-hero-title {font-size:28px; font-weight:950; letter-spacing:-.7px; margin-bottom:8px; color:white;}
    .saas-hero-sub {font-size:14px; line-height:1.58; color:#DBEAFE; max-width:880px;}
    .saas-badge {display:inline-block; background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.22); border-radius:999px; padding:7px 11px; font-size:12px; font-weight:900; margin:0 6px 8px 0; color:white;}
    .saas-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; margin:14px 0 18px 0;}
    .saas-tile {background:rgba(255,255,255,.92); border:1px solid #E2E8F0; border-radius:22px; padding:17px; box-shadow:0 12px 30px rgba(15,23,42,.06); min-height:132px;}
    .saas-tile-k {font-size:12px; color:#64748B; font-weight:950; text-transform:uppercase; letter-spacing:.45px; margin-bottom:9px;}
    .saas-tile-v {font-size:26px; color:#0F172A; font-weight:950; letter-spacing:-.5px; margin-bottom:7px;}
    .saas-tile-n {font-size:13px; color:#64748B; line-height:1.45;}
    .module-grid {display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:14px; margin:12px 0 18px 0;}
    .module-card {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:22px; padding:18px; box-shadow:0 12px 30px rgba(15,23,42,.055); min-height:150px;}
    .module-title {font-size:17px; font-weight:950; color:#0F172A; margin-bottom:8px;}
    .module-desc {font-size:13px; color:#64748B; line-height:1.55;}

    .login-wrap {min-height:82vh; display:grid; grid-template-columns:1.05fr .95fr; gap:28px; align-items:center; padding:26px 0;}
    .login-hero {background:linear-gradient(135deg,#0F172A 0%,#1E3A8A 54%,#2563EB 100%); border-radius:34px; padding:34px; color:white; box-shadow:0 26px 70px rgba(37,99,235,.24); position:relative; overflow:hidden; min-height:520px;}
    .login-hero:after {content:""; position:absolute; width:420px; height:420px; right:-150px; top:-150px; background:radial-gradient(circle,rgba(255,255,255,.22),rgba(255,255,255,0)); border-radius:999px;}
    .login-logo {font-size:42px; font-weight:950; letter-spacing:-1.2px; line-height:1; color:white; margin-bottom:8px; position:relative; z-index:2;}
    .login-logo span {display:block; font-size:20px; font-weight:850; color:#BFDBFE; letter-spacing:.2px; margin-top:8px;}
    .login-title {font-size:28px; font-weight:950; letter-spacing:-.6px; margin:32px 0 10px 0; color:white; position:relative; z-index:2;}
    .login-sub {font-size:15px; line-height:1.65; color:#DBEAFE; max-width:590px; position:relative; z-index:2;}
    .login-features {display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-top:26px; position:relative; z-index:2;}
    .login-feature {background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.20); border-radius:18px; padding:13px 14px; font-size:13px; font-weight:850; color:#F8FAFC;}
    .login-footer-badges {display:flex; flex-wrap:wrap; gap:8px; margin-top:28px; position:relative; z-index:2;}
    .login-badge {background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.22); border-radius:999px; padding:8px 12px; font-size:12px; font-weight:900; color:white;}
    .login-card {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:30px; padding:30px; box-shadow:0 24px 70px rgba(15,23,42,.10);}
    .login-card h2 {margin:0 0 6px 0; font-size:28px; font-weight:950; letter-spacing:-.6px; color:#0F172A;}
    .login-card p {margin:0 0 22px 0; color:#64748B; font-size:14px; line-height:1.55;}
    .demo-box {background:#F8FAFC; border:1px solid #E2E8F0; border-radius:20px; padding:15px; margin-top:16px; color:#334155; font-size:13px; line-height:1.6;}
    .brief-grid {display:grid; grid-template-columns:1.05fr .95fr; gap:16px; margin:16px 0;}
    .brief-hero {background:linear-gradient(135deg,#0F172A 0%,#1E3A8A 56%,#2563EB 100%); border-radius:30px; padding:24px; color:white; box-shadow:0 20px 54px rgba(37,99,235,.18); overflow:hidden; position:relative;}
    .brief-hero:after {content:""; position:absolute; width:320px; height:320px; right:-110px; top:-120px; background:radial-gradient(circle,rgba(255,255,255,.20),rgba(255,255,255,0)); border-radius:999px;}
    .brief-title {font-size:28px; font-weight:950; letter-spacing:-.7px; color:white; margin-bottom:8px; position:relative; z-index:2;}
    .brief-sub {color:#DBEAFE; font-size:14px; line-height:1.55; position:relative; z-index:2; max-width:760px;}
    .brief-score {font-size:64px; font-weight:950; letter-spacing:-2px; color:white; line-height:1; margin-top:18px; position:relative; z-index:2;}
    .brief-score-label {font-size:13px; font-weight:900; color:#BFDBFE; position:relative; z-index:2;}
    .brief-health-box {position:relative; z-index:2; margin-top:18px; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px;}
    .brief-health-col {background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.18); border-radius:16px; padding:12px; backdrop-filter:blur(8px);}
    .brief-health-title {font-size:13px; font-weight:950; color:#FFFFFF; margin-bottom:8px;}
    .brief-health-item {font-size:12px; line-height:1.35; color:#E0F2FE; margin:7px 0; font-weight:780;}
    .brief-health-result {position:relative; z-index:2; margin-top:14px; background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.16); border-radius:16px; padding:12px; color:#FFFFFF; font-size:13px; line-height:1.45; font-weight:800;}
    .copilot-question-grid {display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; margin:12px 0 16px 0;}
    .copilot-question {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:16px; padding:12px; font-weight:900; color:#0F172A; box-shadow:0 8px 22px rgba(15,23,42,.045);}
    @media (max-width:1000px){.brief-health-box,.copilot-question-grid{grid-template-columns:1fr;}}
    .brief-panel {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:28px; padding:20px; box-shadow:0 16px 44px rgba(15,23,42,.07);}
    .group-title {font-size:19px; font-weight:950; color:#0F172A; margin:22px 0 12px 0; letter-spacing:-.25px;}
    .group-grid {display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin-bottom:14px;}
    .compact-card {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:20px; padding:16px; box-shadow:0 10px 28px rgba(15,23,42,.045); min-height:112px;}
    .compact-k {font-size:11px; text-transform:uppercase; letter-spacing:.45px; color:#64748B; font-weight:950; margin-bottom:9px;}
    .compact-v {font-size:25px; color:#0F172A; font-weight:950; letter-spacing:-.45px; margin-bottom:6px;}
    .compact-n {font-size:12px; color:#64748B; line-height:1.4;}
    .task-list {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:24px; padding:18px; box-shadow:0 14px 36px rgba(15,23,42,.055);}
    .task-item {display:flex; gap:10px; align-items:flex-start; border-bottom:1px solid #EEF2F7; padding:11px 0; color:#0F172A; font-size:14px; font-weight:760; line-height:1.45;}
    .task-item:last-child {border-bottom:0;}
    .risk-summary-card {background:#FFFFFF; border:1px solid #E2E8F0; border-radius:24px; padding:18px; box-shadow:0 14px 36px rgba(15,23,42,.055);}
    .risk-line {display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid #EEF2F7; color:#0F172A; font-size:14px; font-weight:850;}
    .risk-line:last-child {border-bottom:0;}
    .risk-line span:last-child {font-size:17px; font-weight:950;}
    @media (max-width:1000px){.login-wrap,.brief-grid{grid-template-columns:1fr;}.login-features,.group-grid{grid-template-columns:1fr 1fr;}}

    @media (max-width:1000px){.saas-grid{grid-template-columns:repeat(2,minmax(0,1fr));}.module-grid{grid-template-columns:1fr;}}
    @media (max-width:1000px){.exec-grid{grid-template-columns:1fr;}.ayca-header{display:block}.header-pill{display:inline-block;margin-top:10px}}
    </style>
    """,
    unsafe_allow_html=True,
)




# ============================================================
# GİRİŞ EKRANI CSS - TAM EKRAN SAAS
# ============================================================
st.markdown(
    """
    <style>
    .login-shell {
        min-height: calc(100vh - 70px);
        width: min(1680px, calc(100vw - 48px));
        margin-left: 50%;
        transform: translateX(-50%);
        display:grid;
        grid-template-columns: 1.12fr .88fr;
        gap:28px;
        align-items:center;
        padding: 24px 18px 42px 18px;
    }
    .login-left {
        background: radial-gradient(circle at top left, rgba(255,255,255,.28), rgba(255,255,255,0) 42%),
                    linear-gradient(135deg,#0F172A 0%,#1E3A8A 52%,#2563EB 100%);
        border-radius:34px;
        min-height:620px;
        padding:42px;
        color:white;
        box-shadow:0 30px 80px rgba(37,99,235,.24);
        position:relative;
        overflow:hidden;
    }
    .login-left:after {
        content:"";
        position:absolute;
        width:420px;
        height:420px;
        right:-130px;
        top:-130px;
        border-radius:999px;
        background:radial-gradient(circle,rgba(255,255,255,.24),rgba(255,255,255,0));
    }
    .login-brand {
        font-size:38px;
        font-weight:950;
        letter-spacing:-1.3px;
        color:#FFFFFF;
        margin-bottom:8px;
    }
    .login-brand span {font-size:20px; font-weight:800; color:#BFDBFE; margin-left:8px;}
    .login-headline {
        margin-top:70px;
        max-width:680px;
        font-size:42px;
        line-height:1.08;
        font-weight:950;
        letter-spacing:-1.45px;
        color:white;
    }
    .login-sub {
        max-width:640px;
        margin-top:18px;
        font-size:16px;
        line-height:1.66;
        color:#DBEAFE;
    }
    .login-feature-grid {
        display:grid;
        grid-template-columns:repeat(2,minmax(0,1fr));
        gap:12px;
        margin-top:34px;
        max-width:720px;
    }
    .login-feature {
        background:rgba(255,255,255,.12);
        border:1px solid rgba(255,255,255,.18);
        border-radius:18px;
        padding:14px 15px;
        color:white;
        font-weight:850;
        font-size:14px;
        backdrop-filter: blur(10px);
    }
    .login-version {
        display:inline-flex;
        margin-top:36px;
        background:rgba(255,255,255,.13);
        border:1px solid rgba(255,255,255,.20);
        color:#FFFFFF;
        border-radius:999px;
        padding:10px 14px;
        font-size:13px;
        font-weight:900;
    }
    .login-right {
        background:#FFFFFF;
        border:1px solid #E2E8F0;
        border-radius:34px;
        padding:34px;
        min-height:520px;
        box-shadow:0 30px 80px rgba(15,23,42,.09);
    }
    .login-card-title {
        font-size:28px;
        font-weight:950;
        letter-spacing:-.7px;
        color:#0F172A;
        margin-bottom:6px;
    }
    .login-card-sub {
        font-size:14px;
        color:#64748B;
        line-height:1.55;
        margin-bottom:22px;
    }
    .demo-box {
        margin-top:22px;
        background:#F8FAFC;
        border:1px solid #E2E8F0;
        border-radius:20px;
        padding:16px;
        color:#334155;
        font-size:13px;
        line-height:1.55;
    }
    .demo-box b {color:#0F172A;}
    @media (max-width: 980px) {
        .login-shell {grid-template-columns:1fr; padding:10px 0 30px 0;}
        .login-left {min-height:auto; padding:28px;}
        .login-headline {margin-top:34px; font-size:32px;}
        .login-feature-grid {grid-template-columns:1fr;}
    }
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
    """Tam ekran ama Streamlit widget uyumlu giriş ekranı.

    Not: Streamlit widget'ları açık HTML div'lerinin içine güvenilir şekilde yerleşmediği için
    giriş formu st.columns ile ayrıldı. Böylece sağ panel boş kalmaz ve form her ortamda görünür.
    """
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {display:none;}
        .block-container {padding-top:0.6rem; max-width:1720px;}
        .login-clean-shell {
            min-height: calc(100vh - 34px);
            display:grid;
            grid-template-columns:1.08fr .92fr;
            gap:28px;
            align-items:center;
            padding:24px 8px 34px 8px;
        }
        .login-clean-left {
            min-height:640px;
            border-radius:36px;
            padding:44px;
            color:#FFFFFF;
            background:
                radial-gradient(circle at 8% 8%, rgba(255,255,255,.28), rgba(255,255,255,0) 34%),
                linear-gradient(135deg,#0F172A 0%,#1E3A8A 54%,#2563EB 100%);
            box-shadow:0 34px 90px rgba(37,99,235,.24);
            position:relative;
            overflow:hidden;
        }
        .login-clean-left:after {
            content:"";
            position:absolute;
            width:440px;
            height:440px;
            right:-140px;
            top:-140px;
            border-radius:999px;
            background:radial-gradient(circle,rgba(255,255,255,.24),rgba(255,255,255,0));
        }
        .login-clean-brand {
            position:relative;
            z-index:2;
            font-size:42px;
            font-weight:950;
            letter-spacing:-1.3px;
            color:#FFFFFF;
            margin-bottom:14px;
        }
        .login-clean-brand span {
            font-size:21px;
            font-weight:850;
            color:#BFDBFE;
            margin-left:8px;
        }
        .login-welcome {
            position:relative;
            z-index:2;
            margin-top:54px;
            font-size:18px;
            color:#DBEAFE;
            font-weight:850;
        }
        .login-clean-headline {
            position:relative;
            z-index:2;
            margin-top:12px;
            max-width:720px;
            font-size:44px;
            line-height:1.08;
            font-weight:950;
            letter-spacing:-1.5px;
            color:#FFFFFF;
        }
        .login-clean-sub {
            position:relative;
            z-index:2;
            max-width:680px;
            margin-top:18px;
            font-size:16px;
            line-height:1.65;
            color:#DBEAFE;
        }
        .login-clean-features {
            position:relative;
            z-index:2;
            display:grid;
            grid-template-columns:repeat(2,minmax(0,1fr));
            gap:12px;
            margin-top:34px;
            max-width:760px;
        }
        .login-clean-feature {
            background:rgba(255,255,255,.12);
            border:1px solid rgba(255,255,255,.20);
            border-radius:18px;
            padding:14px 15px;
            color:#FFFFFF;
            font-weight:850;
            font-size:14px;
            backdrop-filter:blur(10px);
        }
        .login-clean-version {
            position:relative;
            z-index:2;
            display:inline-flex;
            margin-top:36px;
            background:rgba(255,255,255,.13);
            border:1px solid rgba(255,255,255,.22);
            color:#FFFFFF;
            border-radius:999px;
            padding:10px 14px;
            font-size:13px;
            font-weight:900;
        }
        .login-form-card {
            background:#FFFFFF;
            border:1px solid #E2E8F0;
            border-radius:34px;
            padding:34px;
            box-shadow:0 30px 80px rgba(15,23,42,.09);
        }
        .login-form-title {
            font-size:26px;
            font-weight:950;
            letter-spacing:-.6px;
            color:#0F172A;
            margin-bottom:6px;
        }
        .login-form-sub {
            font-size:14px;
            color:#64748B;
            line-height:1.55;
            margin-bottom:18px;
        }
        @media (max-width:980px) {
            .login-clean-shell {grid-template-columns:1fr; padding:8px 0 24px 0;}
            .login-clean-left {min-height:auto; padding:28px;}
            .login-clean-headline {font-size:32px;}
            .login-clean-features {grid-template-columns:1fr;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.08, 0.92], gap="large")

    with left:
        st.markdown(
            """
            <div class="login-clean-left">
                <div class="login-clean-brand">AYÇA <span>Insight</span></div>
                <div class="login-welcome">Hoş geldiniz 👋</div>
                <div class="login-clean-headline">Eczanenizin dijital operasyon merkezi</div>
                <div class="login-clean-sub">
                    Stok, kârlılık, reçete süreçleri ve finansal riskleri tek ekranda yönetin.
                    AYÇA Insight, TEBEOS verilerinizi karar destek katmanına dönüştürür.
                </div>
                <div class="login-clean-features">
                    <div class="login-clean-feature">✓ Sipariş karar desteği</div>
                    <div class="login-clean-feature">✓ KKİ risk takibi</div>
                    <div class="login-clean-feature">✓ Kırmızı / Yeşil reçete yönetimi</div>
                    <div class="login-clean-feature">✓ Kârlılık ve tahsilat analizi</div>
                    <div class="login-clean-feature">✓ Stok ve sermaye kontrolü</div>
                    <div class="login-clean-feature">✓ Yapay zekâ destekli içgörüler</div>
                </div>
                <div class="login-clean-version">AYÇA Insight V10.6 · Copilot Health Analysis</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown(
            """
            <div class="login-form-card">
                <div class="login-form-title">Giriş bilgileri</div>
                <div class="login-form-sub">AYÇA Insight'a giriş yaparak sabah brifingini ve operasyon merkezini açın.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        username = st.text_input("Kullanıcı adı", value="premium")
        password = st.text_input("Şifre", value="premium2026", type="password")
        if st.button("🚀 Giriş Yap", use_container_width=True, type="primary"):
            record = DEMO_USERS.get(username.strip().lower())
            if record and password == record["password"]:
                st.session_state["authenticated"] = True
                st.session_state["auth_user"] = record["name"]
                st.session_state["auth_pharmacy"] = record["pharmacy"]
                st.session_state["membership"] = record["membership"]
                safe_rerun()
            else:
                st.error("Kullanıcı adı veya şifre hatalı. Premium: premium / premium2026")

        st.markdown(
            """
            <div class="demo-box">
                <b>Demo hesaplar</b><br>
                Premium: <b>premium</b> / <b>premium2026</b><br>
                Basic: <b>basic</b> / <b>basic2026</b>
            </div>
            """,
            unsafe_allow_html=True,
        )


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
# RİSK REFERANS MOTORU - KIRMIZI / YEŞİL / EK İZLEM / KKİ
# ============================================================
def normalize_product_key(value: str) -> str:
    value = str(value or "").upper().strip()
    tr_map = str.maketrans("ÇĞİÖŞÜÂÎÛ", "CGIOSUAIU")
    value = value.translate(tr_map)
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _risk_rows_from_keywords(keywords, risk_tipi, risk_seviyesi, kaynak, aciklama=""):
    return [
        {
            "barkod": "",
            "urun_anahtar": normalize_product_key(k),
            "risk_tipi": risk_tipi,
            "risk_seviyesi": risk_seviyesi,
            "kaynak": kaynak,
            "aciklama": aciklama,
            "kki_fark_tl": 0.0,
        }
        for k in keywords
    ]


@st.cache_data(show_spinner=False)
def build_builtin_risk_reference() -> pd.DataFrame:
    """Demo/başlangıç risk sözlüğü.

    Not: Canlı kullanımda bu tablo bir Excel master dosyasından beslenmelidir.
    Buradaki liste PDF'lerden çıkan marka/ürün anahtarlarıyla ilk görünümü oluşturur.
    Barkodlu eşleşme için sol menüde 'Risk Master Excel' yüklenebilir.
    """
    red = [
        "ABSTRAL", "ACTIQ", "ALDINE", "ALDOLAN", "CEDEPTIN", "CONCERTA", "DUROGESIC",
        "EFFENTORA", "FENTANYL", "FENTANYL CITRATE", "FENTANEST", "FENTAVER", "JURNISTA",
        "KONSENIDAT", "MEDIKINET", "M ESLON", "M-ESLON", "MORFIA", "MORFIN", "MORPHINE",
        "OPIVA", "OXOPANE", "PETHIDINE", "PETHOLAN", "RAPIFEN", "RENTANIL", "RITALIN",
        "SPRAVATO", "SUBOXONE", "SUFENTA", "TALINAT", "ULTIVA", "XYREM",
    ]
    green = [
        "AKINETON", "ALYSE", "ANSIOX", "APO ALPRAZ", "AS ALPRALAM", "ATIVAN", "CODEFEN",
        "CONTRAMAL", "DALIZOM", "DEMIZOLAM", "DIAPAM", "DIAZEM", "DIAZEPAM DESITIN",
        "DORMICUM", "DUAMOL", "EKIPENTAL", "FENOKODIN", "FIXDOL", "GALARA", "GERICA",
        "HYPNOMIDATE", "IMOVANE", "KETALAR", "KLIPAKS", "LIBKOL", "LIZAN", "LUMINAL",
        "LUMINALETTEN", "LYPRE", "LYRICA", "MADOL", "MIDOLAM", "MILOZ", "NEOGABA",
        "NERVIUM", "NEURICA", "PADEN", "PAGADIN", "PAGAMAX", "PENTAL", "PERGE",
        "PIREPSIL", "PRECOBAL", "PRELICA", "PREPLUS", "REGAPEN", "RIVOTRIL", "ROLADOL",
        "SEDOZOLAM", "SNAPLINE", "SPESICOR", "STABINA", "STABLON", "SYMRA", "TRADOLEX",
        "TRAMADOLOR", "TRAMOSEL", "TRANXILENE", "ULTRAMEX", "XANAX", "ZALDIAR", "ZENIXA", "ZOLAMID",
    ]
    ek_izlem = [
        "XARELTO", "ELIQUIS", "JANUVIA", "FORXIGA", "JARDIANCE", "OZEMPIC", "RYBELSUS", "SAXENDA",
        "TRULICITY", "KEYTRUDA", "OPDIVO", "TECENTRIQ", "AVASTIN", "HERCEPTIN", "HUMIRA", "COSENTYX",
    ]
    kki_hesaplamali = [
        "ACNEGEN", "ALDARA", "BENZADERM", "BRUNAC", "CONVULEX", "DAFLON", "DECAPEPTYL",
        "DIAZEPAM DESITIN", "ENDOXAN", "FAMODIN", "HEPARINUM", "HIBOR", "KALINOR", "KARBALEX",
        "LEUNASE", "LH RH", "LOKALEN", "LOTEMAX", "MAXICAINE", "METHOTREXATE EBEWE", "NITROLINGUAL",
        "NOKREV", "OCTOSTIM", "OLICLINOMEL", "OVADRIL", "PEN OS", "R X SUSPANSIYON", "SALOFALK",
        "SUBOXONE", "TODAVIT", "TRACUTIL", "TURKFLEKS", "ULTRACAIN", "URSOFALK", "VEINDOCANOL", "VIRGAN",
    ]
    kki_2 = [
        "CALQUENCE", "CETROTIDE", "FASLODEX", "GENOTROPIN", "GONAL F", "IRESSA", "JAKAVI", "LYNPARZA",
        "MEKINIST", "NUBEQA", "OVITRELLE", "TAFINLAR", "TAGRISSO", "TRODELVY", "TYKERB", "VALAMOR",
        "VENCLYXTO", "VERXANT", "ZEJULA", "ZOLADEX",
    ]
    kki_3 = ["IMBRUVICA", "NGENLA"]

    rows = []
    rows += _risk_rows_from_keywords(red, "KIRMIZI_RECETE", "Yüksek", "Kırmızı reçete PDF", "Kontrollü kırmızı reçete ürünü")
    rows += _risk_rows_from_keywords(green, "YESIL_RECETE", "Orta-Yüksek", "Yeşil reçete PDF", "Kontrollü yeşil reçete ürünü")
    rows += _risk_rows_from_keywords(ek_izlem, "EK_IZLEM", "Orta", "TİTCK Ek İzlem", "Ek izlem / özel takip adayı")
    rows += _risk_rows_from_keywords(kki_hesaplamali, "KKI_HESAPLAMALI", "Finansal Risk", "İEO KKİ Listesi", "KKİ eksik / hesaplamalı ilaç")
    rows += _risk_rows_from_keywords(kki_2, "KKI_2_YILDIZ", "Takip", "İEO KKİ Listesi", "Firma KKİ farkını takip eden ay içinde öder")
    rows += _risk_rows_from_keywords(kki_3, "KKI_3_YILDIZ", "Bilgi", "İEO KKİ Listesi", "Firma KKİ farkını eczacıya öder")

    # PDF'de net görülen bazı barkodlar. Barkod varsa isimden daha güvenilir eşleşir.
    rows += [
        {"barkod":"8699769950071", "urun_anahtar":"HIBOR", "risk_tipi":"KKI_HESAPLAMALI", "risk_seviyesi":"Finansal Risk", "kaynak":"İEO KKİ Listesi", "aciklama":"KKİ fark riski", "kki_fark_tl":-320.45},
        {"barkod":"8699510050012", "urun_anahtar":"SUBOXONE", "risk_tipi":"KKI_HESAPLAMALI", "risk_seviyesi":"Finansal Risk", "kaynak":"İEO KKİ Listesi", "aciklama":"KKİ fark riski", "kki_fark_tl":-285.08},
        {"barkod":"8699510050029", "urun_anahtar":"SUBOXONE", "risk_tipi":"KKI_HESAPLAMALI", "risk_seviyesi":"Finansal Risk", "kaynak":"İEO KKİ Listesi", "aciklama":"KKİ fark riski", "kki_fark_tl":-883.19},
        {"barkod":"8699510030175", "urun_anahtar":"CONVULEX", "risk_tipi":"KKI_HESAPLAMALI", "risk_seviyesi":"Finansal Risk", "kaynak":"İEO KKİ Listesi", "aciklama":"KKİ fark riski", "kki_fark_tl":-66.50},
        {"barkod":"8699510030168", "urun_anahtar":"CONVULEX", "risk_tipi":"KKI_HESAPLAMALI", "risk_seviyesi":"Finansal Risk", "kaynak":"İEO KKİ Listesi", "aciklama":"KKİ fark riski", "kki_fark_tl":-108.51},
        {"barkod":"8699786092860", "urun_anahtar":"TAGRISSO", "risk_tipi":"KKI_2_YILDIZ", "risk_seviyesi":"Takip", "kaynak":"İEO KKİ Listesi", "aciklama":"Firma KKİ farkını takip eden ay içinde öder", "kki_fark_tl":0.0},
        {"barkod":"8699786092877", "urun_anahtar":"TAGRISSO", "risk_tipi":"KKI_2_YILDIZ", "risk_seviyesi":"Takip", "kaynak":"İEO KKİ Listesi", "aciklama":"Firma KKİ farkını takip eden ay içinde öder", "kki_fark_tl":0.0},
    ]
    ref = pd.DataFrame(rows).drop_duplicates(["barkod", "urun_anahtar", "risk_tipi"])
    return ref


def standardize_risk_reference(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Kullanıcının yükleyeceği risk master Excel/CSV için esnek kolon okuyucu."""
    cols = list(raw_df.columns)
    barkod_col = find_col(cols, ["Barkod", "Barcode", "GTIN"])
    name_col = find_col(cols, ["Ürün Adı", "Urun Adi", "İlaç Adı", "Ilac Adi", "urun_anahtar", "İlaç"])
    risk_col = find_col(cols, ["Risk Tipi", "risk_tipi", "Kategori", "Liste", "Reçete Tipi", "Recete Tipi"])
    level_col = find_col(cols, ["Risk Seviyesi", "risk_seviyesi", "Seviye", "Önem"])
    source_col = find_col(cols, ["Kaynak", "Source"])
    note_col = find_col(cols, ["Açıklama", "Aciklama", "Not", "Açıklamalar"])
    fark_col = find_col(cols, ["Fark", "KKİ Fark", "KKI Fark", "kki_fark_tl", "Zarar"])

    df = pd.DataFrame()
    df["barkod"] = clean_text_series(raw_df[barkod_col]).str.replace(r"\.0$", "", regex=True).str.strip() if barkod_col else ""
    df["urun_anahtar"] = clean_text_series(raw_df[name_col]).apply(normalize_product_key) if name_col else ""
    df["risk_tipi"] = clean_text_series(raw_df[risk_col]).str.upper().str.strip() if risk_col else "OZEL_RISK"
    df["risk_tipi"] = df["risk_tipi"].replace({
        "KIRMIZI": "KIRMIZI_RECETE", "YEŞİL": "YESIL_RECETE", "YESIL": "YESIL_RECETE",
        "EK IZLEM": "EK_IZLEM", "EK İZLEM": "EK_IZLEM", "KKI": "KKI_HESAPLAMALI", "KKİ": "KKI_HESAPLAMALI",
    })
    df["risk_seviyesi"] = clean_text_series(raw_df[level_col]).str.strip() if level_col else "Takip"
    df["kaynak"] = clean_text_series(raw_df[source_col]).str.strip() if source_col else "Kullanıcı Risk Master"
    df["aciklama"] = clean_text_series(raw_df[note_col]).str.strip() if note_col else ""
    df["kki_fark_tl"] = pd.to_numeric(raw_df[fark_col], errors="coerce").fillna(0) if fark_col else 0.0
    df = df[(df["barkod"].astype(str).str.len() > 0) | (df["urun_anahtar"].astype(str).str.len() > 0)]
    return df.drop_duplicates(["barkod", "urun_anahtar", "risk_tipi"])


@st.cache_data(show_spinner=False)
def load_uploaded_risk_reference(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raw = pd.read_excel(uploaded_file)
    return standardize_risk_reference(raw)


def append_label(existing: str, new_value: str) -> str:
    parts = [p for p in str(existing or "").split(";") if p]
    if new_value not in parts:
        parts.append(new_value)
    return ";".join(parts)


def apply_risk_reference(product_master: pd.DataFrame, risk_ref: pd.DataFrame) -> pd.DataFrame:
    df = product_master.copy()
    for col in ["KIRMIZI_RECETE", "YESIL_RECETE", "EK_IZLEM", "KKI_HESAPLAMALI", "KKI_2_YILDIZ", "KKI_3_YILDIZ", "OZEL_RISK"]:
        df[col.lower() + "_mi"] = False
    df["urun_key"] = df["urun"].apply(normalize_product_key)
    df["risk_tipi"] = ""
    df["risk_kaynak"] = ""
    df["risk_aciklama"] = ""
    df["kki_birim_fark_tl"] = 0.0

    if risk_ref is None or risk_ref.empty:
        return df

    ref = risk_ref.copy()
    for _, row in ref.iterrows():
        risk_type = str(row.get("risk_tipi", "OZEL_RISK") or "OZEL_RISK").upper().strip()
        flag_col = risk_type.lower() + "_mi"
        if flag_col not in df.columns:
            df[flag_col] = False
        barkod = str(row.get("barkod", "") or "").replace(".0", "").strip()
        key = normalize_product_key(row.get("urun_anahtar", ""))
        mask = pd.Series(False, index=df.index)
        if barkod:
            mask = mask | (df["barkod"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() == barkod)
        if key:
            # Marka/anahtar bazlı güvenli içerme. Örn: LYRICA 150 MG 56 KAPSUL.
            mask = mask | df["urun_key"].str.contains(r"(^| )" + re.escape(key) + r"( |$)", regex=True, na=False)
        if not mask.any():
            continue
        df.loc[mask, flag_col] = True
        df.loc[mask, "risk_tipi"] = df.loc[mask, "risk_tipi"].apply(lambda x: append_label(x, risk_type))
        kaynak = str(row.get("kaynak", "") or "")
        aciklama = str(row.get("aciklama", "") or "")
        df.loc[mask, "risk_kaynak"] = df.loc[mask, "risk_kaynak"].apply(lambda x: append_label(x, kaynak))
        df.loc[mask, "risk_aciklama"] = df.loc[mask, "risk_aciklama"].apply(lambda x: append_label(x, aciklama))
        fark = float(row.get("kki_fark_tl", 0) or 0)
        if fark != 0:
            df.loc[mask, "kki_birim_fark_tl"] = np.where(df.loc[mask, "kki_birim_fark_tl"] == 0, fark, df.loc[mask, "kki_birim_fark_tl"])

    df["kontrollu_recete_mi"] = df.get("kirmizi_recete_mi", False) | df.get("yesil_recete_mi", False)
    df["kki_riskli_mi"] = df.get("kki_hesaplamali_mi", False) | df.get("kki_2_yildiz_mi", False) | df.get("kki_3_yildiz_mi", False)
    df["risk_var_mi"] = df["risk_tipi"].astype(str).str.len() > 0
    df["kki_tahmini_fark_tl"] = df["kki_birim_fark_tl"] * pd.to_numeric(df.get("satilan_adet", 0), errors="coerce").fillna(0)

    # Pahalı/seyrek ürün filtresi, kontrollü reçete veya KKİ riskli ürünleri 'reçete geldikçe al' içine gömmesin.
    if "recete_geldikce_al_mi" in df.columns:
        df.loc[df["kontrollu_recete_mi"] | df["kki_riskli_mi"], "recete_geldikce_al_mi"] = False
    df["kontrollu_takip_mi"] = df["kontrollu_recete_mi"] & (df.get("teknik_siparis_onerisi_ham", 0) > 0)

    df["risk_segmenti"] = np.select(
        [
            df.get("kirmizi_recete_mi", False),
            df.get("yesil_recete_mi", False),
            df.get("kki_hesaplamali_mi", False),
            df.get("kki_2_yildiz_mi", False),
            df.get("kki_3_yildiz_mi", False),
            df.get("ek_izlem_mi", False),
        ],
        ["Kırmızı Reçete", "Yeşil Reçete", "KKİ Hesaplamalı", "KKİ 2 Yıldız", "KKİ 3 Yıldız", "Ek İzlem"],
        default="Normal",
    )
    df.loc[df["kontrollu_recete_mi"] & (df["siparis_segmenti"].isin(["Acil Sipariş", "Öncelikli Sipariş"])), "siparis_segmenti"] = "Kontrollü Reçete Takibi"
    df.loc[df["kki_riskli_mi"] & (df["aksiyon"] == "Normal takip"), "aksiyon"] = "KKİ risk merkezi - finansal takip"
    return df


def risk_summary_table(product_master: pd.DataFrame) -> pd.DataFrame:
    risk_cols = [
        ("Kırmızı Reçete", "kirmizi_recete_mi"),
        ("Yeşil Reçete", "yesil_recete_mi"),
        ("Ek İzlem", "ek_izlem_mi"),
        ("KKİ Hesaplamalı", "kki_hesaplamali_mi"),
        ("KKİ 2 Yıldız", "kki_2_yildiz_mi"),
        ("KKİ 3 Yıldız", "kki_3_yildiz_mi"),
    ]
    rows = []
    for label, col in risk_cols:
        if col not in product_master.columns:
            continue
        subset = product_master[product_master[col]].copy()
        rows.append({
            "Risk Grubu": label,
            "Ürün Sayısı": int(len(subset)),
            "Stok Değeri": float(subset["stok_degeri"].sum()) if not subset.empty else 0.0,
            "Satış Tutarı": float(subset["satis_tutari"].sum()) if not subset.empty else 0.0,
            "Satılan Adet": float(subset["satilan_adet"].sum()) if not subset.empty else 0.0,
            "Tahmini KKİ Farkı": float(subset.get("kki_tahmini_fark_tl", pd.Series(dtype=float)).sum()) if not subset.empty else 0.0,
        })
    return pd.DataFrame(rows)


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
        # Bazı TEBEOS satış hareketi çıktılarında ürün/barkod kalemi de bulunabilir.
        # Varsa Doktor Intelligence içinde doktor-ilaç tercihi analizi yapılır. Yoksa modül bilgi notu gösterir.
        "barkod_hareket": find_col(cols, ["Barkod", "Ürün Barkodu", "Urun Barkodu"]),
        "urun_hareket": find_col(cols, ["Ürün Adı", "Urun Adi", "İlaç Adı", "Ilac Adi", "Malzeme Adı"]),
        "adet_hareket": find_col(cols, ["Adet", "Miktar", "Satılan Adet", "Satilan Adet"]),
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
    df["barkod_hareket"] = clean_text_series(raw_df[mapping["barkod_hareket"]]).str.replace(r"\.0$", "", regex=True).str.strip() if mapping["barkod_hareket"] else ""
    df["urun_hareket"] = clean_text_series(raw_df[mapping["urun_hareket"]]).str.strip() if mapping["urun_hareket"] else ""
    df["adet_hareket"] = pd.to_numeric(raw_df[mapping["adet_hareket"]], errors="coerce").fillna(1) if mapping["adet_hareket"] else 1
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

    # Teknik ham öneri: matematiksel olarak stok hedefinin altında kalan tüm ürünler.
    # Eczacı filtresi: pahalı ve seyrek satılan özel/akıllı ilaçlar otomatik acil siparişe düşmesin.
    # Kural: Ürün bu analiz döneminde 10 adet ve üzeri sattıysa VEYA birim satış fiyatı 15.000 TL altındaysa sipariş listesine alınır.
    # Aksi halde ürün ayrı olarak "Reçete geldikçe alınabilir" segmentine düşer.
    master["teknik_siparis_onerisi_ham"] = np.maximum(0, master["hedef_stok"] - master["stok"]).round(0)
    master["teknik_siparis_tahmini_tutar_ham"] = master["teknik_siparis_onerisi_ham"] * master["birim_satis"].fillna(0)
    master["siparis_filtre_gecer_mi"] = (master["satilan_adet"] >= 10) | (master["birim_satis"] < 15000)
    master["recete_geldikce_al_mi"] = (master["teknik_siparis_onerisi_ham"] > 0) & (~master["siparis_filtre_gecer_mi"])
    master["siparis_onerisi_ham"] = np.where(master["siparis_filtre_gecer_mi"], master["teknik_siparis_onerisi_ham"], 0).round(0)
    master["siparis_tahmini_tutar_ham"] = master["siparis_onerisi_ham"] * master["birim_satis"].fillna(0)

    # Varsayılan final öneri, bütçe kontrolünden önce filtrelenmiş ham öneriye eşittir.
    # Aşağıda apply_order_budget() ile mevcut stok değerinin belirli oranına göre orantılı küçültülür.
    master["siparis_onerisi"] = master["siparis_onerisi_ham"].copy()
    master["siparis_tahmini_tutar"] = master["siparis_tahmini_tutar_ham"].copy()
    master["siparis_kisit_katsayisi"] = 1.0
    master["siparis_butce_kisildi_mi"] = False
    master["kritik_mi"] = np.where(master["kritik_stok"] > 0, master["stok"] <= master["kritik_stok"], master["stok"] <= 1)
    master["stok_yok_mu"] = master["stok"] <= 0
    master["hizli_tukeniyor_mu"] = (master["stok_bitis_gunu"] <= 14) & (master["satilan_adet"] > 0)
    master["siparis_gerekli_mi"] = master["siparis_onerisi"] > 0
    master["olu_stok_mu"] = (master["satilan_adet"] <= 0) & (master["stok"] > 0)
    master["yavas_stok_mu"] = (master["satilan_adet"] > 0) & (master["stok_bitis_gunu"] > 90) & (master["stok"] > 0)
    master["sermaye_riski_mi"] = (master["stok_degeri"] > master["stok_degeri"].quantile(0.90)) & (master["stok_bitis_gunu"] > 60)
    master["stokta_yok_satmis_mi"] = (master["satilan_adet"] > 0) & (master["stok"] <= 0)
    master["acil_siparis_mi"] = master["siparis_gerekli_mi"] & master["siparis_filtre_gecer_mi"] & (
        master["stokta_yok_satmis_mi"] | (master["hizli_tukeniyor_mu"] & (master["satilan_adet"] >= 10)) | (master["kritik_mi"] & (master["satilan_adet"] > 0))
    )
    master["oncelikli_siparis_mi"] = master["siparis_gerekli_mi"] & master["siparis_filtre_gecer_mi"] & (~master["acil_siparis_mi"])
    master["siparis_segmenti"] = np.select(
        [master["acil_siparis_mi"], master["oncelikli_siparis_mi"], master["recete_geldikce_al_mi"]],
        ["Acil Sipariş", "Öncelikli Sipariş", "Reçete Geldikçe Al"],
        default="Normal Takip",
    )

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
            master["acil_siparis_mi"],
            master["oncelikli_siparis_mi"],
            master["recete_geldikce_al_mi"],
            master["olu_stok_mu"],
            master["yavas_stok_mu"],
            master["sermaye_riski_mi"],
        ],
        [
            "Acil sipariş / satış kaçırma riski",
            "Öncelikli sipariş",
            "Reçete geldikçe al - pahalı/seyrek ürün",
            "Ölü stok - aksiyon al",
            "Yavaş stok - kampanya / raf kontrolü",
            "Sermaye riski - stok azalt",
        ],
        default="Normal takip",
    )
    return master




def build_patient_loyalty(period_df: pd.DataFrame, all_sales_df: pd.DataFrame) -> dict:
    """Hasta sadakat analizleri. KVKK nedeniyle hasta adı varsayılan olarak maskelenebilir."""
    df = all_sales_df.copy()
    if "hasta" not in df.columns:
        empty = pd.DataFrame()
        return {"summary": {}, "frequency": empty, "basket": empty, "lost": empty, "vip": empty}

    df["hasta_clean"] = df["hasta"].astype(str).str.strip()
    df = df[(df["hasta_clean"] != "") & (df["hasta_clean"].str.lower() != "nan") & (df["hasta_clean"].str.lower() != "bilinmiyor")]
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
    if df is None or df.empty or name_col not in df.columns:
        return df
    out = df.copy()
    unique_names = {name: f"Hasta-{i+1:04d}" for i, name in enumerate(out[name_col].dropna().astype(str).unique())}
    out[name_col] = out[name_col].map(unique_names).fillna("Hasta")
    return out


def build_doctor_intelligence(period_df: pd.DataFrame, all_sales_df: pd.DataFrame, product_master: Optional[pd.DataFrame] = None) -> dict:
    """Doktor bazlı ciro/kârlılık/trend ve varsa doktor-ürün tercih analizi."""
    df = all_sales_df.copy()
    df["doktor_clean"] = df["doktor"].astype(str).str.strip()
    df = df[(df["doktor_clean"] != "") & (df["doktor_clean"].str.lower() != "nan") & (df["doktor_clean"].str.lower() != "bilinmiyor")]
    empty = pd.DataFrame()
    if df.empty:
        return {"doctor_kpi": empty, "doctor_month": empty, "doctor_growth": empty, "doctor_institution": empty, "doctor_products": empty, "doctor_product_note": "Doktor bilgisi bulunamadı."}

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

    # Doktorların en çok tercih ettiği ilaçlar: Sadece satış hareketi dosyasında ürün/barkod bilgisi varsa gerçek hesaplanır.
    doctor_products = pd.DataFrame()
    product_note = "Satış hareketleri dosyasında ürün/barkod kalemi bulunmadığı için doktor-ilaç tercihi hesaplanamadı. TEBEOS'tan reçete kalemi/satış kalemi detay raporu gerekir."
    if "urun_hareket" in period_df.columns and period_df["urun_hareket"].astype(str).str.strip().replace("", np.nan).notna().any():
        dp = period_df.copy()
        dp["doktor_clean"] = dp["doktor"].astype(str).str.strip()
        dp["urun_tercih"] = dp["urun_hareket"].astype(str).str.strip()
        dp = dp[(dp["doktor_clean"] != "") & (dp["doktor_clean"].str.lower() != "nan") & (dp["urun_tercih"] != "") & (dp["urun_tercih"].str.lower() != "nan")]
        if not dp.empty:
            doctor_products = dp.groupby(["doktor_clean", "urun_tercih"], as_index=False).agg(
                adet=("adet_hareket", "sum"),
                ciro=("ciro", "sum"),
                kar=("brut_kar", "sum"),
                recete_sayisi=("satis_no", "nunique"),
            )
            doctor_products["marj"] = np.where(doctor_products["ciro"] > 0, doctor_products["kar"] / doctor_products["ciro"], 0)
            top_docs = doctor_kpi.head(20)["doktor_clean"].tolist()
            doctor_products = doctor_products[doctor_products["doktor_clean"].isin(top_docs)]
            doctor_products = doctor_products.sort_values(["doktor_clean", "ciro", "adet"], ascending=[True, False, False])
            doctor_products["doktor_urun_sira"] = doctor_products.groupby("doktor_clean").cumcount() + 1
            doctor_products = doctor_products[doctor_products["doktor_urun_sira"] <= 10]
            product_note = "Doktor-ilaç tercihleri satış hareketindeki ürün kalemlerinden hesaplandı."

    return {
        "doctor_kpi": doctor_kpi,
        "doctor_month": doctor_month,
        "doctor_growth": doctor_growth,
        "doctor_institution": doctor_institution,
        "doctor_products": doctor_products,
        "doctor_product_note": product_note,
    }



def build_business_insights(product_master: pd.DataFrame, current_stats: dict, previous_stats: dict, order_budget_info: dict, score_items: dict) -> dict:
    """Mevcut 3 dosyayla üretilebilen ileri karar destek içgörüleri."""
    pm = product_master.copy()
    if pm.empty:
        empty = pd.DataFrame()
        return {
            "carrier_50": empty, "carrier_80": empty, "silent_loss": empty, "lost_product_candidates": empty,
            "assistant_cards": [], "health_notes": [], "summary": {}
        }

    # Eczaneyi taşıyan ürünler: ciro kümülatif payına göre %50 ve %80 eşiği.
    carrier = pm[pm["satis_tutari"] > 0].sort_values("satis_tutari", ascending=False).copy()
    total_sales = float(carrier["satis_tutari"].sum())
    if total_sales > 0:
        carrier["kumulatif_ciro"] = carrier["satis_tutari"].cumsum()
        carrier["kumulatif_ciro_payi"] = carrier["kumulatif_ciro"] / total_sales
    else:
        carrier["kumulatif_ciro"] = 0
        carrier["kumulatif_ciro_payi"] = 0
    carrier_50 = carrier[carrier["kumulatif_ciro_payi"] <= 0.50].copy()
    if carrier_50.empty and not carrier.empty:
        carrier_50 = carrier.head(1).copy()
    carrier_80 = carrier[carrier["kumulatif_ciro_payi"] <= 0.80].copy()
    if carrier_80.empty and not carrier.empty:
        carrier_80 = carrier.head(1).copy()

    # Sessiz kâr kaybı: çok satan/ciro yapan ama marjı düşük ürünler.
    active = pm[(pm["satilan_adet"] > 0) & (pm["satis_tutari"] > 0)].copy()
    if not active.empty:
        sales_q75 = active["satis_tutari"].quantile(0.75)
        margin_median = active["kar_marji"].median()
        margin_limit = min(0.12, margin_median) if not pd.isna(margin_median) else 0.12
        silent_loss = active[(active["satis_tutari"] >= sales_q75) & (active["kar_marji"] <= margin_limit)].copy()
        silent_loss["potansiyel_marj_farki"] = np.maximum(0, 0.15 - silent_loss["kar_marji"])
        silent_loss["tahmini_sessiz_kayip"] = silent_loss["satis_tutari"] * silent_loss["potansiyel_marj_farki"]
        silent_loss = silent_loss.sort_values(["tahmini_sessiz_kayip", "satis_tutari"], ascending=False)
    else:
        silent_loss = pd.DataFrame()

    # Kaybedilen ürünler: mevcut tek ürün toplam raporuyla gerçek trend hesaplanamaz.
    # Bu nedenle burada stokta para bağlayan fakat dönem satış hızı zayıf/0 olan ürünler 'kaybedilen ürün adayı' olarak listelenir.
    lost_candidates = pm[((pm["olu_stok_mu"]) | (pm["yavas_stok_mu"])) & (pm["stok_degeri"] > 0)].copy()
    lost_candidates["kayip_urun_notu"] = np.where(
        lost_candidates.get("olu_stok_mu", False),
        "Dönemde satış yok, stok elde bekliyor",
        "Satış hızı zayıf, stok uzun süre yetecek"
    )
    lost_candidates = lost_candidates.sort_values("stok_degeri", ascending=False)

    # Sağlık karnesi yorumları.
    health_notes = []
    if score_items:
        lowest = sorted(score_items.items(), key=lambda x: x[1])[0]
        health_notes.append(f"Bu ay puanı en çok düşüren alan: {lowest[0]} ({lowest[1]}/100).")
    dead_value = float(pm.loc[pm["olu_stok_mu"], "stok_degeri"].sum())
    total_stock = float(pm["stok_degeri"].sum())
    dead_ratio = safe_div(dead_value, total_stock)
    if dead_ratio > 0.20:
        health_notes.append(f"Ölü stok oranı yüksek: stok değerinin {pct_fmt(dead_ratio)} kadarı satış üretmiyor.")
    elif dead_ratio > 0.05:
        health_notes.append(f"Ölü stok izlenmeli: {money_fmt(dead_value)} tutarında hareketsiz stok var.")
    if current_stats.get("marj", 0) < 0.12:
        health_notes.append(f"Brüt marj düşük görünüyor: {pct_fmt(current_stats.get('marj', 0))}.")
    if current_stats.get("tahsilat_acigi", 0) > 0:
        health_notes.append(f"Tahsilat açığı kontrol edilmeli: {money_fmt(current_stats.get('tahsilat_acigi', 0))}.")
    if not silent_loss.empty:
        health_notes.append(f"Sessiz kâr kaybı adayı {len(silent_loss)} ürün var; ilk ürün {silent_loss.iloc[0]['urun']}.")

    # AYÇA Asistan kartları.
    urgent_count = int(pm.get("acil_siparis_mi", pm.get("stokta_yok_satmis_mi", pd.Series(dtype=bool))).sum())
    fast_count = int(pm.get("hizli_tukeniyor_mu", pd.Series(dtype=bool)).sum())
    reorder_count = int(pm.get("siparis_gerekli_mi", pd.Series(dtype=bool)).sum())
    rx_as_needed_count = int(pm.get("recete_geldikce_al_mi", pd.Series(dtype=bool)).sum())
    top_profit = active.sort_values("kar_tutari", ascending=False).head(1) if not active.empty else pd.DataFrame()
    assistant_cards = []
    if urgent_count:
        first = pm[pm.get("acil_siparis_mi", pm["stokta_yok_satmis_mi"])].sort_values("satis_tutari", ascending=False).iloc[0]
        assistant_cards.append(("🔴", "Acil Sipariş Riski", f"{urgent_count} ürün eczacı filtresinden geçerek acil listeye girdi. İlk kontrol: {first['urun']}."))
    if rx_as_needed_count:
        first = pm[pm["recete_geldikce_al_mi"]].sort_values("birim_satis", ascending=False).iloc[0]
        assistant_cards.append(("🧠", "Reçete Geldikçe Al", f"{rx_as_needed_count} pahalı/seyrek ürün acil listeden ayrıldı. İlk örnek: {first['urun']}."))
    if fast_count:
        first = pm[pm["hizli_tukeniyor_mu"]].sort_values("stok_bitis_gunu").iloc[0]
        day_text = num_fmt(first.get("stok_bitis_gunu_goster", 0), 1)
        assistant_cards.append(("🟠", "Hızlı Tükenen Ürün", f"{fast_count} ürün 14 gün içinde bitebilir. En yakın risk: {first['urun']} ({day_text} gün)."))
    if reorder_count:
        assistant_cards.append(("🛒", "Bütçeli Sipariş", f"{reorder_count} ürün için {money_fmt(order_budget_info.get('final_total', 0))} sipariş öneriliyor."))
    if not silent_loss.empty:
        assistant_cards.append(("📉", "Sessiz Kâr Kaybı", f"{len(silent_loss)} çok satan düşük marjlı ürün var. İlk kontrol: {silent_loss.iloc[0]['urun']}."))
    if not lost_candidates.empty:
        assistant_cards.append(("🧊", "Hareketsiz Sermaye", f"{money_fmt(lost_candidates['stok_degeri'].sum())} tutarında ölü/yavaş stok adayı var."))
    if not top_profit.empty:
        row = top_profit.iloc[0]
        assistant_cards.append(("💎", "En Karlı Ürün", f"{row['urun']} bu dönemde {money_fmt(row['kar_tutari'])} kâr bıraktı."))

    summary = {
        "carrier_50_count": int(len(carrier_50)),
        "carrier_80_count": int(len(carrier_80)),
        "silent_loss_count": int(len(silent_loss)),
        "silent_loss_amount": float(silent_loss["tahmini_sessiz_kayip"].sum()) if not silent_loss.empty else 0.0,
        "lost_candidate_count": int(len(lost_candidates)),
        "lost_candidate_value": float(lost_candidates["stok_degeri"].sum()) if not lost_candidates.empty else 0.0,
        "dead_ratio": dead_ratio,
    }
    return {
        "carrier_50": carrier_50,
        "carrier_80": carrier_80,
        "silent_loss": silent_loss,
        "lost_product_candidates": lost_candidates,
        "assistant_cards": assistant_cards,
        "health_notes": health_notes,
        "summary": summary,
    }


def render_assistant_cards(cards):
    if not cards:
        st.markdown("<div class='exec-list-item'>✅ Kritik aksiyon görünmüyor. Günlük ciro, tahsilat ve stok hızını takip et.</div>", unsafe_allow_html=True)
        return
    for icon, title, text in cards:
        st.markdown(
            f"""
            <div class="exec-list-item" style="display:flex;gap:12px;align-items:flex-start;">
                <div style="font-size:22px;line-height:1;">{icon}</div>
                <div><div style="font-weight:950;color:#0F172A;margin-bottom:3px;">{title}</div><div style="font-weight:700;color:#334155;">{text}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def apply_order_budget(product_master: pd.DataFrame, order_budget_ratio: float) -> tuple[pd.DataFrame, dict]:
    """Sipariş önerisini mevcut stok değerinin seçilen oranına göre gerçekçi bütçeye indirir.

    Örnek: stok değeri 1.400.000 TL ve oran %20 ise maksimum sipariş bütçesi 280.000 TL olur.
    Ham öneri 1.200.000 TL ise tüm öneriler öncelik sırası bozulmadan 280/1200 oranında küçültülür.
    """
    df = product_master.copy()
    total_stock_value = float(pd.to_numeric(df.get("stok_degeri", 0), errors="coerce").fillna(0).sum())
    raw_total = float(pd.to_numeric(df.get("siparis_tahmini_tutar_ham", df.get("siparis_tahmini_tutar", 0)), errors="coerce").fillna(0).sum())
    budget_limit = max(0.0, total_stock_value * float(order_budget_ratio))

    if raw_total <= 0 or budget_limit <= 0:
        scale = 0.0 if raw_total > 0 else 1.0
    else:
        scale = min(1.0, budget_limit / raw_total)

    df["siparis_kisit_katsayisi"] = scale
    df["siparis_butce_kisildi_mi"] = scale < 0.999

    # Orantılı küçültme: adetler aşağı yuvarlanır. Çok kritik ve 0'a düşen ürünler için 1 adetlik minimum korunur.
    raw_qty = pd.to_numeric(df["siparis_onerisi_ham"], errors="coerce").fillna(0)
    scaled_qty = np.floor(raw_qty * scale)
    critical_min_mask = (raw_qty > 0) & (scaled_qty < 1) & (df.get("siparis_filtre_gecer_mi", True)) & (df["stokta_yok_satmis_mi"] | df["hizli_tukeniyor_mu"] | df["kritik_mi"])
    scaled_qty = np.where(critical_min_mask, 1, scaled_qty)

    # scaled_qty np.ndarray dönebilir; Series'e çevirmezsek .fillna() hatası oluşur.
    df["siparis_onerisi"] = pd.Series(scaled_qty, index=df.index)
    df["siparis_onerisi"] = pd.to_numeric(df["siparis_onerisi"], errors="coerce").fillna(0).round(0)
    df["siparis_tahmini_tutar"] = df["siparis_onerisi"] * pd.to_numeric(df["birim_satis"], errors="coerce").fillna(0)
    df["siparis_gerekli_mi"] = df["siparis_onerisi"] > 0

    # Eczacı filtresine göre segmentleri güncelle.
    df["acil_siparis_mi"] = df["siparis_gerekli_mi"] & df.get("siparis_filtre_gecer_mi", True) & (
        df["stokta_yok_satmis_mi"] | (df["hizli_tukeniyor_mu"] & (df["satilan_adet"] >= 10)) | (df["kritik_mi"] & (df["satilan_adet"] > 0))
    )
    df["oncelikli_siparis_mi"] = df["siparis_gerekli_mi"] & df.get("siparis_filtre_gecer_mi", True) & (~df["acil_siparis_mi"])
    df["recete_geldikce_al_mi"] = (df.get("teknik_siparis_onerisi_ham", df["siparis_onerisi_ham"]) > 0) & (~df.get("siparis_filtre_gecer_mi", True))
    df["siparis_segmenti"] = np.select(
        [df["acil_siparis_mi"], df["oncelikli_siparis_mi"], df["recete_geldikce_al_mi"]],
        ["Acil Sipariş", "Öncelikli Sipariş", "Reçete Geldikçe Al"],
        default="Normal Takip",
    )

    # Yuvarlama/minimum 1 adet sebebiyle bütçe az aşılırsa en düşük öncelikli ürünlerden kırp.
    final_total = float(df["siparis_tahmini_tutar"].sum())
    if budget_limit > 0 and final_total > budget_limit:
        df = df.sort_values(
            ["acil_siparis_mi", "oncelikli_siparis_mi", "stokta_yok_satmis_mi", "hizli_tukeniyor_mu", "kritik_mi", "satis_tutari", "kar_tutari"],
            ascending=[False, False, False, False, False, False, False],
        ).copy()
        running = 0.0
        keep_qty = []
        for _, row in df.iterrows():
            amount = float(row.get("siparis_tahmini_tutar", 0) or 0)
            qty = float(row.get("siparis_onerisi", 0) or 0)
            unit = float(row.get("birim_satis", 0) or 0)
            if qty <= 0 or unit <= 0:
                keep_qty.append(0)
                continue
            remaining = budget_limit - running
            if remaining <= 0:
                keep_qty.append(0)
                continue
            if amount <= remaining:
                keep_qty.append(qty)
                running += amount
            else:
                max_qty = np.floor(remaining / unit)
                keep_qty.append(max_qty)
                running += max_qty * unit
        df["siparis_onerisi"] = pd.Series(keep_qty, index=df.index)
        df["siparis_onerisi"] = pd.to_numeric(df["siparis_onerisi"], errors="coerce").fillna(0).round(0)
        df["siparis_tahmini_tutar"] = df["siparis_onerisi"] * pd.to_numeric(df["birim_satis"], errors="coerce").fillna(0)
        df["siparis_gerekli_mi"] = df["siparis_onerisi"] > 0
        df["acil_siparis_mi"] = df["siparis_gerekli_mi"] & df.get("siparis_filtre_gecer_mi", True) & (
            df["stokta_yok_satmis_mi"] | (df["hizli_tukeniyor_mu"] & (df["satilan_adet"] >= 10)) | (df["kritik_mi"] & (df["satilan_adet"] > 0))
        )
        df["oncelikli_siparis_mi"] = df["siparis_gerekli_mi"] & df.get("siparis_filtre_gecer_mi", True) & (~df["acil_siparis_mi"])
        df["recete_geldikce_al_mi"] = (df.get("teknik_siparis_onerisi_ham", df["siparis_onerisi_ham"]) > 0) & (~df.get("siparis_filtre_gecer_mi", True))
        df["siparis_segmenti"] = np.select(
            [df["acil_siparis_mi"], df["oncelikli_siparis_mi"], df["recete_geldikce_al_mi"]],
            ["Acil Sipariş", "Öncelikli Sipariş", "Reçete Geldikçe Al"],
            default="Normal Takip",
        )

    # Aksiyon metnini final öneriye göre yeniden yaz.
    df["aksiyon"] = np.select(
        [
            df["stokta_yok_satmis_mi"],
            df["hizli_tukeniyor_mu"] & df["siparis_gerekli_mi"],
            df["olu_stok_mu"],
            df["yavas_stok_mu"],
            df["sermaye_riski_mi"],
        ],
        [
            "Acil sipariş / satış kaçırma riski",
            "Bütçeli sipariş öner",
            "Ölü stok - aksiyon al",
            "Yavaş stok - kampanya / raf kontrolü",
            "Sermaye riski - stok azalt",
        ],
        default="Normal takip",
    )

    final_total = float(df["siparis_tahmini_tutar"].sum())
    info = {
        "stock_value": total_stock_value,
        "budget_ratio": float(order_budget_ratio),
        "budget_limit": budget_limit,
        "raw_total": raw_total,
        "final_total": final_total,
        "scale": scale,
        "cut_amount": max(0.0, raw_total - final_total),
        "budget_used_ratio": safe_div(final_total, budget_limit),
        "was_limited": raw_total > budget_limit > 0,
    }
    return df.sort_values(["satis_tutari", "satilan_adet"], ascending=False), info


def create_action_items(product_master, current_stats, previous_stats, daily_df):
    actions = []
    urgent = product_master[product_master.get("acil_siparis_mi", product_master["stokta_yok_satmis_mi"])].sort_values("satis_tutari", ascending=False)
    reorder = product_master[product_master["siparis_gerekli_mi"]].sort_values(["acil_siparis_mi", "siparis_tahmini_tutar"], ascending=[False, False])
    rx_as_needed = product_master[product_master.get("recete_geldikce_al_mi", False)].sort_values("birim_satis", ascending=False)
    dead = product_master[product_master["olu_stok_mu"]].sort_values("stok_degeri", ascending=False)
    slow = product_master[product_master["yavas_stok_mu"]].sort_values("stok_degeri", ascending=False)
    profitable = product_master[product_master["satilan_adet"] > 0].sort_values("kar_tutari", ascending=False)

    if not urgent.empty:
        r = urgent.iloc[0]
        actions.append(f"⛔ {len(urgent)} ürün acil sipariş filtresinden geçti. İlk risk: {r['urun']} · satış {num_fmt(r['satilan_adet'],0)} adet.")
    if not rx_as_needed.empty:
        r = rx_as_needed.iloc[0]
        actions.append(f"🧠 {len(rx_as_needed)} pahalı/seyrek ürün acil listeden ayrıldı. Reçete geldikçe alınabilir: {r['urun']} · birim {money_fmt(r['birim_satis'])}.")
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


def create_excel_report(product_master, sales_df, period_df, kurum_df, doktor_df, daily_df, weekday_df, hourly_df, doctor_kpi=None, doctor_products=None, patient_frequency=None, patient_lost=None, business_insights=None):
    output = BytesIO()
    export_cols = [
        "barkod", "urun", "urun_grubu", "raf", "stok", "kritik_stok", "psf_final", "stok_degeri",
        "satilan_adet", "satis_tutari", "kar_tutari", "kar_marji", "gunluk_satis_hizi", "stok_bitis_gunu_goster",
        "risk_segmenti", "risk_tipi", "risk_kaynak", "kki_birim_fark_tl", "kki_tahmini_fark_tl",
        "hedef_stok", "siparis_onerisi", "siparis_tahmini_tutar", "abc_sinif", "aksiyon"
    ]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        product_master[export_cols].to_excel(writer, sheet_name="Urun_Zekasi", index=False)
        product_master[product_master["siparis_gerekli_mi"]][export_cols].to_excel(writer, sheet_name="Siparis_Onerisi", index=False)
        if "recete_geldikce_al_mi" in product_master.columns:
            product_master[product_master["recete_geldikce_al_mi"]][export_cols].to_excel(writer, sheet_name="Recete_Geldikce_Al", index=False)
        if "risk_var_mi" in product_master.columns:
            product_master[product_master["risk_var_mi"]][export_cols].to_excel(writer, sheet_name="Risk_Merkezi", index=False)
        if "kki_riskli_mi" in product_master.columns:
            product_master[product_master["kki_riskli_mi"]][export_cols].to_excel(writer, sheet_name="KKI_Risk", index=False)
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
        if doctor_kpi is not None and not doctor_kpi.empty:
            doctor_kpi.to_excel(writer, sheet_name="Doktor_Intelligence", index=False)
        if doctor_products is not None and not doctor_products.empty:
            doctor_products.to_excel(writer, sheet_name="Doktor_Urun_Tercihleri", index=False)
        if patient_frequency is not None and not patient_frequency.empty:
            patient_frequency.to_excel(writer, sheet_name="Hasta_Sadakat", index=False)
        if patient_lost is not None and not patient_lost.empty:
            patient_lost.to_excel(writer, sheet_name="Kaybedilen_Hastalar", index=False)
        if business_insights:
            if not business_insights.get("carrier_80", pd.DataFrame()).empty:
                business_insights["carrier_80"].to_excel(writer, sheet_name="Tasiyan_Urunler", index=False)
            if not business_insights.get("silent_loss", pd.DataFrame()).empty:
                business_insights["silent_loss"].to_excel(writer, sheet_name="Sessiz_Kar_Kaybi", index=False)
            if not business_insights.get("lost_product_candidates", pd.DataFrame()).empty:
                business_insights["lost_product_candidates"].to_excel(writer, sheet_name="Kaybedilen_Urun_Aday", index=False)
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
st.sidebar.caption("V10.7 Copilot Slim · Sağlık Analizi + Bana Sor")
eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")

inventory_file = st.sidebar.file_uploader("1/3) Envanter Exceli - ZORUNLU", type=["xlsx", "xls"], key="inventory_file")
product_file = st.sidebar.file_uploader("2/3) Ürün Bazında Toplamlar Exceli - ZORUNLU", type=["xlsx", "xls"], key="product_file")
sales_file = st.sidebar.file_uploader("3/3) Satış Hareketleri Exceli - ZORUNLU", type=["xlsx", "xls"], key="sales_file")
risk_master_file = st.sidebar.file_uploader("Opsiyonel) Risk Master Excel / CSV", type=["xlsx", "xls", "csv"], key="risk_master_file", help="Kırmızı/yeşil reçete, ek izlem ve KKİ listelerini barkod veya ürün adıyla içeren master tablo.")

st.sidebar.markdown("---")
selected_period = st.sidebar.selectbox("Satış hareket dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Tüm veri"], index=2)
st.session_state["selected_period_label"] = selected_period
period_label = selected_period
target_days = st.sidebar.slider("Sipariş hedef stok günü", 7, 90, 30)
safety_days = st.sidebar.slider("Güvenlik stok günü", 0, 30, 7)
order_policy = st.sidebar.selectbox(
    "Sipariş bütçe politikası",
    ["Konservatif (%10)", "Dengeli (%20)", "Agresif (%30)", "Özel oran"],
    index=1,
    help="AYÇA, toplam sipariş önerisini mevcut stok değerinin seçilen oranıyla sınırlar."
)
if order_policy == "Konservatif (%10)":
    order_budget_ratio = 0.10
elif order_policy == "Dengeli (%20)":
    order_budget_ratio = 0.20
elif order_policy == "Agresif (%30)":
    order_budget_ratio = 0.30
else:
    order_budget_ratio = st.sidebar.slider("Özel sipariş bütçe oranı", 1, 60, 20) / 100
manual_days = st.sidebar.number_input("Ürün raporu kaç günü kapsıyor?", min_value=1, max_value=365, value=30, step=1)
use_sales_date_span = st.sidebar.checkbox("Gün hesabında satış hareket tarih aralığını kullan", value=True)
show_patient_columns = st.sidebar.checkbox("Hasta isim kolonunu göster", value=False)
mask_patient_display = st.sidebar.checkbox("Hasta adlarını maskele", value=True)
st.sidebar.caption("Hasta TC ve kişisel sağlık verisi analiz dışı tutulmalıdır. Hasta adı varsayılan olarak gizlidir.")


# ============================================================
# DOSYA BEKLEME EKRANI
# ============================================================
if inventory_file is None or product_file is None or sales_file is None:
    st.markdown(
        f"""
        <div class="ayca-header">
            <div class="ayca-title">
                <h1>AYÇA Insight V10.7 Copilot Slim Health Analysis</h1>
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
# Risk referansları: gömülü demo sözlüğü + kullanıcı tarafından yüklenen güncel master dosya.
builtin_risk_ref = build_builtin_risk_reference()
uploaded_risk_ref = load_uploaded_risk_reference(risk_master_file) if risk_master_file is not None else pd.DataFrame()
risk_reference_df = pd.concat([builtin_risk_ref, uploaded_risk_ref], ignore_index=True).drop_duplicates(["barkod", "urun_anahtar", "risk_tipi"])
product_master = apply_risk_reference(product_master, risk_reference_df)
product_master, order_budget_info = apply_order_budget(product_master, order_budget_ratio)
product_master = apply_risk_reference(product_master, risk_reference_df)
current_stats = summarize_sales(period_df)
previous_stats = summarize_sales(prev_df)

# Güvenli trend değişkenleri: V10.1 sade ekranında kullanılan kartlar için
# önceki dönem kıyaslarını burada merkezi olarak üretiyoruz.
ciro_trend, ciro_class = rate_fmt(current_stats.get("ciro", 0), previous_stats.get("ciro", 0))
profit_trend, profit_class = rate_fmt(current_stats.get("kar", 0), previous_stats.get("kar", 0))
margin_trend, margin_class = rate_fmt(current_stats.get("marj", 0), previous_stats.get("marj", 0))

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

# Finans Merkezi / Ciro & Tahsilat için güvenli ödeme özeti.
# Eski sürümde payment_df bazı koşullarda hiç oluşmadığı için ekran kırılıyordu.
payment_df = pd.DataFrame()
if period_df is not None and not period_df.empty and {"tahsilat", "ciro"}.issubset(period_df.columns):
    payment_df = period_df.copy()
    payment_df["tahsilat"] = payment_df["tahsilat"].astype(str).replace({"nan": "Bilinmiyor", "None": "Bilinmiyor"}).fillna("Bilinmiyor")
    payment_df = payment_df.groupby("tahsilat", as_index=False).agg(
        ciro=("ciro", "sum"),
        kar=("brut_kar", "sum"),
        islem=("satis_no", "nunique"),
        tahsilat_acigi=("tahsilat_acigi", "sum"),
    )
    payment_df["marj"] = np.where(payment_df["ciro"] > 0, payment_df["kar"] / payment_df["ciro"], 0)
    payment_df = payment_df.sort_values("ciro", ascending=False)

actions = create_action_items(product_master, current_stats, previous_stats, daily_df)
patient_loyalty = build_patient_loyalty(period_df, sales_df)
doctor_intel = build_doctor_intelligence(period_df, sales_df, product_master)
business_insights = build_business_insights(product_master, current_stats, previous_stats, order_budget_info, score_items)

# Segmentler
reorder_df = product_master[product_master["siparis_gerekli_mi"]].sort_values(["acil_siparis_mi", "siparis_tahmini_tutar"], ascending=[False, False])
dead_df = product_master[product_master["olu_stok_mu"]].sort_values("stok_degeri", ascending=False)
slow_df = product_master[product_master["yavas_stok_mu"]].sort_values("stok_degeri", ascending=False)
urgent_df = product_master[product_master.get("acil_siparis_mi", product_master["stokta_yok_satmis_mi"])].sort_values("satis_tutari", ascending=False)
priority_df = product_master[product_master.get("oncelikli_siparis_mi", False)].sort_values("siparis_tahmini_tutar", ascending=False)
rx_as_needed_df = product_master[product_master.get("recete_geldikce_al_mi", False)].sort_values("birim_satis", ascending=False)
fast_df = product_master[product_master["hizli_tukeniyor_mu"]].sort_values("stok_bitis_gunu")
profit_df = product_master[product_master["satilan_adet"] > 0].sort_values("kar_tutari", ascending=False)
capital_df = product_master.sort_values("stok_degeri", ascending=False)
abc_df = product_master.groupby("abc_sinif", as_index=False).agg(urun_sayisi=("barkod", "count"), ciro=("satis_tutari", "sum"), kar=("kar_tutari", "sum"), stok_degeri=("stok_degeri", "sum"))
group_df = product_master.groupby("urun_grubu", as_index=False).agg(urun_sayisi=("barkod", "count"), satilan_adet=("satilan_adet", "sum"), ciro=("satis_tutari", "sum"), kar=("kar_tutari", "sum"), stok_degeri=("stok_degeri", "sum"), siparis_onerisi=("siparis_onerisi", "sum")).sort_values("ciro", ascending=False)
group_df["marj"] = np.where(group_df["ciro"] > 0, group_df["kar"] / group_df["ciro"], 0)

risk_summary_df = risk_summary_table(product_master)
risk_df = product_master[product_master.get("risk_var_mi", False)].sort_values(["risk_segmenti", "satis_tutari"], ascending=[True, False])
red_rx_df = product_master[product_master.get("kirmizi_recete_mi", False)].sort_values("satis_tutari", ascending=False)
green_rx_df = product_master[product_master.get("yesil_recete_mi", False)].sort_values("satis_tutari", ascending=False)
ek_izlem_df = product_master[product_master.get("ek_izlem_mi", False)].sort_values("satis_tutari", ascending=False)
kki_df = product_master[product_master.get("kki_riskli_mi", False)].sort_values("satis_tutari", ascending=False)

# Eşleşme kalitesi
product_barcodes = set(product_df["barkod"])
inv_barcodes = set(inventory_df["barkod"])
matched_count = len(product_barcodes & inv_barcodes)
match_ratio = safe_div(matched_count, len(product_barcodes))


# ============================================================
# HEADER - SADE SAAS ÜST ALAN
# ============================================================
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-header">
        <div class="ayca-title">
            <h1>AYÇA Insight V10.6</h1>
            <p>{eczane_adi} · {period_label} · Gün hesabı: {analysis_days} gün · {today_str}</p>
        </div>
        <div class="header-pill">{get_membership()} Plan · Sağlık Skoru {score}/100</div>
    </div>
    """,
    unsafe_allow_html=True,
)



def render_patient_loyalty_section(patient_loyalty: dict, mask_patient_display: bool = True):
    """Hasta sadakati ekranı: kart + ölçülü grafik + aksiyon tablosu."""
    summary = patient_loyalty.get("summary", {}) if patient_loyalty else {}
    if not summary:
        st.info("Hasta sadakat analizi için satış dosyasında hasta adı bilgisi bulunamadı.")
        return

    h1, h2, h3, h4 = st.columns(4)
    with h1: make_mini_card("Aktif Hasta", str(summary.get("aktif_hasta", 0)), "Son 12 ay verisi", "alert-blue")
    with h2: make_mini_card("VIP Segment", str(summary.get("vip_hasta", 0)), "Sık gelen / yüksek ciro", "alert-green")
    with h3: make_mini_card("Kayıp Riski", str(summary.get("kayip_riski", 0)), "Son 90 gündür gelmeyen", "alert-red" if summary.get("kayip_riski", 0) else "alert-green")
    with h4: make_mini_card("Hasta Kaynaklı Ciro", money_fmt(summary.get("toplam_hasta_ciro", 0)), "Eşleşen hasta kayıtları", "alert-purple")

    freq_df = patient_loyalty.get("frequency", pd.DataFrame()).copy()
    basket_df = patient_loyalty.get("basket", pd.DataFrame()).copy()
    lost_df = patient_loyalty.get("lost", pd.DataFrame()).copy()
    vip_df = patient_loyalty.get("vip", pd.DataFrame()).copy()

    # Görsel yoğunluk: üstte 2 grafik yeterli, detaylar tablolarda.
    g1, g2 = st.columns(2)
    with g1:
        if not vip_df.empty and {"hasta_clean", "ciro"}.issubset(vip_df.columns):
            vip_chart = vip_df.sort_values("ciro", ascending=False).head(15)
            fig = px.bar(vip_chart, x="ciro", y="hasta_clean", orientation="h", title="VIP Hastalar - Ciro Katkısı")
            st.plotly_chart(apply_plot_theme(fig, height=390), use_container_width=True)
        else:
            st.info("VIP hasta grafiği için yeterli veri bulunamadı.")
    with g2:
        if not lost_df.empty and {"hasta_clean", "onceki_ciro"}.issubset(lost_df.columns):
            lost_chart = lost_df.sort_values("onceki_ciro", ascending=False).head(15)
            fig = px.bar(lost_chart, x="onceki_ciro", y="hasta_clean", orientation="h", title="Kayıp Riski - Önceki Ciro")
            st.plotly_chart(apply_plot_theme(fig, height=390), use_container_width=True)
        elif not freq_df.empty and {"ziyaret", "ciro"}.issubset(freq_df.columns):
            fig = px.scatter(freq_df.head(80), x="ziyaret", y="ciro", size="ortalama_sepet" if "ortalama_sepet" in freq_df.columns else None, title="Hasta Frekans / Ciro Haritası")
            st.plotly_chart(apply_plot_theme(fig, height=390), use_container_width=True)
        else:
            st.info("Kayıp riski grafiği için yeterli veri bulunamadı.")

    if mask_patient_display:
        freq_df = mask_patient_names(freq_df)
        basket_df = mask_patient_names(basket_df)
        lost_df = mask_patient_names(lost_df)
        vip_df = mask_patient_names(vip_df)

    st.markdown('<div class="section-title">Aksiyon Listeleri</div>', unsafe_allow_html=True)
    t1, t2, t3, t4 = st.tabs(["VIP Segment", "Kayıp Riski", "En Sık Gelen 50", "En Yüksek Sepet"])
    with t1:
        st.caption("Bu liste eczanenin korunması gereken değerli hasta grubunu gösterir.")
        st.dataframe(vip_df, use_container_width=True, hide_index=True)
    with t2:
        st.caption("Bu liste geçmişte düzenli gelmiş ancak son 90 gündür görünmeyen hastaları gösterir.")
        st.dataframe(lost_df, use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(freq_df, use_container_width=True, hide_index=True)
    with t4:
        st.dataframe(basket_df, use_container_width=True, hide_index=True)

    st.markdown(
        """
        <div class="ai-card">
            <div class="ai-title">KVKK Notu</div>
            <div class="ai-text">
            Bu ekran ticari karar desteği için tasarlanmıştır. Hasta isimleri varsayılan olarak maskelenir.
            Hasta TC, açık sağlık tanısı ve hassas sağlık verisi analiz ekranına alınmamalıdır.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# SAYFALAR
# ============================================================
pages = ["🏠 Sabah Brifingi", "📦 Operasyon Merkezi", "💰 Finans Merkezi", "🚨 Risk Merkezi", "👥 Hasta & Reçete Merkezi", "🤖 AYÇA Copilot", "📊 Raporlar"]
if "active_page" not in st.session_state:
    st.session_state["active_page"] = pages[0]
if st.session_state.get("active_page") not in pages:
    st.session_state["active_page"] = pages[0]
page = st.radio(
    "Bölüm",
    pages,
    horizontal=True,
    label_visibility="collapsed",
    index=pages.index(st.session_state.get("active_page", pages[0])),
    key="page_radio",
)
st.session_state["active_page"] = page

product_cols = [
    "barkod", "urun", "urun_grubu", "raf", "stok", "kritik_stok", "psf_final", "stok_degeri",
    "satilan_adet", "satis_tutari", "kar_tutari", "kar_marji", "gunluk_satis_hizi", "stok_bitis_gunu_goster",
    "risk_segmenti", "risk_tipi", "risk_kaynak", "kki_birim_fark_tl", "kki_tahmini_fark_tl",
    "siparis_segmenti", "siparis_filtre_gecer_mi", "teknik_siparis_onerisi_ham", "siparis_onerisi_ham", "siparis_tahmini_tutar_ham", "siparis_onerisi", "siparis_tahmini_tutar", "siparis_kisit_katsayisi", "abc_sinif", "aksiyon"
]


def _top_product_name(df: pd.DataFrame, sort_col: str = "stok_degeri") -> str:
    try:
        if df is None or df.empty or "urun" not in df.columns:
            return "-"
        if sort_col in df.columns:
            return str(df.sort_values(sort_col, ascending=False).iloc[0]["urun"])
        return str(df.iloc[0]["urun"])
    except Exception:
        return "-"


def build_health_analysis_sections(score: int, score_items: dict, current_stats: dict, product_master: pd.DataFrame,
                                   dead_df: pd.DataFrame, slow_df: pd.DataFrame, urgent_df: pd.DataFrame,
                                   reorder_df: pd.DataFrame, patient_loyalty: dict, business_insights: dict,
                                   kki_df: pd.DataFrame) -> dict:
    """Sabah Brifingi ve Copilot için kısa, aksiyonlu sağlık yorumu üretir."""
    strong, watch, urgent = [], [], []
    marj = float(current_stats.get("marj", 0) or 0)
    ciro = float(current_stats.get("ciro", 0) or 0)
    tahsilat_acigi = float(current_stats.get("tahsilat_acigi", 0) or 0)
    dead_value = float(dead_df.get("stok_degeri", pd.Series(dtype=float)).sum()) if dead_df is not None and not dead_df.empty else 0.0
    slow_value = float(slow_df.get("stok_degeri", pd.Series(dtype=float)).sum()) if slow_df is not None and not slow_df.empty else 0.0
    stock_value = float(product_master.get("stok_degeri", pd.Series(dtype=float)).sum()) if product_master is not None and not product_master.empty else 0.0
    dead_ratio = safe_div(dead_value, stock_value)
    lost_count = int(patient_loyalty.get("summary", {}).get("kayip_riski", 0)) if patient_loyalty else 0
    vip_count = int(patient_loyalty.get("summary", {}).get("vip_hasta", 0)) if patient_loyalty else 0

    if marj >= 0.15:
        strong.append(f"Brüt marj sağlıklı: {pct_fmt(marj)}")
    else:
        watch.append(f"Brüt marj izlenmeli: {pct_fmt(marj)}")
    if ciro > 0:
        strong.append(f"Dönem ciro hacmi güçlü: {money_fmt(ciro)}")
    if len(urgent_df) == 0:
        strong.append("Acil sipariş baskısı düşük")
    else:
        urgent.append(f"{len(urgent_df)} ürün acil sipariş listesinde")
    if dead_ratio > 0.10:
        watch.append(f"Ölü stok sermayesi yüksek: {money_fmt(dead_value)}")
    elif dead_value > 0:
        watch.append(f"Ölü stok takip edilmeli: {money_fmt(dead_value)}")
    if len(slow_df) > 0:
        watch.append(f"{len(slow_df)} yavaş stok raf/kampanya kontrolü istiyor")
    if tahsilat_acigi > 0:
        urgent.append(f"Tahsilat açığı kontrol edilmeli: {money_fmt(tahsilat_acigi)}")
    if lost_count > 0:
        urgent.append(f"{lost_count} hasta kayıp riski grubunda")
    if len(kki_df) > 0:
        watch.append(f"{len(kki_df)} KKİ/riskli ürün finansal takipte")
    if vip_count > 0:
        strong.append(f"{vip_count} VIP hasta segmenti korunmalı")

    # En fazla 3 madde göster, boş kalmasın.
    strong = (strong or ["Genel operasyon izlenebilir seviyede"])[:3]
    watch = (watch or ["Belirgin orta seviye uyarı görünmüyor"])[:3]
    urgent = (urgent or ["Bugün kritik alarm görünmüyor"])[:3]

    if score >= 80:
        result = "Eczaneniz genel olarak sağlıklı görünüyor. Bugünkü odak, iyi performansı koruyup stok ve tahsilat sapmalarını erken yakalamak olmalı."
    elif score >= 60:
        result = "Eczaneniz orta-sağlıklı seviyede. Kârlılık korunurken ölü stok, hasta kaybı ve tahsilat tarafında aksiyon alınırsa skor hızla yükselir."
    else:
        result = "Eczanenizde operasyonel baskı yüksek. İlk öncelik acil sipariş, tahsilat açığı ve sermaye bağlayan stokları azaltmak olmalı."

    return {"strong": strong, "watch": watch, "urgent": urgent, "result": result}


def build_health_analysis_html(sections: dict) -> str:
    def items_html(items):
        return "".join([f"<div class='brief-health-item'>• {html.escape(str(x))}</div>" for x in items])
    return f"""
        <div class="brief-health-box">
            <div class="brief-health-col"><div class="brief-health-title">🟢 Güçlü Yönler</div>{items_html(sections.get('strong', []))}</div>
            <div class="brief-health-col"><div class="brief-health-title">🟡 Dikkat</div>{items_html(sections.get('watch', []))}</div>
            <div class="brief-health-col"><div class="brief-health-title">🔴 Acil Aksiyon</div>{items_html(sections.get('urgent', []))}</div>
        </div>
        <div class="brief-health-result">📌 {html.escape(str(sections.get('result', '')))}</div>
    """


def copilot_answer(question: str, product_master: pd.DataFrame, current_stats: dict, reorder_df: pd.DataFrame,
                   dead_df: pd.DataFrame, slow_df: pd.DataFrame, business_insights: dict,
                   patient_loyalty: dict, doctor_intel: dict, payment_df: Optional[pd.DataFrame]) -> list[str]:
    q = normalize_col_name(question or "")
    answers = []
    if any(k in q for k in ["kar", "marj", "zarar", "karlilik", "karli"]):
        answers.append(f"Dönem brüt kârınız {money_fmt(current_stats.get('kar', 0))}, brüt marjınız {pct_fmt(current_stats.get('marj', 0))}.")
        silent = business_insights.get("silent_loss", pd.DataFrame())
        if silent is not None and not silent.empty:
            r = silent.iloc[0]
            answers.append(f"İlk kontrol edilecek sessiz kâr kaybı adayı: {r.get('urun','-')} · tahmini etki {money_fmt(r.get('tahmini_sessiz_kayip',0))}.")
        profit = product_master[product_master.get("satilan_adet", 0) > 0].sort_values("kar_tutari", ascending=False) if product_master is not None and not product_master.empty else pd.DataFrame()
        if not profit.empty:
            r = profit.iloc[0]
            answers.append(f"En çok kâr getiren ürün: {r.get('urun','-')} · {money_fmt(r.get('kar_tutari',0))}.")
    elif any(k in q for k in ["stok", "sermaye", "olu", "yavas", "azalt"]):
        answers.append(f"Stokta bağlı toplam sermaye yaklaşık {money_fmt(product_master.get('stok_degeri', pd.Series(dtype=float)).sum())}.")
        if dead_df is not None and not dead_df.empty:
            r = dead_df.sort_values("stok_degeri", ascending=False).iloc[0]
            answers.append(f"İlk azaltılacak/aksiyon alınacak ölü stok: {r.get('urun','-')} · {money_fmt(r.get('stok_degeri',0))}.")
        if slow_df is not None and not slow_df.empty:
            r = slow_df.sort_values("stok_degeri", ascending=False).iloc[0]
            answers.append(f"Yavaş stokta raf/kampanya kontrolü: {r.get('urun','-')}.")
    elif any(k in q for k in ["siparis", "almal", "bitecek", "oner"]):
        answers.append(f"Bütçeli sipariş öneriniz {money_fmt(reorder_df.get('siparis_tahmini_tutar', pd.Series(dtype=float)).sum())}; listede {len(reorder_df)} ürün var.")
        if reorder_df is not None and not reorder_df.empty:
            r = reorder_df.sort_values("siparis_tahmini_tutar", ascending=False).iloc[0]
            answers.append(f"En yüksek tutarlı sipariş önerisi: {r.get('urun','-')} · {num_fmt(r.get('siparis_onerisi',0),0)} adet · {money_fmt(r.get('siparis_tahmini_tutar',0))}.")
    elif any(k in q for k in ["hasta", "vip", "kayip", "musteri"]):
        summary = patient_loyalty.get("summary", {}) if patient_loyalty else {}
        answers.append(f"Aktif hasta {summary.get('aktif_hasta',0)}, VIP segment {summary.get('vip_hasta',0)}, kayıp riski {summary.get('kayip_riski',0)}.")
        lost = patient_loyalty.get("lost", pd.DataFrame()) if patient_loyalty else pd.DataFrame()
        if lost is not None and not lost.empty:
            r = lost.iloc[0]
            answers.append(f"İlk geri kazanım adayı: {r.get('hasta_clean','Hasta')} · önceki ciro {money_fmt(r.get('onceki_ciro',0))}.")
    elif any(k in q for k in ["doktor", "recete", "reçete"]):
        doc = doctor_intel.get("doctor_kpi", pd.DataFrame()) if doctor_intel else pd.DataFrame()
        if doc is not None and not doc.empty:
            r = doc.sort_values("ciro", ascending=False).iloc[0]
            answers.append(f"En çok ciro getiren doktor: {r.get('doktor_clean','-')} · {money_fmt(r.get('ciro',0))}.")
            answers.append("Doktor trend düşüşleri için Hasta & Reçete Merkezi > Doktor Analizi tablosunu kontrol edin.")
        else:
            answers.append("Doktor analizi için satış hareketlerinde doktor alanı bulunmalı.")
    elif any(k in q for k in ["tahsilat", "nakit", "odeme", "ödeme", "finans"]):
        answers.append(f"Tahsilat açığı {money_fmt(current_stats.get('tahsilat_acigi',0))}; ciroya oranı {pct_fmt(safe_div(current_stats.get('tahsilat_acigi',0), current_stats.get('ciro',0)))}.")
        if payment_df is not None and not payment_df.empty and "ciro" in payment_df.columns:
            r = payment_df.sort_values("ciro", ascending=False).iloc[0]
            answers.append(f"En büyük tahsilat kanalı: {r.get('tahsilat','-')} · {money_fmt(r.get('ciro',0))}.")
    else:
        answers.append("Bu soru için genel analiz: önce finansal marjı, sonra stokta bağlı sermayeyi, ardından hasta/doktor kayıp riskini kontrol edin.")
        answers.append(f"Bugünkü genel sağlık skoru ve aksiyonlar Sabah Brifingi'nde özetlenmiştir; sipariş önerisi {money_fmt(reorder_df.get('siparis_tahmini_tutar', pd.Series(dtype=float)).sum())}.")
    return answers[:4]

# Güvenli dönem etiketi: bazı Streamlit yeniden çalıştırmalarında sidebar değeri erişilemezse hata vermesin.
period_label = st.session_state.get("selected_period_label", globals().get("selected_period", "Son 30 gün"))

if page == "🏠 Sabah Brifingi":
    st.markdown('<div class="section-title">🏠 Sabah Brifingi</div>', unsafe_allow_html=True)
    st.caption("Ana ekran sadeleştirildi: ekonomik sağlık, envanter sağlığı, risk özeti ve yapılacaklar.")

    top_actions = actions[:5] if actions else ["Kritik aksiyon görünmüyor. Bugün genel takip yeterli."]
    task_html = "".join([f"<div class='task-item'><span>☑</span><span>{html.escape(str(item))}</span></div>" for item in top_actions])
    risk_total = len(red_rx_df) + len(green_rx_df) + len(ek_izlem_df) + len(kki_df)
    health_sections = build_health_analysis_sections(score, score_items, current_stats, product_master, dead_df, slow_df, urgent_df, reorder_df, patient_loyalty, business_insights, kki_df)
    health_html = build_health_analysis_html(health_sections)

    brief_html = (
        f'<div class="brief-grid">'
        f'<div class="brief-hero">'
        f'<div class="brief-title">Günaydın {html.escape(str(kullanici_adi))} 👋</div>'
        f'<div class="brief-sub">AYÇA bugün eczanenizin ekonomik, envanter, hasta sadakati ve risk durumunu sağlık analizine dönüştürdü.</div>'
        f'<div class="brief-score">{score}</div>'
        f'<div class="brief-score-label">Eczane Sağlık Skoru · {html.escape(str(score_status(score)))}</div>'
        f'{health_html}'
        f'</div>'
        f'<div class="brief-panel">'
        f'<div class="group-title" style="margin-top:0;">🤖 AYÇA Bugün Ne Diyor?</div>'
        f'<div class="task-list">{task_html}</div>'
        f'</div>'
        f'</div>'
    )
    st.markdown(brief_html, unsafe_allow_html=True)

    st.markdown('<div class="group-title">💰 Ekonomik Sağlık</div>', unsafe_allow_html=True)
    e1, e2, e3, e4 = st.columns(4)
    with e1: make_metric_card("Ciro", money_fmt(current_stats["ciro"]), period_label, ciro_trend, ciro_class)
    with e2: make_metric_card("Brüt Kâr", money_fmt(current_stats["kar"]), "Satış hareketleri", profit_trend, profit_class)
    with e3: make_metric_card("Marj", pct_fmt(current_stats["marj"]), "Brüt kâr / ciro", margin_trend, margin_class)
    with e4: make_metric_card("Tahsilat Açığı", money_fmt(current_stats["tahsilat_acigi"]), f"Oran {pct_fmt(safe_div(current_stats['tahsilat_acigi'], current_stats['ciro']))}", None, "metric-down")

    st.markdown('<div class="group-title">📦 Envanter Sağlığı</div>', unsafe_allow_html=True)
    v1, v2, v3, v4 = st.columns(4)
    with v1: make_mini_card("Stok Değeri", money_fmt(product_master["stok_degeri"].sum()), "Envanterde bağlı sermaye", "alert-blue")
    with v2: make_mini_card("Kritik Stok", str(len(urgent_df)), "Satılmış / kritik seviyede ürün", "alert-red" if len(urgent_df) else "alert-green")
    with v3: make_mini_card("Ölü Stok", str(len(dead_df)), money_fmt(dead_df["stok_degeri"].sum()), "alert-orange" if len(dead_df) else "alert-green")
    with v4: make_mini_card("Yavaş Stok", str(len(slow_df)), money_fmt(slow_df["stok_degeri"].sum()), "alert-purple" if len(slow_df) else "alert-green")

    st.markdown('<div class="group-title">🚨 Risk Özeti</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([.95, 1.05])
    with c1:
        st.markdown(
            f"""
            <div class="risk-summary-card">
                <div class="risk-line"><span>🔴 Kırmızı Reçete</span><span>{len(red_rx_df)}</span></div>
                <div class="risk-line"><span>🟢 Yeşil Reçete</span><span>{len(green_rx_df)}</span></div>
                <div class="risk-line"><span>🟣 Mor Reçete</span><span>0</span></div>
                <div class="risk-line"><span>▼ Ek İzlem</span><span>{len(ek_izlem_df)}</span></div>
                <div class="risk-line"><span>💰 KKİ Risk</span><span>{len(kki_df)}</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="risk-summary-card">
                <div class="group-title" style="margin-top:0;">🤖 AYÇA Önerileri</div>
                <div class="task-item"><span>📦</span><span>{len(reorder_df)} ürün için bütçeli sipariş önerisi mevcut. Toplam öneri: <b>{money_fmt(order_budget_info['final_total'])}</b>.</span></div>
                <div class="task-item"><span>💰</span><span>{len(kki_df)} KKİ riskli ürün tespit edildi. Satış tutarı: <b>{money_fmt(kki_df['satis_tutari'].sum())}</b>.</span></div>
                <div class="task-item"><span>⚠️</span><span>Tahsilat açığı: <b>{money_fmt(current_stats['tahsilat_acigi'])}</b>.</span></div>
                <div class="task-item"><span>🚨</span><span>Toplam risk etiketi bulunan ürün sayısı: <b>{risk_total}</b>.</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

elif page == "📦 Operasyon Merkezi":
    st.markdown('<div class="section-title">📦 Operasyon Merkezi</div>', unsafe_allow_html=True)
    op_tabs = st.tabs(["🛒 Sipariş Motoru", "📦 Ürün Zekası", "🏆 Taşıyan Ürünler", "🧊 Ölü / Yavaş Stok"])
    with op_tabs[0]:
        s1, s2, s3, s4 = st.columns(4)
        with s1: make_mini_card("Bütçeli Sipariş", money_fmt(order_budget_info["final_total"]), f"Limit {money_fmt(order_budget_info['budget_limit'])}", "alert-green")
        with s2: make_mini_card("Acil Sipariş", str(len(urgent_df)), "Satış kaçırma riski", "alert-red" if len(urgent_df) else "alert-green")
        with s3: make_mini_card("Reçete Geldikçe Al", str(int(product_master.get("recete_geldikce_al_mi", pd.Series(False, index=product_master.index)).sum())), "Pahalı / seyrek ürün", "alert-purple")
        with s4: make_mini_card("Bütçe Kullanımı", pct_fmt(order_budget_info["budget_used_ratio"]), "Sipariş önerisi", "alert-blue")
        st.dataframe(reorder_df[product_cols].head(300), use_container_width=True, hide_index=True)
    with op_tabs[1]:
        st.dataframe(product_master[product_cols].head(500), use_container_width=True, hide_index=True)
    with op_tabs[2]:
        st.dataframe(business_insights.get("carrier_80", pd.DataFrame()).head(300), use_container_width=True, hide_index=True)
    with op_tabs[3]:
        t1, t2 = st.tabs(["Ölü Stok", "Yavaş Stok"])
        with t1: st.dataframe(dead_df[product_cols].head(500), use_container_width=True, hide_index=True)
        with t2: st.dataframe(slow_df[product_cols].head(500), use_container_width=True, hide_index=True)

elif page == "💰 Finans Merkezi":
    st.markdown('<div class="section-title">💰 Finans Merkezi</div>', unsafe_allow_html=True)

    # Streamlit Cloud ortamında px değişkeni beklenmedik şekilde scope dışı kalırsa
    # finans ekranı kırılmasın diye bu bölümde lokal ve güvenli import kullanıyoruz.
    try:
        import plotly.express as _px_finance
    except Exception:
        _px_finance = None

    fin_tabs = st.tabs(["Karlılık", "Sessiz Kâr Kaybı", "Ciro & Tahsilat"])

    with fin_tabs[0]:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            make_metric_card("Ciro", money_fmt(current_stats.get("ciro", 0)), period_label, ciro_trend, ciro_class)
        with k2:
            make_metric_card("Brüt Kâr", money_fmt(current_stats.get("kar", 0)), "Satış hareketleri", profit_trend, profit_class)
        with k3:
            make_metric_card("Marj", pct_fmt(current_stats.get("marj", 0)), "Brüt kâr / ciro", margin_trend, margin_class)
        with k4:
            make_metric_card("Ortalama Sepet", money_fmt(current_stats.get("ortalama_sepet", 0)), f"{current_stats.get('islem', 0)} işlem")

        if _px_finance is not None and daily_df is not None and not daily_df.empty:
            fig = _px_finance.line(daily_df, x="gun", y=["ciro", "kar"], markers=True, title="Günlük Ciro ve Kâr")
            st.plotly_chart(apply_plot_theme(fig, height=420), use_container_width=True)
        else:
            st.info("Günlük ciro/kâr grafiği için yeterli veri bulunamadı.")

    with fin_tabs[1]:
        silent_df = business_insights.get("silent_loss", pd.DataFrame())
        silent_cols = [c for c in product_cols + ["tahmini_sessiz_kayip"] if c in silent_df.columns]
        if not silent_df.empty and silent_cols:
            st.dataframe(silent_df[silent_cols].head(500), use_container_width=True, hide_index=True)
        else:
            st.info("Sessiz kâr kaybı için uygun ürün verisi bulunamadı.")

    with fin_tabs[2]:
        c1, c2 = st.columns(2)
        with c1:
            if _px_finance is not None and daily_df is not None and not daily_df.empty:
                fig = _px_finance.line(daily_df, x="gun", y="ciro", markers=True, title="Günlük Ciro")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
            else:
                st.info("Günlük ciro grafiği için yeterli veri bulunamadı.")
        with c2:
            if _px_finance is not None and payment_df is not None and not payment_df.empty and {"tahsilat", "ciro"}.issubset(payment_df.columns):
                fig = _px_finance.bar(payment_df, x="tahsilat", y="ciro", title="Tahsilat Tipine Göre Ciro")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
            else:
                st.info("Tahsilat tipi grafiği için yeterli veri bulunamadı.")

        if payment_df is not None and not payment_df.empty:
            p1, p2, p3 = st.columns(3)
            toplam_tahsilat_ciro = float(payment_df.get("ciro", pd.Series(dtype=float)).sum())
            en_buyuk_tahsilat = payment_df.sort_values("ciro", ascending=False).iloc[0] if toplam_tahsilat_ciro > 0 else None
            tahsilat_acigi_toplam = float(payment_df.get("tahsilat_acigi", pd.Series(dtype=float)).sum())
            with p1: make_mini_card("Tahsilat Kanalı", str(len(payment_df)), "Ödeme tipi kırılımı", "alert-blue")
            with p2: make_mini_card("En Büyük Kanal", str(en_buyuk_tahsilat["tahsilat"]) if en_buyuk_tahsilat is not None else "-", money_fmt(en_buyuk_tahsilat["ciro"]) if en_buyuk_tahsilat is not None else "₺0", "alert-green")
            with p3: make_mini_card("Tahsilat Açığı", money_fmt(tahsilat_acigi_toplam), "Ciro - ödenen farkı", "alert-red" if tahsilat_acigi_toplam > 0 else "alert-green")
            st.markdown('<div class="section-title">Tahsilat Aksiyon Tablosu</div>', unsafe_allow_html=True)
            st.dataframe(payment_df, use_container_width=True, hide_index=True)
        else:
            st.info("Tahsilat özeti için veri bulunamadı.")

elif page == "👥 Hasta & Reçete Merkezi":
    st.markdown('<div class="section-title">👥 Hasta & Reçete Merkezi</div>', unsafe_allow_html=True)
    hr_tabs = st.tabs(["👨‍⚕️ Doktor Analizi", "👥 Hasta Sadakati", "🏥 Kurum Analizi", "🔐 Reçete Takibi"])
    with hr_tabs[0]:
        if doctor_intel.get("doctor_kpi") is not None and not doctor_intel.get("doctor_kpi").empty:
            st.dataframe(doctor_intel.get("doctor_kpi"), use_container_width=True, hide_index=True)
        else:
            st.info("Doktor analizi için satış hareketlerinde doktor alanı gerekir.")
    with hr_tabs[1]:
        render_patient_loyalty_section(patient_loyalty, mask_patient_display)
    with hr_tabs[2]:
        st.dataframe(kurum_df, use_container_width=True, hide_index=True)
    with hr_tabs[3]:
        rec_tabs = st.tabs(["🔴 Kırmızı", "🟢 Yeşil", "🟣 Mor", "▼ Ek İzlem"])
        with rec_tabs[0]: st.dataframe(red_rx_df[product_cols].head(300), use_container_width=True, hide_index=True)
        with rec_tabs[1]: st.dataframe(green_rx_df[product_cols].head(300), use_container_width=True, hide_index=True)
        with rec_tabs[2]: st.info("Mor reçete master listesi eklendiğinde bu bölüm otomatik dolacak.")
        with rec_tabs[3]: st.dataframe(ek_izlem_df[product_cols].head(300), use_container_width=True, hide_index=True)

elif page == "🤖 AYÇA Copilot":
    st.markdown('<div class="section-title">🤖 AYÇA Copilot</div>', unsafe_allow_html=True)
    st.caption("Copilot, özet ekran değil; eczacıya ne yapacağını söyleyen yönetim danışmanı olarak çalışır.")

    health_sections = build_health_analysis_sections(score, score_items, current_stats, product_master, dead_df, slow_df, urgent_df, reorder_df, patient_loyalty, business_insights, kki_df)
    copilot_tabs = st.tabs(["📋 Genel Durum", "🧠 Bana Sor"])

    with copilot_tabs[0]:
        st.markdown(
            f"""
            <div class="exec-grid">
                <div class="exec-card">
                    <div class="exec-title">Eczane CEO Özeti</div>
                    <div class="exec-sub">AYÇA, sağlık skorunu yalnızca puan olarak değil; güçlü yön, dikkat alanı ve aksiyon planı olarak yorumlar.</div>
                    {''.join([f'<div class="exec-list-item">🟢 {html.escape(str(x))}</div>' for x in health_sections.get('strong', [])])}
                    {''.join([f'<div class="exec-list-item">🟡 {html.escape(str(x))}</div>' for x in health_sections.get('watch', [])])}
                    {''.join([f'<div class="exec-list-item">🔴 {html.escape(str(x))}</div>' for x in health_sections.get('urgent', [])])}
                </div>
                <div class="exec-card">
                    <div class="exec-sub">Genel Sağlık Skoru</div>
                    <div class="score-big">{score}</div>
                    <div class="exec-sub">{html.escape(str(score_status(score)))}</div>
                    <div class="exec-list-item">📌 {html.escape(str(health_sections.get('result','')))}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        for name, val in score_items.items():
            st.markdown(
                f"""
                <div class="health-row"><div class="health-head"><span>{html.escape(str(name))}</span><span>{val}/100</span></div>
                <div class="health-bar-bg"><div class="health-bar-fill" style="width:{val}%;"></div></div></div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown('<div class="section-title">Bu Hafta Yapılacaklar</div>', unsafe_allow_html=True)
        for item in actions[:7]:
            st.markdown(f"<div class='exec-list-item'>☑ {html.escape(str(item))}</div>", unsafe_allow_html=True)

    with copilot_tabs[1]:
        st.markdown('<div class="section-title">Hazır Sorular</div>', unsafe_allow_html=True)
        ready_questions = [
            "Bu ay neden kâr düştü?",
            "Hangi ürünler sermaye bağlıyor?",
            "Bu hafta ne sipariş etmeliyim?",
            "Kaybettiğim hastalar kimler?",
            "En çok ciro getiren doktorlar kimler?",
            "Tahsilat yapım sağlıklı mı?",
        ]
        q_cols = st.columns(3)
        for i, q in enumerate(ready_questions):
            with q_cols[i % 3]:
                if st.button(q, use_container_width=True, key=f"copilot_q_{i}"):
                    st.session_state["copilot_question"] = q
        custom_q = st.text_input("AYÇA'ya sor", value=st.session_state.get("copilot_question", "Hangi ürünler paramı bağlıyor?"))
        if st.button("Yanıtla", type="primary", use_container_width=True):
            st.session_state["copilot_question"] = custom_q
        selected_q = st.session_state.get("copilot_question", custom_q)
        st.markdown(
            f"""
            <div class="ai-card">
                <div class="ai-title">Soru: {html.escape(str(selected_q))}</div>
                <div class="ai-text">AYÇA'nın veri bazlı yanıtı aşağıdadır.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        for ans in copilot_answer(selected_q, product_master, current_stats, reorder_df, dead_df, slow_df, business_insights, patient_loyalty, doctor_intel, payment_df):
            st.markdown(f"<div class='exec-list-item'>💡 {html.escape(str(ans))}</div>", unsafe_allow_html=True)
        st.caption("Bu bölüm şu an kural tabanlı çalışır. İleride gerçek LLM/API bağlantısı ile serbest soru-cevap motoruna dönüşebilir.")

elif page == "🩺 Sağlık Karnesi":
    st.markdown('<div class="section-title">🩺 Eczane Sağlık Karnesi</div>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="exec-grid">
            <div class="exec-card">
                <div class="exec-title">Genel Puan: {score}/100 · {score_status(score)}</div>
                <div class="exec-sub">Bu puan kârlılık, tahsilat, ürün bulunurluğu, stok verimliliği ve büyüme bileşenlerinden oluşur.</div>
                {''.join([f'<div class="exec-list-item">{note}</div>' for note in business_insights.get('health_notes', [])])}
            </div>
            <div class="exec-card">
                <div class="score-big">{score}</div>
                <div class="exec-sub">AYÇA Sağlık Puanı</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    for name, val in score_items.items():
        st.markdown(
            f"""
            <div class="health-row"><div class="health-head"><span>{name}</span><span>{val}/100</span></div>
            <div class="health-bar-bg"><div class="health-bar-fill" style="width:{val}%;"></div></div></div>
            """,
            unsafe_allow_html=True,
        )

    h1, h2, h3, h4 = st.columns(4)
    with h1: make_mini_card("Ölü Stok Oranı", pct_fmt(business_insights['summary']['dead_ratio']), money_fmt(dead_df['stok_degeri'].sum()), "alert-orange" if business_insights['summary']['dead_ratio'] > .05 else "alert-green")
    with h2: make_mini_card("Tahsilat Açığı", money_fmt(current_stats['tahsilat_acigi']), f"Ciro oranı {pct_fmt(safe_div(current_stats['tahsilat_acigi'], current_stats['ciro']))}", "alert-red" if current_stats['tahsilat_acigi'] else "alert-green")
    with h3: make_mini_card("Kârlılık", pct_fmt(current_stats['marj']), "Brüt kâr / ciro", "alert-green" if current_stats['marj'] >= .15 else "alert-orange")
    with h4: make_mini_card("Bütçeli Sipariş", money_fmt(order_budget_info['final_total']), f"Bütçe kullanımı {pct_fmt(order_budget_info['budget_used_ratio'])}", "alert-blue")

elif page == "🏆 Taşıyan Ürünler":
    st.markdown('<div class="section-title">🏆 Eczaneyi Taşıyan Ürünler</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1: make_mini_card("Cironun %50'si", f"{len(business_insights['carrier_50'])} ürün", "Eczanenin ana motoru", "alert-green")
    with c2: make_mini_card("Cironun %80'i", f"{len(business_insights['carrier_80'])} ürün", "A sınıfı odak alan", "alert-blue")
    with c3: make_mini_card("Toplam Ürün", str(len(product_master)), "Ürün raporu", "alert-purple")
    c4, c5 = st.columns(2)
    with c4:
        fig = px.bar(business_insights["carrier_50"].head(25), x="satis_tutari", y="urun", orientation="h", title="Cironun İlk %50'sini Taşıyan Ürünler")
        st.plotly_chart(apply_plot_theme(fig, height=600), use_container_width=True)
    with c5:
        fig = px.line(business_insights["carrier_80"].head(100), x=np.arange(1, len(business_insights["carrier_80"].head(100)) + 1), y="kumulatif_ciro_payi", title="Kümülatif Ciro Payı")
        fig.update_xaxes(title="Ürün sırası")
        fig.update_yaxes(title="Kümülatif pay", tickformat=".0%")
        st.plotly_chart(apply_plot_theme(fig, height=600), use_container_width=True)
    t1, t2 = st.tabs(["%50 Ürünleri", "%80 Ürünleri"])
    with t1: st.dataframe(business_insights["carrier_50"][[c for c in product_cols + ["kumulatif_ciro_payi"] if c in business_insights["carrier_50"].columns]], use_container_width=True, hide_index=True)
    with t2: st.dataframe(business_insights["carrier_80"][[c for c in product_cols + ["kumulatif_ciro_payi"] if c in business_insights["carrier_80"].columns]], use_container_width=True, hide_index=True)

elif page == "📉 Sessiz Kâr Kaybı":
    st.markdown('<div class="section-title">📉 Sessiz Kâr Kaybı</div>', unsafe_allow_html=True)
    st.caption("Çok satan/ciro yapan fakat marjı düşük kalan ürünleri yakalar. Bu ekran fiyat, iskonto, alış maliyeti ve muadil stratejisi kontrolü için kullanılır.")
    silent_loss = business_insights["silent_loss"]
    l1, l2, l3 = st.columns(3)
    with l1: make_mini_card("Ürün Sayısı", str(len(silent_loss)), "Çok satan düşük marj", "alert-orange" if len(silent_loss) else "alert-green")
    with l2: make_mini_card("Tahmini Etki", money_fmt(business_insights['summary']['silent_loss_amount']), "15% hedef marj varsayımı", "alert-red" if business_insights['summary']['silent_loss_amount'] else "alert-green")
    with l3: make_mini_card("Kontrol Amacı", "Fiyat / İskonto", "Maliyet ve muadil kontrolü", "alert-blue")
    if silent_loss.empty:
        st.success("Çok satan ama belirgin düşük marjlı ürün bulunmadı.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(silent_loss.head(20), x="tahmini_sessiz_kayip", y="urun", orientation="h", title="Tahmini Sessiz Kâr Kaybı")
            st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
        with c2:
            fig = px.scatter(silent_loss, x="satis_tutari", y="kar_marji", size="satilan_adet", hover_name="urun", title="Ciro / Marj Haritası")
            st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
        cols = [c for c in product_cols + ["potansiyel_marj_farki", "tahmini_sessiz_kayip"] if c in silent_loss.columns]
        st.dataframe(silent_loss[cols], use_container_width=True, hide_index=True)

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
        fig = px.bar(top_reorder, x="siparis_tahmini_tutar", y="urun", orientation="h", title="Tutar Bazlı İlk Bütçeli Sipariş Önerileri")
        st.plotly_chart(apply_plot_theme(fig), use_container_width=True)

    st.markdown('<div class="section-title">Bugünün Aksiyon Listesi</div>', unsafe_allow_html=True)
    for item in actions:
        st.markdown(f"<div class='exec-list-item'>{item}</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">En Acil 20 Ürün</div>', unsafe_allow_html=True)
    acil = pd.concat([urgent_df, priority_df, reorder_df], ignore_index=True).drop_duplicates("barkod").head(20)
    st.dataframe(acil[product_cols], use_container_width=True, hide_index=True)

elif page == "🛒 Sipariş Motoru":
    st.markdown('<div class="section-title">🛒 AYÇA Sipariş Motoru</div>', unsafe_allow_html=True)

    top_order_names = reorder_df.head(4)["urun"].tolist() if not reorder_df.empty else []
    top_order_html = "".join([f"<div class='exec-list-item'>🛒 {x}</div>" for x in top_order_names]) or "<div class='exec-list-item'>✅ Acil sipariş önerisi yok.</div>"
    st.markdown(
        f"""
        <div class="exec-grid">
            <div class="exec-card">
                <div class="exec-title">🤖 AYÇA Öneriyor</div>
                <div class="exec-sub">Teknik ham öneri yerine eczacının uygulanabilir sipariş listesi gösterilir.</div>
                <div class="exec-list-item">📦 Mevcut stok değeri: <b>{money_fmt(order_budget_info['stock_value'])}</b></div>
                <div class="exec-list-item">🎯 Bu dönem önerilen sipariş: <b>{money_fmt(order_budget_info['final_total'])}</b> · stokun <b>%{int(order_budget_info['budget_ratio']*100)}</b>'si</div>
                <div class="exec-list-item">🚨 Sipariş bekleyen kritik ürün: <b>{len(reorder_df)}</b></div>
                {top_order_html}
            </div>
            <div class="exec-card">
                <div class="exec-sub">Sipariş Bütçesi Kullanımı</div>
                <div class="score-big">{pct_fmt(order_budget_info['budget_used_ratio'])}</div>
                <div class="exec-sub">Bütçe: <b>{money_fmt(order_budget_info['budget_limit'])}</b> · Öneri: <b>{money_fmt(order_budget_info['final_total'])}</b></div>
                <div class="health-bar-bg"><div class="health-bar-fill" style="width:{min(100, max(0, order_budget_info['budget_used_ratio']*100))}%;"></div></div>
                <div class="exec-list-item">Limit öncesi hesap teknik detaydır; ana ekranda sadece uygulanabilir öneri gösterilir.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1: make_mini_card("Sipariş Ürünü", str(len(reorder_df)), "Bütçe sonrası önerilen ürün", "alert-blue")
    with c2: make_mini_card("Önerilen Tutar", money_fmt(reorder_df["siparis_tahmini_tutar"].sum()), "Uygulanabilir liste", "alert-green")
    with c3: make_mini_card("Reçete Geldikçe Al", str(len(rx_as_needed_df)), "Pahalı ve seyrek ürünler", "alert-purple" if len(rx_as_needed_df) else "alert-green")

    c5, c6 = st.columns(2)
    with c5:
        fig = px.bar(reorder_df.head(20), x="siparis_tahmini_tutar", y="urun", orientation="h", title="İlk 20 Bütçeli Sipariş Önerisi")
        st.plotly_chart(apply_plot_theme(fig, height=560), use_container_width=True)
    with c6:
        critical_matrix = reorder_df[["urun", "stok", "satilan_adet", "stok_bitis_gunu_goster", "kar_tutari", "siparis_onerisi", "siparis_tahmini_tutar", "aksiyon"]].head(15).copy()
        st.markdown('<div class="section-title">🚨 Kritik Ürün Matrisi</div>', unsafe_allow_html=True)
        st.dataframe(critical_matrix, use_container_width=True, hide_index=True)

    with st.expander("Teknik sipariş detayı"):
        st.write(f"Ham öneri: {money_fmt(order_budget_info['raw_total'])}")
        st.write(f"Kısılan tutar: {money_fmt(order_budget_info['cut_amount'])}")
        st.write(f"Orantı katsayısı: {num_fmt(order_budget_info['scale'], 2)}")

    t1, t2, t3, t4, t5 = st.tabs(["Bütçeli Sipariş Listesi", "Teknik Karşılaştırma", "Acil Sipariş", "Hızlı Tükenen", "Reçete Geldikçe Al"])
    with t1: st.dataframe(reorder_df[product_cols], use_container_width=True, hide_index=True)
    with t2:
        compare_cols = ["barkod", "urun", "stok", "satilan_adet", "birim_satis", "stok_bitis_gunu_goster", "siparis_filtre_gecer_mi", "teknik_siparis_onerisi_ham", "siparis_onerisi_ham", "siparis_tahmini_tutar_ham", "siparis_onerisi", "siparis_tahmini_tutar", "siparis_kisit_katsayisi", "siparis_segmenti", "aksiyon"]
        st.dataframe(product_master[product_master["teknik_siparis_onerisi_ham"] > 0][compare_cols].sort_values("teknik_siparis_onerisi_ham", ascending=False), use_container_width=True, hide_index=True)
    with t3: st.dataframe(urgent_df[product_cols], use_container_width=True, hide_index=True)
    with t4: st.dataframe(fast_df[product_cols], use_container_width=True, hide_index=True)
    with t5: st.dataframe(rx_as_needed_df[product_cols], use_container_width=True, hide_index=True)

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


elif page == "👨‍⚕️ Doktor Intelligence":
    st.markdown('<div class="section-title">👨‍⚕️ Doktor Intelligence</div>', unsafe_allow_html=True)
    doctor_kpi = doctor_intel["doctor_kpi"]
    doctor_month = doctor_intel["doctor_month"]
    doctor_growth = doctor_intel["doctor_growth"]
    doctor_institution = doctor_intel["doctor_institution"]
    doctor_products = doctor_intel["doctor_products"]

    if doctor_kpi.empty:
        st.info("Doktor analizi için satış hareketleri dosyasında doktor bilgisi bulunamadı.")
    else:
        d1, d2, d3, d4 = st.columns(4)
        with d1: make_mini_card("Doktor Sayısı", str(doctor_kpi["doktor_clean"].nunique()), "Seçili dönem", "alert-blue")
        with d2: make_mini_card("İlk 10 Doktor Ciro", money_fmt(doctor_kpi.head(10)["ciro"].sum()), "Ciro liderleri", "alert-green")
        with d3: make_mini_card("En Karlı Doktor", doctor_kpi.sort_values("kar", ascending=False).iloc[0]["doktor_clean"], money_fmt(doctor_kpi.sort_values("kar", ascending=False).iloc[0]["kar"]), "alert-purple")
        with d4: make_mini_card("En Yüksek Ortalama Reçete", doctor_kpi.sort_values("ortalama_recete", ascending=False).iloc[0]["doktor_clean"], money_fmt(doctor_kpi.sort_values("ortalama_recete", ascending=False).iloc[0]["ortalama_recete"]), "alert-orange")

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(doctor_kpi.head(15), x="ciro", y="doktor_clean", orientation="h", title="En Çok Ciro Yapan Doktorlar")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)
        with c2:
            fig = px.bar(doctor_kpi.sort_values("kar", ascending=False).head(15), x="kar", y="doktor_clean", orientation="h", title="En Çok Kâr Getiren Doktorlar")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)

        st.markdown('<div class="section-title">En Çok Ciro Yapan Doktorların Tercih Ettiği İlaçlar</div>', unsafe_allow_html=True)
        st.caption(doctor_intel["doctor_product_note"])
        if doctor_products.empty:
            st.warning("Bu tablo için satış hareketi dosyasında ürün/barkod kalemi olmalı. Mevcut dosya yalnızca reçete/satış özeti ise doktor bazlı ilaç tercihi gerçekçi hesaplanamaz.")
        else:
            st.dataframe(doctor_products, use_container_width=True, hide_index=True)

        if not doctor_growth.empty:
            st.markdown('<div class="section-title">Son 3 Aylık Doktor Trendleri</div>', unsafe_allow_html=True)
            st.dataframe(doctor_growth.head(50), use_container_width=True, hide_index=True)

        st.markdown('<div class="section-title">Doktor Detay Kartı</div>', unsafe_allow_html=True)
        selected_doctor = st.selectbox("Doktor seç", doctor_kpi["doktor_clean"].tolist())
        selected_doc_month = doctor_month[doctor_month["doktor_clean"] == selected_doctor].copy()
        selected_doc_inst = doctor_institution[doctor_institution["doktor"] == selected_doctor].head(10).copy()
        selected_doc_products = doctor_products[doctor_products["doktor_clean"] == selected_doctor].copy() if not doctor_products.empty else pd.DataFrame()
        c3, c4 = st.columns(2)
        with c3:
            if not selected_doc_month.empty:
                fig = px.line(selected_doc_month, x="ay", y="ciro", markers=True, title=f"{selected_doctor} - Aylık Ciro")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
        with c4:
            if not selected_doc_inst.empty:
                fig = px.bar(selected_doc_inst, x="ciro", y="kurum", orientation="h", title=f"{selected_doctor} - Kurum Dağılımı")
                st.plotly_chart(apply_plot_theme(fig), use_container_width=True)
        if not selected_doc_products.empty:
            fig = px.bar(selected_doc_products, x="ciro", y="urun_tercih", orientation="h", title=f"{selected_doctor} - En Çok Tercih Edilen İlaçlar")
            st.plotly_chart(apply_plot_theme(fig, height=460), use_container_width=True)
            st.dataframe(selected_doc_products, use_container_width=True, hide_index=True)

        st.markdown('<div class="section-title">Doktor KPI Tablosu</div>', unsafe_allow_html=True)
        st.dataframe(doctor_kpi, use_container_width=True, hide_index=True)

elif page == "🧑‍🤝‍🧑 Hasta Sadakat":
    st.markdown('<div class="section-title">🧑‍🤝‍🧑 Hasta Sadakat Merkezi</div>', unsafe_allow_html=True)
    summary = patient_loyalty["summary"]
    if not summary:
        st.info("Hasta sadakat analizi için satış dosyasında hasta adı bilgisi bulunamadı.")
    else:
        h1, h2, h3, h4 = st.columns(4)
        with h1: make_mini_card("Aktif Hasta", str(summary["aktif_hasta"]), "Son 12 ay verisi", "alert-blue")
        with h2: make_mini_card("VIP Segment", str(summary["vip_hasta"]), "Sık gelen / yüksek ciro", "alert-green")
        with h3: make_mini_card("Kaybedilen Risk", str(summary["kayip_riski"]), "Son 90 gündür gelmeyen", "alert-red" if summary["kayip_riski"] else "alert-green")
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

        t1, t2, t3, t4 = st.tabs(["En Sık Gelen 50", "En Yüksek Sepet", "Kaybedilen Müşteriler", "VIP Segment"])
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
                <div class="ai-title">KVKK Notu</div>
                <div class="ai-text">
                Bu ekran ticari karar desteği için tasarlanmıştır. Hasta isimleri varsayılan olarak maskelenir.
                Hasta TC, açık sağlık tanısı ve hassas sağlık verisi analiz ekranına alınmamalıdır.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

elif page == "🏥 Kurum Intelligence":
    st.markdown('<div class="section-title">🏥 Kurum Intelligence</div>', unsafe_allow_html=True)
    st.caption("Bu ekran doktor ekranından ayrıldı. Kurum tarafında finansal performans, tahsilat açığı ve marj odağı gösterilir.")

    if kurum_df.empty:
        st.info("Kurum analizi için satış hareketleri dosyasında kurum bilgisi bulunamadı.")
    else:
        k_top = kurum_df.sort_values("ciro", ascending=False).iloc[0]
        k_gap = kurum_df.sort_values("tahsilat_acigi", ascending=False).iloc[0]
        k_marj = kurum_df[kurum_df["ciro"] > 0].sort_values("marj", ascending=False).iloc[0] if not kurum_df[kurum_df["ciro"] > 0].empty else k_top
        c1, c2, c3, c4 = st.columns(4)
        with c1: make_mini_card("Kurum Sayısı", str(kurum_df["kurum"].nunique()), "Seçili dönem", "alert-blue")
        with c2: make_mini_card("En Büyük Kurum", str(k_top["kurum"]), money_fmt(k_top["ciro"]), "alert-green")
        with c3: make_mini_card("Tahsilat Riski", str(k_gap["kurum"]), money_fmt(k_gap["tahsilat_acigi"]), "alert-red" if k_gap["tahsilat_acigi"] > 0 else "alert-green")
        with c4: make_mini_card("En Yüksek Marj", str(k_marj["kurum"]), pct_fmt(k_marj["marj"]), "alert-purple")

        c5, c6 = st.columns(2)
        with c5:
            fig = px.bar(kurum_df.head(15), x="ciro", y="kurum", orientation="h", title="Kurum Bazlı Ciro")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)
        with c6:
            gap_df = kurum_df.sort_values("tahsilat_acigi", ascending=False).head(15)
            fig = px.bar(gap_df, x="tahsilat_acigi", y="kurum", orientation="h", title="Kurum Bazlı Tahsilat Açığı")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)

        c7, c8 = st.columns(2)
        with c7:
            margin_df = kurum_df[kurum_df["ciro"] > 0].sort_values("marj", ascending=False).head(15)
            fig = px.bar(margin_df, x="marj", y="kurum", orientation="h", title="Kurum Bazlı Marj")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)
        with c8:
            fig = px.scatter(kurum_df, x="ciro", y="kar", size="islem", hover_name="kurum", title="Kurum Ciro / Kâr Haritası")
            st.plotly_chart(apply_plot_theme(fig, height=520), use_container_width=True)

        st.markdown('<div class="section-title">Kurum Detay Tablosu</div>', unsafe_allow_html=True)
        st.dataframe(kurum_df, use_container_width=True, hide_index=True)

elif page == "🚨 Risk Merkezi":
    st.markdown('<div class="section-title">🧯 Eczacı Risk Merkezi</div>', unsafe_allow_html=True)
    st.caption("Kırmızı/yeşil reçete, TİTCK ek izlem ve KKİ risklerini ürün bazlı satış-stok verisiyle birleştirir.")

    a1, a2, a3, a4 = st.columns(4)
    with a1: make_mini_card("🔴 Kırmızı Reçete", str(len(red_rx_df)), f"Stok: {money_fmt(red_rx_df['stok_degeri'].sum())}", "alert-red" if len(red_rx_df) else "")
    with a2: make_mini_card("🟢 Yeşil Reçete", str(len(green_rx_df)), f"Son dönem satış: {money_fmt(green_rx_df['satis_tutari'].sum())}", "alert-green" if len(green_rx_df) else "")
    with a3: make_mini_card("🟣 Ek İzlem", str(len(ek_izlem_df)), f"Stok: {money_fmt(ek_izlem_df['stok_degeri'].sum())}", "alert-purple" if len(ek_izlem_df) else "")
    with a4: make_mini_card("💰 KKİ Risk", str(len(kki_df)), f"Tahmini fark: {money_fmt(kki_df.get('kki_tahmini_fark_tl', pd.Series(dtype=float)).sum())}", "alert-orange" if len(kki_df) else "")

    if risk_summary_df.empty:
        st.info("Yüklenen eczane verisinde risk listeleriyle eşleşen ürün bulunamadı. Barkodlu risk master yüklersen eşleşme daha net olur.")
    else:
        show_summary = risk_summary_df.copy()
        for col in ["Stok Değeri", "Satış Tutarı", "Tahmini KKİ Farkı"]:
            if col in show_summary.columns:
                show_summary[col] = show_summary[col].map(money_fmt)
        st.markdown('<div class="section-title">Risk Özeti</div>', unsafe_allow_html=True)
        st.dataframe(show_summary, use_container_width=True, hide_index=True)

    risk_cols_show = [c for c in ["barkod", "urun", "risk_segmenti", "risk_tipi", "risk_kaynak", "stok", "stok_degeri", "satilan_adet", "satis_tutari", "birim_satis", "kki_birim_fark_tl", "kki_tahmini_fark_tl", "siparis_segmenti", "aksiyon"] if c in product_master.columns]
    risk_view_options = ["🔴 Kırmızı", "🟢 Yeşil", "🟣 Ek İzlem", "💰 KKİ", "📋 Tüm Riskler"]
    if st.session_state.get("risk_view") not in risk_view_options:
        st.session_state["risk_view"] = risk_view_options[0]
    risk_view = st.radio("Risk detay görünümü", risk_view_options, horizontal=True, key="risk_view")

    if risk_view == "🔴 Kırmızı":
        st.markdown("### 🔴 Kırmızı Reçeteli Ürünler")
        st.dataframe(red_rx_df[risk_cols_show], use_container_width=True, hide_index=True)
    elif risk_view == "🟢 Yeşil":
        st.markdown("### 🟢 Yeşil Reçeteli Ürünler")
        st.dataframe(green_rx_df[risk_cols_show], use_container_width=True, hide_index=True)
    elif risk_view == "🟣 Ek İzlem":
        st.markdown("### 🟣 Ek İzlem Ürünleri")
        st.dataframe(ek_izlem_df[risk_cols_show], use_container_width=True, hide_index=True)
    elif risk_view == "💰 KKİ":
        st.markdown("### 💰 KKİ Riskli Ürünler")
        if kki_df.empty:
            st.info("Bu veri setinde KKİ listesiyle eşleşen ürün bulunamadı.")
        else:
            kki_show = kki_df[risk_cols_show].sort_values(["kki_tahmini_fark_tl", "satis_tutari"], ascending=[True, False])
            st.dataframe(kki_show, use_container_width=True, hide_index=True)
            fig = px.bar(kki_df.sort_values("kki_tahmini_fark_tl").head(15), x="urun", y="kki_tahmini_fark_tl", title="Tahmini KKİ Farkı - İlk 15 Ürün")
            st.plotly_chart(apply_plot_theme(fig, height=430), use_container_width=True)
            st.warning("KKİ tutarları risk listesi ve satış adedi üzerinden tahmini gösterilir. Nihai kontrol eczane/SGK mutabakatıyla yapılmalıdır.")
    else:
        st.markdown("### 📋 Tüm Riskli Ürünler")
        st.dataframe(risk_df[risk_cols_show], use_container_width=True, hide_index=True)

elif page == "🔐 Reçete Merkezi":
    st.markdown('<div class="section-title">🔐 Kontrollü Reçete Merkezi</div>', unsafe_allow_html=True)
    st.caption("Bu ekran hasta tanısı veya açık hassas sağlık verisi işlemeden, sadece ürün bazlı stok/satış ve operasyon riski gösterir.")

    c1, c2, c3, c4 = st.columns(4)
    with c1: make_mini_card("🔴 Kırmızı Reçete", str(len(red_rx_df)), f"Stok: {money_fmt(red_rx_df['stok_degeri'].sum())}", "alert-red" if len(red_rx_df) else "")
    with c2: make_mini_card("🟢 Yeşil Reçete", str(len(green_rx_df)), f"Satış: {money_fmt(green_rx_df['satis_tutari'].sum())}", "alert-green" if len(green_rx_df) else "")
    with c3:
        controlled_df = product_master[product_master.get("kontrollu_recete_mi", False)].copy()
        make_mini_card("Kontrollü Toplam", str(len(controlled_df)), f"Stok değeri: {money_fmt(controlled_df['stok_degeri'].sum()) if not controlled_df.empty else money_fmt(0)}", "alert-purple" if len(controlled_df) else "")
    with c4:
        controlled_order_df = controlled_df[controlled_df.get("kontrollu_takip_mi", False)].copy() if not controlled_df.empty else pd.DataFrame()
        make_mini_card("Takip Gereken", str(len(controlled_order_df)), "Teknik sipariş/stok sinyali var", "alert-orange" if len(controlled_order_df) else "alert-green")

    recete_cols_show = [c for c in [
        "barkod", "urun", "risk_segmenti", "risk_tipi", "risk_kaynak",
        "stok", "kritik_stok", "stok_degeri", "satilan_adet", "satis_tutari",
        "birim_satis", "stok_bitis_gunu_goster", "siparis_segmenti", "aksiyon"
    ] if c in product_master.columns]

    recete_view_options = ["🔴 Kırmızı", "🟢 Yeşil", "📋 Tüm Kontrollü", "⚠️ Takip Gereken"]
    if st.session_state.get("recete_view") not in recete_view_options:
        st.session_state["recete_view"] = recete_view_options[0]
    recete_view = st.radio("Reçete merkezi görünümü", recete_view_options, horizontal=True, key="recete_view")

    if recete_view == "🔴 Kırmızı":
        st.markdown("### 🔴 Kırmızı Reçeteli Ürünler")
        if red_rx_df.empty:
            st.info("Bu eczane verisinde kırmızı reçete listesiyle eşleşen ürün bulunamadı.")
        else:
            st.dataframe(red_rx_df[recete_cols_show].sort_values(["stok_degeri", "satis_tutari"], ascending=False), use_container_width=True, hide_index=True)
    elif recete_view == "🟢 Yeşil":
        st.markdown("### 🟢 Yeşil Reçeteli Ürünler")
        if green_rx_df.empty:
            st.info("Bu eczane verisinde yeşil reçete listesiyle eşleşen ürün bulunamadı.")
        else:
            st.dataframe(green_rx_df[recete_cols_show].sort_values(["satis_tutari", "satilan_adet"], ascending=False), use_container_width=True, hide_index=True)
    elif recete_view == "📋 Tüm Kontrollü":
        st.markdown("### 📋 Tüm Kontrollü Reçeteler")
        if controlled_df.empty:
            st.info("Kırmızı/yeşil reçete eşleşmesi yok. Risk master dosyasını barkodlu yüklersen eşleşme oranı artar.")
        else:
            st.dataframe(controlled_df[recete_cols_show].sort_values(["risk_segmenti", "satis_tutari"], ascending=[True, False]), use_container_width=True, hide_index=True)
    else:
        st.markdown("### ⚠️ Takip Gereken Kontrollü Ürünler")
        if controlled_order_df.empty:
            st.success("Kontrollü reçete ürünlerinde teknik sipariş/stok takip sinyali görünmüyor.")
        else:
            st.dataframe(controlled_order_df[recete_cols_show].sort_values(["stok_bitis_gunu_goster", "satis_tutari"], ascending=[True, False]), use_container_width=True, hide_index=True)

    st.markdown("---")
    make_mini_card("Not", "KVKK uyumlu kullanım", "Bu modül hasta adı/TC/tanı göstermeden, sadece ürün bazlı kontrollü reçete stok ve satış takibi yapar.", "alert-purple")

elif page == "📊 Raporlar":
    st.markdown('<div class="section-title">Excel Raporu</div>', unsafe_allow_html=True)
    report = create_excel_report(
        product_master, sales_df, period_df, kurum_df, doktor_df, daily_df, weekday_df, hourly_df,
        doctor_intel.get("doctor_kpi"), doctor_intel.get("doctor_products"),
        patient_loyalty.get("frequency"), patient_loyalty.get("lost"), business_insights
    )
    st.download_button(
        "📥 AYÇA Insight V10.7 Copilot Slim Health Analysis Raporunu İndir",
        data=report,
        file_name=f"ayca_insight_v9_2_kki_risk_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
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
