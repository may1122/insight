# ============================================================
# AYÇA LOGO + WEBSITE ENTEGRASYONU
# Dosya yapısı önerisi:
# app.py
# assets/ayca_logo.png
# ============================================================

import os
import base64
import streamlit as st

LOGO_PATH = "assets/ayca_logo.png"
AYCA_WEBSITE = "www.aycayazilim.com"


def image_to_base64(path: str) -> str:
    """Streamlit markdown içinde logo göstermek için resmi base64'e çevirir."""
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def show_ayca_brand_header(version: str = "V10.7.4"):
    logo_b64 = image_to_base64(LOGO_PATH)
    logo_html = (
        f"<img src='data:image/png;base64,{logo_b64}' style='height:46px; object-fit:contain;'>"
        if logo_b64 else
        "<div style='font-weight:950;font-size:28px;color:#0F172A;'>AYÇA</div>"
    )

    st.markdown(
        f"""
        <div class="ayca-brand-top">
            <div class="ayca-brand-left">
                {logo_html}
                <div>
                    <div class="ayca-product-name">AYÇA Insight <span>{version}</span></div>
                    <div class="ayca-product-sub">Eczane Yönetim Zekâsı</div>
                </div>
            </div>
            <div class="ayca-brand-right">
                <div>AYÇA Yazılım</div>
                <a href="https://{AYCA_WEBSITE}" target="_blank">{AYCA_WEBSITE}</a>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_ayca_footer():
    st.markdown(
        f"""
        <div class="ayca-footer">
            Powered by <b>AYÇA Yazılım</b> · <a href="https://{AYCA_WEBSITE}" target="_blank">{AYCA_WEBSITE}</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


# CSS BLOĞUNA EKLE
st.markdown(
    """
    <style>
    .ayca-brand-top {
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:18px;
        margin:4px 0 18px 0;
        padding:14px 18px;
        background:rgba(255,255,255,.86);
        border:1px solid #E2E8F0;
        border-radius:22px;
        box-shadow:0 12px 34px rgba(15,23,42,.06);
        backdrop-filter:blur(12px);
    }
    .ayca-brand-left {
        display:flex;
        align-items:center;
        gap:14px;
    }
    .ayca-product-name {
        color:#0F172A;
        font-size:24px;
        font-weight:950;
        letter-spacing:-.6px;
        line-height:1.05;
    }
    .ayca-product-name span {
        font-size:13px;
        color:#64748B;
        font-weight:900;
        margin-left:6px;
    }
    .ayca-product-sub {
        margin-top:4px;
        color:#64748B;
        font-size:13px;
        font-weight:800;
    }
    .ayca-brand-right {
        text-align:right;
        color:#0F172A;
        font-size:13px;
        font-weight:900;
    }
    .ayca-brand-right a,
    .ayca-footer a {
        color:#2563EB;
        text-decoration:none;
        font-weight:900;
    }
    .ayca-footer {
        margin:30px 0 8px 0;
        padding:14px 0;
        text-align:center;
        color:#64748B;
        font-size:12px;
        border-top:1px solid #E2E8F0;
    }
    @media (max-width:900px){
        .ayca-brand-top{display:block;}
        .ayca-brand-right{text-align:left;margin-top:10px;}
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# KULLANIM:
# 1) Giriş başarılı olduktan sonra, ana başlık yerine veya hemen üstüne çağır:
# show_ayca_brand_header("V10.7.4")
#
# 2) Sayfanın en altına çağır:
# show_ayca_footer()
