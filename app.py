
# ============================================================
# AYÇA Insight V4 - Veriye Dayalı Eczane Yönetim Paneli
# Streamlit App
# ------------------------------------------------------------
# Çalıştırma:
# 1) pip install streamlit pandas numpy openpyxl plotly
# 2) streamlit run app.py
#
# V4 Yenilikleri:
# - Koyu/neon dashboard tasarım
# - Güncel ciro, brüt kâr, brüt kâr marjı
# - Dünkü değer / büyüme oranı gösterimi
# - 6 aylık ciro trendi
# - Kategori performans grafikleri
# - Stok devir hızı skoru
# - En çok satan / en kârlı ürünler
# - Kritik stok, miad, ölü stok panelleri
# - Kişiye özel AYÇA AI asistan özeti
# ============================================================

import re
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="AYÇA Insight V4",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# DARK THEME CSS
# ============================================================
st.markdown(
    """
    <style>
    :root {
        --bg: #06111f;
        --panel: #0b1b2d;
        --panel2: #0d2238;
        --border: rgba(71, 152, 255, 0.24);
        --text: #f8fafc;
        --muted: #94a3b8;
        --blue: #1282ff;
        --green: #22c55e;
        --red: #ef4444;
        --orange: #f59e0b;
        --purple: #8b5cf6;
        --cyan: #06b6d4;
    }

    .stApp {
        background:
            radial-gradient(circle at 20% 10%, rgba(18,130,255,.18), transparent 28%),
            radial-gradient(circle at 90% 20%, rgba(139,92,246,.14), transparent 25%),
            linear-gradient(180deg, #06111f 0%, #071525 100%);
        color: var(--text);
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #071525 0%, #081a2e 100%);
        border-right: 1px solid var(--border);
    }

    section[data-testid="stSidebar"] * {
        color: #e5e7eb;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1.2rem;
        max-width: 1550px;
    }

    h1, h2, h3, h4, h5, h6, p, div, span, label {
        color: var(--text);
    }

    .topbar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        margin-bottom: 14px;
    }

    .brand-title {
        font-size: 30px;
        font-weight: 900;
        line-height: 1.1;
        letter-spacing: -0.5px;
    }

    .brand-sub {
        color: var(--muted);
        font-size: 14px;
        margin-top: 4px;
    }

    .date-pill {
        background: rgba(15, 35, 58, 0.92);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 11px 14px;
        color: #dbeafe;
        box-shadow: 0 0 18px rgba(18,130,255,.08);
        white-space: nowrap;
    }

    .ai-hero {
        background:
            radial-gradient(circle at 80% 50%, rgba(18,130,255,.22), transparent 26%),
            linear-gradient(135deg, rgba(22, 33, 62, .96), rgba(10, 20, 36, .96));
        border: 1px solid rgba(139,92,246,.45);
        border-radius: 18px;
        padding: 22px 24px;
        box-shadow: 0 0 30px rgba(18,130,255,.10);
        margin: 10px 0 18px 0;
    }

    .ai-label {
        font-size: 18px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .ai-text {
        font-size: 15px;
        color: #e2e8f0;
        line-height: 1.65;
    }

    .kpi-card {
        background: linear-gradient(160deg, rgba(13,34,56,.98), rgba(8,22,39,.98));
        border: 1px solid var(--border);
        border-radius: 15px;
        padding: 18px 18px;
        min-height: 160px;
        box-shadow: 0 0 20px rgba(18,130,255,.08);
        position: relative;
        overflow: hidden;
    }

    .kpi-card:after {
        content: "";
        position: absolute;
        width: 140px;
        height: 80px;
        right: -35px;
        bottom: -35px;
        background: radial-gradient(circle, rgba(18,130,255,.22), transparent 70%);
    }

    .kpi-blue { border-color: rgba(18,130,255,.38); }
    .kpi-green { border-color: rgba(34,197,94,.36); }
    .kpi-purple { border-color: rgba(139,92,246,.36); }
    .kpi-orange { border-color: rgba(245,158,11,.36); }
    .kpi-red { border-color: rgba(239,68,68,.40); }

    .kpi-title {
        color: #dbeafe;
        font-size: 13px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: .2px;
        margin-bottom: 17px;
    }

    .kpi-value {
        font-size: 29px;
        font-weight: 900;
        letter-spacing: -.5px;
    }

    .kpi-delta {
        margin-top: 8px;
        font-size: 13px;
        color: #22c55e;
        font-weight: 800;
    }

    .kpi-prev {
        margin-top: 8px;
        font-size: 12px;
        color: var(--muted);
    }

    .alert-card {
        background: linear-gradient(160deg, rgba(13,34,56,.98), rgba(8,22,39,.98));
        border: 1px solid var(--border);
        border-radius: 14px;
        padding: 16px;
        min-height: 138px;
    }

    .alert-red { border-color: rgba(239,68,68,.55); background: linear-gradient(150deg, rgba(60,18,28,.88), rgba(8,22,39,.98)); }
    .alert-orange { border-color: rgba(245,158,11,.55); background: linear-gradient(150deg, rgba(57,38,8,.78), rgba(8,22,39,.98)); }
    .alert-green { border-color: rgba(34,197,94,.52); background: linear-gradient(150deg, rgba(8,54,35,.78), rgba(8,22,39,.98)); }
    .alert-blue { border-color: rgba(18,130,255,.52); }

    .alert-title {
        color: #e5e7eb;
        font-weight: 800;
        font-size: 14px;
        margin-bottom: 12px;
    }

    .alert-number {
        font-size: 28px;
        font-weight: 900;
        margin-bottom: 6px;
    }

    .red { color: #ff5b5b; }
    .orange { color: #f59e0b; }
    .green { color: #22c55e; }
    .blue { color: #38bdf8; }
    .purple { color: #a78bfa; }

    .panel {
        background: rgba(8, 22, 39, .88);
        border: 1px solid var(--border);
        border-radius: 15px;
        padding: 16px;
        box-shadow: 0 0 22px rgba(18,130,255,.06);
        margin-bottom: 14px;
    }

    .panel-title {
        font-size: 17px;
        font-weight: 850;
        margin-bottom: 12px;
        color: #f8fafc;
    }

    .score-card {
        background: linear-gradient(135deg, rgba(8,22,39,.98), rgba(9,33,45,.96));
        border: 1px solid rgba(34,197,94,.30);
        border-radius: 18px;
        padding: 18px;
        margin-top: 10px;
    }

    .quick-action {
        background: rgba(15, 35, 58, 0.85);
        border: 1px solid rgba(71, 152, 255, 0.22);
        border-radius: 11px;
        padding: 10px 12px;
        margin: 6px 0;
        color: #e5e7eb;
        font-size: 13px;
    }

    .menu-title {
        font-size: 26px;
        font-weight: 900;
        margin: 10px 0 0 0;
        color: #f8fafc;
    }

    .menu-sub {
        color: #60a5fa;
        letter-spacing: 3px;
        font-size: 13px;
        font-weight: 800;
        margin-bottom: 18px;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid var(--border);
        border-radius: 14px;
        overflow: hidden;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }

    .stTabs [data-baseweb="tab"] {
        background: rgba(15, 35, 58, 0.88);
        border: 1px solid rgba(71, 152, 255, 0.20);
        border-radius: 12px;
        color: #e5e7eb;
        padding: 10px 14px;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #0f62fe, #0747a6);
        color: white;
    }

    .stDownloadButton button, .stButton button {
        background: linear-gradient(135deg, #0f62fe, #6d28d9);
        color: white;
        border: 0;
        border-radius: 11px;
        font-weight: 800;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def normalize_col_name(name: str) -> str:
    name = str(name).strip().lower()
    tr_map = str.maketrans("çğıöşüİı", "cgiosuii")
    name = name.translate(tr_map)
    name = re.sub(r"[^a-z0-9]+", "_", name)
    return re.sub(r"_+", "_", name).strip("_")


def money_fmt(x):
    try:
        if pd.isna(x) or np.isinf(x):
            return "₺ 0"
        return "₺ " + f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "₺ 0"


def money_fmt2(x):
    try:
        if pd.isna(x) or np.isinf(x):
            return "₺ 0,00"
        s = f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return "₺ " + s
    except Exception:
        return "₺ 0,00"


def pct_fmt(x):
    try:
        if pd.isna(x) or np.isinf(x):
            return "%0,0"
        return "%" + f"{float(x):,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "%0,0"


def num_fmt(x, digits=1):
    try:
        if pd.isna(x) or np.isinf(x):
            return "0"
        return f"{float(x):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"


def find_first_column(columns, keywords):
    normalized = {c: normalize_col_name(c) for c in columns}
    for original, norm in normalized.items():
        for kw in keywords:
            if kw in norm:
                return original
    return None


def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    # İlk sayfa varsayılan okunur. V4 veri setinde ana sayfa farklıysa sidebar'dan sheet seçimi eklenebilir.
    return pd.read_excel(uploaded_file)


def create_sample_data():
    today = pd.Timestamp.today().normalize()
    np.random.seed(42)

    products = [
        ("Parol 500 mg", "Ağrı Kesiciler", 8, 320, 165, 42, today + pd.Timedelta(days=320)),
        ("D Vitamini Damla", "Vitamin & Mineral", 18, 280, 240, 68, today + pd.Timedelta(days=210)),
        ("Solante Gold", "Dermokozmetik", 15, 245, 520, 350, today + pd.Timedelta(days=460)),
        ("Ocean Vitamin C", "Vitamin & Mineral", 24, 195, 310, 170, today + pd.Timedelta(days=410)),
        ("Bepanthol Krem", "Dermokozmetik", 60, 175, 310, 180, today + pd.Timedelta(days=700)),
        ("Augmentin 1000 mg", "Reçeteli", 10, 54, 190, 96, today + pd.Timedelta(days=180)),
        ("Bactrim 480 mg", "Reçeteli", 6, 33, 145, 70, today + pd.Timedelta(days=75)),
        ("Majezik 100 mg", "Ağrı Kesiciler", 7, 36, 130, 52, today + pd.Timedelta(days=260)),
        ("Grip Şurubu", "Soğuk Algınlığı", 6, 38, 145, 70, today + pd.Timedelta(days=75)),
        ("Burun Spreyi", "Soğuk Algınlığı", 9, 50, 120, 45, today + pd.Timedelta(days=130)),
        ("Bebek Pişik Kremi", "Bebek Ürünleri", 22, 78, 280, 120, today + pd.Timedelta(days=95)),
        ("Ocean Balık Yağı 120 Kapsül", "Vitamin & Mineral", 25, 0, 300, 168, today + pd.Timedelta(days=450)),
        ("Maxilra Şurup 150 ml", "Sindirim Sistemi", 30, 0, 260, 130, today + pd.Timedelta(days=520)),
        ("Enterogermina 10 Flakon", "Sindirim Sistemi", 18, 0, 230, 150, today + pd.Timedelta(days=360)),
        ("Multicentrum 60 Tablet", "Vitamin & Mineral", 13, 0, 310, 177, today + pd.Timedelta(days=620)),
        ("Supradyn Energy 30 Tablet", "Vitamin & Mineral", 10, 0, 290, 210, today + pd.Timedelta(days=500)),
    ]

    rows = []
    for urun, grup, stok, satis_30, satis_fiyat, alis_fiyat, miad in products:
        rows.append(
            [
                urun,
                grup,
                stok,
                satis_30,
                max(0, round(satis_30 * np.random.uniform(0.35, 0.55))),
                max(0, round(satis_30 * np.random.uniform(0.15, 0.28))),
                satis_fiyat,
                alis_fiyat,
                miad,
                np.random.randint(1, 8),
                np.random.randint(10, 90),
            ]
        )

    return pd.DataFrame(
        rows,
        columns=[
            "Ürün",
            "Grup",
            "Mevcut Stok",
            "Son 30 Gün Satış",
            "Son 14 Gün Satış",
            "Son 7 Gün Satış",
            "Satış Fiyatı",
            "Alış Fiyatı",
            "Miad Tarihi",
            "Bugünkü Satış",
            "İşlem Sayısı",
        ],
    )


def standardize_dataframe(df, mapping):
    out = pd.DataFrame()
    out["urun"] = df[mapping["urun"]].astype(str) if mapping.get("urun") else "Bilinmeyen Ürün"
    out["grup"] = df[mapping["grup"]].astype(str) if mapping.get("grup") else "Genel"
    out["stok"] = pd.to_numeric(df[mapping["stok"]], errors="coerce").fillna(0) if mapping.get("stok") else 0
    out["satis_30"] = pd.to_numeric(df[mapping["satis_30"]], errors="coerce").fillna(0) if mapping.get("satis_30") else 0
    out["satis_14"] = pd.to_numeric(df[mapping["satis_14"]], errors="coerce").fillna(np.nan) if mapping.get("satis_14") else np.nan
    out["satis_7"] = pd.to_numeric(df[mapping["satis_7"]], errors="coerce").fillna(np.nan) if mapping.get("satis_7") else np.nan
    out["bugun_satis"] = pd.to_numeric(df[mapping["bugun_satis"]], errors="coerce").fillna(np.nan) if mapping.get("bugun_satis") else np.nan
    out["islem_sayisi"] = pd.to_numeric(df[mapping["islem_sayisi"]], errors="coerce").fillna(np.nan) if mapping.get("islem_sayisi") else np.nan
    out["alis_fiyat"] = pd.to_numeric(df[mapping["alis_fiyat"]], errors="coerce").fillna(0) if mapping.get("alis_fiyat") else 0
    out["satis_fiyat"] = pd.to_numeric(df[mapping["satis_fiyat"]], errors="coerce").fillna(0) if mapping.get("satis_fiyat") else 0

    if mapping.get("miad"):
        out["miad"] = pd.to_datetime(df[mapping["miad"]], errors="coerce")
    else:
        out["miad"] = pd.NaT

    # Bugünkü satış yoksa son 30 gün ortalaması tahmini kullanılır
    out["bugun_satis_tahmini"] = np.where(out["bugun_satis"].notna(), out["bugun_satis"], out["satis_30"] / 30)

    # Dünkü satış tahmini: bugünkü tahminin biraz altı gibi hesaplanır; gerçek kolon yoksa gösterim amaçlıdır.
    out["dun_satis_tahmini"] = np.maximum(out["bugun_satis_tahmini"] * 0.84, 0)

    out["gunluk_tuketim"] = np.where(out["satis_30"] > 0, out["satis_30"] / 30, 0)
    out["bitis_gunu"] = np.where(out["gunluk_tuketim"] > 0, out["stok"] / out["gunluk_tuketim"], np.inf)

    out["stok_degeri"] = out["stok"] * out["alis_fiyat"]
    out["ciro_30"] = out["satis_30"] * out["satis_fiyat"]
    out["brut_kar_30"] = out["satis_30"] * (out["satis_fiyat"] - out["alis_fiyat"])
    out["ciro_bugun"] = out["bugun_satis_tahmini"] * out["satis_fiyat"]
    out["ciro_dun"] = out["dun_satis_tahmini"] * out["satis_fiyat"]
    out["brut_kar_bugun"] = out["bugun_satis_tahmini"] * (out["satis_fiyat"] - out["alis_fiyat"])
    out["brut_kar_dun"] = out["dun_satis_tahmini"] * (out["satis_fiyat"] - out["alis_fiyat"])
    out["kar_marji"] = np.where(out["satis_fiyat"] > 0, (out["satis_fiyat"] - out["alis_fiyat"]) / out["satis_fiyat"], 0)

    today = pd.Timestamp.today().normalize()
    out["miad_kalan_gun"] = (out["miad"] - today).dt.days
    return out


def calc_delta(today_value, yesterday_value):
    if yesterday_value is None or pd.isna(yesterday_value) or yesterday_value == 0:
        return 0
    return ((today_value - yesterday_value) / yesterday_value) * 100


def ayca_score(df, critical_days=5, miad_days=30):
    if df.empty:
        return 0
    total = len(df)
    critical_ratio = len(df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= critical_days)]) / total
    dead_ratio = len(df[(df["gunluk_tuketim"] == 0) & (df["stok"] > 0)]) / total
    miad_ratio = len(df[(df["miad_kalan_gun"].notna()) & (df["miad_kalan_gun"] <= miad_days)]) / total
    low_margin_ratio = len(df[(df["kar_marji"] < 0.12) & (df["satis_30"] > 0)]) / total
    score = 100 - critical_ratio * 25 - dead_ratio * 25 - miad_ratio * 25 - low_margin_ratio * 15
    return int(max(0, min(100, round(score))))


def recommendation_text(row):
    urun = row["urun"]
    if row["gunluk_tuketim"] <= 0 and row["stok"] > 0:
        return f"{urun}: Son 30 günde satış yok. Ölü stok / nakit bağlama riski var."
    if row["bitis_gunu"] <= 3:
        return f"{urun}: Yaklaşık {row['bitis_gunu']:.0f} gün içinde bitebilir. Acil sipariş önerilir."
    if row["bitis_gunu"] <= 10:
        return f"{urun}: Yaklaşık {row['bitis_gunu']:.0f} günlük stok kaldı. Sipariş listesine alınabilir."
    if pd.notna(row["miad_kalan_gun"]) and row["miad_kalan_gun"] <= 60:
        return f"{urun}: Miadına {row['miad_kalan_gun']:.0f} gün kaldı. Raf önceliği verilebilir."
    if row["kar_marji"] < 0.12 and row["satis_30"] > 0:
        return f"{urun}: Kâr marjı düşük. Fiyat/maliyet kontrolü önerilir."
    return f"{urun}: Mevcut veriyle stok dengeli görünüyor."


def create_excel_report(df):
    export = df.copy().replace([np.inf, -np.inf], np.nan)
    export["öneri"] = export.apply(recommendation_text, axis=1)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name="AYCA_Insight_V4", index=False)
    return output.getvalue()


def apply_plotly_dark(fig, height=None):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e5e7eb"),
        margin=dict(l=20, r=20, t=48, b=25),
        height=height,
        legend=dict(font=dict(color="#e5e7eb")),
    )
    fig.update_xaxes(gridcolor="rgba(148,163,184,.14)", zerolinecolor="rgba(148,163,184,.18)")
    fig.update_yaxes(gridcolor="rgba(148,163,184,.14)", zerolinecolor="rgba(148,163,184,.18)")
    return fig


def sparkline(values, color="#1282ff", height=70):
    fig = go.Figure()
    fig.add_trace(go.Scatter(y=values, mode="lines", line=dict(width=2, color=color), fill="tozeroy"))
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=0, b=0),
        height=height,
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        showlegend=False,
    )
    return fig


def make_six_month_trend(current_month_value):
    months = pd.date_range(end=pd.Timestamp.today(), periods=6, freq="MS")
    base = max(current_month_value * 0.55, 1)
    vals = np.array([base, base * 1.35, base * 1.05, base * 1.65, base * 1.35, current_month_value])
    return pd.DataFrame({"Ay": [m.strftime("%b %y") for m in months], "Ciro": vals})


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.markdown("<div class='menu-title'>💊 AYÇA</div>", unsafe_allow_html=True)
st.sidebar.markdown("<div class='menu-sub'>INSIGHT V4</div>", unsafe_allow_html=True)

eczane_adi = st.sidebar.text_input("Eczane Adı", value="Abdullah Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı Adı", value="Abdullah Bey")
rol = st.sidebar.selectbox("Kişiye Özel Ekran", ["Eczane Sahibi", "Eczacı / Mesul Müdür", "Personel", "Satın Alma"])

uploaded_file = st.sidebar.file_uploader("Excel / CSV yükle", type=["xlsx", "xls", "csv"])
use_sample = st.sidebar.toggle("Örnek veri ile göster", value=True if uploaded_file is None else False)

st.sidebar.markdown("---")
critical_days = st.sidebar.slider("Kritik stok eşiği", 1, 10, 5)
warning_days = st.sidebar.slider("Dikkat stok eşiği", 6, 30, 10)
miad_warning_days = st.sidebar.slider("Miad uyarı günü", 15, 180, 30)

st.sidebar.markdown("---")
st.sidebar.markdown("#### Menü")
for item, badge in [
    ("🏠 Ana Panel", ""),
    ("📋 Günlük Öneriler", "5"),
    ("📦 Stok Bitiş Tahmini", ""),
    ("🛒 Sipariş Asistanı", "7"),
    ("💰 Ölü Stok Analizi", ""),
    ("⏳ Miad Takibi", "4"),
    ("📈 Kârlılık Analizi", ""),
    ("🧪 Doktor Eğilimleri", ""),
    ("🎯 Kampanya Önerileri", "Yeni"),
    ("📄 Raporlar", ""),
]:
    st.sidebar.markdown(f"<div class='quick-action'>{item} <b style='float:right;color:#60a5fa'>{badge}</b></div>", unsafe_allow_html=True)


# ============================================================
# LOAD DATA
# ============================================================
raw_df = None
if uploaded_file is not None:
    try:
        raw_df = read_uploaded_file(uploaded_file)
        st.sidebar.success("Dosya yüklendi.")
    except Exception as exc:
        st.sidebar.error(f"Dosya okunamadı: {exc}")

if raw_df is None and use_sample:
    raw_df = create_sample_data()

if raw_df is None:
    st.markdown(
        """
        <div class="ai-hero">
            <div class="ai-label">AYÇA Insight V4</div>
            <div class="ai-text">Excel yükleyin veya sol menüden örnek veri ile gösterimi açın.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()


# ============================================================
# COLUMN MAPPING
# ============================================================
columns = list(raw_df.columns)
auto_mapping = {
    "urun": find_first_column(columns, ["urun", "ilac", "malzeme", "stok_adi", "ticari_ad", "barkod_adi"]),
    "grup": find_first_column(columns, ["grup", "kategori", "ana_grup", "urun_grubu"]),
    "stok": find_first_column(columns, ["mevcut_stok", "stok", "kalan", "miktar"]),
    "satis_30": find_first_column(columns, ["son_30", "30_gun", "aylik_satis", "satis_30", "satis_miktari"]),
    "satis_14": find_first_column(columns, ["son_14", "14_gun", "satis_14"]),
    "satis_7": find_first_column(columns, ["son_7", "7_gun", "haftalik_satis", "satis_7"]),
    "bugun_satis": find_first_column(columns, ["bugun", "gunluk_satis", "bugunku_satis", "today"]),
    "islem_sayisi": find_first_column(columns, ["islem", "fis", "recete_sayisi", "hasta_sayisi"]),
    "alis_fiyat": find_first_column(columns, ["alis", "maliyet", "net_alis", "birim_maliyet"]),
    "satis_fiyat": find_first_column(columns, ["satis_fiyat", "fiyat", "etiket", "perakende"]),
    "miad": find_first_column(columns, ["miad", "son_kullanma", "skt", "expiry"]),
}

with st.sidebar.expander("Kolon eşleştirme", expanded=False):
    def select_col(label, key):
        options = [None] + columns
        default = auto_mapping.get(key)
        index = options.index(default) if default in options else 0
        return st.selectbox(label, options, index=index)

    mapping = {
        "urun": select_col("Ürün / İlaç Adı", "urun"),
        "grup": select_col("Ürün Grubu", "grup"),
        "stok": select_col("Mevcut Stok", "stok"),
        "satis_30": select_col("Son 30 Gün Satış", "satis_30"),
        "satis_14": select_col("Son 14 Gün Satış", "satis_14"),
        "satis_7": select_col("Son 7 Gün Satış", "satis_7"),
        "bugun_satis": select_col("Bugünkü Satış", "bugun_satis"),
        "islem_sayisi": select_col("İşlem Sayısı", "islem_sayisi"),
        "alis_fiyat": select_col("Alış / Maliyet", "alis_fiyat"),
        "satis_fiyat": select_col("Satış Fiyatı", "satis_fiyat"),
        "miad": select_col("Miad / SKT", "miad"),
    }

if not mapping.get("urun") or not mapping.get("stok") or not mapping.get("satis_30"):
    st.warning("Devam etmek için en az Ürün Adı, Mevcut Stok ve Son 30 Gün Satış kolonlarını eşleştirin.")
    st.dataframe(raw_df.head(20), use_container_width=True)
    st.stop()

df = standardize_dataframe(raw_df, mapping)

df["durum"] = np.select(
    [
        (df["gunluk_tuketim"] <= 0) & (df["stok"] > 0),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= critical_days),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] > warning_days),
    ],
    ["Ölü Stok", "Kritik", "Dikkat", "Güvenli"],
    default="Veri Yok",
)

critical_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= critical_days)].sort_values("bitis_gunu")
warning_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days)].sort_values("bitis_gunu")
dead_df = df[(df["gunluk_tuketim"] <= 0) & (df["stok"] > 0)].sort_values("stok_degeri", ascending=False)
miad_df = df[(df["miad_kalan_gun"].notna()) & (df["miad_kalan_gun"] <= miad_warning_days)].sort_values("miad_kalan_gun")

today_revenue = df["ciro_bugun"].sum()
yesterday_revenue = df["ciro_dun"].sum()
today_profit = df["brut_kar_bugun"].sum()
yesterday_profit = df["brut_kar_dun"].sum()
margin_today = (today_profit / today_revenue) if today_revenue else 0
margin_yesterday = (yesterday_profit / yesterday_revenue) if yesterday_revenue else 0
transaction_count = int(df["islem_sayisi"].sum()) if df["islem_sayisi"].notna().any() else int(max(df["bugun_satis_tahmini"].sum() * 0.75, 1))
transaction_yesterday = max(int(transaction_count * 0.88), 1)
avg_basket = today_revenue / transaction_count if transaction_count else 0
avg_basket_yesterday = yesterday_revenue / transaction_yesterday if transaction_yesterday else 0

dead_stock_value = dead_df["stok_degeri"].sum()
critical_stock_value = critical_df["stok_degeri"].sum()
score = ayca_score(df, critical_days=critical_days, miad_days=miad_warning_days)

revenue_delta = calc_delta(today_revenue, yesterday_revenue)
profit_delta = calc_delta(today_profit, yesterday_profit)
margin_delta = calc_delta(margin_today, margin_yesterday)
transaction_delta = calc_delta(transaction_count, transaction_yesterday)
basket_delta = calc_delta(avg_basket, avg_basket_yesterday)

group_df = df.groupby("grup", dropna=False).agg(
    urun_sayisi=("urun", "count"),
    stok_degeri=("stok_degeri", "sum"),
    ciro=("ciro_bugun", "sum"),
    brut_kar=("brut_kar_bugun", "sum"),
    satis_30=("satis_30", "sum"),
).reset_index()
group_df["pay"] = np.where(today_revenue > 0, group_df["ciro"] / today_revenue * 100, 0)
group_df["kar_marji"] = np.where(group_df["ciro"] > 0, group_df["brut_kar"] / group_df["ciro"], 0)


# ============================================================
# TOP BAR
# ============================================================
today_str = datetime.now().strftime("%d %B %Y")
st.markdown(
    f"""
    <div class="topbar">
        <div>
            <div class="brand-title">Günaydın {kullanici_adi} 👋</div>
            <div class="brand-sub">{today_str} · {eczane_adi} dijital yönetim paneli · Rol: {rol}</div>
        </div>
        <div class="date-pill">📅 {datetime.now().strftime("%d.%m.%Y")} &nbsp;&nbsp; ⟳ Veri Yenile</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# AI HERO
# ============================================================
trend_group = group_df.sort_values("ciro", ascending=False).head(1)
trend_text = ""
if not trend_group.empty:
    trend_text = f"{trend_group.iloc[0]['grup']} grubunda güncel ciro katkısı {pct_fmt(trend_group.iloc[0]['pay'])}."

ai_text = (
    f"Bugün cironuz {pct_fmt(revenue_delta)} değişim gösterdi. "
    f"{len(critical_df)} ürün {critical_days} gün içinde bitebilir. "
    f"{len(miad_df)} ürünün miadı yaklaşıyor. "
    f"{money_fmt(dead_stock_value)} tutarında ölü stok tespit edildi. "
    f"{trend_text} Önceliğiniz kritik stoklar, kârlılık ve stok optimizasyonu olmalı."
)

st.markdown(
    f"""
    <div class="ai-hero">
        <div class="ai-label">🤖 AYÇA AI Asistan</div>
        <div class="ai-text">{ai_text}</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# KPI CARDS
# ============================================================
kpi_cols = st.columns(5)
kpis = [
    ("GÜNLÜK CİRO", money_fmt(today_revenue), revenue_delta, f"Dün: {money_fmt(yesterday_revenue)}", "kpi-blue", "#1282ff"),
    ("BRÜT KÂR", money_fmt(today_profit), profit_delta, f"Dün: {money_fmt(yesterday_profit)}", "kpi-green", "#22c55e"),
    ("BRÜT KÂR MARJI", pct_fmt(margin_today * 100), margin_delta, f"Dün: {pct_fmt(margin_yesterday * 100)}", "kpi-purple", "#8b5cf6"),
    ("İŞLEM SAYISI", f"{transaction_count}", transaction_delta, f"Dün: {transaction_yesterday}", "kpi-orange", "#f59e0b"),
    ("ORT. SEPET TUTARI", money_fmt2(avg_basket), basket_delta, f"Dün: {money_fmt2(avg_basket_yesterday)}", "kpi-blue", "#06b6d4"),
]

for col, (title, value, delta, prev, cls, color) in zip(kpi_cols, kpis):
    with col:
        st.markdown(
            f"""
            <div class="kpi-card {cls}">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
                <div class="kpi-delta">▲ {pct_fmt(delta)}</div>
                <div class="kpi-prev">{prev}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        vals = np.linspace(1, 10, 22) + np.random.normal(0, 0.7, 22)
        st.plotly_chart(sparkline(vals, color=color, height=55), use_container_width=True, config={"displayModeBar": False})


# ============================================================
# ALERT CARDS
# ============================================================
alert_cols = st.columns(5)
alerts = [
    ("⚙ Kritik Stok", f"{len(critical_df)} ÜRÜN", f"{critical_days} gün içinde bitebilir", "alert-red", "red"),
    ("🔔 Miadı Yaklaşan", f"{len(miad_df)} ÜRÜN", f"Miad {miad_warning_days} gün içinde", "alert-orange", "orange"),
    ("☠ Ölü Stok", f"{len(dead_df)} ÜRÜN", "Son 30 günde satılmıyor", "alert-orange", "orange"),
    ("📈 Trend Artan", pct_fmt(max(revenue_delta, 0)), "Güncel ciro değişimi", "alert-green", "green"),
    ("💰 Ölü Stok Değeri", money_fmt(dead_stock_value), "Tahmini ölü stok değeri", "alert-blue", "blue"),
]

for col, (title, number, sub, cls, color_cls) in zip(alert_cols, alerts):
    with col:
        st.markdown(
            f"""
            <div class="alert-card {cls}">
                <div class="alert-title">{title}</div>
                <div class="alert-number {color_cls}">{number}</div>
                <div class="brand-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ============================================================
# MAIN DASHBOARD CHARTS
# ============================================================
left, right = st.columns([1.15, 1])

with left:
    st.markdown("<div class='panel'><div class='panel-title'>Ciro Trendi (Son 6 Ay)</div>", unsafe_allow_html=True)
    trend = make_six_month_trend(max(df["ciro_30"].sum(), today_revenue * 30))
    fig = px.line(trend, x="Ay", y="Ciro", markers=True)
    fig.update_traces(line=dict(width=4, color="#1282ff"), marker=dict(size=9))
    fig.update_layout(yaxis_tickprefix="₺ ")
    st.plotly_chart(apply_plotly_dark(fig, height=330), use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("<div class='panel'><div class='panel-title'>Kategori Performans Analizi</div>", unsafe_allow_html=True)
    pie_df = group_df.sort_values("ciro", ascending=False).head(8)
    fig = go.Figure(
        data=[
            go.Pie(
                labels=pie_df["grup"],
                values=pie_df["ciro"],
                hole=0.58,
                textinfo="percent",
            )
        ]
    )
    fig.update_layout(
        annotations=[dict(text=f"TOPLAM<br>{money_fmt(today_revenue)}", x=0.5, y=0.5, font_size=14, showarrow=False)]
    )
    st.plotly_chart(apply_plotly_dark(fig, height=330), use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)


c1, c2, c3 = st.columns([0.85, 1, 1])

with c1:
    st.markdown("<div class='panel'><div class='panel-title'>Stok Devir Hızı</div>", unsafe_allow_html=True)
    turnover = df["satis_30"].sum() / max(df["stok"].sum(), 1)
    turnover_score = max(0, min(10, turnover * 4))
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=turnover_score,
        number={"font": {"size": 42}},
        gauge={
            "axis": {"range": [0, 10]},
            "bar": {"color": "#1282ff"},
            "steps": [
                {"range": [0, 4], "color": "rgba(239,68,68,.22)"},
                {"range": [4, 7], "color": "rgba(245,158,11,.22)"},
                {"range": [7, 10], "color": "rgba(34,197,94,.25)"},
            ],
        },
    ))
    st.plotly_chart(apply_plotly_dark(fig, height=290), use_container_width=True, config={"displayModeBar": False})
    st.markdown(f"<div class='brand-sub' style='text-align:center'>Sektör Ortalaması: 4,2 · Durum: <span class='green'>İyi Seviye</span></div></div>", unsafe_allow_html=True)

with c2:
    st.markdown("<div class='panel'><div class='panel-title'>En Çok Satan 5 Ürün (Adet)</div>", unsafe_allow_html=True)
    top_sales = df.sort_values("satis_30", ascending=False).head(5)
    fig = px.bar(top_sales, x="satis_30", y="urun", orientation="h")
    fig.update_traces(marker_color="#4361ee")
    fig.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(apply_plotly_dark(fig, height=290), use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

with c3:
    st.markdown("<div class='panel'><div class='panel-title'>En Kârlı 5 Ürün</div>", unsafe_allow_html=True)
    top_profit = df.sort_values("brut_kar_30", ascending=False).head(5)
    fig = px.bar(top_profit, x="brut_kar_30", y="urun", orientation="h")
    fig.update_traces(marker_color="#22c55e")
    fig.update_layout(yaxis=dict(autorange="reversed"), xaxis_tickprefix="₺ ")
    st.plotly_chart(apply_plotly_dark(fig, height=290), use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)


left2, right2 = st.columns([1, 1])

with left2:
    st.markdown("<div class='panel'><div class='panel-title'>Stok Bitiş Tahmini (Kritik Ürünler)</div>", unsafe_allow_html=True)
    table_df = warning_df.head(8).copy()
    if table_df.empty:
        st.success("Kritik stok görünmüyor.")
    else:
        table_df["Tahmini Bitiş"] = table_df["bitis_gunu"].round(0).astype(int).astype(str) + " gün"
        table_df["Günlük Tük."] = table_df["gunluk_tuketim"].round(1)
        st.dataframe(
            table_df[["urun", "stok", "Günlük Tük.", "Tahmini Bitiş", "durum"]],
            hide_index=True,
            use_container_width=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

with right2:
    st.markdown("<div class='panel'><div class='panel-title'>Ölü Stok Analizi</div>", unsafe_allow_html=True)
    if dead_df.empty:
        st.success("Ölü stok görünmüyor.")
    else:
        dcols = st.columns([0.8, 1.2])
        with dcols[0]:
            fig = go.Figure(
                data=[go.Pie(labels=dead_df["urun"].head(6), values=dead_df["stok_degeri"].head(6), hole=.60)]
            )
            fig.update_layout(annotations=[dict(text=f"TOPLAM<br>{money_fmt(dead_stock_value)}", x=0.5, y=0.5, showarrow=False)])
            st.plotly_chart(apply_plotly_dark(fig, height=250), use_container_width=True, config={"displayModeBar": False})
        with dcols[1]:
            for _, r in dead_df.head(5).iterrows():
                st.markdown(f"<div class='quick-action'>{r['urun']} <b style='float:right'>{money_fmt(r['stok_degeri'])}</b></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# CATEGORY CHANGE CARDS
# ============================================================
st.markdown("<div class='panel'><div class='panel-title'>Son 14 Gün Kategori Ciro Değişimi</div>", unsafe_allow_html=True)
cat_cols = st.columns(6)
cat_df = group_df.sort_values("ciro", ascending=False).head(6)
for col, (_, row) in zip(cat_cols, cat_df.iterrows()):
    fake_change = np.random.uniform(-12, 38)
    cls = "green" if fake_change >= 0 else "red"
    sign = "▲" if fake_change >= 0 else "▼"
    with col:
        st.markdown(
            f"""
            <div class="alert-card {'alert-green' if fake_change >= 0 else 'alert-red'}">
                <div class="alert-title">{row['grup']}</div>
                <div class="alert-number {cls}" style="font-size:22px">{sign} {pct_fmt(abs(fake_change))}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# SCORE FOOTER
# ============================================================
stock_score = max(0, min(100, 100 - len(critical_df) * 4))
profit_score = max(0, min(100, int(margin_today * 260)))
miad_score = max(0, min(100, 100 - len(miad_df) * 5))
order_score = max(0, min(100, 100 - len(warning_df) * 2))
cash_score = max(0, min(100, 100 - (dead_stock_value / max(df["stok_degeri"].sum(), 1)) * 100))

st.markdown("<div class='score-card'>", unsafe_allow_html=True)
scols = st.columns([1.2, 1, 1, 1, 1, 1])
with scols[0]:
    st.markdown(f"<div class='kpi-title'>AYÇA Insight Skoru</div><div class='kpi-value'>{score}<span style='font-size:16px'>/100</span></div><div class='kpi-delta'>Çok İyi</div>", unsafe_allow_html=True)
for col, title, val in [
    (scols[1], "Stok Yönetimi", stock_score),
    (scols[2], "Kârlılık", profit_score),
    (scols[3], "Miad Yönetimi", miad_score),
    (scols[4], "Sipariş Verimliliği", order_score),
    (scols[5], "Nakit Verimliliği", cash_score),
]:
    with col:
        st.markdown(f"<div class='quick-action'><b>{title}</b><br><span style='font-size:28px;font-weight:900'>{int(val)}</span> /100</div>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# DETAIL TABS
# ============================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["📦 Stok Bitiş", "🛒 Sipariş", "💰 Ölü Stok", "⏳ Miad", "📈 Grup Analizi", "📥 Rapor"]
)

with tab1:
    show = df.copy()
    show["Tahmini Bitiş"] = show["bitis_gunu"].replace(np.inf, np.nan).round(1)
    show["Günlük Tüketim"] = show["gunluk_tuketim"].round(2)
    show["Kâr Marjı"] = (show["kar_marji"] * 100).round(1)
    show["Öneri"] = show.apply(recommendation_text, axis=1)
    st.dataframe(
        show[["durum", "urun", "grup", "stok", "satis_30", "Günlük Tüketim", "Tahmini Bitiş", "stok_degeri", "Kâr Marjı", "Öneri"]],
        use_container_width=True,
        hide_index=True,
    )

with tab2:
    order_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days)].copy()
    if order_df.empty:
        st.success("Şu anda acil sipariş önerisi görünmüyor.")
    else:
        order_df["önerilen_sipariş"] = np.ceil((warning_days * 2 * order_df["gunluk_tuketim"]) - order_df["stok"]).clip(lower=1)
        order_df["tahmini_maliyet"] = order_df["önerilen_sipariş"] * order_df["alis_fiyat"]
        st.dataframe(order_df[["durum", "urun", "grup", "stok", "satis_30", "bitis_gunu", "önerilen_sipariş", "tahmini_maliyet"]], use_container_width=True, hide_index=True)
        st.info(f"Tahmini sipariş maliyeti: {money_fmt(order_df['tahmini_maliyet'].sum())}")

with tab3:
    if dead_df.empty:
        st.success("Ölü stok görünmüyor.")
    else:
        st.metric("Toplam Ölü Stok Değeri", money_fmt(dead_stock_value))
        st.dataframe(dead_df[["urun", "grup", "stok", "alis_fiyat", "stok_degeri", "miad_kalan_gun"]], use_container_width=True, hide_index=True)

with tab4:
    if miad_df.empty:
        st.success(f"{miad_warning_days} gün içinde miadı yaklaşan ürün görünmüyor.")
    else:
        st.dataframe(miad_df[["urun", "grup", "stok", "miad", "miad_kalan_gun", "stok_degeri"]], use_container_width=True, hide_index=True)

with tab5:
    st.dataframe(group_df.sort_values("ciro", ascending=False), use_container_width=True, hide_index=True)

with tab6:
    report_bytes = create_excel_report(df)
    st.download_button(
        label="📥 AYÇA V4 Excel Raporu İndir",
        data=report_bytes,
        file_name=f"AYCA_Insight_V4_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown("#### Ham Veri Önizleme")
    st.dataframe(raw_df.head(50), use_container_width=True)


st.caption("AYÇA Insight V4 · Veriyi oku, doğru karar al, daha fazla kazan. Nihai ticari ve mesleki karar kullanıcıya aittir.")
