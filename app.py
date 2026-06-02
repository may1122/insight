
# ============================================================
# AYÇA Insight V3 - Eczanenin Dijital Müdürü
# Streamlit App
# ------------------------------------------------------------
# Çalıştırma:
# 1) pip install streamlit pandas numpy openpyxl plotly
# 2) streamlit run app.py
#
# Not:
# Bu uygulama TEBEOS/Medula benzeri Excel raporlarını yorumlamak için
# tasarlanmıştır. Hasta TC, reçete no, kişisel sağlık verisi gibi alanları
# analiz dışı bırakmanız önerilir.
# ============================================================

import re
from datetime import datetime, date
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


# -----------------------------
# Sayfa ayarları
# -----------------------------
st.set_page_config(
    page_title="AYÇA Insight V3",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------
# Stil
# -----------------------------
st.markdown(
    """
    <style>
    .main {
        background-color: #f7f9fc;
    }
    .ayca-hero {
        padding: 22px 24px;
        border-radius: 22px;
        background: linear-gradient(135deg, #103b5b 0%, #1677a8 45%, #20b486 100%);
        color: white;
        margin-bottom: 18px;
        box-shadow: 0 10px 25px rgba(16, 59, 91, 0.18);
    }
    .ayca-hero h1 {
        margin: 0;
        font-size: 34px;
        font-weight: 800;
    }
    .ayca-hero p {
        margin: 8px 0 0 0;
        font-size: 17px;
        opacity: 0.95;
    }
    .metric-card {
        background: white;
        border-radius: 18px;
        padding: 18px 18px;
        box-shadow: 0 5px 18px rgba(15, 23, 42, 0.07);
        border: 1px solid #edf2f7;
        min-height: 122px;
    }
    .metric-title {
        color: #64748b;
        font-size: 14px;
        margin-bottom: 8px;
    }
    .metric-value {
        color: #0f172a;
        font-size: 27px;
        font-weight: 800;
    }
    .metric-sub {
        color: #64748b;
        font-size: 13px;
        margin-top: 6px;
    }
    .assistant-card {
        background: #ffffff;
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 5px 18px rgba(15, 23, 42, 0.07);
        border-left: 7px solid #20b486;
        margin-bottom: 16px;
    }
    .warning-red {
        background: #fff1f2;
        color: #9f1239;
        border-left: 5px solid #e11d48;
        padding: 12px 14px;
        border-radius: 12px;
        margin-bottom: 8px;
        font-weight: 600;
    }
    .warning-orange {
        background: #fff7ed;
        color: #9a3412;
        border-left: 5px solid #f97316;
        padding: 12px 14px;
        border-radius: 12px;
        margin-bottom: 8px;
        font-weight: 600;
    }
    .warning-blue {
        background: #eff6ff;
        color: #1d4ed8;
        border-left: 5px solid #3b82f6;
        padding: 12px 14px;
        border-radius: 12px;
        margin-bottom: 8px;
        font-weight: 600;
    }
    .warning-green {
        background: #ecfdf5;
        color: #047857;
        border-left: 5px solid #10b981;
        padding: 12px 14px;
        border-radius: 12px;
        margin-bottom: 8px;
        font-weight: 600;
    }
    .section-title {
        font-size: 24px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 12px;
        margin-bottom: 10px;
    }
    .small-muted {
        color: #64748b;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Yardımcı fonksiyonlar
# -----------------------------
def normalize_col_name(name: str) -> str:
    name = str(name).strip().lower()
    tr_map = str.maketrans("çğıöşüİ", "cgiosui")
    name = name.translate(tr_map)
    name = re.sub(r"[^a-z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name


def money_fmt(x):
    try:
        if pd.isna(x):
            return "0 TL"
        return f"{float(x):,.0f} TL".replace(",", ".")
    except Exception:
        return "0 TL"


def num_fmt(x, digits=1):
    try:
        if pd.isna(x):
            return "0"
        return f"{float(x):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"


def safe_div(a, b):
    try:
        if b == 0 or pd.isna(b):
            return 0
        return a / b
    except Exception:
        return 0


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

    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def create_sample_data():
    today = pd.Timestamp.today().normalize()
    data = [
        ["Parol 500 mg", "Reçeteli", 8, 72, 165, 42, 55, 2.4, today + pd.Timedelta(days=320)],
        ["D Vitamini Damla", "Vitamin", 18, 90, 240, 68, 95, 3.0, today + pd.Timedelta(days=210)],
        ["Solante Gold", "Güneş Ürünleri", 15, 45, 520, 350, 470, 1.5, today + pd.Timedelta(days=460)],
        ["Bepanthol Krem", "Dermokozmetik", 60, 12, 310, 180, 260, 0.4, today + pd.Timedelta(days=700)],
        ["Augmentin 1000 mg", "Reçeteli", 4, 41, 190, 96, 128, 1.37, today + pd.Timedelta(days=180)],
        ["Bebek Pişik Kremi", "Bebek", 22, 16, 280, 120, 190, 0.53, today + pd.Timedelta(days=95)],
        ["Grip Şurubu", "Soğuk Algınlığı", 6, 38, 145, 70, 105, 1.27, today + pd.Timedelta(days=75)],
        ["Omega 3", "Vitamin", 35, 7, 410, 210, 320, 0.23, today + pd.Timedelta(days=540)],
        ["Burun Spreyi", "OTC", 9, 50, 120, 45, 80, 1.67, today + pd.Timedelta(days=130)],
        ["Eski Stok Ürün A", "Dermokozmetik", 25, 0, 250, 140, 190, 0, today + pd.Timedelta(days=45)],
        ["Eski Stok Ürün B", "OTC", 14, 0, 175, 95, 120, 0, today + pd.Timedelta(days=25)],
        ["Ateş Ölçer", "Medikal", 5, 2, 650, 430, 590, 0.07, today + pd.Timedelta(days=1200)],
    ]
    return pd.DataFrame(
        data,
        columns=[
            "Ürün",
            "Grup",
            "Mevcut Stok",
            "Son 30 Gün Satış",
            "Satış Fiyatı",
            "Alış Fiyatı",
            "Son Satış Fiyatı",
            "Günlük Ortalama Satış",
            "Miad Tarihi",
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
    out["alis_fiyat"] = pd.to_numeric(df[mapping["alis_fiyat"]], errors="coerce").fillna(0) if mapping.get("alis_fiyat") else 0
    out["satis_fiyat"] = pd.to_numeric(df[mapping["satis_fiyat"]], errors="coerce").fillna(0) if mapping.get("satis_fiyat") else 0

    if mapping.get("miad"):
        out["miad"] = pd.to_datetime(df[mapping["miad"]], errors="coerce")
    else:
        out["miad"] = pd.NaT

    out["gunluk_tuketim"] = np.where(
        out["satis_30"] > 0,
        out["satis_30"] / 30,
        0,
    )

    out["bitis_gunu"] = np.where(
        out["gunluk_tuketim"] > 0,
        out["stok"] / out["gunluk_tuketim"],
        np.inf,
    )

    out["stok_degeri"] = out["stok"] * out["alis_fiyat"]
    out["ciro_30"] = out["satis_30"] * out["satis_fiyat"]
    out["brut_kar_30"] = out["satis_30"] * (out["satis_fiyat"] - out["alis_fiyat"])
    out["kar_marji"] = np.where(out["satis_fiyat"] > 0, (out["satis_fiyat"] - out["alis_fiyat"]) / out["satis_fiyat"], 0)

    today = pd.Timestamp.today().normalize()
    out["miad_kalan_gun"] = (out["miad"] - today).dt.days

    return out


def classify_stock(row):
    if row["gunluk_tuketim"] <= 0:
        return "⚫ Hareketsiz"
    if row["bitis_gunu"] <= 3:
        return "🔴 Kritik"
    if row["bitis_gunu"] <= 10:
        return "🟠 Dikkat"
    return "🟢 Güvenli"


def recommendation_text(row):
    urun = row["urun"]
    if row["gunluk_tuketim"] <= 0 and row["stok"] > 0:
        return f"{urun}: Son 30 günde satış yok. Rafta para bağlıyor olabilir."
    if row["bitis_gunu"] <= 3:
        return f"{urun}: Mevcut satış hızına göre yaklaşık {row['bitis_gunu']:.0f} gün içinde bitebilir. Acil sipariş önerilir."
    if row["bitis_gunu"] <= 10:
        return f"{urun}: Yaklaşık {row['bitis_gunu']:.0f} günlük stok kaldı. Sipariş listesine alınabilir."
    if pd.notna(row["miad_kalan_gun"]) and row["miad_kalan_gun"] <= 60:
        return f"{urun}: Miadına {row['miad_kalan_gun']:.0f} gün kaldı. Raf önceliği verilebilir."
    if row["kar_marji"] < 0.12 and row["satis_30"] > 0:
        return f"{urun}: Kâr marjı düşük görünüyor. Fiyat/maliyet kontrolü yapılmalı."
    return f"{urun}: Stok seviyesi mevcut satış hızına göre dengeli görünüyor."


def ayca_score(df):
    if df.empty:
        return 0

    total = len(df)
    critical_ratio = len(df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= 3)]) / total
    dead_ratio = len(df[(df["gunluk_tuketim"] == 0) & (df["stok"] > 0)]) / total
    miad_ratio = len(df[(df["miad_kalan_gun"].notna()) & (df["miad_kalan_gun"] <= 60)]) / total
    low_margin_ratio = len(df[(df["kar_marji"] < 0.12) & (df["satis_30"] > 0)]) / total

    score = 100
    score -= critical_ratio * 30
    score -= dead_ratio * 25
    score -= miad_ratio * 25
    score -= low_margin_ratio * 20

    return int(max(0, min(100, round(score))))


def create_excel_report(df):
    export = df.copy()
    export = export.replace([np.inf, -np.inf], np.nan)
    export["durum"] = export.apply(classify_stock, axis=1)
    export["öneri"] = export.apply(recommendation_text, axis=1)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export.to_excel(writer, sheet_name="AYCA_Insight_V3", index=False)
    return output.getvalue()


# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("💊 AYÇA Insight V3")
st.sidebar.caption("Eczanenin Dijital Müdürü")

eczane_adi = st.sidebar.text_input("Eczane Adı", value="İdil Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı Adı", value="Abdullah Bey")
rol = st.sidebar.selectbox(
    "Kişiye Özel Ekran",
    ["Eczane Sahibi", "Eczacı / Mesul Müdür", "Personel", "Satın Alma"],
)

uploaded_file = st.sidebar.file_uploader(
    "Excel / CSV yükle",
    type=["xlsx", "xls", "csv"],
)

use_sample = st.sidebar.toggle("Örnek veri ile göster", value=True if uploaded_file is None else False)

st.sidebar.markdown("---")
critical_days = st.sidebar.slider("Kritik stok eşiği", 1, 10, 3)
warning_days = st.sidebar.slider("Dikkat stok eşiği", 4, 30, 10)
dead_stock_days = st.sidebar.slider("Hareketsiz stok gün varsayımı", 30, 180, 90)
miad_warning_days = st.sidebar.slider("Miad uyarı günü", 15, 180, 60)


# -----------------------------
# Veri yükleme
# -----------------------------
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
        f"""
        <div class="ayca-hero">
            <h1>AYÇA Insight V3</h1>
            <p>Excel yükleyin veya sol menüden örnek veri ile önizleme açın.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()


# -----------------------------
# Kolon eşleştirme
# -----------------------------
columns = list(raw_df.columns)

auto_mapping = {
    "urun": find_first_column(columns, ["urun", "ilac", "malzeme", "stok_adi", "ticari_ad", "barkod_adi"]),
    "grup": find_first_column(columns, ["grup", "kategori", "ana_grup", "urun_grubu"]),
    "stok": find_first_column(columns, ["mevcut_stok", "stok", "kalan", "miktar"]),
    "satis_30": find_first_column(columns, ["son_30", "30_gun", "aylik_satis", "satis_30", "satis_miktari"]),
    "satis_14": find_first_column(columns, ["son_14", "14_gun", "satis_14"]),
    "satis_7": find_first_column(columns, ["son_7", "7_gun", "haftalik_satis", "satis_7"]),
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
        "alis_fiyat": select_col("Alış / Maliyet", "alis_fiyat"),
        "satis_fiyat": select_col("Satış Fiyatı", "satis_fiyat"),
        "miad": select_col("Miad / SKT", "miad"),
    }

if not mapping.get("urun") or not mapping.get("stok") or not mapping.get("satis_30"):
    st.warning("Devam etmek için en az Ürün Adı, Mevcut Stok ve Son 30 Gün Satış kolonlarını eşleştirin.")
    st.dataframe(raw_df.head(20), use_container_width=True)
    st.stop()

df = standardize_dataframe(raw_df, mapping)

# Dinamik sınıflandırma eşikleri
df["durum"] = np.select(
    [
        (df["gunluk_tuketim"] <= 0) & (df["stok"] > 0),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= critical_days),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days),
        (df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] > warning_days),
    ],
    ["⚫ Hareketsiz", "🔴 Kritik", "🟠 Dikkat", "🟢 Güvenli"],
    default="Veri Yok",
)

df["öneri"] = df.apply(recommendation_text, axis=1)

score = ayca_score(df)

critical_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= critical_days)].sort_values("bitis_gunu")
warning_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days)].sort_values("bitis_gunu")
dead_df = df[(df["gunluk_tuketim"] <= 0) & (df["stok"] > 0)].sort_values("stok_degeri", ascending=False)
miad_df = df[(df["miad_kalan_gun"].notna()) & (df["miad_kalan_gun"] <= miad_warning_days)].sort_values("miad_kalan_gun")
low_margin_df = df[(df["satis_30"] > 0) & (df["kar_marji"] < 0.12)].sort_values("kar_marji")

total_ciro_30 = df["ciro_30"].sum()
total_profit_30 = df["brut_kar_30"].sum()
daily_ciro_est = total_ciro_30 / 30
daily_profit_est = total_profit_30 / 30
dead_stock_value = dead_df["stok_degeri"].sum()
critical_stock_value = critical_df["stok_degeri"].sum()


# -----------------------------
# Hero
# -----------------------------
today_str = datetime.now().strftime("%d.%m.%Y")
st.markdown(
    f"""
    <div class="ayca-hero">
        <h1>AYÇA Insight V3</h1>
        <p>{eczane_adi} için kişiye özel eczane karar asistanı · {today_str} · AYÇA Skoru: {score}/100</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(f"### Günaydın {kullanici_adi} 👋")
st.caption(f"Rol bazlı ekran: {rol}")


# -----------------------------
# Metrik kartları
# -----------------------------
c1, c2, c3, c4, c5 = st.columns(5)

cards = [
    ("Günlük Ciro Tahmini", money_fmt(daily_ciro_est), "Son 30 gün satış hızına göre"),
    ("Brüt Kâr Tahmini", money_fmt(daily_profit_est), "Yaklaşık günlük brüt kâr"),
    ("Kritik Stok", f"{len(critical_df)} ürün", f"{critical_days} gün ve altı"),
    ("Ölü Stok", money_fmt(dead_stock_value), "Son 30 günde satışı olmayan"),
    ("AYÇA Skoru", f"{score}/100", "Eczane sağlığı"),
]

for col, (title, value, sub) in zip([c1, c2, c3, c4, c5], cards):
    with col:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-title">{title}</div>
                <div class="metric-value">{value}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# -----------------------------
# AYÇA AI yorum kutusu
# -----------------------------
st.markdown('<div class="section-title">🧠 AYÇA AI - Günlük Karar Özeti</div>', unsafe_allow_html=True)

main_messages = []

if len(critical_df) > 0:
    first = critical_df.iloc[0]
    main_messages.append(
        f"En kritik ürün: {first['urun']}. Mevcut satış hızına göre yaklaşık {first['bitis_gunu']:.0f} gün içinde bitebilir."
    )

if len(miad_df) > 0:
    first = miad_df.iloc[0]
    main_messages.append(
        f"Miad uyarısı: {first['urun']} için miada {first['miad_kalan_gun']:.0f} gün kaldı."
    )

if len(dead_df) > 0:
    main_messages.append(
        f"Yaklaşık {money_fmt(dead_stock_value)} değerinde hareketsiz stok görünüyor."
    )

if len(low_margin_df) > 0:
    first = low_margin_df.iloc[0]
    main_messages.append(
        f"Kârlılık kontrolü: {first['urun']} ürününde kâr marjı düşük görünüyor."
    )

if not main_messages:
    main_messages.append("Genel tablo dengeli görünüyor. Kritik stok, miad veya ölü stok baskısı düşük.")

assistant_text = " ".join(main_messages)

st.markdown(
    f"""
    <div class="assistant-card">
        <b>AYÇA AI:</b><br><br>
        {kullanici_adi}, {assistant_text}
        Bugün önceliğiniz stok, miad ve nakit verimliliğini birlikte kontrol etmek olmalı.
    </div>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Rol bazlı aksiyonlar
# -----------------------------
st.markdown('<div class="section-title">📋 Bugün Yapılacaklar</div>', unsafe_allow_html=True)

role_actions = []

if rol == "Eczane Sahibi":
    role_actions = [
        f"Ölü stok değeri: {money_fmt(dead_stock_value)}. Nakit bağlayan ürünleri kontrol edin.",
        f"Günlük brüt kâr tahmini: {money_fmt(daily_profit_est)}.",
        f"Kritik stok değeri: {money_fmt(critical_stock_value)}.",
    ]
elif rol == "Eczacı / Mesul Müdür":
    role_actions = [
        f"{len(critical_df)} kritik ürün için muadil/stok kontrolü yapılmalı.",
        f"{len(miad_df)} ürünün miadı {miad_warning_days} gün içinde yaklaşıyor.",
        "Reçeteli ürünlerde bitiş süresi kısa olanlar önceliklendirilmeli.",
    ]
elif rol == "Personel":
    role_actions = [
        f"{len(miad_df)} miadı yaklaşan ürün rafta öne alınmalı.",
        f"{len(critical_df)} azalan ürün raf ve depo kontrolünden geçirilmeli.",
        "Hareketsiz ürünler için raf düzeni gözden geçirilmeli.",
    ]
else:
    role_actions = [
        f"{len(critical_df)} ürün acil sipariş listesine alınmalı.",
        f"{len(warning_df)} ürün için 10 gün içi stok planı yapılmalı.",
        "Yavaş dönen ürünlerde yeni sipariş verilmeden önce satış hızı kontrol edilmeli.",
    ]

for i, action in enumerate(role_actions, start=1):
    st.markdown(f"<div class='warning-blue'>{i}. {action}</div>", unsafe_allow_html=True)


# -----------------------------
# Sekmeler
# -----------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📦 Stok Bitiş Tahmini",
        "🛒 Sipariş Asistanı",
        "💰 Ölü Stok",
        "⏳ Miad Takibi",
        "📈 Grup & Trend",
        "📥 Rapor",
    ]
)

with tab1:
    st.subheader("Stok Bitiş Tahmini")
    show = df.copy()
    show["Tahmini Bitiş"] = show["bitis_gunu"].replace(np.inf, np.nan).round(1)
    show["Günlük Tüketim"] = show["gunluk_tuketim"].round(2)
    show["Kar Marjı %"] = (show["kar_marji"] * 100).round(1)

    st.dataframe(
        show[
            [
                "durum",
                "urun",
                "grup",
                "stok",
                "satis_30",
                "Günlük Tüketim",
                "Tahmini Bitiş",
                "stok_degeri",
                "Kar Marjı %",
                "öneri",
            ]
        ].sort_values(["durum", "Tahmini Bitiş"], ascending=[True, True]),
        use_container_width=True,
        hide_index=True,
    )

    chart_df = show[(show["gunluk_tuketim"] > 0)].sort_values("bitis_gunu").head(15)
    if not chart_df.empty:
        fig = px.bar(
            chart_df,
            x="urun",
            y="bitis_gunu",
            title="En Yakın Bitecek Ürünler",
            labels={"urun": "Ürün", "bitis_gunu": "Tahmini Bitiş Günü"},
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Sipariş Asistanı")
    order_df = df[(df["gunluk_tuketim"] > 0) & (df["bitis_gunu"] <= warning_days)].copy()
    order_df["önerilen_sipariş"] = np.ceil((warning_days * 2 * order_df["gunluk_tuketim"]) - order_df["stok"]).clip(lower=1)
    order_df["tahmini_maliyet"] = order_df["önerilen_sipariş"] * order_df["alis_fiyat"]

    if order_df.empty:
        st.success("Şu anda acil sipariş önerisi görünmüyor.")
    else:
        st.dataframe(
            order_df[
                [
                    "durum",
                    "urun",
                    "grup",
                    "stok",
                    "satis_30",
                    "gunluk_tuketim",
                    "bitis_gunu",
                    "önerilen_sipariş",
                    "tahmini_maliyet",
                ]
            ].sort_values("bitis_gunu"),
            use_container_width=True,
            hide_index=True,
        )
        st.info(f"Tahmini sipariş maliyeti: {money_fmt(order_df['tahmini_maliyet'].sum())}")

with tab3:
    st.subheader("Ölü Stok Merkezi")
    st.metric("Toplam Hareketsiz Stok Değeri", money_fmt(dead_stock_value))

    if dead_df.empty:
        st.success("Son 30 gün satışına göre hareketsiz stok görünmüyor.")
    else:
        st.dataframe(
            dead_df[
                ["urun", "grup", "stok", "alis_fiyat", "stok_degeri", "miad_kalan_gun", "öneri"]
            ],
            use_container_width=True,
            hide_index=True,
        )

        fig = px.pie(
            dead_df.head(10),
            names="urun",
            values="stok_degeri",
            title="En Yüksek Hareketsiz Stoklar",
        )
        st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.subheader("Miad Takibi")

    if miad_df.empty:
        st.success(f"{miad_warning_days} gün içinde miadı yaklaşan ürün görünmüyor.")
    else:
        st.dataframe(
            miad_df[
                ["urun", "grup", "stok", "miad", "miad_kalan_gun", "stok_degeri", "öneri"]
            ],
            use_container_width=True,
            hide_index=True,
        )

        fig = px.bar(
            miad_df.head(15),
            x="urun",
            y="miad_kalan_gun",
            title="Miadı En Yakın Ürünler",
            labels={"urun": "Ürün", "miad_kalan_gun": "Kalan Gün"},
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

with tab5:
    st.subheader("Ürün Grubu Analizi")

    group_df = df.groupby("grup", dropna=False).agg(
        urun_sayisi=("urun", "count"),
        stok_degeri=("stok_degeri", "sum"),
        ciro_30=("ciro_30", "sum"),
        brut_kar_30=("brut_kar_30", "sum"),
        satis_30=("satis_30", "sum"),
    ).reset_index()

    group_df["kar_marji"] = np.where(group_df["ciro_30"] > 0, group_df["brut_kar_30"] / group_df["ciro_30"], 0)

    st.dataframe(group_df.sort_values("ciro_30", ascending=False), use_container_width=True, hide_index=True)

    c_left, c_right = st.columns(2)
    with c_left:
        fig = px.bar(
            group_df.sort_values("ciro_30", ascending=False),
            x="grup",
            y="ciro_30",
            title="Grup Bazlı 30 Günlük Ciro",
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

    with c_right:
        fig = px.bar(
            group_df.sort_values("brut_kar_30", ascending=False),
            x="grup",
            y="brut_kar_30",
            title="Grup Bazlı Brüt Kâr",
        )
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)

with tab6:
    st.subheader("Rapor İndir")

    st.write("AYÇA Insight V3 analiz sonucunu Excel olarak indirebilirsiniz.")

    report_bytes = create_excel_report(df)
    st.download_button(
        label="📥 AYÇA V3 Excel Raporu İndir",
        data=report_bytes,
        file_name=f"AYCA_Insight_V3_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("#### Ham Veri Önizleme")
    st.dataframe(raw_df.head(50), use_container_width=True)


# -----------------------------
# Alt bilgi
# -----------------------------
st.markdown("---")
st.caption(
    "AYÇA Insight V3 · Raporları okumaz, eczanenizi yönetir. "
    "Bu uygulama karar desteği sunar; nihai ticari ve mesleki karar kullanıcıya aittir."
)
