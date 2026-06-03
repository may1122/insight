# ============================================================
# AYÇA Insight V4.1 - Dinamik Dosya Yorumlayan Dashboard
# ------------------------------------------------------------
# Çalıştırma:
# pip install streamlit pandas numpy openpyxl plotly
# streamlit run AYCA_Insight_V4_1_dinamik_app.py
# ============================================================

import re
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="AYÇA Insight V4.1",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Dark dashboard CSS
# -----------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: radial-gradient(circle at top left, #0e2a46 0%, #061321 38%, #030914 100%); color: #e5f0ff; }
[data-testid="stSidebar"] { background: linear-gradient(180deg, #07192b 0%, #020813 100%); border-right: 1px solid rgba(76,154,255,.2); }
[data-testid="stSidebar"] * { color: #dbeafe !important; }
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1550px; }
.ayca-topbar { display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; }
.ayca-title h1 { margin:0; color:#ffffff; font-size:30px; font-weight:800; letter-spacing:-.5px; }
.ayca-title p { margin:2px 0 0; color:#94a3b8; font-size:14px; }
.glass-card { background: linear-gradient(145deg, rgba(12,32,55,.92), rgba(4,15,30,.94)); border:1px solid rgba(80,150,255,.22); border-radius:18px; padding:18px; box-shadow:0 16px 40px rgba(0,0,0,.32), inset 0 0 22px rgba(0,140,255,.04); }
.ai-card { background: linear-gradient(135deg, rgba(85,46,180,.50), rgba(8,25,52,.94)); border:1px solid rgba(155,113,255,.55); border-radius:18px; padding:18px 22px; box-shadow:0 0 30px rgba(104,69,255,.17); margin-bottom:14px; }
.ai-card h3 { margin:0 0 10px; color:#fff; }
.ai-card p { margin:6px 0; color:#edf2ff; line-height:1.45; }
.kpi { min-height:142px; border-radius:16px; padding:16px; border:1px solid rgba(100,180,255,.22); background: linear-gradient(145deg, rgba(9,30,56,.98), rgba(4,13,25,.97)); position:relative; overflow:hidden; box-shadow: inset 0 0 30px rgba(0,133,255,.08); }
.kpi:after { content:""; position:absolute; right:-30px; bottom:-30px; width:120px; height:120px; background:radial-gradient(circle, rgba(0,145,255,.22), transparent 68%); }
.kpi .label { color:#cbd5e1; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:.4px; }
.kpi .value { color:#fff; font-size:30px; font-weight:800; margin-top:12px; }
.kpi .delta { color:#22c55e; font-size:13px; font-weight:700; margin-top:8px; }
.kpi.blue { border-color:rgba(0,132,255,.45); }
.kpi.green { border-color:rgba(16,185,129,.45); }
.kpi.purple { border-color:rgba(168,85,247,.45); }
.kpi.orange { border-color:rgba(245,158,11,.42); }
.kpi.red { border-color:rgba(239,68,68,.48); }
.alert-card { border-radius:16px; padding:16px; min-height:126px; background:rgba(4,15,30,.9); border:1px solid rgba(100,180,255,.18); }
.alert-card .name { font-size:14px; color:#e2e8f0; font-weight:700; }
.alert-card .big { font-size:30px; font-weight:800; margin:8px 0 2px; }
.alert-red { border-color:rgba(239,68,68,.55); background:linear-gradient(145deg, rgba(70,16,24,.55), rgba(5,14,25,.95)); }
.alert-orange { border-color:rgba(245,158,11,.55); background:linear-gradient(145deg, rgba(75,45,10,.45), rgba(5,14,25,.95)); }
.alert-green { border-color:rgba(34,197,94,.55); background:linear-gradient(145deg, rgba(14,70,38,.45), rgba(5,14,25,.95)); }
.alert-blue { border-color:rgba(14,165,233,.55); background:linear-gradient(145deg, rgba(10,54,85,.45), rgba(5,14,25,.95)); }
.section-title { color:#f8fafc; font-weight:800; font-size:19px; margin:12px 0 10px; }
.score-wrap { display:flex; gap:18px; align-items:center; }
.score-circle { width:118px; height:118px; border-radius:50%; display:flex; flex-direction:column; align-items:center; justify-content:center; background:conic-gradient(#22c55e var(--p), rgba(30,41,59,.9) 0); box-shadow:0 0 25px rgba(34,197,94,.23); }
.score-circle-inner { width:88px; height:88px; border-radius:50%; background:#071321; display:flex; flex-direction:column; align-items:center; justify-content:center; }
.score-number { font-size:34px; font-weight:900; color:#fff; line-height:1; }
.score-small { color:#cbd5e1; font-size:12px; }
.badge { display:inline-block; padding:4px 9px; border-radius:999px; font-size:12px; font-weight:800; }
.badge-red { background:rgba(239,68,68,.18); color:#fca5a5; }
.badge-orange { background:rgba(245,158,11,.18); color:#fcd34d; }
.badge-green { background:rgba(34,197,94,.18); color:#86efac; }
div[data-testid="stDataFrame"] { border: 1px solid rgba(80,150,255,.2); border-radius: 14px; overflow:hidden; }
.stTabs [data-baseweb="tab-list"] { gap:8px; }
.stTabs [data-baseweb="tab"] { background:rgba(15,36,64,.8); border-radius:12px; color:#dbeafe; padding:8px 14px; }
.stTabs [aria-selected="true"] { background:#1452d9 !important; }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Helpers
# -----------------------------
def normalize_col_name(name: str) -> str:
    name = str(name).strip().lower()
    tr_map = str.maketrans("çğıöşüİ", "cgiosui")
    name = name.translate(tr_map)
    name = re.sub(r"[^a-z0-9]+", "_", name)
    return re.sub(r"_+", "_", name).strip("_")

def money_fmt(x):
    try:
        if pd.isna(x): return "₺ 0"
        return "₺ " + f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "₺ 0"

def pct_fmt(x):
    try:
        if pd.isna(x): return "%0,0"
        return "%" + f"{float(x)*100:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "%0,0"

def num_fmt(x, digits=0):
    try:
        if pd.isna(x): return "0"
        return f"{float(x):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"

def excel_date_to_datetime(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_datetime(s, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def find_first_column(columns, keywords):
    normalized = {c: normalize_col_name(c) for c in columns}
    for original, norm in normalized.items():
        for kw in keywords:
            if kw in norm:
                return original
    return None

def pick_best_sheet(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file), "CSV"
    xl = pd.ExcelFile(uploaded_file)
    best_name, best_df, best_score = None, None, -1
    keys = ["urun", "ürün", "adet", "ciro", "stok", "kar", "maliyet", "tarih", "miad", "kategori"]
    for sheet in xl.sheet_names:
        tmp = pd.read_excel(uploaded_file, sheet_name=sheet)
        cols = " ".join(map(str, tmp.columns)).lower()
        score = sum(k in cols for k in keys) + min(len(tmp), 3000) / 10000
        if score > best_score:
            best_name, best_df, best_score = sheet, tmp, score
    return best_df, best_name

def create_sample_data():
    today = pd.Timestamp.today().normalize()
    rng = np.random.default_rng(42)
    products = [
        ("PAROL 500 MG 20 TB", "İlaç", "Ağrı Kesici", 27, 17.94, 24.68, 214),
        ("AUGMENTIN 1000 MG 14 TB", "İlaç", "Antibiyotik", 12, 95.56, 132.28, 140),
        ("D VİTAMİNİ DAMLA", "Takviye", "Vitamin", 18, 68, 95, 180),
        ("SOLANTE GOLD", "Dermokozmetik", "Güneş", 15, 350, 520, 85),
        ("BEPANTHOL KREM", "Dermokozmetik", "Bakım", 60, 180, 310, 22),
        ("OMEGA 3", "Takviye", "Vitamin", 35, 210, 410, 7),
        ("ESKİ STOK ÜRÜN A", "OTC", "Diğer", 25, 140, 250, 0),
        ("ESKİ STOK ÜRÜN B", "Bebek", "Diğer", 18, 95, 175, 0),
    ]
    rows = []
    start = today - pd.Timedelta(days=180)
    for p, cat, sub, stock, cost, price, total60 in products:
        days = rng.choice(np.arange(180), size=max(1, int(total60)), replace=True) if total60 else []
        if total60 == 0:
            rows.append([today - pd.Timedelta(days=170), "Demo", p, cat, sub, 0, cost, price, 0, 0, 0, stock, today + pd.Timedelta(days=int(rng.integers(30, 520)))])
        for d in days:
            adet = int(rng.integers(1, 4))
            tarih = start + pd.Timedelta(days=int(d))
            rows.append([tarih, "Demo", p, cat, sub, adet, cost, price, adet*price, adet*cost, adet*(price-cost), stock, today + pd.Timedelta(days=int(rng.integers(30, 520)))])
    return pd.DataFrame(rows, columns=["Tarih","Kaynak","Ürün Adı","Kategori","Alt Kategori","Adet","Alış Birim TL","Satış Birim TL","Ciro TL","Maliyet TL","Brüt Kar TL","Mevcut Stok","Miad Tarihi"])

def read_data(uploaded_file):
    if uploaded_file is None:
        return create_sample_data(), "Örnek Veri"
    return pick_best_sheet(uploaded_file)

def standardize_transactions(raw_df, mapping):
    out = pd.DataFrame()
    n = len(raw_df)
    out["urun"] = raw_df[mapping["urun"]].astype(str) if mapping.get("urun") else "Bilinmeyen Ürün"
    out["kategori"] = raw_df[mapping["kategori"]].astype(str) if mapping.get("kategori") else "Genel"
    out["alt_kategori"] = raw_df[mapping["alt_kategori"]].astype(str) if mapping.get("alt_kategori") else out["kategori"]
    out["adet"] = pd.to_numeric(raw_df[mapping["adet"]], errors="coerce").fillna(0) if mapping.get("adet") else 0
    out["stok"] = pd.to_numeric(raw_df[mapping["stok"]], errors="coerce").fillna(0) if mapping.get("stok") else 0
    out["alis"] = pd.to_numeric(raw_df[mapping["alis"]], errors="coerce").fillna(0) if mapping.get("alis") else 0
    out["satis"] = pd.to_numeric(raw_df[mapping["satis"]], errors="coerce").fillna(0) if mapping.get("satis") else 0
    out["ciro"] = pd.to_numeric(raw_df[mapping["ciro"]], errors="coerce").fillna(0) if mapping.get("ciro") else out["adet"] * out["satis"]
    out["maliyet"] = pd.to_numeric(raw_df[mapping["maliyet"]], errors="coerce").fillna(0) if mapping.get("maliyet") else out["adet"] * out["alis"]
    out["brut_kar"] = pd.to_numeric(raw_df[mapping["brut_kar"]], errors="coerce").fillna(np.nan) if mapping.get("brut_kar") else np.nan
    out["brut_kar"] = out["brut_kar"].fillna(out["ciro"] - out["maliyet"])
    out["son_60_cikis"] = pd.to_numeric(raw_df[mapping["son_60"]], errors="coerce").fillna(np.nan) if mapping.get("son_60") else np.nan
    if mapping.get("tarih"):
        out["tarih"] = excel_date_to_datetime(raw_df[mapping["tarih"]])
    else:
        out["tarih"] = pd.Timestamp.today().normalize()
    if mapping.get("miad"):
        out["miad"] = excel_date_to_datetime(raw_df[mapping["miad"]])
    else:
        out["miad"] = pd.NaT
    return out

def make_product_summary(tx):
    today = pd.Timestamp.today().normalize()
    max_date = tx["tarih"].max()
    if pd.isna(max_date): max_date = today
    d30 = max_date - pd.Timedelta(days=30)
    d14 = max_date - pd.Timedelta(days=14)
    d7 = max_date - pd.Timedelta(days=7)

    g = tx.groupby("urun", dropna=False)
    prod = g.agg(
        kategori=("kategori", lambda x: x.mode().iloc[0] if len(x.mode()) else "Genel"),
        alt_kategori=("alt_kategori", lambda x: x.mode().iloc[0] if len(x.mode()) else "Genel"),
        stok=("stok", "max"),
        alis=("alis", "mean"),
        satis=("satis", "mean"),
        ciro_toplam=("ciro", "sum"),
        maliyet_toplam=("maliyet", "sum"),
        brut_kar_toplam=("brut_kar", "sum"),
        adet_toplam=("adet", "sum"),
        miad=("miad", "min"),
        son_60_cikis_raw=("son_60_cikis", "max"),
    ).reset_index()

    def sum_since(days_ago):
        return tx[tx["tarih"] >= days_ago].groupby("urun")["adet"].sum()
    prod = prod.merge(sum_since(d30).rename("satis_30"), on="urun", how="left")
    prod = prod.merge(sum_since(d14).rename("satis_14"), on="urun", how="left")
    prod = prod.merge(sum_since(d7).rename("satis_7"), on="urun", how="left")
    prod[["satis_30","satis_14","satis_7"]] = prod[["satis_30","satis_14","satis_7"]].fillna(0)

    # If file has ready "Son 60 Gün Çıkış" but tarih range is not ideal, use it as a fallback signal.
    prod["satis_60"] = np.where(prod["son_60_cikis_raw"].notna(), prod["son_60_cikis_raw"], prod["satis_30"] * 2)
    prod["gunluk_tuketim"] = np.where(prod["satis_60"] > 0, prod["satis_60"] / 60, np.where(prod["satis_30"] > 0, prod["satis_30"] / 30, 0))
    prod["bitis_gunu"] = np.where(prod["gunluk_tuketim"] > 0, prod["stok"] / prod["gunluk_tuketim"], np.inf)
    prod["stok_degeri"] = prod["stok"] * prod["alis"]
    prod["kar_marji"] = np.where(prod["ciro_toplam"] > 0, prod["brut_kar_toplam"] / prod["ciro_toplam"], 0)
    prod["miad_kalan_gun"] = (prod["miad"] - today).dt.days
    prod["trend_14_30"] = np.where(prod["satis_30"] > 0, (prod["satis_14"]*2 - prod["satis_30"]) / prod["satis_30"], 0)
    return prod, max_date

def classify(row, critical_days, warning_days):
    if row["gunluk_tuketim"] <= 0 and row["stok"] > 0: return "Hareketsiz"
    if row["bitis_gunu"] <= critical_days: return "Kritik"
    if row["bitis_gunu"] <= warning_days: return "Dikkat"
    return "Güvenli"

def ayca_score(prod, critical_days, miad_days):
    if prod.empty: return 0
    total = len(prod)
    critical = len(prod[(prod["gunluk_tuketim"]>0)&(prod["bitis_gunu"]<=critical_days)]) / total
    dead = len(prod[(prod["gunluk_tuketim"]<=0)&(prod["stok"]>0)]) / total
    miad = len(prod[(prod["miad_kalan_gun"].notna())&(prod["miad_kalan_gun"]<=miad_days)]) / total
    margin = len(prod[(prod["kar_marji"]<0.12)&(prod["adet_toplam"]>0)]) / total
    return int(max(0, min(100, round(100 - critical*30 - dead*25 - miad*25 - margin*20))))

def plot_layout(fig, height=None):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#dbeafe", family="Inter"),
        margin=dict(l=10,r=10,t=45,b=10),
        height=height,
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_xaxes(gridcolor="rgba(148,163,184,.16)")
    fig.update_yaxes(gridcolor="rgba(148,163,184,.16)")
    return fig

def spark(values, color="#00a3ff"):
    fig = go.Figure(go.Scatter(y=values, mode="lines", line=dict(color=color, width=2), fill="tozeroy", fillcolor="rgba(0,130,255,.18)"))
    fig.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=54, margin=dict(l=0,r=0,t=0,b=0), xaxis=dict(visible=False), yaxis=dict(visible=False))
    return fig

def excel_report(tx, prod):
    output = BytesIO()
    export_prod = prod.replace([np.inf, -np.inf], np.nan).copy()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_prod.to_excel(writer, sheet_name="Urun_Yorumlari", index=False)
        tx.to_excel(writer, sheet_name="Standart_Islemler", index=False)
    return output.getvalue()

# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.markdown("## 💊 AYÇA Insight V4.1")
st.sidebar.caption("Dinamik dosya yorumlayan karar asistanı")
eczane_adi = st.sidebar.text_input("Eczane Adı", value="Abdullah Eczanesi")
kullanici_adi = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")
uploaded_file = st.sidebar.file_uploader("Excel / CSV yükle", type=["xlsx", "xls", "csv"])
critical_days = st.sidebar.slider("Kritik stok eşiği", 1, 10, 5)
warning_days = st.sidebar.slider("Dikkat stok eşiği", 6, 30, 10)
miad_warning_days = st.sidebar.slider("Miad uyarı günü", 15, 180, 60)

raw_df, sheet_name = read_data(uploaded_file)
columns = list(raw_df.columns)

auto = {
    "tarih": find_first_column(columns, ["tarih", "date", "islem_tarihi"]),
    "urun": find_first_column(columns, ["urun_adi", "urun", "ilac", "malzeme", "stok_adi", "ticari_ad"]),
    "kategori": find_first_column(columns, ["kategori", "ana_grup", "grup"]),
    "alt_kategori": find_first_column(columns, ["alt_kategori", "alt_grup"]),
    "adet": find_first_column(columns, ["adet", "miktar", "satis_miktari", "qty"]),
    "alis": find_first_column(columns, ["alis_birim", "alis", "maliyet_birim", "net_alis"]),
    "satis": find_first_column(columns, ["satis_birim", "satis_fiyat", "etiket", "perakende"]),
    "ciro": find_first_column(columns, ["ciro", "tutar", "satis_tl"]),
    "maliyet": find_first_column(columns, ["maliyet_tl", "maliyet"]),
    "brut_kar": find_first_column(columns, ["brut_kar", "brüt_kar", "kar_tl"]),
    "stok": find_first_column(columns, ["mevcut_stok", "stok", "kalan"]),
    "miad": find_first_column(columns, ["miad", "skt", "son_kullanma", "expiry"]),
    "son_60": find_first_column(columns, ["son_60", "60_gun", "son60"]),
}

with st.sidebar.expander("Kolon eşleştirme", expanded=False):
    def sel(label, key):
        opts = [None] + columns
        default = auto.get(key)
        idx = opts.index(default) if default in opts else 0
        return st.selectbox(label, opts, index=idx)
    mapping = {
        "tarih": sel("Tarih", "tarih"),
        "urun": sel("Ürün Adı", "urun"),
        "kategori": sel("Kategori", "kategori"),
        "alt_kategori": sel("Alt Kategori", "alt_kategori"),
        "adet": sel("Adet / Satış Miktarı", "adet"),
        "alis": sel("Alış Birim", "alis"),
        "satis": sel("Satış Birim", "satis"),
        "ciro": sel("Ciro TL", "ciro"),
        "maliyet": sel("Maliyet TL", "maliyet"),
        "brut_kar": sel("Brüt Kâr TL", "brut_kar"),
        "stok": sel("Mevcut Stok", "stok"),
        "miad": sel("Miad Tarihi", "miad"),
        "son_60": sel("Son 60 Gün Çıkış", "son_60"),
    }

required = [mapping.get("urun"), mapping.get("adet"), mapping.get("stok")]
if not all(required):
    st.warning("Ürün Adı, Adet ve Mevcut Stok kolonları eşleşmeli. Sol menüden kolon eşleştirmeyi kontrol edin.")
    st.dataframe(raw_df.head(30), use_container_width=True)
    st.stop()

tx = standardize_transactions(raw_df, mapping)
prod, max_date = make_product_summary(tx)
prod["durum"] = prod.apply(lambda r: classify(r, critical_days, warning_days), axis=1)
score = ayca_score(prod, critical_days, miad_warning_days)

# Main dynamic metrics
period_start = tx["tarih"].min()
period_end = tx["tarih"].max()
if pd.isna(period_start): period_start = pd.Timestamp.today().normalize()
if pd.isna(period_end): period_end = pd.Timestamp.today().normalize()

current_ciro = float(tx["ciro"].sum())
current_profit = float(tx["brut_kar"].sum())
current_margin = current_profit / current_ciro if current_ciro else 0
transactions = int(tx[tx["adet"] > 0].shape[0])
total_units = float(tx["adet"].sum())
avg_basket = current_ciro / transactions if transactions else 0

critical_df = prod[(prod["gunluk_tuketim"]>0)&(prod["bitis_gunu"]<=critical_days)].sort_values("bitis_gunu")
miad_df = prod[(prod["miad_kalan_gun"].notna())&(prod["miad_kalan_gun"]<=miad_warning_days)].sort_values("miad_kalan_gun")
dead_df = prod[(prod["gunluk_tuketim"]<=0)&(prod["stok"]>0)].sort_values("stok_degeri", ascending=False)
trend_df = prod.sort_values("trend_14_30", ascending=False)
trend_positive = trend_df[trend_df["trend_14_30"] > 0]
dead_stock_value = float(dead_df["stok_degeri"].sum())
stock_turnover = total_units / max(prod["stok"].sum(), 1)

# Deltas from file itself: last 30 days vs previous 30 days
last30 = tx[tx["tarih"] >= period_end - pd.Timedelta(days=30)]
prev30 = tx[(tx["tarih"] < period_end - pd.Timedelta(days=30)) & (tx["tarih"] >= period_end - pd.Timedelta(days=60))]
def delta(now, prev):
    if prev == 0: return 0 if now == 0 else 1
    return (now - prev) / prev
ciro_delta = delta(last30["ciro"].sum(), prev30["ciro"].sum())
profit_delta = delta(last30["brut_kar"].sum(), prev30["brut_kar"].sum())
margin_delta = current_margin - (prev30["brut_kar"].sum()/prev30["ciro"].sum() if prev30["ciro"].sum() else current_margin)
trx_delta = delta(last30[last30["adet"]>0].shape[0], prev30[prev30["adet"]>0].shape[0])
basket_delta = delta((last30["ciro"].sum()/max(last30[last30["adet"]>0].shape[0],1)), (prev30["ciro"].sum()/max(prev30[prev30["adet"]>0].shape[0],1)))

# -----------------------------
# Header
# -----------------------------
st.markdown(f"""
<div class="ayca-topbar">
  <div class="ayca-title">
    <h1>Günaydın {kullanici_adi} 👋</h1>
    <p>{period_start.strftime('%d.%m.%Y')} - {period_end.strftime('%d.%m.%Y')} · {eczane_adi} · Okunan sayfa: {sheet_name}</p>
  </div>
  <div style="display:flex; gap:10px; align-items:center;">
    <div class="glass-card" style="padding:10px 14px;">📅 {datetime.now().strftime('%d %B %Y')}</div>
    <div class="glass-card" style="padding:10px 14px;">🔄 Dinamik Veri</div>
  </div>
</div>
""", unsafe_allow_html=True)

# AI assistant dynamic
ai_lines = []
if len(critical_df):
    r = critical_df.iloc[0]
    ai_lines.append(f"{len(critical_df)} ürün {critical_days} gün içinde bitebilir. En kritik ürün: <b>{r['urun']}</b> ({r['bitis_gunu']:.0f} gün).")
if len(miad_df):
    r = miad_df.iloc[0]
    ai_lines.append(f"{len(miad_df)} ürünün miadı yaklaşıyor. En yakın miad: <b>{r['urun']}</b> ({r['miad_kalan_gun']:.0f} gün).")
if len(dead_df):
    ai_lines.append(f"Hareketsiz stok değeri yaklaşık <b>{money_fmt(dead_stock_value)}</b>.")
if len(trend_positive):
    r = trend_positive.iloc[0]
    ai_lines.append(f"Son 14 günde yükselen ürün: <b>{r['urun']}</b> ({pct_fmt(r['trend_14_30'])}).")
if not ai_lines:
    ai_lines.append("Dosyada kritik bir stok veya miad baskısı düşük görünüyor.")

st.markdown(f"""
<div class="ai-card">
  <h3>🤖 AYÇA AI Asistan</h3>
  <p>{' '.join(ai_lines)}</p>
  <p>Bu ekran hazır rakam göstermez; yüklenen dosyadaki işlem satırlarını, stokları, miadları ve kâr alanlarını yorumlayarak hesaplar.</p>
</div>
""", unsafe_allow_html=True)

# Alert cards
acols = st.columns(5)
alerts = [
    ("Kritik Stok", f"{len(critical_df)} ÜRÜN", f"{critical_days} gün içinde bitebilir", "alert-red"),
    ("Miadı Yaklaşan", f"{len(miad_df)} ÜRÜN", f"{miad_warning_days} gün içinde", "alert-orange"),
    ("Ölü Stok", f"{len(dead_df)} ÜRÜN", "Satış hızı yok", "alert-orange"),
    ("Trend Artan", pct_fmt(trend_positive['trend_14_30'].max()) if len(trend_positive) else "%0,0", "En yüksek artış", "alert-green"),
    ("Ölü Stok Değeri", money_fmt(dead_stock_value), "Tahmini stok maliyeti", "alert-blue"),
]
for col, (name, big, sub, klass) in zip(acols, alerts):
    with col:
        st.markdown(f"<div class='alert-card {klass}'><div class='name'>{name}</div><div class='big'>{big}</div><div style='color:#cbd5e1'>{sub}</div></div>", unsafe_allow_html=True)

# KPI cards
st.markdown("<div class='section-title'>Finansal Durum</div>", unsafe_allow_html=True)
kcols = st.columns(5)
kpis = [
    ("Güncel Ciro", money_fmt(current_ciro), ciro_delta, "blue"),
    ("Brüt Kâr", money_fmt(current_profit), profit_delta, "green"),
    ("Brüt Kâr Marjı", pct_fmt(current_margin), margin_delta, "purple"),
    ("İşlem Sayısı", num_fmt(transactions), trx_delta, "orange"),
    ("Ort. Sepet Tutarı", money_fmt(avg_basket), basket_delta, "blue"),
]
for col, (label, val, d, klass) in zip(kcols, kpis):
    with col:
        arrow = "▲" if d >= 0 else "▼"
        st.markdown(f"<div class='kpi {klass}'><div class='label'>{label}</div><div class='value'>{val}</div><div class='delta'>{arrow} {pct_fmt(abs(d))}</div><div style='height:54px'></div></div>", unsafe_allow_html=True)

# Charts row 1
row1_left, row1_mid, row1_right = st.columns([1.55, 1.25, 1])
with row1_left:
    st.markdown("<div class='glass-card'><div class='section-title'>Ciro Trendi</div>", unsafe_allow_html=True)
    monthly = tx.dropna(subset=["tarih"]).copy()
    monthly["ay"] = monthly["tarih"].dt.to_period("M").dt.to_timestamp()
    m = monthly.groupby("ay", as_index=False).agg(ciro=("ciro","sum"), kar=("brut_kar","sum"))
    if len(m):
        fig = px.line(m, x="ay", y="ciro", markers=True, labels={"ay":"Ay", "ciro":"Ciro"})
        fig.update_traces(line=dict(width=4, color="#0b84ff"), fill="tozeroy", fillcolor="rgba(11,132,255,.22)")
        st.plotly_chart(plot_layout(fig, 320), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
with row1_mid:
    st.markdown("<div class='glass-card'><div class='section-title'>Kategori Performans Analizi</div>", unsafe_allow_html=True)
    cat = tx.groupby("kategori", as_index=False).agg(ciro=("ciro","sum"), kar=("brut_kar","sum"), adet=("adet","sum"))
    if len(cat):
        fig = px.pie(cat.sort_values("ciro", ascending=False), names="kategori", values="ciro", hole=.55)
        fig.update_traces(textinfo="none")
        st.plotly_chart(plot_layout(fig, 320), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
with row1_right:
    st.markdown("<div class='glass-card'><div class='section-title'>Stok Devir Hızı</div>", unsafe_allow_html=True)
    gauge_val = float(min(10, stock_turnover))
    fig = go.Figure(go.Indicator(mode="gauge+number", value=gauge_val, number={"suffix":"", "font":{"size":44}}, gauge={"axis":{"range":[0,10]}, "bar":{"color":"#0b84ff"}, "steps":[{"range":[0,4],"color":"rgba(239,68,68,.2)"},{"range":[4,7],"color":"rgba(245,158,11,.22)"},{"range":[7,10],"color":"rgba(34,197,94,.24)"}]}))
    st.plotly_chart(plot_layout(fig, 320), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# Charts row 2
r2a, r2b, r2c = st.columns(3)
with r2a:
    st.markdown("<div class='glass-card'><div class='section-title'>En Çok Satan 5 Ürün</div>", unsafe_allow_html=True)
    top_sold = prod.sort_values("adet_toplam", ascending=False).head(5)
    fig = px.bar(top_sold, y="urun", x="adet_toplam", orientation="h", labels={"adet_toplam":"Adet", "urun":""})
    fig.update_traces(marker_color="#3b82f6")
    st.plotly_chart(plot_layout(fig, 300), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
with r2b:
    st.markdown("<div class='glass-card'><div class='section-title'>En Kârlı 5 Ürün</div>", unsafe_allow_html=True)
    top_profit = prod.sort_values("brut_kar_toplam", ascending=False).head(5)
    fig = px.bar(top_profit, y="urun", x="brut_kar_toplam", orientation="h", labels={"brut_kar_toplam":"Kâr", "urun":""})
    fig.update_traces(marker_color="#22c55e")
    st.plotly_chart(plot_layout(fig, 300), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
with r2c:
    st.markdown("<div class='glass-card'><div class='section-title'>Ölü Stok Analizi</div>", unsafe_allow_html=True)
    if len(dead_df):
        fig = px.pie(dead_df.head(8), names="urun", values="stok_degeri", hole=.58)
        fig.update_traces(textinfo="none")
        st.plotly_chart(plot_layout(fig, 300), use_container_width=True)
    else:
        st.success("Hareketsiz stok görünmüyor.")
    st.markdown("</div>", unsafe_allow_html=True)

# Tables + trend mini cards
st.markdown("<div class='section-title'>Operasyonel Yorumlar</div>", unsafe_allow_html=True)
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Stok Bitiş Tahmini", "Sipariş Asistanı", "Miad Takibi", "Ölü Stok", "Rapor"])

with tab1:
    show = prod.copy().replace([np.inf, -np.inf], np.nan)
    show["Tahmini Bitiş"] = show["bitis_gunu"].round(1)
    show["Günlük Tüketim"] = show["gunluk_tuketim"].round(2)
    show["Kâr Marjı"] = (show["kar_marji"]*100).round(1)
    st.dataframe(show[["durum","urun","kategori","stok","satis_30","Günlük Tüketim","Tahmini Bitiş","stok_degeri","Kâr Marjı"]].sort_values(["durum","Tahmini Bitiş"]), use_container_width=True, hide_index=True)
with tab2:
    order = prod[(prod["gunluk_tuketim"]>0)&(prod["bitis_gunu"]<=warning_days)].copy()
    order["önerilen_sipariş"] = np.ceil((warning_days*2*order["gunluk_tuketim"]) - order["stok"]).clip(lower=1)
    order["tahmini_maliyet"] = order["önerilen_sipariş"] * order["alis"]
    st.dataframe(order[["durum","urun","kategori","stok","gunluk_tuketim","bitis_gunu","önerilen_sipariş","tahmini_maliyet"]].sort_values("bitis_gunu"), use_container_width=True, hide_index=True)
    st.info(f"Tahmini sipariş maliyeti: {money_fmt(order['tahmini_maliyet'].sum())}")
with tab3:
    st.dataframe(miad_df[["urun","kategori","stok","miad","miad_kalan_gun","stok_degeri"]].replace([np.inf,-np.inf],np.nan), use_container_width=True, hide_index=True)
with tab4:
    st.dataframe(dead_df[["urun","kategori","stok","alis","stok_degeri","miad_kalan_gun"]].replace([np.inf,-np.inf],np.nan), use_container_width=True, hide_index=True)
with tab5:
    report = excel_report(tx, prod)
    st.download_button("📥 Dinamik AYÇA V4.1 Raporu İndir", data=report, file_name=f"AYCA_Insight_V4_1_Rapor_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Okunan ham veri önizlemesi")
    st.dataframe(raw_df.head(50), use_container_width=True)

# Bottom score
st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
score_pct = f"{score}%"
st.markdown(f"""
<div class="score-wrap">
  <div class="score-circle" style="--p:{score_pct};"><div class="score-circle-inner"><div class="score-number">{score}</div><div class="score-small">/100</div></div></div>
  <div><h3 style="margin:0;color:white;">AYÇA Insight Skoru</h3><p style="color:#94a3b8; margin-top:6px;">Bu skor yüklenen dosyadaki kritik stok, miad, ölü stok ve kârlılık dengesinden hesaplandı.</p></div>
</div>
""", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

st.caption("AYÇA Insight V4.1 · Veriyi okur, yorumlar, ekrana dinamik sonuç basar. Hazır/hard-coded rakam kullanılmaz.")
