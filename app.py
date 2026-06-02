import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import numpy as np

st.set_page_config(
    page_title="AYÇA Insight v2.1",
    page_icon="💊",
    layout="wide"
)

# =====================================================
# CSS
# =====================================================
st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, #f6f9ff 0%, #eef3f8 100%);
}
.big-title {
    font-size: 38px;
    font-weight: 900;
    color: #12304d;
    margin-bottom: 4px;
}
.sub-title {
    font-size: 17px;
    color: #607084;
    margin-bottom: 22px;
}
.hero-box {
    padding: 24px;
    border-radius: 24px;
    background: linear-gradient(135deg, #15395b, #2f80ed);
    color: white;
    box-shadow: 0 8px 30px rgba(47,128,237,0.28);
    margin-bottom: 18px;
}
.hero-score {
    font-size: 48px;
    font-weight: 900;
}
.hero-small {
    font-size: 15px;
    opacity: 0.9;
}
.card {
    background: white;
    padding: 20px;
    border-radius: 18px;
    border: 1px solid #e7edf5;
    box-shadow: 0 5px 18px rgba(22, 50, 79, 0.06);
    margin-bottom: 12px;
}
.task-card {
    background: #ffffff;
    border-left: 7px solid #2f80ed;
    padding: 17px 18px;
    border-radius: 16px;
    margin-bottom: 12px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.05);
}
.success-card {
    background: #effdf5;
    border-left: 7px solid #27ae60;
    padding: 17px 18px;
    border-radius: 16px;
    margin-bottom: 12px;
}
.warning-card {
    background: #fff7ec;
    border-left: 7px solid #f2994a;
    padding: 17px 18px;
    border-radius: 16px;
    margin-bottom: 12px;
}
.danger-card {
    background: #fff1f1;
    border-left: 7px solid #eb5757;
    padding: 17px 18px;
    border-radius: 16px;
    margin-bottom: 12px;
}
.small-muted {
    color: #6b7c90;
    font-size: 13px;
}
.section-title {
    font-size: 23px;
    font-weight: 800;
    color: #12304d;
    margin-top: 12px;
    margin-bottom: 10px;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# Helpers
# =====================================================
def normalize_col(col):
    return (
        str(col).strip().lower()
        .replace("ı", "i")
        .replace("ğ", "g")
        .replace("ü", "u")
        .replace("ş", "s")
        .replace("ö", "o")
        .replace("ç", "c")
    )


def find_column(df, keywords):
    normalized = {normalize_col(c): c for c in df.columns}
    for key in keywords:
        key = normalize_col(key)
        for n_col, original in normalized.items():
            if key in n_col:
                return original
    return None


def money(x):
    try:
        return f"{float(x):,.0f} ₺".replace(",", ".")
    except Exception:
        return "0 ₺"


def number(x):
    try:
        return f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "0"


def to_numeric(df, col):
    if col and col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def safe_sum(df, col):
    if col is None or col not in df.columns:
        return 0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def score_label(score):
    if score >= 80:
        return "Çok İyi"
    if score >= 60:
        return "Geliştirilebilir"
    return "Dikkat Gerekli"


def score_color_class(score):
    if score >= 80:
        return "success-card"
    if score >= 60:
        return "warning-card"
    return "danger-card"


def calculate_scores(total_sales, total_qty, total_profit, dead_count, critical_count, low_profit_count, total_rows):
    profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

    stock_score = 100
    if total_rows > 0:
        stock_score -= min((critical_count / total_rows) * 180, 45)
        stock_score -= min((dead_count / total_rows) * 120, 35)
    stock_score = int(max(0, min(100, stock_score)))

    profit_score = 100
    if profit_margin < 10:
        profit_score = 45
    elif profit_margin < 15:
        profit_score = 60
    elif profit_margin < 25:
        profit_score = 78
    else:
        profit_score = 90
    profit_score -= min(low_profit_count * 1.2, 20)
    profit_score = int(max(0, min(100, profit_score)))

    sales_score = 85 if total_sales > 0 and total_qty > 0 else 40
    if total_qty < 20:
        sales_score -= 15
    sales_score = int(max(0, min(100, sales_score)))

    risk_score = 100
    risk_score -= min(critical_count * 2, 40)
    risk_score -= min(dead_count * 1, 35)
    risk_score = int(max(0, min(100, risk_score)))

    general = int((stock_score * 0.35) + (profit_score * 0.25) + (sales_score * 0.20) + (risk_score * 0.20))
    return general, stock_score, profit_score, sales_score, risk_score, profit_margin


def build_order_recommendations(df, product_col, qty_col, stock_col, sales_col, profit_col, barcode_col=None):
    data = df.copy()
    data["AYCA_Hiz_Skoru"] = data[qty_col] * 2 + (data[sales_col] / 1000)
    data["AYCA_Siparis_Onceligi"] = np.select(
        [
            (data[stock_col] <= 2) & (data[qty_col] >= 5),
            (data[stock_col] <= 5) & (data[qty_col] >= 3),
            (data[stock_col] <= 2),
        ],
        ["Acil Sipariş", "Takip Et", "Kontrol Et"],
        default="Siparişe Gerek Yok"
    )
    data["AYCA_Onerilen_Siparis"] = np.where(
        data["AYCA_Siparis_Onceligi"] == "Acil Sipariş",
        np.maximum(10, data[qty_col] * 2 - data[stock_col]),
        np.where(data["AYCA_Siparis_Onceligi"] == "Takip Et", np.maximum(5, data[qty_col] - data[stock_col]), 0)
    )
    cols = []
    if barcode_col:
        cols.append(barcode_col)
    cols += [product_col, stock_col, qty_col, sales_col, profit_col, "AYCA_Siparis_Onceligi", "AYCA_Onerilen_Siparis"]
    return data[cols].sort_values(["AYCA_Siparis_Onceligi", qty_col], ascending=[True, False])


def build_campaign_recommendations(df, product_col, qty_col, stock_col, sales_col, profit_col, category_col=None, barcode_col=None):
    data = df.copy()
    data["AYCA_Kampanya_Tipi"] = np.select(
        [
            (data[stock_col] >= 10) & (data[qty_col] == 0),
            (data[stock_col] >= 5) & (data[qty_col] <= 1),
            (data[stock_col] >= 3) & (data[sales_col] == 0),
        ],
        ["Kasa Önü / Sepet", "Raf Öne Alma", "Fiyat-Kampanya Kontrol"],
        default="Kampanya Gerekmez"
    )
    data["AYCA_Kampanya_Notu"] = np.where(
        data["AYCA_Kampanya_Tipi"] != "Kampanya Gerekmez",
        "Stokta bekleyen ürün. Görünürlük veya kampanya ile hareketlendirilebilir.",
        ""
    )
    cols = []
    if barcode_col:
        cols.append(barcode_col)
    cols += [product_col]
    if category_col:
        cols.append(category_col)
    cols += [stock_col, qty_col, sales_col, "AYCA_Kampanya_Tipi", "AYCA_Kampanya_Notu"]
    return data[data["AYCA_Kampanya_Tipi"] != "Kampanya Gerekmez"][cols].sort_values(stock_col, ascending=False)


def cross_sell_suggestions(category_text):
    t = normalize_col(category_text)
    if any(k in t for k in ["bebek", "baby", "anne"]):
        return "Bebek ürünü alan müşteriye pişik kremi, ıslak mendil veya bebek şampuanı hatırlatılabilir."
    if any(k in t for k in ["gunes", "dermo", "kozmetik", "cilt"]):
        return "Güneş koruyucu alan müşteriye nemlendirici, leke karşıtı ürün veya dudak koruyucu önerilebilir."
    if any(k in t for k in ["vitamin", "takviye", "mineral"]):
        return "Vitamin alan müşteriye kullanım düzeni, tamamlayıcı mineral veya bağışıklık destek ürünü danışmanlığı verilebilir."
    if any(k in t for k in ["agri", "ağrı", "analjezik"]):
        return "Ağrı kesici alan müşteriye doğru kullanım ve mide hassasiyeti konusunda danışmanlık verilebilir."
    if any(k in t for k in ["sac", "saç", "hair"]):
        return "Saç ürünü alan müşteriye şampuan, serum veya saç vitamini kombinasyonu önerilebilir."
    return "Kategoriye uygun tamamlayıcı ürün önerisi hazırlanabilir."


def build_daily_tasks(order_df, campaign_df, critical_count, dead_count, top_category, profit_margin):
    tasks = []
    if not order_df.empty:
        urgent = order_df[order_df["AYCA_Siparis_Onceligi"] == "Acil Sipariş"]
        if not urgent.empty:
            p = urgent.iloc[0]
            tasks.append(f"Acil sipariş kontrolü yap: {p.iloc[1] if len(p) > 1 else 'yüksek satışlı ürün'} kritik stokta görünüyor.")
    if critical_count > 0:
        tasks.append(f"Kritik stok listesini kontrol et: {critical_count} ürün satış kaybı riski taşıyor.")
    if not campaign_df.empty:
        p = campaign_df.iloc[0]
        product_name = p.iloc[1] if len(p) > 1 else p.iloc[0]
        tasks.append(f"Kampanya/raf aksiyonu al: {product_name} stokta bekliyor, kasa önü veya raf öne alma denenebilir.")
    if top_category:
        tasks.append(f"Bugünün kategori odağı: {top_category}. Bu kategori için çapraz satış konuşması hazırlayın.")
    if profit_margin < 15:
        tasks.append("Kârlılığı gözden geçir: Çok satan ama düşük marjlı ürünleri kontrol edin.")
    if dead_count > 0:
        tasks.append(f"Ölü stok temizliği başlat: {dead_count} ürün hareket görmüyor.")

    default_tasks = [
        "Gün sonunda en çok satan 10 ürünün stok durumunu kontrol edin.",
        "Kasa önü ürünlerini bugün en az bir kez yenileyin.",
        "Yüksek kârlı ürünleri görünür rafa alın.",
        "Reçete dışı satış fırsatlarını ekip içinde paylaşın.",
        "Günün sonunda AYÇA skorunu tekrar kontrol edin."
    ]
    for t in default_tasks:
        if len(tasks) >= 5:
            break
        tasks.append(t)
    return tasks[:5]

# =====================================================
# Header
# =====================================================
st.markdown('<div class="big-title">💊 AYÇA Insight v2.1</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Rapor okumaz; eczaneniz için bugünkü aksiyonu gösterir.</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Tebeos / eczane satış-stok Excel dosyanızı yükleyin", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Excel dosyasını yükleyince AYÇA günlük öneriler, sipariş önerileri, kampanya fırsatları ve skor ekranını oluşturacaktır.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.sidebar.selectbox("Excel Sayfası", xls.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

if df.empty:
    st.warning("Excel dosyası boş görünüyor.")
    st.stop()

# =====================================================
# Column Detection
# =====================================================
product_col = find_column(df, ["urun", "ürün", "mal", "stok adi", "stok adı", "ilac", "ilaç", "product"])
barcode_col = find_column(df, ["barkod", "barcode", "gtin"])
category_col = find_column(df, ["kategori", "grup", "ana grup", "category"])
qty_col = find_column(df, ["miktar", "adet", "satis miktari", "satış miktarı", "qty", "quantity"])
sales_col = find_column(df, ["ciro", "satis tutari", "satış tutarı", "tutar", "net satis", "net satış", "sales"])
cost_col = find_column(df, ["maliyet", "alis", "alış", "cost"])
profit_col = find_column(df, ["kar", "kâr", "kazanc", "kazanç", "profit"])
stock_col = find_column(df, ["stok", "kalan", "mevcut", "stock"])
date_col = find_column(df, ["tarih", "date"])

if product_col is None:
    product_col = df.columns[0]
if qty_col is None:
    df["_AYCA_MIKTAR"] = 1
    qty_col = "_AYCA_MIKTAR"
if sales_col is None:
    df["_AYCA_CIRO"] = 0
    sales_col = "_AYCA_CIRO"
if stock_col is None:
    df["_AYCA_STOK"] = 0
    stock_col = "_AYCA_STOK"
if profit_col is None:
    df["_AYCA_KAR"] = pd.to_numeric(df[sales_col], errors="coerce").fillna(0) * 0.18
    profit_col = "_AYCA_KAR"

for col in [qty_col, sales_col, stock_col, profit_col, cost_col]:
    if col:
        df = to_numeric(df, col)

# =====================================================
# Main KPIs
# =====================================================
total_sales = safe_sum(df, sales_col)
total_qty = safe_sum(df, qty_col)
total_profit = safe_sum(df, profit_col)
avg_basket = total_sales / total_qty if total_qty > 0 else 0

critical_stock = df[df[stock_col] <= 2].copy()
dead_stock = df[(df[stock_col] > 0) & (df[qty_col] == 0)].copy()

product_summary = df.groupby(product_col, as_index=False).agg({
    qty_col: "sum",
    sales_col: "sum",
    profit_col: "sum",
    stock_col: "sum"
})
product_summary["Kar Marjı %"] = product_summary.apply(
    lambda r: (r[profit_col] / r[sales_col] * 100) if r[sales_col] > 0 else 0,
    axis=1
)
low_profit_count = len(product_summary[(product_summary[qty_col] > 0) & (product_summary["Kar Marjı %"] < 15)])

general_score, stock_score, profit_score, sales_score, risk_score, profit_margin = calculate_scores(
    total_sales, total_qty, total_profit, len(dead_stock), len(critical_stock), low_profit_count, len(df)
)

# Top category
if category_col:
    category_summary = df.groupby(category_col, as_index=False).agg({sales_col: "sum", qty_col: "sum"}).sort_values(sales_col, ascending=False)
    top_category = category_summary.iloc[0][category_col] if not category_summary.empty else None
else:
    category_summary = pd.DataFrame()
    top_category = None

order_df = build_order_recommendations(df, product_col, qty_col, stock_col, sales_col, profit_col, barcode_col)
campaign_df = build_campaign_recommendations(df, product_col, qty_col, stock_col, sales_col, profit_col, category_col, barcode_col)
daily_tasks = build_daily_tasks(order_df, campaign_df, len(critical_stock), len(dead_stock), top_category, profit_margin)

# =====================================================
# Hero
# =====================================================
left, right = st.columns([2.1, 1])
with left:
    st.markdown(f"""
    <div class="hero-box">
        <div class="hero-small">Bugünkü AYÇA yorumu</div>
        <div style="font-size:26px;font-weight:800;margin-top:6px;">Eczanenizin önceliği: {"stok riski" if len(critical_stock) > 0 else "satış ve kârlılık takibi"}</div>
        <div style="margin-top:8px;font-size:16px;">{len(critical_stock)} kritik stok, {len(dead_stock)} ölü stok ve %{profit_margin:.1f} kâr marjı tespit edildi.</div>
    </div>
    """, unsafe_allow_html=True)
with right:
    st.markdown(f"""
    <div class="hero-box">
        <div class="hero-small">AYÇA Genel Skor</div>
        <div class="hero-score">{general_score}/100</div>
        <div>{score_label(general_score)}</div>
    </div>
    """, unsafe_allow_html=True)

# KPI cards
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Toplam Ciro", money(total_sales))
k2.metric("Satış Adedi", number(total_qty))
k3.metric("Brüt Kâr", money(total_profit))
k4.metric("Kâr Marjı", f"%{profit_margin:.1f}")
k5.metric("Ortalama Sepet", money(avg_basket))

# =====================================================
# Daily Tasks
# =====================================================
st.markdown('<div class="section-title">✅ Bugünkü 5 Görev</div>', unsafe_allow_html=True)
cols = st.columns(5)
for i, task in enumerate(daily_tasks, start=1):
    with cols[i-1]:
        st.markdown(f"""
        <div class="task-card">
            <b>{i}. Görev</b><br>{task}
        </div>
        """, unsafe_allow_html=True)

# =====================================================
# Tabs
# =====================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🧠 Günlük Öneriler",
    "🛒 Sipariş Önerisi",
    "🎯 Kampanya",
    "🔗 Çapraz Satış",
    "📊 Skor Detayı",
    "📈 Analizler",
    "⚙️ Veri / Kolon"
])

with tab1:
    st.subheader("AYÇA Günlük Yönetici Özeti")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="{score_color_class(general_score)}">
        <b>Genel durum:</b><br>
        AYÇA skorunuz <b>{general_score}/100</b>. {score_label(general_score)} seviyesinde.
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="warning-card">
        <b>Stok riski:</b><br>
        Stok miktarı 2 veya altında olan <b>{len(critical_stock)}</b> ürün var. Bunlar satış kaybı oluşturabilir.
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="danger-card">
        <b>Ölü stok:</b><br>
        Stokta olup satış hareketi görünmeyen <b>{len(dead_stock)}</b> ürün tespit edildi.
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="success-card">
        <b>Kârlılık:</b><br>
        Toplam brüt kâr <b>{money(total_profit)}</b>, kâr marjı <b>%{profit_margin:.1f}</b>.
        </div>
        """, unsafe_allow_html=True)
        if top_category:
            st.markdown(f"""
            <div class="task-card">
            <b>Kategori odağı:</b><br>
            En güçlü kategori <b>{top_category}</b>. Bugün bu kategori için raf görünürlüğü artırılabilir.
            </div>
            """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="task-card">
        <b>Bugünkü aksiyon:</b><br>
        Önce acil sipariş listesini, sonra kampanya önerilerini kontrol edin.
        </div>
        """, unsafe_allow_html=True)

with tab2:
    st.subheader("Sipariş Öneri Motoru")
    st.caption("Kural: Stok düşük + satış hareketi yüksek olan ürünler acil siparişe düşer.")
    filter_status = st.multiselect(
        "Sipariş önceliği filtrele",
        options=list(order_df["AYCA_Siparis_Onceligi"].unique()),
        default=[x for x in ["Acil Sipariş", "Takip Et", "Kontrol Et"] if x in list(order_df["AYCA_Siparis_Onceligi"].unique())]
    )
    filtered_order = order_df[order_df["AYCA_Siparis_Onceligi"].isin(filter_status)] if filter_status else order_df
    st.dataframe(filtered_order, use_container_width=True)

    urgent_count = len(order_df[order_df["AYCA_Siparis_Onceligi"] == "Acil Sipariş"])
    st.info(f"Acil sipariş önerilen ürün sayısı: {urgent_count}")

with tab3:
    st.subheader("Kampanya / Raf Aksiyonu Önerileri")
    st.caption("Kural: Stokta bekleyen ve az/hic satmayan ürünler kasa önü, raf öne alma veya kampanya kontrolüne düşer.")
    if campaign_df.empty:
        st.success("Kampanya önerilecek belirgin ölü stok bulunamadı.")
    else:
        st.dataframe(campaign_df, use_container_width=True)
        fig_campaign = px.bar(
            campaign_df.head(20),
            x=product_col,
            y=stock_col,
            color="AYCA_Kampanya_Tipi",
            title="Kampanya Adayı Ürünler - Stok Miktarı"
        )
        fig_campaign.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_campaign, use_container_width=True)

with tab4:
    st.subheader("Çapraz Satış Önerileri")
    if category_col and not category_summary.empty:
        cross = category_summary.copy()
        cross["AYÇA Çapraz Satış Önerisi"] = cross[category_col].apply(cross_sell_suggestions)
        st.dataframe(cross, use_container_width=True)
        selected_cat = st.selectbox("Kategori seç", cross[category_col].tolist())
        st.markdown(f"""
        <div class="task-card">
        <b>{selected_cat} için öneri:</b><br>
        {cross_sell_suggestions(selected_cat)}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.warning("Çapraz satış önerileri için Excel'de kategori/grup kolonu gerekir.")

with tab5:
    st.subheader("Eczane Sağlık Skoru Detayı")
    score_df = pd.DataFrame({
        "Skor Alanı": ["Stok Skoru", "Kârlılık Skoru", "Satış Skoru", "Risk Skoru", "Genel Skor"],
        "Puan": [stock_score, profit_score, sales_score, risk_score, general_score]
    })
    fig_score = px.bar(score_df, x="Skor Alanı", y="Puan", text="Puan", title="AYÇA Skor Kırılımı", range_y=[0, 100])
    st.plotly_chart(fig_score, use_container_width=True)
    st.dataframe(score_df, use_container_width=True)

with tab6:
    st.subheader("Satış ve Kârlılık Analizleri")
    top_products = product_summary.sort_values(qty_col, ascending=False).head(20)
    fig_top = px.bar(top_products, x=product_col, y=qty_col, text=qty_col, title="En Çok Satan İlk 20 Ürün")
    fig_top.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_top, use_container_width=True)

    profit_products = product_summary.sort_values(profit_col, ascending=False).head(20)
    fig_profit = px.bar(profit_products, x=product_col, y=profit_col, text=profit_col, title="En Çok Kâr Bırakan Ürünler")
    fig_profit.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_profit, use_container_width=True)

    if category_col and not category_summary.empty:
        fig_cat = px.pie(category_summary.head(10), names=category_col, values=sales_col, title="Kategori Ciro Dağılımı")
        st.plotly_chart(fig_cat, use_container_width=True)

with tab7:
    st.subheader("Yüklenen Veri")
    st.dataframe(df, use_container_width=True)
    st.subheader("Otomatik Kolon Eşleşmeleri")
    mapping = pd.DataFrame({
        "AYÇA Alanı": ["Ürün Adı", "Barkod", "Kategori", "Satış Miktarı", "Ciro", "Maliyet", "Kâr", "Stok", "Tarih"],
        "Excel Kolonu": [product_col, barcode_col, category_col, qty_col, sales_col, cost_col, profit_col, stock_col, date_col]
    })
    st.dataframe(mapping, use_container_width=True)

st.divider()
st.caption("AYÇA Insight v2.1 | Veriler kalıcı olarak saklanmaz. Excel dosyası anlık analiz edilir.")
