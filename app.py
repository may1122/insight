import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(
    page_title="AYÇA Insight v2.0",
    page_icon="💊",
    layout="wide"
)

# -----------------------------
# CSS
# -----------------------------
st.markdown("""
<style>
.main {
    background-color: #f7f9fc;
}
.metric-card {
    background: linear-gradient(135deg, #ffffff, #eef4ff);
    padding: 22px;
    border-radius: 18px;
    box-shadow: 0 4px 18px rgba(0,0,0,0.06);
    border: 1px solid #e6ecf5;
}
.big-title {
    font-size: 34px;
    font-weight: 800;
    color: #16324f;
}
.sub-title {
    font-size: 17px;
    color: #5d6d7e;
}
.insight-card {
    background-color: #ffffff;
    border-left: 6px solid #2f80ed;
    padding: 18px;
    border-radius: 14px;
    margin-bottom: 12px;
    box-shadow: 0 3px 14px rgba(0,0,0,0.05);
}
.warning-card {
    background-color: #fff7ec;
    border-left: 6px solid #f2994a;
    padding: 18px;
    border-radius: 14px;
    margin-bottom: 12px;
}
.danger-card {
    background-color: #fff1f1;
    border-left: 6px solid #eb5757;
    padding: 18px;
    border-radius: 14px;
    margin-bottom: 12px;
}
.success-card {
    background-color: #f0fff5;
    border-left: 6px solid #27ae60;
    padding: 18px;
    border-radius: 14px;
    margin-bottom: 12px;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Helpers
# -----------------------------
def normalize_col(col):
    return str(col).strip().lower().replace("ı", "i").replace("ğ", "g").replace("ü", "u").replace("ş", "s").replace("ö", "o").replace("ç", "c")


def find_column(df, keywords):
    normalized = {normalize_col(c): c for c in df.columns}
    for key in keywords:
        key = normalize_col(key)
        for n_col, original in normalized.items():
            if key in n_col:
                return original
    return None


def safe_sum(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).sum()


def safe_mean(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).mean()


def ayca_score(total_sales, dead_stock_count, critical_stock_count, profit_margin):
    score = 100

    if total_sales <= 0:
        score -= 35

    score -= min(dead_stock_count * 1.2, 25)
    score -= min(critical_stock_count * 1.5, 20)

    if profit_margin < 10:
        score -= 20
    elif profit_margin < 20:
        score -= 10

    return max(0, min(100, int(score)))


def score_label(score):
    if score >= 80:
        return "Çok İyi"
    elif score >= 60:
        return "Geliştirilebilir"
    else:
        return "Dikkat Gerekli"


# -----------------------------
# Header
# -----------------------------
st.markdown('<div class="big-title">💊 AYÇA Insight v2.0</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Eczanenizin Excel raporlarını okuyup karar destek ekranına dönüştürür.</div>', unsafe_allow_html=True)
st.divider()

uploaded_file = st.file_uploader("Tebeos / eczane satış-stok Excel dosyanızı yükleyin", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Excel dosyasını yükleyince analiz ekranı otomatik oluşacaktır.")
    st.stop()

# -----------------------------
# Read Excel
# -----------------------------
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

if df.empty:
    st.warning("Excel dosyası boş görünüyor.")
    st.stop()

# -----------------------------
# Auto column detection
# -----------------------------
product_col = find_column(df, ["urun", "ürün", "mal", "stok adi", "stok adı", "ilac", "ilaç"])
barcode_col = find_column(df, ["barkod", "barcode"])
category_col = find_column(df, ["kategori", "grup", "ana grup"])
qty_col = find_column(df, ["miktar", "adet", "satis miktari", "satış miktarı"])
sales_col = find_column(df, ["ciro", "satis tutari", "satış tutarı", "tutar", "net satis", "net satış"])
cost_col = find_column(df, ["maliyet", "alis", "alış"])
profit_col = find_column(df, ["kar", "kâr", "kazanc", "kazanç"])
stock_col = find_column(df, ["stok", "kalan", "mevcut"])
date_col = find_column(df, ["tarih", "date"])

# Fallbacks
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

df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
df[sales_col] = pd.to_numeric(df[sales_col], errors="coerce").fillna(0)
df[stock_col] = pd.to_numeric(df[stock_col], errors="coerce").fillna(0)

if profit_col:
    df[profit_col] = pd.to_numeric(df[profit_col], errors="coerce").fillna(0)
else:
    df["_AYCA_KAR"] = df[sales_col] * 0.18
    profit_col = "_AYCA_KAR"

total_sales = safe_sum(df[sales_col])
total_qty = safe_sum(df[qty_col])
total_profit = safe_sum(df[profit_col])
avg_basket = total_sales / total_qty if total_qty > 0 else 0
profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

critical_stock = df[df[stock_col] <= 2]
dead_stock = df[(df[stock_col] > 0) & (df[qty_col] == 0)]

score = ayca_score(total_sales, len(dead_stock), len(critical_stock), profit_margin)

# -----------------------------
# Dashboard Cards
# -----------------------------
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric("Toplam Ciro", f"{total_sales:,.0f} ₺")

with col2:
    st.metric("Satış Adedi", f"{total_qty:,.0f}")

with col3:
    st.metric("Brüt Kâr", f"{total_profit:,.0f} ₺")

with col4:
    st.metric("Kâr Marjı", f"%{profit_margin:.1f}")

with col5:
    st.metric("AYÇA Skoru", f"{score}/100", score_label(score))

st.divider()

# -----------------------------
# AYÇA Insight Cards
# -----------------------------
st.subheader("🧠 AYÇA Günlük Yorumları")

c1, c2 = st.columns(2)

with c1:
    if score >= 80:
        st.markdown(f"""
        <div class="success-card">
        <b>Eczane sağlığı güçlü görünüyor.</b><br>
        AYÇA skoru {score}/100. Stok ve satış dengeniz genel olarak iyi.
        </div>
        """, unsafe_allow_html=True)
    elif score >= 60:
        st.markdown(f"""
        <div class="warning-card">
        <b>Geliştirme fırsatı var.</b><br>
        AYÇA skoru {score}/100. Özellikle stok ve kârlılık tarafında iyileştirme yapılabilir.
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="danger-card">
        <b>Dikkat gerekli.</b><br>
        AYÇA skoru {score}/100. Ölü stok, kritik stok veya düşük kârlılık eczaneyi zorluyor olabilir.
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="insight-card">
    <b>Kritik stok uyarısı</b><br>
    Stok miktarı 2 veya altında olan <b>{len(critical_stock)}</b> ürün var.
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div class="insight-card">
    <b>Ortalama sepet değeri</b><br>
    Ortalama satış değeri yaklaşık <b>{avg_basket:,.2f} ₺</b>.
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="warning-card">
    <b>Ölü stok kontrolü</b><br>
    Stokta olup satış hareketi görünmeyen <b>{len(dead_stock)}</b> ürün tespit edildi.
    </div>
    """, unsafe_allow_html=True)

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Satış Analizi",
    "📦 Stok Analizi",
    "💰 Kârlılık",
    "🧾 Veri Önizleme",
    "⚙️ Kolon Eşleşmeleri"
])

with tab1:
    st.subheader("En Çok Satan Ürünler")

    top_products = (
        df.groupby(product_col, as_index=False)
        .agg({
            qty_col: "sum",
            sales_col: "sum",
            profit_col: "sum"
        })
        .sort_values(by=qty_col, ascending=False)
        .head(20)
    )

    fig = px.bar(
        top_products,
        x=product_col,
        y=qty_col,
        title="En Çok Satan İlk 20 Ürün",
        text=qty_col
    )
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(top_products, use_container_width=True)

    if category_col:
        st.subheader("Kategori Bazlı Satış")
        cat_sales = (
            df.groupby(category_col, as_index=False)
            .agg({
                sales_col: "sum",
                qty_col: "sum"
            })
            .sort_values(by=sales_col, ascending=False)
        )

        fig2 = px.pie(
            cat_sales.head(10),
            names=category_col,
            values=sales_col,
            title="Kategori Ciro Dağılımı"
        )
        st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(cat_sales, use_container_width=True)

with tab2:
    st.subheader("Kritik Stok Listesi")

    critical_view = critical_stock[[product_col, stock_col]].copy()
    if barcode_col:
        critical_view.insert(0, barcode_col, critical_stock[barcode_col])

    st.dataframe(
        critical_view.sort_values(by=stock_col, ascending=True),
        use_container_width=True
    )

    st.subheader("Ölü Stok / Hareket Görmeyen Ürünler")

    dead_view = dead_stock[[product_col, stock_col, qty_col]].copy()
    if barcode_col:
        dead_view.insert(0, barcode_col, dead_stock[barcode_col])

    st.dataframe(dead_view, use_container_width=True)

    st.subheader("Stok Dağılımı")

    stock_summary = pd.DataFrame({
        "Durum": ["Kritik Stok", "Ölü Stok", "Diğer"],
        "Adet": [
            len(critical_stock),
            len(dead_stock),
            max(len(df) - len(critical_stock) - len(dead_stock), 0)
        ]
    })

    fig3 = px.bar(stock_summary, x="Durum", y="Adet", title="Stok Sağlığı")
    st.plotly_chart(fig3, use_container_width=True)

with tab3:
    st.subheader("En Kârlı Ürünler")

    profit_products = (
        df.groupby(product_col, as_index=False)
        .agg({
            sales_col: "sum",
            profit_col: "sum",
            qty_col: "sum"
        })
        .sort_values(by=profit_col, ascending=False)
        .head(20)
    )

    profit_products["Kar Marjı %"] = profit_products.apply(
        lambda row: (row[profit_col] / row[sales_col] * 100) if row[sales_col] > 0 else 0,
        axis=1
    )

    fig4 = px.bar(
        profit_products,
        x=product_col,
        y=profit_col,
        title="En Çok Kâr Bırakan Ürünler",
        text=profit_col
    )
    fig4.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig4, use_container_width=True)

    st.dataframe(profit_products, use_container_width=True)

    st.subheader("Çok Satan Ama Düşük Kârlı Ürünler")

    low_profit = profit_products[
        profit_products["Kar Marjı %"] < 15
    ].sort_values(by=qty_col, ascending=False)

    st.dataframe(low_profit, use_container_width=True)

with tab4:
    st.subheader("Yüklenen Excel Verisi")
    st.dataframe(df, use_container_width=True)

with tab5:
    st.subheader("AYÇA Otomatik Kolon Eşleştirme")

    mapping = pd.DataFrame({
        "AYÇA Alanı": [
            "Ürün Adı",
            "Barkod",
            "Kategori",
            "Satış Miktarı",
            "Ciro",
            "Maliyet",
            "Kâr",
            "Stok",
            "Tarih"
        ],
        "Excel Kolonu": [
            product_col,
            barcode_col,
            category_col,
            qty_col,
            sales_col,
            cost_col,
            profit_col,
            stock_col,
            date_col
        ]
    })

    st.dataframe(mapping, use_container_width=True)

    st.info("Kolon isimleri Tebeos raporuna göre değişirse bu alanlardan hangi kolonların algılandığını kontrol edebilirsiniz.")

# -----------------------------
# Footer
# -----------------------------
st.divider()
st.caption("AYÇA Insight v2.0 | Veriler kalıcı olarak saklanmaz. Yüklenen Excel sadece anlık analiz edilir.")
