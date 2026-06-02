import streamlit as st
import pandas as pd
import plotly.express as px

# =========================================================
# AYÇA Insight - TEBEOS Rapor Okuyucu MVP
# Amaç:
# - Hasta bilgisi / TC / reçete kişisel verisi tutmadan
# - TEBEOS veya benzeri eczane otomasyon raporlarını okuyup
# - Eczacının bilgisayarı açar açmaz görmek isteyeceği ilk 5 şeyi göstermek
# =========================================================

st.set_page_config(
    page_title="AYÇA Insight | TEBEOS Rapor Okuyucu",
    page_icon="💊",
    layout="wide"
)

st.title("💊 AYÇA Insight")
st.caption("TEBEOS raporlarını yorumlayan akıllı eczane karar destek ekranı")

st.info(
    "Bu demo hasta adı, TC kimlik numarası, reçete kişisel detayı işlemez. "
    "Sadece eczanenin ticari verileri üzerinden analiz yapar."
)

# =========================================================
# Yardımcı Fonksiyonlar
# =========================================================

def money(value):
    try:
        return f"{value:,.0f} TL".replace(",", ".")
    except Exception:
        return "0 TL"

def percent(value):
    try:
        return f"%{value * 100:.1f}"
    except Exception:
        return "%0.0"

@st.cache_data
def load_excel(file):
    sales = pd.read_excel(file, sheet_name="Satis_Recete")
    stock = pd.read_excel(file, sheet_name="Stok")
    purchase = pd.read_excel(file, sheet_name="Alis")

    sales["Tarih"] = pd.to_datetime(sales["Tarih"], errors="coerce")
    purchase["Tarih"] = pd.to_datetime(purchase["Tarih"], errors="coerce")
    stock["Miad Tarihi"] = pd.to_datetime(stock["Miad Tarihi"], errors="coerce")

    return sales, stock, purchase

def build_order_recommendation(sales, stock):
    last_day = sales["Tarih"].max()
    last_60 = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=60)]

    demand = (
        last_60
        .groupby(["Barkod", "Ürün Adı"], as_index=False)
        .agg({"Adet": "sum"})
        .rename(columns={"Adet": "Son 60 Gün Çıkış"})
    )

    demand["Aylık Ortalama Çıkış"] = demand["Son 60 Gün Çıkış"] / 2

    order = stock.merge(
        demand[["Barkod", "Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]],
        on="Barkod",
        how="left"
    )

    order[["Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]] = (
        order[["Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]].fillna(0)
    )

    order["Stok Ay Karşılığı"] = order.apply(
        lambda r: 99 if r["Aylık Ortalama Çıkış"] == 0
        else r["Mevcut Stok"] / r["Aylık Ortalama Çıkış"],
        axis=1
    )

    def recommendation(row):
        if row["Mevcut Stok"] < row["Minimum Stok"]:
            return "Acil sipariş"
        if row["Stok Ay Karşılığı"] < 0.5:
            return "Acil sipariş"
        if row["Stok Ay Karşılığı"] < 1:
            return "Sipariş önerilir"
        if row["Miad Kalan Gün"] <= 180:
            return "Miad takip"
        if row["Stok Ay Karşılığı"] > 4 and row["Mevcut Stok"] > 20:
            return "Fazla stok"
        return "Normal"

    order["Öneri"] = order.apply(recommendation, axis=1)

    priority = {
        "Acil sipariş": 0,
        "Sipariş önerilir": 1,
        "Miad takip": 2,
        "Fazla stok": 3,
        "Normal": 4
    }

    order["Sıra"] = order["Öneri"].map(priority)
    return order.sort_values(["Sıra", "Stok Ay Karşılığı"])

def build_dead_stock(sales, stock, days=180):
    last_day = sales["Tarih"].max()
    recent_sales = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=days)]

    sold_barcodes = set(recent_sales["Barkod"].dropna().unique())
    dead = stock[~stock["Barkod"].isin(sold_barcodes)].copy()

    return dead.sort_values("Stok Değeri TL", ascending=False)

def ayca_commentary(total_revenue, gross_profit, margin, stock_value,
                    near_expiry_value, dead_stock_value, urgent_count,
                    top_category, weak_category):
    comments = []

    comments.append(
        f"Bu dönem toplam ciro {money(total_revenue)}, brüt kâr {money(gross_profit)} "
        f"ve brüt kâr marjı {percent(margin)} seviyesinde gerçekleşti."
    )

    if near_expiry_value > 0:
        comments.append(
            f"Miad riski taşıyan ürünlerde yaklaşık {money(near_expiry_value)} sermaye bağlı. "
            "Bu ürünler için öncelikli aksiyon alınması önerilir."
        )

    if dead_stock_value > 0:
        comments.append(
            f"Son 180 günde hareket görmeyen stok değeri yaklaşık {money(dead_stock_value)}. "
            "Bu alan eczanenin nakit akışını yavaşlatabilir."
        )

    if urgent_count > 0:
        comments.append(
            f"{urgent_count} ürün için acil sipariş uyarısı oluştu. "
            "Bu ürünler satış kaybı yaşamamak için öncelikli kontrol edilmeli."
        )

    if top_category:
        comments.append(
            f"En güçlü kategori şu anda '{top_category}' görünüyor. "
            "Bu kategori kârlılık ve ciro açısından ayrıca takip edilmeli."
        )

    if weak_category:
        comments.append(
            f"'{weak_category}' kategorisinde performans zayıf görünüyor. "
            "Stok ve satış stratejisi gözden geçirilebilir."
        )

    return comments

# =========================================================
# Dosya Yükleme
# =========================================================

uploaded_file = st.sidebar.file_uploader(
    "TEBEOS / Eczane rapor Excel dosyasını yükleyin",
    type=["xlsx"]
)

if uploaded_file is None:
    st.warning("Analiz için Excel raporunu yükleyin.")
    st.stop()

try:
    sales, stock, purchase = load_excel(uploaded_file)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

# =========================================================
# Filtreler
# =========================================================

st.sidebar.header("Filtreler")

min_date = sales["Tarih"].min().date()
max_date = sales["Tarih"].max().date()

date_range = st.sidebar.date_input(
    "Satış tarih aralığı",
    [min_date, max_date],
    min_value=min_date,
    max_value=max_date
)

if len(date_range) == 2:
    start_date = pd.to_datetime(date_range[0])
    end_date = pd.to_datetime(date_range[1])
    sales_f = sales[(sales["Tarih"] >= start_date) & (sales["Tarih"] <= end_date)]
else:
    sales_f = sales.copy()

categories = sorted(sales["Kategori"].dropna().unique())
selected_categories = st.sidebar.multiselect(
    "Kategori seçin",
    categories,
    default=categories
)

sales_f = sales_f[sales_f["Kategori"].isin(selected_categories)]

# =========================================================
# Ana Hesaplamalar
# =========================================================

total_revenue = sales_f["Ciro TL"].sum()
total_cost = sales_f["Maliyet TL"].sum()
gross_profit = sales_f["Brüt Kar TL"].sum()
gross_margin = gross_profit / total_revenue if total_revenue else 0

stock_value = stock["Stok Değeri TL"].sum()

near_expiry = stock[stock["Miad Kalan Gün"] <= 180].copy()
near_expiry_value = near_expiry["Stok Değeri TL"].sum()

order = build_order_recommendation(sales, stock)
urgent_order = order[order["Öneri"] == "Acil sipariş"]
urgent_count = len(urgent_order)

dead_stock = build_dead_stock(sales, stock, days=180)
dead_stock_value = dead_stock["Stok Değeri TL"].sum()

category_perf = (
    sales_f
    .groupby("Kategori", as_index=False)
    .agg({
        "Ciro TL": "sum",
        "Brüt Kar TL": "sum",
        "Adet": "sum"
    })
)

category_perf["Kar %"] = category_perf["Brüt Kar TL"] / category_perf["Ciro TL"]

top_category = None
weak_category = None

if not category_perf.empty:
    top_category = category_perf.sort_values("Brüt Kar TL", ascending=False).iloc[0]["Kategori"]
    weak_category = category_perf.sort_values("Kar %", ascending=True).iloc[0]["Kategori"]

# =========================================================
# İlk 5 Şey: Eczacının Açılış Ekranı
# =========================================================

st.subheader("📌 Eczacının Bilgisayarı Açar Açmaz Görmesi Gereken İlk 5 Şey")

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric(
    "1. Bugünkü / Dönem Cirosu",
    money(total_revenue)
)

k2.metric(
    "2. Brüt Kâr",
    money(gross_profit),
    percent(gross_margin)
)

k3.metric(
    "3. Miad Riski",
    money(near_expiry_value),
    f"{len(near_expiry)} ürün"
)

k4.metric(
    "4. Ölü Stok",
    money(dead_stock_value),
    f"{len(dead_stock)} ürün"
)

k5.metric(
    "5. Acil Sipariş",
    f"{urgent_count} ürün"
)

# =========================================================
# AYÇA Yönetici Yorumu
# =========================================================

st.subheader("🧠 AYÇA Yorumu")

comments = ayca_commentary(
    total_revenue=total_revenue,
    gross_profit=gross_profit,
    margin=gross_margin,
    stock_value=stock_value,
    near_expiry_value=near_expiry_value,
    dead_stock_value=dead_stock_value,
    urgent_count=urgent_count,
    top_category=top_category,
    weak_category=weak_category
)

for comment in comments:
    st.write("• " + comment)

# =========================================================
# Sekmeler
# =========================================================

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Genel Bakış",
    "Karlılık",
    "Miad Riski",
    "Ölü Stok",
    "Sipariş Radarı",
    "Ham Veri"
])

with tab1:
    st.subheader("Günlük Ciro ve Brüt Kâr")

    daily = (
        sales_f
        .groupby("Tarih", as_index=False)
        .agg({
            "Ciro TL": "sum",
            "Brüt Kar TL": "sum"
        })
    )

    fig = px.line(
        daily,
        x="Tarih",
        y=["Ciro TL", "Brüt Kar TL"],
        markers=True,
        title="Günlük Ciro ve Brüt Kâr Trendi"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Kategori Performansı")
    st.dataframe(
        category_perf.sort_values("Ciro TL", ascending=False),
        use_container_width=True
    )

with tab2:
    st.subheader("Ürün Bazlı Karlılık")

    product_profit = (
        sales_f
        .groupby(["Barkod", "Ürün Adı", "Kategori"], as_index=False)
        .agg({
            "Adet": "sum",
            "Ciro TL": "sum",
            "Maliyet TL": "sum",
            "Brüt Kar TL": "sum"
        })
    )

    product_profit["Kar %"] = product_profit["Brüt Kar TL"] / product_profit["Ciro TL"]

    st.dataframe(
        product_profit.sort_values("Brüt Kar TL", ascending=False),
        use_container_width=True
    )

    top_profit = product_profit.sort_values("Brüt Kar TL", ascending=False).head(10)

    fig = px.bar(
        top_profit,
        x="Ürün Adı",
        y="Brüt Kar TL",
        title="En Çok Brüt Kâr Bırakan 10 Ürün"
    )

    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("Miad Riski")

    near_expiry_show = near_expiry.sort_values("Miad Kalan Gün")

    st.metric("Miad Riski Taşıyan Stok Değeri", money(near_expiry_value))

    st.dataframe(
        near_expiry_show[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Mevcut Stok",
                "Stok Değeri TL",
                "Miad Tarihi",
                "Miad Kalan Gün",
                "Raf Lokasyonu"
            ]
        ],
        use_container_width=True
    )

with tab4:
    st.subheader("Ölü Stok Analizi")

    st.metric("Son 180 Günde Hareket Görmeyen Stok Değeri", money(dead_stock_value))

    st.dataframe(
        dead_stock[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Mevcut Stok",
                "Stok Değeri TL",
                "Miad Tarihi",
                "Miad Kalan Gün",
                "Tedarikçi"
            ]
        ],
        use_container_width=True
    )

with tab5:
    st.subheader("Sipariş Radarı")

    st.dataframe(
        order[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Mevcut Stok",
                "Minimum Stok",
                "Son 60 Gün Çıkış",
                "Aylık Ortalama Çıkış",
                "Stok Ay Karşılığı",
                "Miad Kalan Gün",
                "Öneri"
            ]
        ],
        use_container_width=True
    )

with tab6:
    st.subheader("Ham Veri")

    data_choice = st.radio(
        "Tablo seçin",
        ["Satış/Reçete", "Stok", "Alış"],
        horizontal=True
    )

    if data_choice == "Satış/Reçete":
        st.dataframe(sales, use_container_width=True)
    elif data_choice == "Stok":
        st.dataframe(stock, use_container_width=True)
    else:
        st.dataframe(purchase, use_container_width=True)
