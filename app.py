import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# =========================================================
# AYÇA Insight - TEBEOS Rapor Okuyucu Gelişmiş MVP
# Versiyon: V1.2
#
# Amaç:
# - TEBEOS / eczane otomasyon raporlarını yüklemek
# - Hasta adı, TC, reçete kişisel verisi işlemeden
# - Eczacının ilk bakışta görmesi gereken riskleri ve aksiyonları göstermek
# =========================================================

st.set_page_config(
    page_title="AYÇA Insight | TEBEOS Rapor Okuyucu",
    page_icon="💊",
    layout="wide"
)

# =========================================================
# Stil
# =========================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 0px;
    }
    .sub-title {
        color: #666;
        font-size: 16px;
        margin-bottom: 20px;
    }
    .risk-high {
        background-color: #ffe5e5;
        padding: 14px;
        border-radius: 10px;
        border-left: 6px solid #d62828;
        margin-bottom: 10px;
    }
    .risk-medium {
        background-color: #fff4d6;
        padding: 14px;
        border-radius: 10px;
        border-left: 6px solid #f4a261;
        margin-bottom: 10px;
    }
    .risk-good {
        background-color: #e7f7ec;
        padding: 14px;
        border-radius: 10px;
        border-left: 6px solid #2a9d8f;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="main-title">💊 AYÇA Insight</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">TEBEOS raporlarını yorumlayan akıllı eczane karar destek ekranı</div>',
    unsafe_allow_html=True
)

st.info(
    "Bu sistem hasta adı, TC kimlik numarası veya kişisel reçete verisi kullanmadan; "
    "stok, satış, maliyet, miad ve kârlılık verilerini yorumlar."
)

# =========================================================
# Yardımcı Fonksiyonlar
# =========================================================

def money(value):
    try:
        return f"{float(value):,.0f} TL".replace(",", ".")
    except Exception:
        return "0 TL"

def percent(value):
    try:
        return f"%{float(value) * 100:.1f}"
    except Exception:
        return "%0.0"

def safe_div(a, b):
    return a / b if b not in [0, None] and not pd.isna(b) else 0

def normalize_columns(df):
    """Kolon adlarında baş/son boşlukları temizler."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data
def load_excel(file):
    sales = pd.read_excel(file, sheet_name="Satis_Recete")
    stock = pd.read_excel(file, sheet_name="Stok")
    purchase = pd.read_excel(file, sheet_name="Alis")

    sales = normalize_columns(sales)
    stock = normalize_columns(stock)
    purchase = normalize_columns(purchase)

    sales["Tarih"] = pd.to_datetime(sales["Tarih"], errors="coerce")
    purchase["Tarih"] = pd.to_datetime(purchase["Tarih"], errors="coerce")
    stock["Miad Tarihi"] = pd.to_datetime(stock["Miad Tarihi"], errors="coerce")

    return sales, stock, purchase

def validate_columns(sales, stock, purchase):
    required_sales = [
        "Tarih", "Barkod", "Ürün Adı", "Kategori", "Adet",
        "Ciro TL", "Maliyet TL", "Brüt Kar TL"
    ]

    required_stock = [
        "Barkod", "Ürün Adı", "Kategori", "Mevcut Stok",
        "Minimum Stok", "Stok Değeri TL", "Miad Tarihi",
        "Miad Kalan Gün"
    ]

    required_purchase = [
        "Tarih", "Barkod", "Ürün Adı", "Adet", "Alış Tutarı TL"
    ]

    missing = {
        "Satış/Reçete": [c for c in required_sales if c not in sales.columns],
        "Stok": [c for c in required_stock if c not in stock.columns],
        "Alış": [c for c in required_purchase if c not in purchase.columns]
    }

    return missing

def build_order_recommendation(sales, stock, days=60):
    last_day = sales["Tarih"].max()
    last_period = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=days)]

    demand = (
        last_period
        .groupby(["Barkod", "Ürün Adı"], as_index=False)
        .agg({"Adet": "sum"})
        .rename(columns={"Adet": f"Son {days} Gün Çıkış"})
    )

    demand["Aylık Ortalama Çıkış"] = demand[f"Son {days} Gün Çıkış"] / (days / 30)

    order = stock.merge(
        demand[["Barkod", f"Son {days} Gün Çıkış", "Aylık Ortalama Çıkış"]],
        on="Barkod",
        how="left"
    )

    order[[f"Son {days} Gün Çıkış", "Aylık Ortalama Çıkış"]] = (
        order[[f"Son {days} Gün Çıkış", "Aylık Ortalama Çıkış"]].fillna(0)
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
        if row["Miad Kalan Gün"] <= 180 and row["Mevcut Stok"] > 0:
            return "Miad takip"
        if row["Stok Ay Karşılığı"] > 4 and row["Mevcut Stok"] > 20:
            return "Fazla stok"
        if row["Aylık Ortalama Çıkış"] == 0 and row["Mevcut Stok"] > 0:
            return "Hareketsiz stok"
        return "Normal"

    order["Öneri"] = order.apply(recommendation, axis=1)

    priority = {
        "Acil sipariş": 0,
        "Sipariş önerilir": 1,
        "Miad takip": 2,
        "Fazla stok": 3,
        "Hareketsiz stok": 4,
        "Normal": 5
    }

    order["Sıra"] = order["Öneri"].map(priority).fillna(9)

    return order.sort_values(["Sıra", "Stok Ay Karşılığı"])

def build_dead_stock(sales, stock, days=180):
    last_day = sales["Tarih"].max()
    recent_sales = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=days)]
    sold_barcodes = set(recent_sales["Barkod"].dropna().unique())

    dead = stock[~stock["Barkod"].isin(sold_barcodes)].copy()
    return dead.sort_values("Stok Değeri TL", ascending=False)

def build_slow_moving_stock(sales, stock, days=90):
    last_day = sales["Tarih"].max()
    period_sales = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=days)]

    demand = (
        period_sales
        .groupby("Barkod", as_index=False)
        .agg({"Adet": "sum"})
        .rename(columns={"Adet": f"Son {days} Gün Satış"})
    )

    merged = stock.merge(demand, on="Barkod", how="left")
    merged[f"Son {days} Gün Satış"] = merged[f"Son {days} Gün Satış"].fillna(0)

    slow = merged[
        (merged["Mevcut Stok"] > 0) &
        (merged[f"Son {days} Gün Satış"] <= 2)
    ].copy()

    return slow.sort_values("Stok Değeri TL", ascending=False)

def build_trend_analysis(sales):
    last_day = sales["Tarih"].max()

    recent_30 = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=30)]
    prev_30 = sales[
        (sales["Tarih"] < last_day - pd.Timedelta(days=30)) &
        (sales["Tarih"] >= last_day - pd.Timedelta(days=60))
    ]

    r = (
        recent_30
        .groupby(["Barkod", "Ürün Adı", "Kategori"], as_index=False)
        .agg({"Adet": "sum", "Ciro TL": "sum"})
        .rename(columns={"Adet": "Son 30 Gün Adet", "Ciro TL": "Son 30 Gün Ciro"})
    )

    p = (
        prev_30
        .groupby(["Barkod"], as_index=False)
        .agg({"Adet": "sum", "Ciro TL": "sum"})
        .rename(columns={"Adet": "Önceki 30 Gün Adet", "Ciro TL": "Önceki 30 Gün Ciro"})
    )

    trend = r.merge(p, on="Barkod", how="left")
    trend[["Önceki 30 Gün Adet", "Önceki 30 Gün Ciro"]] = (
        trend[["Önceki 30 Gün Adet", "Önceki 30 Gün Ciro"]].fillna(0)
    )

    trend["Adet Değişim %"] = trend.apply(
        lambda row: 1 if row["Önceki 30 Gün Adet"] == 0 and row["Son 30 Gün Adet"] > 0
        else safe_div(row["Son 30 Gün Adet"] - row["Önceki 30 Gün Adet"], row["Önceki 30 Gün Adet"]),
        axis=1
    )

    rising = trend[trend["Adet Değişim %"] > 0.3].sort_values("Adet Değişim %", ascending=False)
    falling = trend[trend["Adet Değişim %"] < -0.3].sort_values("Adet Değişim %", ascending=True)

    return rising, falling, trend

def calculate_risk_score(near_expiry_value, dead_stock_value, urgent_count, stock_value):
    score = 0

    expiry_ratio = safe_div(near_expiry_value, stock_value)
    dead_ratio = safe_div(dead_stock_value, stock_value)

    if expiry_ratio > 0.15:
        score += 30
    elif expiry_ratio > 0.07:
        score += 18
    elif expiry_ratio > 0.03:
        score += 8

    if dead_ratio > 0.20:
        score += 35
    elif dead_ratio > 0.10:
        score += 22
    elif dead_ratio > 0.05:
        score += 10

    if urgent_count >= 20:
        score += 30
    elif urgent_count >= 10:
        score += 20
    elif urgent_count >= 5:
        score += 10

    return min(score, 100)

def risk_label(score):
    if score >= 70:
        return "Yüksek Risk"
    if score >= 40:
        return "Orta Risk"
    return "Kontrollü"

def build_action_list(order, near_expiry, dead_stock, slow_stock):
    actions = []

    urgent = order[order["Öneri"] == "Acil sipariş"].head(10)
    for _, row in urgent.iterrows():
        actions.append({
            "Öncelik": "Yüksek",
            "Aksiyon": "Acil sipariş ver",
            "Ürün": row["Ürün Adı"],
            "Sebep": f"Mevcut stok {row['Mevcut Stok']}, minimum stok {row['Minimum Stok']}",
            "Tahmini Etki": "Satış kaybını önler"
        })

    exp = near_expiry.sort_values("Stok Değeri TL", ascending=False).head(10)
    for _, row in exp.iterrows():
        actions.append({
            "Öncelik": "Yüksek",
            "Aksiyon": "Miad aksiyonu al",
            "Ürün": row["Ürün Adı"],
            "Sebep": f"{int(row['Miad Kalan Gün'])} gün kaldı, stok değeri {money(row['Stok Değeri TL'])}",
            "Tahmini Etki": "Fire riskini azaltır"
        })

    dead = dead_stock.head(10)
    for _, row in dead.iterrows():
        actions.append({
            "Öncelik": "Orta",
            "Aksiyon": "Ölü stok kontrolü",
            "Ürün": row["Ürün Adı"],
            "Sebep": f"Son 180 günde hareket yok, stok değeri {money(row['Stok Değeri TL'])}",
            "Tahmini Etki": "Nakit bağlanmasını azaltır"
        })

    slow = slow_stock.head(10)
    for _, row in slow.iterrows():
        actions.append({
            "Öncelik": "Orta",
            "Aksiyon": "Yavaş ürün takibi",
            "Ürün": row["Ürün Adı"],
            "Sebep": "Satış hızı düşük, stok var",
            "Tahmini Etki": "Fazla stok riskini azaltır"
        })

    if not actions:
        return pd.DataFrame(columns=["Öncelik", "Aksiyon", "Ürün", "Sebep", "Tahmini Etki"])

    return pd.DataFrame(actions)

def ayca_commentary(total_revenue, gross_profit, margin, stock_value,
                    near_expiry_value, dead_stock_value, urgent_count,
                    top_category, weak_category, risk_score_value):
    comments = []

    comments.append(
        f"Bu dönem toplam ciro {money(total_revenue)}, brüt kâr {money(gross_profit)} "
        f"ve brüt kâr marjı {percent(margin)} seviyesinde."
    )

    if margin < 0.18:
        comments.append(
            "Kâr marjı düşük görünüyor. Alış fiyatları, iskonto oranları ve düşük kârlı ürün satışları incelenmeli."
        )
    elif margin > 0.30:
        comments.append(
            "Kâr marjı güçlü görünüyor. Bu yapıyı korumak için yüksek kârlı kategoriler ayrıca takip edilmeli."
        )

    if near_expiry_value > 0:
        comments.append(
            f"Miad riski taşıyan ürünlerde yaklaşık {money(near_expiry_value)} sermaye bağlı. "
            "Bu ürünler için kampanya, iade veya satış önceliği planlanmalı."
        )

    if dead_stock_value > 0:
        comments.append(
            f"Son 180 günde hareket görmeyen stok değeri yaklaşık {money(dead_stock_value)}. "
            "Bu tutar eczanenin nakit akışını yavaşlatabilir."
        )

    if urgent_count > 0:
        comments.append(
            f"{urgent_count} ürün için acil sipariş uyarısı oluştu. "
            "Bu ürünler satış kaybı yaşamamak için öncelikli kontrol edilmeli."
        )

    if top_category:
        comments.append(
            f"En güçlü kategori '{top_category}' görünüyor. Bu kategori büyütülebilir veya kampanya ile desteklenebilir."
        )

    if weak_category:
        comments.append(
            f"'{weak_category}' kategorisinde performans zayıf görünüyor. Raf, stok ve fiyat stratejisi gözden geçirilebilir."
        )

    comments.append(
        f"Genel risk skoru {risk_score_value}/100 - {risk_label(risk_score_value)}."
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

missing = validate_columns(sales, stock, purchase)
missing_total = sum(len(v) for v in missing.values())

if missing_total > 0:
    st.error("Excel kolon yapısı beklenen formatta değil.")
    st.write(missing)
    st.stop()

# =========================================================
# Sidebar Filtreler
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

analysis_days = st.sidebar.slider(
    "Sipariş analizi gün aralığı",
    min_value=30,
    max_value=120,
    value=60,
    step=30
)

dead_stock_days = st.sidebar.slider(
    "Ölü stok gün sınırı",
    min_value=90,
    max_value=365,
    value=180,
    step=30
)

# =========================================================
# Ana Hesaplamalar
# =========================================================

total_revenue = sales_f["Ciro TL"].sum()
total_cost = sales_f["Maliyet TL"].sum()
gross_profit = sales_f["Brüt Kar TL"].sum()
gross_margin = safe_div(gross_profit, total_revenue)

stock_value = stock["Stok Değeri TL"].sum()

near_expiry = stock[(stock["Miad Kalan Gün"] <= 180) & (stock["Mevcut Stok"] > 0)].copy()
near_expiry_value = near_expiry["Stok Değeri TL"].sum()

order = build_order_recommendation(sales, stock, days=analysis_days)
urgent_order = order[order["Öneri"] == "Acil sipariş"]
urgent_count = len(urgent_order)

dead_stock = build_dead_stock(sales, stock, days=dead_stock_days)
dead_stock_value = dead_stock["Stok Değeri TL"].sum()

slow_stock = build_slow_moving_stock(sales, stock, days=90)

rising_products, falling_products, trend_products = build_trend_analysis(sales)

category_perf = (
    sales_f
    .groupby("Kategori", as_index=False)
    .agg({
        "Ciro TL": "sum",
        "Brüt Kar TL": "sum",
        "Adet": "sum"
    })
)

category_perf["Kar %"] = category_perf.apply(
    lambda row: safe_div(row["Brüt Kar TL"], row["Ciro TL"]),
    axis=1
)

top_category = None
weak_category = None

if not category_perf.empty:
    top_category = category_perf.sort_values("Brüt Kar TL", ascending=False).iloc[0]["Kategori"]
    weak_category = category_perf.sort_values("Kar %", ascending=True).iloc[0]["Kategori"]

risk_score_value = calculate_risk_score(
    near_expiry_value=near_expiry_value,
    dead_stock_value=dead_stock_value,
    urgent_count=urgent_count,
    stock_value=stock_value
)

action_list = build_action_list(order, near_expiry, dead_stock, slow_stock)

# =========================================================
# Açılış: İlk 5 Kart
# =========================================================

st.subheader("📌 Eczacının Bilgisayarı Açar Açmaz Görmesi Gereken İlk 5 Şey")

k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("1. Dönem Cirosu", money(total_revenue))
k2.metric("2. Brüt Kâr", money(gross_profit), percent(gross_margin))
k3.metric("3. Miad Riski", money(near_expiry_value), f"{len(near_expiry)} ürün")
k4.metric("4. Ölü Stok", money(dead_stock_value), f"{len(dead_stock)} ürün")
k5.metric("5. Acil Sipariş", f"{urgent_count} ürün")

# =========================================================
# AYÇA Radar
# =========================================================

st.subheader("🚨 AYÇA Radar")

r1, r2, r3 = st.columns([1, 2, 2])

with r1:
    st.metric("Genel Risk Skoru", f"{risk_score_value}/100", risk_label(risk_score_value))

with r2:
    if risk_score_value >= 70:
        st.markdown(
            f"<div class='risk-high'><b>Yüksek risk:</b> Miad, ölü stok veya sipariş tarafında hızlı aksiyon gerekiyor.</div>",
            unsafe_allow_html=True
        )
    elif risk_score_value >= 40:
        st.markdown(
            f"<div class='risk-medium'><b>Orta risk:</b> Eczane genel olarak yönetilebilir durumda; ancak bazı ürün grupları takip edilmeli.</div>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f"<div class='risk-good'><b>Kontrollü:</b> Genel tablo dengeli görünüyor. Riskli ürünler düzenli takip edilmeli.</div>",
            unsafe_allow_html=True
        )

with r3:
    st.write("**Bugünkü öncelik sırası:**")
    st.write("1. Acil sipariş ürünlerini kontrol et")
    st.write("2. Miadı yaklaşan yüksek değerli ürünleri incele")
    st.write("3. Ölü stoktaki sermaye bağlayan ürünleri ayır")

# =========================================================
# AYÇA Yorumu
# =========================================================

st.subheader("🧠 AYÇA Yönetici Yorumu")

comments = ayca_commentary(
    total_revenue=total_revenue,
    gross_profit=gross_profit,
    margin=gross_margin,
    stock_value=stock_value,
    near_expiry_value=near_expiry_value,
    dead_stock_value=dead_stock_value,
    urgent_count=urgent_count,
    top_category=top_category,
    weak_category=weak_category,
    risk_score_value=risk_score_value
)

for comment in comments:
    st.write("• " + comment)

# =========================================================
# Aksiyon Listesi
# =========================================================

st.subheader("✅ AYÇA Aksiyon Listesi")

if action_list.empty:
    st.success("Bugün için kritik aksiyon görünmüyor.")
else:
    st.dataframe(action_list.head(25), use_container_width=True)

# =========================================================
# Sekmeler
# =========================================================

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "Genel Bakış",
    "Karlılık",
    "Miad Riski",
    "Ölü Stok",
    "Sipariş Radarı",
    "Trendler",
    "Kategori Analizi",
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

    st.subheader("Stok Değerinin Risk Dağılımı")

    risk_summary = pd.DataFrame({
        "Risk Tipi": ["Miad Riski", "Ölü Stok", "Kontrollü Stok"],
        "Tutar": [
            near_expiry_value,
            dead_stock_value,
            max(stock_value - near_expiry_value - dead_stock_value, 0)
        ]
    })

    fig2 = px.pie(
        risk_summary,
        names="Risk Tipi",
        values="Tutar",
        title="Stok Değeri Risk Dağılımı"
    )

    st.plotly_chart(fig2, use_container_width=True)

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

    product_profit["Kar %"] = product_profit.apply(
        lambda row: safe_div(row["Brüt Kar TL"], row["Ciro TL"]),
        axis=1
    )

    st.dataframe(
        product_profit.sort_values("Brüt Kar TL", ascending=False),
        use_container_width=True
    )

    top_profit = product_profit.sort_values("Brüt Kar TL", ascending=False).head(10)

    fig = px.bar(
        top_profit,
        x="Ürün Adı",
        y="Brüt Kar TL",
        color="Kategori",
        title="En Çok Brüt Kâr Bırakan 10 Ürün"
    )

    st.plotly_chart(fig, use_container_width=True)

    low_margin = product_profit[
        (product_profit["Ciro TL"] > 0) &
        (product_profit["Kar %"] < 0.10)
    ].sort_values("Kar %")

    st.subheader("Düşük Kâr Marjlı Ürünler")
    st.dataframe(low_margin, use_container_width=True)

with tab3:
    st.subheader("Miad Riski")

    near_expiry_show = near_expiry.sort_values(["Miad Kalan Gün", "Stok Değeri TL"])

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

    if not near_expiry_show.empty:
        fig = px.bar(
            near_expiry_show.head(15),
            x="Ürün Adı",
            y="Stok Değeri TL",
            color="Miad Kalan Gün",
            title="Miad Riski En Yüksek Ürünler"
        )
        st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.subheader("Ölü Stok Analizi")

    st.metric(f"Son {dead_stock_days} Günde Hareket Görmeyen Stok Değeri", money(dead_stock_value))

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

    st.subheader("Yavaş Hareket Eden Stok")
    st.dataframe(
        slow_stock[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Mevcut Stok",
                "Stok Değeri TL",
                "Son 90 Gün Satış",
                "Miad Kalan Gün"
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
                f"Son {analysis_days} Gün Çıkış",
                "Aylık Ortalama Çıkış",
                "Stok Ay Karşılığı",
                "Miad Kalan Gün",
                "Öneri"
            ]
        ],
        use_container_width=True
    )

    order_summary = order["Öneri"].value_counts().reset_index()
    order_summary.columns = ["Öneri", "Ürün Sayısı"]

    fig = px.bar(
        order_summary,
        x="Öneri",
        y="Ürün Sayısı",
        title="Sipariş Önerisi Dağılımı"
    )

    st.plotly_chart(fig, use_container_width=True)

with tab6:
    st.subheader("Satış Trendleri")

    st.write("Son 30 gün ve önceki 30 gün karşılaştırması.")

    st.subheader("Yükselen Ürünler")
    st.dataframe(
        rising_products[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Son 30 Gün Adet",
                "Önceki 30 Gün Adet",
                "Adet Değişim %",
                "Son 30 Gün Ciro"
            ]
        ].head(25),
        use_container_width=True
    )

    st.subheader("Düşen Ürünler")
    st.dataframe(
        falling_products[
            [
                "Barkod",
                "Ürün Adı",
                "Kategori",
                "Son 30 Gün Adet",
                "Önceki 30 Gün Adet",
                "Adet Değişim %",
                "Son 30 Gün Ciro"
            ]
        ].head(25),
        use_container_width=True
    )

with tab7:
    st.subheader("Kategori Analizi")

    st.dataframe(
        category_perf.sort_values("Ciro TL", ascending=False),
        use_container_width=True
    )

    if not category_perf.empty:
        fig = px.bar(
            category_perf.sort_values("Ciro TL", ascending=False),
            x="Kategori",
            y=["Ciro TL", "Brüt Kar TL"],
            title="Kategori Bazlı Ciro ve Brüt Kâr"
        )
        st.plotly_chart(fig, use_container_width=True)

        fig2 = px.bar(
            category_perf.sort_values("Kar %", ascending=False),
            x="Kategori",
            y="Kar %",
            title="Kategori Bazlı Kâr Marjı"
        )
        st.plotly_chart(fig2, use_container_width=True)

with tab8:
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

# =========================================================
# Footer
# =========================================================

st.caption(
    "AYÇA Insight V1.2 - Bu demo eczanenin ticari verilerini yorumlar; hasta kişisel verisi işlemez."
)
