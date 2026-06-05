import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="AYÇA Insight",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AYÇA Insight")
st.caption("Satış ve stok Excel raporlarıyla eczane analiz ekranı")

# -------------------------------------------------
# Yardımcı Fonksiyonlar
# -------------------------------------------------

def read_excel(file):
    return pd.read_excel(file)

def clean_money_columns(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df

def prepare_sales(df):
    df = df.copy()

    if "İşlem Tarihi" in df.columns:
        df["İşlem Tarihi"] = pd.to_datetime(df["İşlem Tarihi"], errors="coerce")
        df["Tarih"] = df["İşlem Tarihi"].dt.date
    elif "Alım Tarih" in df.columns:
        df["Alım Tarih"] = pd.to_datetime(df["Alım Tarih"], errors="coerce")
        df["Tarih"] = df["Alım Tarih"].dt.date

    df = clean_money_columns(df, [
        "Toplam Tutar",
        "Ödenen Tutar",
        "İskonto",
        "Eld. Top. Tut",
        "Kar Tutarı",
        "Maliyet Tutarı",
        "Fiy. Farkı"
    ])

    return df

def prepare_stock(df):
    df = df.copy()

    df = clean_money_columns(df, [
        "Psf",
        "Kamu",
        "Stok",
        "Kritik Stok",
        "Psf Toplam",
        "Kamu Toplam",
        "Mal Top(Kdv Hariç)",
        "Mal Top(Kdv Dahil)"
    ])

    return df

# -------------------------------------------------
# Dosya Yükleme
# -------------------------------------------------

st.sidebar.header("📁 Excel Yükle")

sales_file = st.sidebar.file_uploader(
    "1) Satışlar Excel dosyasını yükle",
    type=["xlsx"],
    key="sales_file"
)

stock_file = st.sidebar.file_uploader(
    "2) Envanter / Stok Excel dosyasını yükle",
    type=["xlsx"],
    key="stock_file"
)

if not sales_file and not stock_file:
    st.info("Lütfen önce satış ve stok Excel dosyalarını yükleyin.")
    st.stop()

sales_df = None
stock_df = None

if sales_file:
    sales_df = prepare_sales(read_excel(sales_file))

if stock_file:
    stock_df = prepare_stock(read_excel(stock_file))

# -------------------------------------------------
# Satış Analizi
# -------------------------------------------------

if sales_df is not None:
    st.subheader("💰 Satış Özeti")

    total_sales = sales_df["Toplam Tutar"].sum() if "Toplam Tutar" in sales_df.columns else 0
    total_paid = sales_df["Ödenen Tutar"].sum() if "Ödenen Tutar" in sales_df.columns else 0
    total_discount = sales_df["İskonto"].sum() if "İskonto" in sales_df.columns else 0
    transaction_count = len(sales_df)

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Toplam Ciro", f"{total_sales:,.2f} TL")
    col2.metric("Ödenen Tutar", f"{total_paid:,.2f} TL")
    col3.metric("İskonto", f"{total_discount:,.2f} TL")
    col4.metric("İşlem Sayısı", f"{transaction_count:,}")

    if "Tarih" in sales_df.columns and "Toplam Tutar" in sales_df.columns:
        daily_sales = sales_df.groupby("Tarih", as_index=False)["Toplam Tutar"].sum()

        st.subheader("📈 Günlük Satış Trendi")
        fig = px.line(
            daily_sales,
            x="Tarih",
            y="Toplam Tutar",
            markers=True,
            title="Günlük Ciro"
        )
        st.plotly_chart(fig, use_container_width=True)

    col_a, col_b = st.columns(2)

    with col_a:
        if "Tahsilat" in sales_df.columns:
            st.subheader("💳 Tahsilat Dağılımı")
            tahsilat = sales_df.groupby("Tahsilat", as_index=False)["Toplam Tutar"].sum()
            fig = px.pie(
                tahsilat,
                names="Tahsilat",
                values="Toplam Tutar",
                title="Tahsilat Tipine Göre Ciro"
            )
            st.plotly_chart(fig, use_container_width=True)

    with col_b:
        if "Kurum Adı" in sales_df.columns:
            st.subheader("🏥 Kurum Dağılımı")
            kurum = (
                sales_df.groupby("Kurum Adı", as_index=False)["Toplam Tutar"]
                .sum()
                .sort_values("Toplam Tutar", ascending=False)
                .head(10)
            )
            fig = px.bar(
                kurum,
                x="Toplam Tutar",
                y="Kurum Adı",
                orientation="h",
                title="İlk 10 Kurum"
            )
            st.plotly_chart(fig, use_container_width=True)

# -------------------------------------------------
# Stok Analizi
# -------------------------------------------------

if stock_df is not None:
    st.subheader("📦 Stok Özeti")

    total_stock_value = stock_df["Psf Toplam"].sum() if "Psf Toplam" in stock_df.columns else 0
    total_cost_value = stock_df["Mal Top(Kdv Dahil)"].sum() if "Mal Top(Kdv Dahil)" in stock_df.columns else 0
    product_count = len(stock_df)
    total_stock_qty = stock_df["Stok"].sum() if "Stok" in stock_df.columns else 0

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Ürün Sayısı", f"{product_count:,}")
    col2.metric("Toplam Stok Adedi", f"{total_stock_qty:,.0f}")
    col3.metric("PSF Stok Değeri", f"{total_stock_value:,.2f} TL")
    col4.metric("Maliyet Değeri", f"{total_cost_value:,.2f} TL")

    if "Stok" in stock_df.columns and "Kritik Stok" in stock_df.columns:
        critical_df = stock_df[
            (stock_df["Kritik Stok"] > 0) &
            (stock_df["Stok"] <= stock_df["Kritik Stok"])
        ]

        st.warning(f"Kritik stok seviyesinde olan ürün sayısı: {len(critical_df)}")

        if len(critical_df) > 0:
            st.dataframe(
                critical_df[["Barkod", "Ürün Adı", "Stok", "Kritik Stok", "Psf"]]
                .sort_values("Stok", ascending=True),
                use_container_width=True
            )

    if "Psf Toplam" in stock_df.columns:
        st.subheader("💎 En Çok Stok Sermayesi Bağlayan Ürünler")

        top_stock = (
            stock_df.sort_values("Psf Toplam", ascending=False)
            .head(20)
        )

        fig = px.bar(
            top_stock,
            x="Psf Toplam",
            y="Ürün Adı",
            orientation="h",
            title="Stok Değeri En Yüksek 20 Ürün"
        )
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            top_stock[["Barkod", "Ürün Adı", "Stok", "Psf", "Psf Toplam", "Mal Top(Kdv Dahil)"]],
            use_container_width=True
        )

# -------------------------------------------------
# Satış + Stok Birleşik Yorum
# -------------------------------------------------

if sales_df is not None and stock_df is not None:
    st.subheader("🧠 AYÇA Insight Yorumu")

    total_sales = sales_df["Toplam Tutar"].sum() if "Toplam Tutar" in sales_df.columns else 0
    total_stock_value = stock_df["Psf Toplam"].sum() if "Psf Toplam" in stock_df.columns else 0

    if "Tarih" in sales_df.columns:
        min_date = sales_df["Tarih"].min()
        max_date = sales_df["Tarih"].max()
        day_count = max((pd.to_datetime(max_date) - pd.to_datetime(min_date)).days, 1)
    else:
        day_count = 1

    daily_avg_sales = total_sales / day_count
    estimated_stock_days = total_stock_value / daily_avg_sales if daily_avg_sales > 0 else 0

    col1, col2, col3 = st.columns(3)

    col1.metric("Analiz Gün Sayısı", f"{day_count} gün")
    col2.metric("Günlük Ortalama Ciro", f"{daily_avg_sales:,.2f} TL")
    col3.metric("Tahmini Stok Karşılama Süresi", f"{estimated_stock_days:,.0f} gün")

    st.info(
        f"""
        Bu veri setiyle AYÇA Insight temel ticari analiz yapabilir.

        Şu anda toplam stok değeriniz yaklaşık **{total_stock_value:,.2f} TL**.

        Satış raporundaki ortalama günlük ciroya göre mevcut stok,
        yaklaşık **{estimated_stock_days:,.0f} günlük** satış hacmini karşılıyor görünüyor.

        Ancak ürün bazlı satış hareketi olmadığı için:
        - En çok satan ürünler
        - Ürün bazlı stok devir hızı
        - Otomatik sipariş önerisi

        henüz tam hesaplanamaz.
        """
    )

# -------------------------------------------------
# Ham Veri Önizleme
# -------------------------------------------------

with st.expander("📄 Ham Verileri Göster"):
    if sales_df is not None:
        st.write("Satış Verisi")
        st.dataframe(sales_df.head(100), use_container_width=True)

    if stock_df is not None:
        st.write("Stok Verisi")
        st.dataframe(stock_df.head(100), use_container_width=True)
