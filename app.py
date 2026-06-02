import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="AYÇA Insight Demo", page_icon="💊", layout="wide")

st.title("💊 AYÇA Insight Demo")
st.caption("Sample eczane verisi üzerinden ciro, kârlılık, stok, miad ve sipariş önerisi analizi")

DEFAULT_FILE = "ayca_insight_sample_veri.xlsx"

uploaded_file = st.sidebar.file_uploader("Excel dosyası yükleyin", type=["xlsx"])
file = uploaded_file if uploaded_file is not None else DEFAULT_FILE

@st.cache_data
def load_data(file):
    sales = pd.read_excel(file, sheet_name="Satis_Recete")
    stock = pd.read_excel(file, sheet_name="Stok")
    purchase = pd.read_excel(file, sheet_name="Alis")
    sales["Tarih"] = pd.to_datetime(sales["Tarih"])
    purchase["Tarih"] = pd.to_datetime(purchase["Tarih"])
    stock["Miad Tarihi"] = pd.to_datetime(stock["Miad Tarihi"])
    return sales, stock, purchase

try:
    sales, stock, purchase = load_data(file)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

# Sidebar filters
min_date = sales["Tarih"].min().date()
max_date = sales["Tarih"].max().date()
date_range = st.sidebar.date_input("Tarih aralığı", [min_date, max_date], min_value=min_date, max_value=max_date)
if len(date_range) == 2:
    start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    sales_f = sales[(sales["Tarih"] >= start_date) & (sales["Tarih"] <= end_date)]
else:
    sales_f = sales.copy()

categories = sorted(sales["Kategori"].dropna().unique())
selected_categories = st.sidebar.multiselect("Kategori", categories, default=categories)
sales_f = sales_f[sales_f["Kategori"].isin(selected_categories)]

# KPI metrics
total_revenue = sales_f["Ciro TL"].sum()
total_cost = sales_f["Maliyet TL"].sum()
gross_profit = sales_f["Brüt Kar TL"].sum()
gross_margin = gross_profit / total_revenue if total_revenue else 0
stock_value = stock["Stok Değeri TL"].sum()
near_expiry_count = (stock["Miad Kalan Gün"] <= 180).sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Toplam Ciro", f"{total_revenue:,.0f} TL")
c2.metric("Brüt Kar", f"{gross_profit:,.0f} TL")
c3.metric("Brüt Kar %", f"{gross_margin:.1%}")
c4.metric("Stok Değeri", f"{stock_value:,.0f} TL")
c5.metric("Yaklaşan Miad", f"{near_expiry_count} ürün")

# Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Genel Bakış", "Karlılık", "Stok & Miad", "Sipariş Önerisi", "Veri"])

with tab1:
    daily = sales_f.groupby("Tarih", as_index=False).agg({"Ciro TL":"sum", "Brüt Kar TL":"sum"})
    st.subheader("Günlük Ciro ve Brüt Kar")
    fig = px.line(daily, x="Tarih", y=["Ciro TL", "Brüt Kar TL"], markers=True)
    st.plotly_chart(fig, use_container_width=True)

    cat = sales_f.groupby("Kategori", as_index=False).agg({"Ciro TL":"sum", "Brüt Kar TL":"sum", "Adet":"sum"})
    cat["Kar %"] = cat["Brüt Kar TL"] / cat["Ciro TL"]
    st.subheader("Kategori Performansı")
    st.dataframe(cat.sort_values("Ciro TL", ascending=False), use_container_width=True)

with tab2:
    st.subheader("Ürün Bazlı Karlılık")
    prod = sales_f.groupby(["Barkod", "Ürün Adı", "Kategori"], as_index=False).agg(
        Adet=("Adet", "sum"),
        Ciro=("Ciro TL", "sum"),
        Maliyet=("Maliyet TL", "sum"),
        BrutKar=("Brüt Kar TL", "sum"),
    )
    prod["Kar %"] = prod["BrutKar"] / prod["Ciro"]
    st.dataframe(prod.sort_values("BrutKar", ascending=False), use_container_width=True)

    top_profit = prod.sort_values("BrutKar", ascending=False).head(10)
    fig = px.bar(top_profit, x="Ürün Adı", y="BrutKar", title="En Çok Brüt Kar Bırakan 10 Ürün")
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.subheader("Stok ve Miad Riskleri")
    risk = stock.copy()
    risk["Miad Risk"] = pd.cut(
        risk["Miad Kalan Gün"],
        bins=[-9999, 90, 180, 365, 99999],
        labels=["Çok Yakın", "Yakın", "Orta", "Normal"]
    )
    st.dataframe(risk.sort_values("Miad Kalan Gün"), use_container_width=True)

    fig = px.pie(risk, names="Miad Risk", values="Stok Değeri TL", title="Miad Riskine Göre Stok Değeri")
    st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.subheader("Basit Sipariş Önerisi")
    last_day = sales["Tarih"].max()
    last_60 = sales[sales["Tarih"] >= last_day - pd.Timedelta(days=60)]
    demand = last_60.groupby(["Barkod", "Ürün Adı"], as_index=False).agg({"Adet":"sum"})
    demand.rename(columns={"Adet":"Son 60 Gün Çıkış"}, inplace=True)
    demand["Aylık Ortalama Çıkış"] = demand["Son 60 Gün Çıkış"] / 2

    order = stock.merge(demand[["Barkod", "Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]], on="Barkod", how="left")
    order[["Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]] = order[["Son 60 Gün Çıkış", "Aylık Ortalama Çıkış"]].fillna(0)
    order["Stok Ay Karşılığı"] = order.apply(lambda r: 99 if r["Aylık Ortalama Çıkış"] == 0 else r["Mevcut Stok"] / r["Aylık Ortalama Çıkış"], axis=1)

    def recommendation(row):
        if row["Mevcut Stok"] < row["Minimum Stok"] or row["Stok Ay Karşılığı"] < 0.5:
            return "Acil sipariş"
        if row["Stok Ay Karşılığı"] < 1:
            return "Sipariş önerilir"
        if row["Stok Ay Karşılığı"] > 4 and row["Mevcut Stok"] > 20:
            return "Fazla stok"
        if row["Miad Kalan Gün"] <= 180:
            return "Miad takip"
        return "Normal"

    order["Öneri"] = order.apply(recommendation, axis=1)
    priority = {"Acil sipariş": 0, "Sipariş önerilir": 1, "Miad takip": 2, "Fazla stok": 3, "Normal": 4}
    order["Sıra"] = order["Öneri"].map(priority)
    st.dataframe(order.sort_values(["Sıra", "Stok Ay Karşılığı"]), use_container_width=True)

with tab5:
    st.subheader("Ham Veri")
    data_tab = st.radio("Tablo seç", ["Satış/Reçete", "Stok", "Alış"], horizontal=True)
    if data_tab == "Satış/Reçete":
        st.dataframe(sales, use_container_width=True)
    elif data_tab == "Stok":
        st.dataframe(stock, use_container_width=True)
    else:
        st.dataframe(purchase, use_container_width=True)
