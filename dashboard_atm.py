import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from gspread_pandas import spread
import json
import sys

# --- KONFIGURASI KONEKSI (V35 - METODE JSON STRING) ---
try:
    # Kita mengambil satu blok teks penuh yang berisi seluruh credential
    # Ini jauh lebih aman daripada memecah variable satu per satu
    json_text = st.secrets["google_credentials"]["file_content"]
    
    # Python akan mengubah teks string menjadi Dictionary (JSON Object)
    # Ini otomatis menangani format \n dan karakter spesial
    creds_dict = json.loads(json_text)

    # Gunakan dictionary hasil parsing untuk koneksi
    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error(f"GAGAL KONEKSI KE GOOGLE SHEETS. Cek Secrets. Error Detail: {e}")
    # Stop eksekusi agar tidak muncul error beruntun
    sys.exit()

# --- LOAD DATA ---
st.set_page_config(layout='wide', page_title="Dashboard ATM")
SHEET_URL = "https://docs.google.com/spreadsheets/d/1G-Fp1l_4p5x6i9W0h-zH3Zl2M_F5q_8g_R8jWbWlG0o/edit?usp=sharing"
SHEET_NAME = 'data-atm-brilian'

@st.cache_data(ttl=600)
def load_data():
    try:
        sp = spread.Spread("dummy-name", sheet=SHEET_NAME, url=SHEET_URL, client=gc)
        df = sp.sheet_to_df(index=False)
        
        # Cleaning Data
        if 'Tanggal' in df.columns:
            df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        if 'Jumlah Transaksi' in df.columns:
            df['Jumlah Transaksi'] = pd.to_numeric(df['Jumlah Transaksi'], errors='coerce').fillna(0).astype(int)
        
        df = df.dropna(subset=['Tanggal'])
        return df
    except Exception as e:
        st.error(f"Gagal memuat data dari Sheet. Pastikan Sheet dishare ke client_email. Error: {e}")
        return pd.DataFrame()

# --- TAMPILAN DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data kosong atau gagal dimuat.")
else:
    st.title("💸 Dashboard Kinerja ATM Brilian")

    # --- SIDEBAR FILTER ---
    st.sidebar.header("Filter Data")
    
    # Filter Bank
    bank_options = ['Semua Bank'] + sorted(df['Bank'].unique().tolist())
    selected_bank = st.sidebar.selectbox('Pilih Bank:', bank_options)
    
    # Filter Lokasi
    lokasi_options = ['Semua Lokasi'] + sorted(df['Lokasi ATM'].unique().tolist())
    selected_location = st.sidebar.selectbox('Pilih Lokasi:', lokasi_options)

    # Filter Tanggal
    min_date = df['Tanggal'].min().date()
    max_date = df['Tanggal'].max().date()
    date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)

    # Logika Filter
    filtered_df = df.copy()
    if selected_bank != 'Semua Bank':
        filtered_df = filtered_df[filtered_df['Bank'] == selected_bank]
    if selected_location != 'Semua Lokasi':
        filtered_df = filtered_df[filtered_df['Lokasi ATM'] == selected_location]
    
    if isinstance(date_input, list) and len(date_input) == 2:
        start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
        filtered_df = filtered_df[(filtered_df['Tanggal'] >= start_date) & (filtered_df['Tanggal'] <= end_date)]

    # --- VISUALISASI UTAMA ---
    
    # Metrics
    col1, col2, col3 = st.columns(3)
    total_trx = filtered_df['Jumlah Transaksi'].sum()
    avg_trx = filtered_df['Jumlah Transaksi'].mean() if not filtered_df.empty else 0
    total_loc = filtered_df['Lokasi ATM'].nunique()
    
    col1.metric("Total Transaksi", f"{total_trx:,}")
    col2.metric("Rata-rata Transaksi", f"{avg_trx:,.0f}")
    col3.metric("Total Lokasi", total_loc)
    
    st.markdown("---")

    # Chart 1: Line Chart Tren
    st.subheader("Tren Transaksi Harian")
    daily_trend = filtered_df.groupby('Tanggal')['Jumlah Transaksi'].sum().reset_index()
    if not daily_trend.empty:
        fig_line = px.line(daily_trend, x='Tanggal', y='Jumlah Transaksi', template='plotly_dark')
        st.plotly_chart(fig_line, use_container_width=True)

    # Chart 2: Pie Chart Bank
    st.subheader("Distribusi per Bank")
    bank_dist = filtered_df.groupby('Bank')['Jumlah Transaksi'].sum().reset_index()
    if not bank_dist.empty:
        fig_pie = px.pie(bank_dist, values='Jumlah Transaksi', names='Bank', hole=0.4, template='plotly_dark')
        st.plotly_chart(fig_pie, use_container_width=True)
    
    # Tabel Data
    st.subheader("Detail Data")
    st.dataframe(filtered_df.sort_values('Tanggal', ascending=False), use_container_width=True)
