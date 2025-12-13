import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from gspread_pandas import spread
import json
import sys

# --- CONFIGURASI INI WAJIB UNTUK STREAMLIT CLOUD ---
try:
    # 1. Mengambil kunci rahasia dari Streamlit Secrets (hasil input TOML)
    secrets = st.secrets["gspread_service_account"]
    secrets_dict = dict(secrets)

    # 2. Menggunakan kunci rahasia untuk otorisasi gspread
    gc = gspread.service_account_from_dict(secrets_dict)
    
except Exception as e:
    # MENUTUP BLOK TRY: Memberikan pesan error yang jelas jika koneksi Secrets gagal
    st.error(f"GAGAL KONEKSI G-SHEET: Error konfigurasi kunci rahasia (Secrets). Cek kembali format TOML. Error: {e}")
    sys.exit() # Menghentikan aplikasi jika gagal koneksi Secrets

# --- PENGATURAN UMUM ---
st.set_page_config(layout='wide')
SHEET_URL = "https://docs.google.com/spreadsheets/d/1G-Fp1l_4p5x6i9W0h-zH3Zl2M_F5q_8g_R8jWbWlG0o/edit?usp=sharing"
SHEET_NAME = 'data-atm-brilian'

@st.cache_data(ttl=600)
def load_data():
    try:
        # Menggunakan Spreadsheet API untuk membaca data
        sp = spread.Spread("dummy-name", sheet=SHEET_NAME, url=SHEET_URL, client=gc)
        df = sp.sheet_to_df(index=False)
        
        # Data Cleaning dan Konversi
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        df['Jumlah Transaksi'] = df['Jumlah Transaksi'].astype(int, errors='ignore')
        df = df.dropna(subset=['Tanggal'])
        
        return df
    
    except Exception as e:
        st.error(f"GAGAL MEMUAT DATA GOOGLE SHEETS. Cek kunci Secrets dan URL Sheet. Error: {e}")
        return pd.DataFrame()

# --- EKSEKUSI APLIKASI UTAMA ---

df = load_data()

if df.empty:
    st.warning("Dashboard tidak dapat menampilkan data. Periksa log error di atas dan pastikan Google Sheet sudah diizinkan oleh Service Account.")
else:
    # Tampilan Dashboard
    st.title("💸 Dashboard Kinerja ATM Brilian")

    # Sidebar Filter
    st.sidebar.header("Filter Data")
    
    # Filter Bank
    bank_list = ['Semua Bank'] + sorted(df['Bank'].unique().tolist())
    selected_bank = st.sidebar.selectbox('Pilih Bank:', bank_list)
    
    # Filter Lokasi
    location_list = ['Semua Lokasi'] + sorted(df['Lokasi ATM'].unique().tolist())
    selected_location = st.sidebar.selectbox('Pilih Lokasi:', location_list)

    # Filter Transaksi
    min_date = df['Tanggal'].min().date()
    max_date = df['Tanggal'].max().date()
    date_range = st.sidebar.date_input("Pilih Rentang Tanggal:", [min_date, max_date], min_value=min_date, max_value=max_date)

    # Aplikasi Filter
    filtered_df = df.copy()
    
    if selected_bank != 'Semua Bank':
        filtered_df = filtered_df[filtered_df['Bank'] == selected_bank]
        
    if selected_location != 'Semua Lokasi':
        filtered_df = filtered_df[filtered_df['Lokasi ATM'] == selected_location]
        
    if len(date_range) == 2:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])
        filtered_df = filtered_df[(filtered_df['Tanggal'] >= start_date) & (filtered_df['Tanggal'] <= end_date)]

    # --- METRICS DAN VISUALISASI ---
    
    total_transaksi = filtered_df['Jumlah Transaksi'].sum()
    total_atm = filtered_df['Lokasi ATM'].nunique()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Total Jumlah Transaksi", value=f"{total_transaksi:,}")
    with col2:
        st.metric(label="Rata-rata Transaksi Harian", value=f"{filtered_df.groupby('Tanggal')['Jumlah Transaksi'].sum().mean():,.0f}")
    with col3:
        st.metric(label="Total Lokasi ATM Terpantau", value=f"{total_atm:,}")
        
    st.markdown("---")

    # 1. Tren Transaksi Harian
    st.subheader("Tren Jumlah Transaksi Harian")
    daily_trend = filtered_df.groupby('Tanggal')['Jumlah Transaksi'].sum().reset_index()
    fig_trend = px.line(daily_trend, x='Tanggal', y='Jumlah Transaksi', 
                        title='Tren Transaksi dari Waktu ke Waktu', 
                        template='plotly_dark')
    st.plotly_chart(fig_trend, use_container_width=True)

    # 2. Transaksi per Bank (Pie Chart)
    st.subheader("Distribusi Transaksi Berdasarkan Bank")
    bank_summary = filtered_df.groupby('Bank')['Jumlah Transaksi'].sum().reset_index()
    fig_bank = px.pie(bank_summary, values='Jumlah Transaksi', names='Bank', 
                      title='Persentase Transaksi per Bank', 
                      hole=.3, template='plotly_dark')
    st.plotly_chart(fig_bank, use_container_width=True)

    # 3. Kinerja per Lokasi (Tabel)
    st.subheader("Detail Kinerja Transaksi per Lokasi ATM")
    location_summary = filtered_df.groupby('Lokasi ATM')['Jumlah Transaksi'].agg(['sum', 'count']).reset_index()
    location_summary.columns = ['Lokasi ATM', 'Total Transaksi', 'Jumlah Hari Data']
    
    st.dataframe(location_summary.sort_values(by='Total Transaksi', ascending=False), use_container_width=True)
