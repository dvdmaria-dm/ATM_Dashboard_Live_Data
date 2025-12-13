import streamlit as st
import pandas as pd
import plotly.express as px
from gspread_pandas import spread
import gspread
import sys

# --- KONFIGURASI KONEKSI V41 (FINAL MATCHING) ---
try:
    # KITA PAKAI LOGIKA SEDERHANA:
    # Kita cari kunci bernama "gcp_service_account" di Secrets.
    # Jika error sebelumnya bilang "no key google_credentials", itu karena script lama.
    # Script ini PASTI mencari "gcp_service_account".
    
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS. Mohon cek ejaan header di Secrets.")
        st.stop()
    
    # Ambil data dari secrets
    creds = st.secrets["gcp_service_account"]

    # SIASAT KHUSUS PRIVATE KEY:
    # Mengembalikan karakter \n yang mungkin rusak saat copy-paste
    private_key_fixed = creds["private_key"].replace("\\n", "\n")

    # Susun ulang dictionary agar dimengerti Google
    creds_dict = {
        "type": creds["type"],
        "project_id": creds["project_id"],
        "private_key_id": creds["private_key_id"],
        "private_key": private_key_fixed,
        "client_email": creds["client_email"],
        "client_id": creds["client_id"],
        "auth_uri": creds["auth_uri"],
        "token_uri": creds["token_uri"],
        "auth_provider_x509_cert_url": creds["auth_provider_x509_cert_url"],
        "client_x509_cert_url": creds["client_x509_cert_url"],
        "universe_domain": creds["universe_domain"]
    }
    
    # Buka koneksi
    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error(f"GAGAL KONEKSI. Detail Error: {e}")
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
        st.error(f"Gagal memuat data dari Google Sheet. Error: {e}")
        return pd.DataFrame()

# --- TAMPILAN DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data kosong atau gagal dimuat.")
else:
    st.title("💸 Dashboard Kinerja ATM Brilian")

    # Sidebar
    st.sidebar.header("Filter Data")
    bank_options = ['Semua Bank'] + sorted(df['Bank'].unique().tolist())
    selected_bank = st.sidebar.selectbox('Pilih Bank:', bank_options)
    
    lokasi_options = ['Semua Lokasi'] + sorted(df['Lokasi ATM'].unique().tolist())
    selected_location = st.sidebar.selectbox('Pilih Lokasi:', lokasi_options)

    min_date = df['Tanggal'].min().date()
    max_date = df['Tanggal'].max().date()
    date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)

    # Filter Logic
    filtered_df = df.copy()
    if selected_bank != 'Semua Bank':
        filtered_df = filtered_df[filtered_df['Bank'] == selected_bank]
    if selected_location != 'Semua Lokasi':
        filtered_df = filtered_df[filtered_df['Lokasi ATM'] == selected_location]
    
    if isinstance(date_input, list) and len(date_input) == 2:
        start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
        filtered_df = filtered_df[(filtered_df['Tanggal'] >= start_date) & (filtered_df['Tanggal'] <= end_date)]

    # Visualisasi
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Transaksi", f"{filtered_df['Jumlah Transaksi'].sum():,}")
    col2.metric("Rata-rata Transaksi", f"{filtered_df['Jumlah Transaksi'].mean():,.0f}")
    col3.metric("Total Lokasi", filtered_df['Lokasi ATM'].nunique())
    
    st.markdown("---")

    st.subheader("Tren Transaksi Harian")
    daily_trend = filtered_df.groupby('Tanggal')['Jumlah Transaksi'].sum().reset_index()
    if not daily_trend.empty:
        st.plotly_chart(px.line(daily_trend, x='Tanggal', y='Jumlah Transaksi', template='plotly_dark'), use_container_width=True)

    st.subheader("Distribusi per Bank")
    bank_dist = filtered_df.groupby('Bank')['Jumlah Transaksi'].sum().reset_index()
    if not bank_dist.empty:
        st.plotly_chart(px.pie(bank_dist, values='Jumlah Transaksi', names='Bank', hole=0.4, template='plotly_dark'), use_container_width=True)

    st.subheader("Detail Data")
    st.dataframe(filtered_df.sort_values('Tanggal', ascending=False), use_container_width=True)
