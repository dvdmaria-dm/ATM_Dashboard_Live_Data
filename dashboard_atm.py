import streamlit as st
import pandas as pd
import plotly.express as px
from gspread_pandas import spread
import gspread
import sys
import re

# --- KONFIGURASI KONEKSI V44 (FIXED LIBRARY USAGE) ---
try:
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]

    # --- FITUR PEMBERSIH KUNCI ---
    raw_key = creds["private_key"]
    key_clean = raw_key.strip() 
    key_clean = key_clean.replace("\\n", "\n")
    
    if "-----BEGIN PRIVATE KEY-----" in key_clean:
        content = key_clean.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
        content = re.sub(r'\s+', '', content)
        final_private_key = f"-----BEGIN PRIVATE KEY-----\n{content}\n-----END PRIVATE KEY-----"
    else:
        final_private_key = key_clean

    # Susun dictionary
    creds_dict = {
        "type": creds["type"],
        "project_id": creds["project_id"],
        "private_key_id": creds["private_key_id"],
        "private_key": final_private_key,
        "client_email": creds["client_email"],
        "client_id": creds["client_id"],
        "auth_uri": creds["auth_uri"],
        "token_uri": creds["token_uri"],
        "auth_provider_x509_cert_url": creds["auth_provider_x509_cert_url"],
        "client_x509_cert_url": creds["client_x509_cert_url"],
        "universe_domain": creds["universe_domain"]
    }
    
    # KONEKSI BERHASIL DI SINI
    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error(f"GAGAL KONEKSI. Masalahnya ada di Private Key. Detail Error: {e}")
    sys.exit()

# --- LOAD DATA ---
st.set_page_config(layout='wide', page_title="Dashboard ATM")
SHEET_URL = "https://docs.google.com/spreadsheets/d/1G-Fp1l_4p5x6i9W0h-zH3Zl2M_F5q_8g_R8jWbWlG0o/edit?usp=sharing"
SHEET_NAME = 'data-atm-brilian'

@st.cache_data(ttl=600)
def load_data():
    try:
        # PERBAIKAN V44:
        # Kita buka Spreadsheet-nya dulu pakai 'gc' (karena kita punya URL-nya)
        # Ini metode paling anti-gagal dibanding menebak nama file.
        sh = gc.open_by_url(SHEET_URL)
        
        # Lalu kita serahkan objek spreadsheet 'sh' ke gspread_pandas untuk dijadikan DataFrame
        # HAPUS parameter 'url=' yang bikin error tadi!
        sp = spread.Spread(sh, sheet=SHEET_NAME)
        
        df = sp.sheet_to_df(index=False)
        
        # Cleaning Data
        if 'Tanggal' in df.columns:
            df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        if 'Jumlah Transaksi' in df.columns:
            df['Jumlah Transaksi'] = pd.to_numeric(df['Jumlah Transaksi'], errors='coerce').fillna(0).astype(int)
        
        df = df.dropna(subset=['Tanggal'])
        return df
    except Exception as e:
        # Jika error di sini, kemungkinan besar email service account belum di-invite ke GSheet
        st.error(f"GAGAL MEMBUKA SPREADSHEET. Pastikan email '{creds['client_email']}' sudah dijadikan 'Viewer/Editor' di Google Sheet Anda. Error: {e}")
        return pd.DataFrame()

# --- TAMPILAN DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data kosong atau gagal dimuat.")
else:
    st.title("💸 Dashboard Kinerja ATM Brilian")

    st.sidebar.header("Filter Data")
    # Cek apakah kolom ada sebelum membuat filter
    if 'Bank' in df.columns:
        bank_options = ['Semua Bank'] + sorted(df['Bank'].unique().tolist())
        selected_bank = st.sidebar.selectbox('Pilih Bank:', bank_options)
    else:
        selected_bank = 'Semua Bank'
        
    if 'Lokasi ATM' in df.columns:
        lokasi_options = ['Semua Lokasi'] + sorted(df['Lokasi ATM'].unique().tolist())
        selected_location = st.sidebar.selectbox('Pilih Lokasi:', lokasi_options)
    else:
        selected_location = 'Semua Lokasi'

    if 'Tanggal' in df.columns:
        min_date = df['Tanggal'].min().date()
        max_date = df['Tanggal'].max().date()
        date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)
    else:
        date_input = []

    # Filter Logic
    filtered_df = df.copy()
    if selected_bank != 'Semua Bank':
        filtered_df = filtered_df[filtered_df['Bank'] == selected_bank]
    if selected_location != 'Semua Lokasi':
        filtered_df = filtered_df[filtered_df['Lokasi ATM'] == selected_location]
    
    if isinstance(date_input, list) and len(date_input) == 2:
        start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
        filtered_df = filtered_df[(filtered_df['Tanggal'] >= start_date) & (filtered_df['Tanggal'] <= end_date)]

    # Metrics
    if not filtered_df.empty:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Transaksi", f"{filtered_df['Jumlah Transaksi'].sum():,}")
        col2.metric("Rata-rata Transaksi", f"{filtered_df['Jumlah Transaksi'].mean():,.0f}")
        col3.metric("Total Lokasi", filtered_df['Lokasi ATM'].nunique())
        
        st.markdown("---")

        st.subheader("Tren Transaksi Harian")
        daily_trend = filtered_df.groupby('Tanggal')['Jumlah Transaksi'].sum().reset_index()
        st.plotly_chart(px.line(daily_trend, x='Tanggal', y='Jumlah Transaksi', template='plotly_dark'), use_container_width=True)

        st.subheader("Distribusi per Bank")
        bank_dist = filtered_df.groupby('Bank')['Jumlah Transaksi'].sum().reset_index()
        st.plotly_chart(px.pie(bank_dist, values='Jumlah Transaksi', names='Bank', hole=0.4, template='plotly_dark'), use_container_width=True)

        st.subheader("Detail Data")
        st.dataframe(filtered_df.sort_values('Tanggal', ascending=False), use_container_width=True)
