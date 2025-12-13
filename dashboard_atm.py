import streamlit as st
import pandas as pd
import plotly.express as px
from gspread_pandas import spread
import gspread
import sys
import re # Kita butuh ini untuk pembersihan total

# --- KONFIGURASI KONEKSI V43 (AUTO-CLEANING) ---
try:
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]

    # --- FITUR PEMBERSIH KUNCI (THE CLEANER) ---
    raw_key = creds["private_key"]
    
    # 1. Hapus semua spasi dan tab yang tidak sengaja terbawa
    key_clean = raw_key.strip() 
    
    # 2. Koreksi karakter \n yang sering rusak
    # Jika ada double slash \\n, ubah jadi \n
    key_clean = key_clean.replace("\\n", "\n")
    
    # 3. Trik Paling Jitu: Jika kunci masih dianggap error, kita bangun ulang
    # Kadang copy paste membuat header terpisah spasi. Kita rapihkan paksa.
    if "-----BEGIN PRIVATE KEY-----" in key_clean:
        # Ambil isinya saja
        content = key_clean.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
        # Buang semua spasi/enter di dalam konten
        content = re.sub(r'\s+', '', content)
        # Susun ulang dengan header yang sempurna
        final_private_key = f"-----BEGIN PRIVATE KEY-----\n{content}\n-----END PRIVATE KEY-----"
    else:
        final_private_key = key_clean

    # Susun dictionary
    creds_dict = {
        "type": creds["type"],
        "project_id": creds["project_id"],
        "private_key_id": creds["private_key_id"],
        "private_key": final_private_key, # Pakai kunci yang sudah dicuci bersih
        "client_email": creds["client_email"],
        "client_id": creds["client_id"],
        "auth_uri": creds["auth_uri"],
        "token_uri": creds["token_uri"],
        "auth_provider_x509_cert_url": creds["auth_provider_x509_cert_url"],
        "client_x509_cert_url": creds["client_x509_cert_url"],
        "universe_domain": creds["universe_domain"]
    }
    
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
        st.error(f"Gagal memuat data. Error: {e}")
        return pd.DataFrame()

# --- TAMPILAN DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data kosong atau gagal dimuat.")
else:
    st.title("💸 Dashboard Kinerja ATM Brilian")

    st.sidebar.header("Filter Data")
    bank_options = ['Semua Bank'] + sorted(df['Bank'].unique().tolist())
    selected_bank = st.sidebar.selectbox('Pilih Bank:', bank_options)
    
    lokasi_options = ['Semua Lokasi'] + sorted(df['Lokasi ATM'].unique().tolist())
    selected_location = st.sidebar.selectbox('Pilih Lokasi:', lokasi_options)

    min_date = df['Tanggal'].min().date()
    max_date = df['Tanggal'].max().date()
    date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)

    filtered_df = df.copy()
    if selected_bank != 'Semua Bank':
        filtered_df = filtered_df[filtered_df['Bank'] == selected_bank]
    if selected_location != 'Semua Lokasi':
        filtered_df = filtered_df[filtered_df['Lokasi ATM'] == selected_location]
    
    if isinstance(date_input, list) and len(date_input) == 2:
        start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
        filtered_df = filtered_df[(filtered_df['Tanggal'] >= start_date) & (filtered_df['Tanggal'] <= end_date)]

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
