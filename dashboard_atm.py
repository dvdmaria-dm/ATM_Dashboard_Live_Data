import streamlit as st
import pandas as pd
import plotly.express as px
from gspread_pandas import spread
import gspread
import sys
import re

# --- DATA CONFIGURATION ---
# URL ASLI DARI BANG DAVID (Sudah dipasang)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
# SHEET UTAMA
SHEET_NAME = 'AIMS_Master'

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="Dashboard ATM AIMS")

# --- KONEKSI SECRETS (V46 - STABLE) ---
try:
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]
    
    # Auto-Cleaning Key (Pembersih Spasi & Enter)
    raw_key = creds["private_key"]
    key_clean = raw_key.strip().replace("\\n", "\n")
    if "-----BEGIN PRIVATE KEY-----" in key_clean:
        content = key_clean.replace("-----BEGIN PRIVATE KEY-----", "").replace("-----END PRIVATE KEY-----", "")
        content = re.sub(r'\s+', '', content)
        final_private_key = f"-----BEGIN PRIVATE KEY-----\n{content}\n-----END PRIVATE KEY-----"
    else:
        final_private_key = key_clean

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
    
    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error(f"GAGAL KONEKSI SECRETS. Error: {e}")
    sys.exit()

# --- FUNGSI LOAD DATA ---
@st.cache_data(ttl=600)
def load_data():
    try:
        # Buka Spreadsheet via URL
        sh = gc.open_by_url(SHEET_URL)
        # Buka Tab AIMS_Master
        sp = spread.Spread(sh, sheet=SHEET_NAME)
        df = sp.sheet_to_df(index=False)
        
        # --- DATA CLEANING (SESUAI SCREENSHOT AIMS_MASTER) ---
        # 1. Bersihkan nama kolom (Hapus spasi, jadikan huruf besar semua)
        df.columns = df.columns.str.strip().str.upper()
        
        # 2. Mapping Kolom (Deteksi otomatis kolom yang ada)
        # Mencari kolom TANGGAL
        col_tgl = 'TANGGAL' if 'TANGGAL' in df.columns else None
        # Mencari kolom CABANG (sebagai Bank/Unit)
        col_cabang = 'CABANG' if 'CABANG' in df.columns else None
        # Mencari kolom LOKASI
        col_lokasi = 'LOKASI' if 'LOKASI' in df.columns else None
        # Mencari kolom JUMLAH_COMPLAIN (sebagai Metric)
        col_metric = 'JUMLAH_COMPLAIN' if 'JUMLAH_COMPLAIN' in df.columns else None

        # 3. Konversi Tipe Data
        if col_tgl:
            # dayfirst=True untuk format Indonesia (DD/MM/YYYY)
            df[col_tgl] = pd.to_datetime(df[col_tgl], dayfirst=True, errors='coerce')
            df = df.dropna(subset=[col_tgl])
            
        if col_metric:
            # Pastikan angka, ubah strip (-) jadi 0
            df[col_metric] = pd.to_numeric(df[col_metric], errors='coerce').fillna(0).astype(int)
        else:
            # Jika tidak ada kolom jumlah, kita hitung baris saja (Tiket Count)
            df['Total Tiket'] = 1
            col_metric = 'Total Tiket'

        return df, col_tgl, col_metric, col_lokasi, col_cabang

    except Exception as e:
        st.error(f"❌ GAGAL MEMUAT DATA: {e}")
        return pd.DataFrame(), None, None, None, None

# --- TAMPILAN DASHBOARD ---
data_result = load_data()

# Unpacking hasil return
if isinstance(data_result, tuple):
    df, tgl_col, metric_col, lok_col, cabang_col = data_result
else:
    df = data_result
    tgl_col, metric_col, lok_col, cabang_col = None, None, None, None

if df.empty:
    st.warning("Data belum tersedia atau gagal dimuat. Cek apakah Sheet 'AIMS_Master' ada datanya.")
else:
    # Judul Dashboard
    st.title("📊 Dashboard Monitoring ATM (AIMS)")
    st.markdown(f"**Data Source:** {SHEET_NAME} | **Total Data:** {len(df)} Baris")
    
    # --- SIDEBAR FILTER ---
    st.sidebar.header("Filter Data")
    
    # Filter Cabang
    if cabang_col:
        cabang_opts = ['Semua'] + sorted(df[cabang_col].astype(str).unique().tolist())
        sel_cabang = st.sidebar.selectbox(f'Pilih {cabang_col}:', cabang_opts)
    else:
        sel_cabang = 'Semua'

    # Filter Lokasi
    if lok_col:
        lok_opts = ['Semua'] + sorted(df[lok_col].astype(str).unique().tolist())
        sel_lok = st.sidebar.selectbox(f'Pilih {lok_col}:', lok_opts)
    else:
        sel_lok = 'Semua'

    # Filter Tanggal
    if tgl_col:
        min_date = df[tgl_col].min().date()
        max_date = df[tgl_col].max().date()
        date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)
    else:
        date_input = []

    # --- LOGIKA FILTER ---
    fil_df = df.copy()
    
    if sel_cabang != 'Semua' and cabang_col:
        fil_df = fil_df[fil_df[cabang_col].astype(str) == sel_cabang]
        
    if sel_lok != 'Semua' and lok_col:
        fil_df = fil_df[fil_df[lok_col].astype(str) == sel_lok]
    
    if tgl_col and isinstance(date_input, list) and len(date_input) == 2:
        start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
        fil_df = fil_df[(fil_df[tgl_col] >= start_date) & (fil_df[tgl_col] <= end_date)]

    # --- KPI METRICS ---
    st.markdown("---")
    if not fil_df.empty:
        col1, col2, col3 = st.columns(3)
        
        # Metric 1: Total Volume (Complain / Tiket)
        total_vol = fil_df[metric_col].sum()
        col1.metric("Total Volume (Tiket/Complain)", f"{total_vol:,}")
        
        # Metric 2: Total Lokasi Terdampak
        if lok_col:
            total_loc = fil_df[lok_col].nunique()
            col2.metric("Lokasi Terdampak", f"{total_loc}")
            
        # Metric 3: Rata-rata per Hari
        if tgl_col:
            days = (fil_df[tgl_col].max() - fil_df[tgl_col].min()).days + 1
            avg_daily = total_vol / days if days > 0 else total_vol
            col3.metric("Rata-rata per Hari", f"{avg_daily:,.1f}")

        # --- CHARTS ---
        st.markdown("---")
        
        # Chart 1: Tren Harian
        if tgl_col:
            st.subheader("Tren Harian")
            daily_trend = fil_df.groupby(tgl_col)[metric_col].sum().reset_index()
            fig_line = px.line(daily_trend, x=tgl_col, y=metric_col, markers=True, template='plotly_dark')
            st.plotly_chart(fig_line, use_container_width=True)

        # Chart 2: Top 10 Lokasi/Cabang
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            if cabang_col:
                st.subheader(f"Top 5 {cabang_col}")
                top_cabang = fil_df.groupby(cabang_col)[metric_col].sum().reset_index().sort_values(metric_col, ascending=False).head(5)
                fig_bar = px.bar(top_cabang, x=metric_col, y=cabang_col, orientation='h', template='plotly_dark')
                st.plotly_chart(fig_bar, use_container_width=True)
                
        with col_chart2:
            if lok_col:
                st.subheader(f"Top 5 {lok_col}")
                top_lok = fil_df.groupby(lok_col)[metric_col].sum().reset_index().sort_values(metric_col, ascending=False).head(5)
                fig_bar2 = px.bar(top_lok, x=metric_col, y=lok_col, orientation='h', template='plotly_dark')
                st.plotly_chart(fig_bar2, use_container_width=True)

        # --- DATA TABLE ---
        st.subheader("Detail Data")
        st.dataframe(fil_df, use_container_width=True)
    else:
        st.info("Tidak ada data yang sesuai dengan filter.")
