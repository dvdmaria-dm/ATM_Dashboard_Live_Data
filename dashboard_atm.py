import streamlit as st
import pandas as pd
import plotly.express as px
import gspread  # Kita pakai ini saja, jangan pakai gspread_pandas lagi!
import sys
import re

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="Dashboard ATM AIMS")

# --- DATA CONFIGURATION ---
# URL dan Nama Sheet ASLI milik Bang David
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_NAME = 'AIMS_Master'

# --- KONEKSI SECRETS (SUDAH TERUJI BERHASIL) ---
try:
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]
    
    # Auto-Cleaning Key (Pembersih Kunci)
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
    
    # Login ke Google
    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error(f"GAGAL KONEKSI SECRETS. Error: {e}")
    sys.exit()

# --- FUNGSI LOAD DATA (METODE BARU - LEBIH STABIL) ---
@st.cache_data(ttl=600)
def load_data():
    try:
        # 1. Buka Spreadsheet
        sh = gc.open_by_url(SHEET_URL)
        
        # 2. Buka Tab (Worksheet)
        ws = sh.worksheet(SHEET_NAME)
        
        # 3. Ambil Semua Data (Ini metode gspread murni, tanpa library tambahan)
        data = ws.get_all_records()
        
        # 4. Masukkan ke Pandas DataFrame
        df = pd.DataFrame(data)
        
        # --- DATA CLEANING ---
        if df.empty:
            return pd.DataFrame(), None, None, None, None

        # Standarisasi Nama Kolom (Huruf Besar & Hapus Spasi)
        df.columns = df.columns.str.strip().str.upper()
        
        # Deteksi Kolom Otomatis
        col_tgl = 'TANGGAL' if 'TANGGAL' in df.columns else None
        col_cabang = 'CABANG' if 'CABANG' in df.columns else None
        col_lokasi = 'LOKASI' if 'LOKASI' in df.columns else None
        col_metric = 'JUMLAH_COMPLAIN' if 'JUMLAH_COMPLAIN' in df.columns else None

        # Konversi Tipe Data
        if col_tgl:
            df[col_tgl] = pd.to_datetime(df[col_tgl], dayfirst=True, errors='coerce')
            df = df.dropna(subset=[col_tgl])
            
        if col_metric:
            # Pastikan data berupa string dulu sebelum di-replace, lalu convert ke angka
            df[col_metric] = df[col_metric].astype(str).str.replace('-', '0')
            df[col_metric] = pd.to_numeric(df[col_metric], errors='coerce').fillna(0).astype(int)
        else:
            df['Total Tiket'] = 1
            col_metric = 'Total Tiket'

        return df, col_tgl, col_metric, col_lokasi, col_cabang

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ TAB TIDAK DITEMUKAN: '{SHEET_NAME}'. Pastikan nama tab (Sheet) di bawah Excel sama persis.")
        return pd.DataFrame(), None, None, None, None
    except Exception as e:
        st.error(f"❌ GAGAL MEMUAT DATA. Error: {e}")
        return pd.DataFrame(), None, None, None, None

# --- TAMPILAN DASHBOARD ---
data_result = load_data()

# Unpack hasil
if isinstance(data_result, tuple):
    df, tgl_col, metric_col, lok_col, cabang_col = data_result
else:
    df = data_result
    tgl_col, metric_col, lok_col, cabang_col = None, None, None, None

if df.empty:
    st.warning("Data kosong atau tidak terbaca. Mohon cek URL dan Nama Sheet.")
else:
    st.title("📊 Dashboard Monitoring ATM (AIMS)")
    st.markdown(f"**Data Source:** {SHEET_NAME} | **Total Data:** {len(df)} Baris")
    
    # --- SIDEBAR FILTER ---
    st.sidebar.header("Filter Data")
    fil_df = df.copy()

    # Filter Cabang
    if cabang_col:
        cabang_opts = ['Semua'] + sorted(df[cabang_col].astype(str).unique().tolist())
        sel_cabang = st.sidebar.selectbox(f'Pilih {cabang_col}:', cabang_opts)
        if sel_cabang != 'Semua':
            fil_df = fil_df[fil_df[cabang_col].astype(str) == sel_cabang]

    # Filter Lokasi
    if lok_col:
        lok_opts = ['Semua'] + sorted(df[lok_col].astype(str).unique().tolist())
        sel_lok = st.sidebar.selectbox(f'Pilih {lok_col}:', lok_opts)
        if sel_lok != 'Semua':
            fil_df = fil_df[fil_df[lok_col].astype(str) == sel_lok]

    # Filter Tanggal
    if tgl_col:
        min_date = df[tgl_col].min().date()
        max_date = df[tgl_col].max().date()
        date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)
        if isinstance(date_input, list) and len(date_input) == 2:
            start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
            fil_df = fil_df[(fil_df[tgl_col] >= start_date) & (fil_df[tgl_col] <= end_date)]

    # --- KPI METRICS ---
    st.markdown("---")
    if not fil_df.empty:
        col1, col2, col3 = st.columns(3)
        
        total_vol = fil_df[metric_col].sum()
        col1.metric("Total Volume", f"{total_vol:,}")
        
        if lok_col:
            total_loc = fil_df[lok_col].nunique()
            col2.metric("Lokasi Terdampak", f"{total_loc}")
            
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

        # Chart 2: Top Lokasi
        if cabang_col:
            st.subheader(f"Top 10 {cabang_col} Tertinggi")
            top_cabang = fil_df.groupby(cabang_col)[metric_col].sum().reset_index().sort_values(metric_col, ascending=False).head(10)
            fig_bar = px.bar(top_cabang, x=metric_col, y=cabang_col, orientation='h', template='plotly_dark')
            st.plotly_chart(fig_bar, use_container_width=True)

        # --- DATA TABLE ---
        st.subheader("Detail Data")
        st.dataframe(fil_df, use_container_width=True)
    else:
        st.info("Tidak ada data yang sesuai dengan filter.")
