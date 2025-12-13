import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="Dashboard ATM AIMS")

# --- DATA CONFIGURATION ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_NAME = 'AIMS_Master'

# --- KONEKSI SECRETS (SUDAH STABIL) ---
try:
    if "gcp_service_account" not in st.secrets:
        st.error("KUNCI 'gcp_service_account' TIDAK DITEMUKAN DI SECRETS.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]
    
    # Auto-Cleaning Key
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

# --- FUNGSI LOAD DATA (METODE RAW - ANTI DUPLICATE HEADER) ---
@st.cache_data(ttl=600)
def load_data():
    try:
        sh = gc.open_by_url(SHEET_URL)
        ws = sh.worksheet(SHEET_NAME)
        
        # PERUBAHAN V49: Ambil semua data sebagai List MENTAH (Matriks)
        # Ini menghindari error "duplicates header" dari library gspread
        all_values = ws.get_all_values()
        
        if not all_values:
            return pd.DataFrame(), None, None, None, None

        # Baris pertama adalah Header, sisanya Data
        headers = all_values[0]
        rows = all_values[1:]

        # Buat DataFrame
        df = pd.DataFrame(rows, columns=headers)
        
        # --- DATA CLEANING EKSTRA ---
        # 1. Buang kolom yang namanya Kosong (penyebab error sebelumnya)
        df = df.loc[:, df.columns != '']
        
        # 2. Standarisasi Header (Huruf Besar & Trim Spasi)
        df.columns = df.columns.str.strip().str.upper()

        # 3. Deteksi Kolom
        col_tgl = 'TANGGAL' if 'TANGGAL' in df.columns else None
        col_cabang = 'CABANG' if 'CABANG' in df.columns else None
        col_lokasi = 'LOKASI' if 'LOKASI' in df.columns else None
        col_metric = 'JUMLAH_COMPLAIN' if 'JUMLAH_COMPLAIN' in df.columns else None

        # 4. Konversi Tipe Data
        if col_tgl:
            # Gunakan dayfirst=True untuk format DD/MM/YYYY
            df[col_tgl] = pd.to_datetime(df[col_tgl], dayfirst=True, errors='coerce')
            df = df.dropna(subset=[col_tgl])
            
        if col_metric:
            # Bersihkan angka dari teks aneh, ubah '-' jadi 0
            df[col_metric] = df[col_metric].astype(str).str.replace(r'[^\d\-]', '', regex=True)
            df[col_metric] = pd.to_numeric(df[col_metric].replace({'': '0', '-': '0'}), errors='coerce').fillna(0).astype(int)
        else:
            df['Total Tiket'] = 1
            col_metric = 'Total Tiket'

        return df, col_tgl, col_metric, col_lokasi, col_cabang

    except Exception as e:
        st.error(f"❌ GAGAL MEMROSES DATA. Error: {e}")
        return pd.DataFrame(), None, None, None, None

# --- TAMPILAN DASHBOARD ---
data_result = load_data()

if isinstance(data_result, tuple):
    df, tgl_col, metric_col, lok_col, cabang_col = data_result
else:
    df = data_result
    tgl_col, metric_col, lok_col, cabang_col = None, None, None, None

if df.empty:
    st.warning("Data berhasil terkoneksi, tapi isinya kosong atau format tidak terbaca.")
else:
    st.title("📊 Dashboard Monitoring ATM (AIMS)")
    st.markdown(f"**Status:** Terkoneksi ✅ | **Total Data:** {len(df)} Baris")
    
    # --- SIDEBAR ---
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
        try:
            min_date = fil_df[tgl_col].min().date()
            max_date = fil_df[tgl_col].max().date()
            date_input = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date], min_value=min_date, max_value=max_date)
            if isinstance(date_input, list) and len(date_input) == 2:
                start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
                fil_df = fil_df[(fil_df[tgl_col] >= start_date) & (fil_df[tgl_col] <= end_date)]
        except:
            st.sidebar.info("Filter tanggal non-aktif (Format tanggal di Excel mungkin beragam).")

    # --- VISUALISASI ---
    st.markdown("---")
    if not fil_df.empty:
        col1, col2, col3 = st.columns(3)
        
        total_vol = fil_df[metric_col].sum()
        col1.metric("Total Volume", f"{total_vol:,}")
        
        if lok_col:
            total_loc = fil_df[lok_col].nunique()
            col2.metric("Lokasi Terdampak", f"{total_loc}")
            
        # Chart 1: Tren
        if tgl_col:
            st.subheader("Tren Harian")
            daily_trend = fil_df.groupby(tgl_col)[metric_col].sum().reset_index()
            st.plotly_chart(px.line(daily_trend, x=tgl_col, y=metric_col, markers=True, template='plotly_dark'), use_container_width=True)

        # Chart 2: Top Cabang/Lokasi
        if cabang_col:
            st.subheader(f"Top 10 {cabang_col}")
            top_cabang = fil_df.groupby(cabang_col)[metric_col].sum().reset_index().sort_values(metric_col, ascending=False).head(10)
            st.plotly_chart(px.bar(top_cabang, x=metric_col, y=cabang_col, orientation='h', template='plotly_dark'), use_container_width=True)

        # Tabel
        st.subheader("Detail Data")
        st.dataframe(fil_df, use_container_width=True)
    else:
        st.info("Tidak ada data yang sesuai filter.")
