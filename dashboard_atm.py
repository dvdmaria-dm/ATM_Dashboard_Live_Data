import streamlit as st
import pandas as pd
import plotly.express as px
from gspread_pandas import spread
import gspread
import sys
import re

# --- LOAD DATA (BAGIAN INI WAJIB KAU ISI, BANG!) ---
# 1. Masukkan Link Google Sheet Asli Milikmu di bawah ini:
#    Contoh: "https://docs.google.com/spreadsheets/d/1234567890abcdefg/edit"
SHEET_URL = https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit?gid=98670277#gid=98670277  # <--- GANTI TEKS INI DENGAN LINK ASLIMU

# 2. Masukkan Nama Tab (Sheet) yang ada datanya:
#    Lihat di bagian bawah Google Sheet-mu.
#    Berdasarkan screenshotmu, kemungkinannya adalah 'AIMS_Master'
SHEET_NAME = 'AIMS_Master'  # <--- PASTIKAN INI SAMA PERSIS (Huruf Besar/Kecil Pengaruh!)

# -------------------------------------------------------
# --- KONFIGURASI KONEKSI V45 (NO EDIT NEEDED HERE) ---
st.set_page_config(layout='wide', page_title="Dashboard ATM")

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

@st.cache_data(ttl=600)
def load_data():
    try:
        # Cek apakah user sudah mengganti URL dummy
        if "TEMPEL_URL" in SHEET_URL:
             st.error("⚠️ URL BELUM DIISI! Tolong edit file dashboard_atm.py dan masukkan Link Google Sheet asli Anda di baris 'SHEET_URL'.")
             return pd.DataFrame()

        sh = gc.open_by_url(SHEET_URL)
        sp = spread.Spread(sh, sheet=SHEET_NAME)
        df = sp.sheet_to_df(index=False)
        
        # Cleaning Data (Sesuaikan nama kolom dengan Excel aslimu!)
        # Aku lihat di screenshot ada kolom: 'TANGGAL', 'LOKASI', 'JUMLAH_COMPLAIN', dll.
        # Kita harus sesuaikan agar tidak error kolom tidak ditemukan.
        
        # Standardisasi nama kolom (ubah ke huruf kecil semua biar aman)
        df.columns = df.columns.str.strip() # Hapus spasi nama kolom
        
        # Mapping kolom (Sesuaikan dengan data aslimu di AIMS_Master)
        # Ganti 'TANGGAL' jika di excelmu namanya beda
        col_tanggal = 'TANGGAL' if 'TANGGAL' in df.columns else 'Tanggal'
        col_jumlah = 'JUMLAH_COMPLAIN' if 'JUMLAH_COMPLAIN' in df.columns else 'Jumlah Transaksi' 
        col_lokasi = 'LOKASI' if 'LOKASI' in df.columns else 'Lokasi ATM'
        col_bank = 'CABANG' if 'CABANG' in df.columns else 'Bank' # Asumsi Cabang sebagai grouping

        if col_tanggal in df.columns:
            df[col_tanggal] = pd.to_datetime(df[col_tanggal], errors='coerce')
            df = df.dropna(subset=[col_tanggal])
            
        if col_jumlah in df.columns:
            df[col_jumlah] = pd.to_numeric(df[col_jumlah], errors='coerce').fillna(0).astype(int)

        return df, col_tanggal, col_jumlah, col_lokasi, col_bank

    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ TAB TIDAK DITEMUKAN: '{SHEET_NAME}'. Cek nama tab di bawah Google Sheet Anda (Contoh: AIMS_Master) dan ganti di script.")
        return pd.DataFrame(), None, None, None, None
    except Exception as e:
        st.error(f"❌ GAGAL MEMBUKA SPREADSHEET: {e}")
        return pd.DataFrame(), None, None, None, None

# --- TAMPILAN DASHBOARD ---
data_result = load_data()

# Handle return value (karena fungsi bisa return tuple atau dataframe kosong)
if isinstance(data_result, tuple):
    df, tgl_col, jml_col, lok_col, bank_col = data_result
else:
    df = data_result
    tgl_col, jml_col, lok_col, bank_col = None, None, None, None

if df.empty:
    st.warning("Data belum tersedia. Cek URL dan Nama Tab.")
else:
    st.title("💸 Dashboard ATM Monitoring")
    
    # Sidebar Filter
    st.sidebar.header("Filter Data")
    
    # Filter Bank/Cabang
    if bank_col and bank_col in df.columns:
        bank_opts = ['Semua'] + sorted(df[bank_col].astype(str).unique().tolist())
        sel_bank = st.sidebar.selectbox(f'Pilih {bank_col}:', bank_opts)
    else:
        sel_bank = 'Semua'

    # Filter Lokasi
    if lok_col and lok_col in df.columns:
        lok_opts = ['Semua'] + sorted(df[lok_col].astype(str).unique().tolist())
        sel_lok = st.sidebar.selectbox(f'Pilih {lok_col}:', lok_opts)
    else:
        sel_lok = 'Semua'

    # Filter Logic
    fil_df = df.copy()
    if sel_bank != 'Semua' and bank_col:
        fil_df = fil_df[fil_df[bank_col].astype(str) == sel_bank]
    if sel_lok != 'Semua' and lok_col:
        fil_df = fil_df[fil_df[lok_col].astype(str) == sel_lok]

    # Visualisasi Sederhana
    st.write(f"Menampilkan {len(fil_df)} data.")
    st.dataframe(fil_df, use_container_width=True)
    
    if tgl_col and jml_col and not fil_df.empty:
        st.subheader("Tren Data")
        trend = fil_df.groupby(tgl_col)[jml_col].sum().reset_index()
        st.plotly_chart(px.line(trend, x=tgl_col, y=jml_col), use_container_width=True)
