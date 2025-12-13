import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
from gspread_pandas import spread
import json
import sys

# --- CONFIGURASI INI WAJIB UNTUK STREAMLIT CLOUD (V33) ---
try:
    # 1. Mengambil kunci rahasia dari Streamlit Secrets (dengan header)
    secrets = st.secrets["gspread_service_account"]
    
    # 2. Kumpulkan semua variabel dari secrets menjadi satu dictionary
    secrets_dict = {
        "type": secrets["type"],
        "project_id": secrets["project_id"],
        "private_key_id": secrets["private_key_id"],
        "client_email": secrets["client_email"],
        "client_id": secrets["client_id"],
        "auth_uri": secrets["auth_uri"],
        "token_uri": secrets["token_uri"],
        "auth_provider_x509_cert_url": secrets["auth_provider_x509_cert_url"],
        "client_x509_cert_url": secrets["client_x509_cert_url"],
        "universe_domain": secrets["universe_domain"],
        # KOREKSI AKHIR: Pastikan line break \n di private_key diproses.
        "private_key": secrets["private_key"].replace('\\n', '\n') 
    }

    # Gunakan dictionary yang sudah dikoreksi
    gc = gspread.service_account_from_dict(secrets_dict)
    
except Exception as e:
    st.error(f"GAGAL KONEKSI G-SHEET: Error konfigurasi kunci rahasia (Secrets). Cek kembali format TOML. Error: {e}")
    sys.exit() 

# ... (Lanjutan skrip dashboard utama V26/V31/V32 dari st.set_page_config sampai akhir)
