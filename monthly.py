import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
import base64
import io
import urllib.parse
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ==========================================
# 0. GLOBAL CONFIGURATION & CONSTANTS
# ==========================================
SPREADSHEET_ID = "1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340"

# ==========================================
# 1. KONFIGURASI HALAMAN UTAMA
# ==========================================
# Fungsi untuk Icon Tab Browser saja (Tetap butuh lokal agar tidak error di set_page_config)
def get_image_base64(filepath):
    try:
        img = Image.open(filepath)
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode()
    except FileNotFoundError:
        return ""

bri_b64 = get_image_base64('logo_BRI.PNG')

st.set_page_config(
    page_title="BRI ATM Performance Review",
    page_icon=Image.open('logo_BRI.PNG') if bri_b64 else "🏦", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 2. INJEKSI CUSTOM CSS BRUTAL & ANIMASI
# ==========================================
st.markdown("""
    <style>
    .block-container { padding-top: 0.2rem !important; padding-bottom: 1rem !important; padding-left: 3rem !important; padding-right: 3rem !important; max-width: 98% !important; }
    [data-testid="stHeader"] { display: none; }
    .main {background-color: #F8F9FA; font-family: 'Segoe UI', Tahoma, sans-serif;}
    h1, h2, h3 {color: #003366; font-weight: 700;}
    
    /* HEADER KORPORAT DENGAN RUNNING TEXT */
    .corporate-header { background-color: #003366; color: white; padding: 5px 16px; height: 55px; border-radius: 4px; display: flex; align-items: center; margin-bottom: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .corporate-header h2 { margin: 0; color: white !important; font-size: 18px; font-weight: bold; letter-spacing: 0.5px; display: flex; align-items: center; white-space: nowrap; z-index: 10; }
    
    /* RUNNING TEXT STYLE (Setiap 1 Menit, Muncul & Hilang Lembut) */
    .ticker-wrapper {
        flex-grow: 1;
        overflow: hidden;
        white-space: nowrap;
        margin-left: 40px; 
        display: flex;
        align-items: center;
        /* Efek Gradient agar teks pudar sebelum menabrak batas */
        -webkit-mask-image: linear-gradient(to right, transparent 0%, black 5%, black 95%, transparent 100%);
        mask-image: linear-gradient(to right, transparent 0%, black 5%, black 95%, transparent 100%);
    }
    .ticker-text {
        display: inline-block;
        font-size: 13px;
        font-weight: 600;
        color: #e2e8f0;
        letter-spacing: 0.5px;
        padding-left: 100%; 
        /* Siklus 60 detik (1 Menit) */
        animation: slideTicker 60s linear infinite; 
    }
    @keyframes slideTicker {
        0% { transform: translateX(0); }
        60% { transform: translateX(-100%); } /* Berjalan normal dan santai selama 36 detik */
        100% { transform: translateX(-100%); } /* Sembunyi selama 24 detik sisanya */
    }
    .blink-dot {
        display: inline-block;
        width: 9px;
        height: 9px;
        background-color: #10b981;
        border-radius: 50%;
        margin-right: 10px;
        box-shadow: 0 0 6px #10b981, 0 0 10px #10b981;
        animation: blink 1s cubic-bezier(0.5, 0, 1, 1) infinite alternate;
        vertical-align: middle;
        margin-bottom: 2px;
    }
    @keyframes blink {
        0% { opacity: 1; transform: scale(1); }
        100% { opacity: 0.4; transform: scale(0.8); }
    }

    div[data-testid="stSelectbox"]:first-of-type { position: absolute; right: 0.3rem; top: 0.5rem; width: 180px; z-index: 999; }
    
    /* FIX DROPDOWN (-10px) */
    .stTabs div[data-testid="stSelectbox"] {
        margin-top: -10px !important; 
        float: right !important;
        width: 140px !important; 
        position: relative !important; 
        z-index: 999 !important;
    }

    .stTabs { margin-top: 0rem !important; }

    .custom-table { width: 100%; border-collapse: collapse; font-family: 'Segoe UI', Tahoma, sans-serif; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); background-color: white; border: 1px solid #ddd; }
    .custom-table th { background-color: #003366; color: white; text-align: center; padding: 7px; font-weight: bold; border: 1px solid #ddd; font-size: 12px; }
    .custom-table td { text-align: center; padding: 7px 8px; border: 1px solid #ddd; color: #333; font-size: 12px; }
    
    details > summary { list-style: none; }
    details summary::-webkit-details-marker { display:none; }

    .table-title { color: #003366; font-weight: bold; margin-bottom: 5px; font-size: 14px; margin-top: 8px !important; }
    .inline-title { color: #003366; font-weight: bold; font-size: 14px; margin-top: 10px; margin-bottom: 0px; }

    .action-box { background-color: white; padding: 10px 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; color: #333; box-shadow: 0 1px 3px rgba(0,0,0,0.05); white-space: pre-wrap; }
    
    .stTabs [data-baseweb="tab-list"] { gap: 5px; margin-bottom: 2px !important; } 
    .stTabs [data-baseweb="tab"] { height: 36px; white-space: pre-wrap; background-color: #e9ecef; border-radius: 4px 4px 0px 0px; gap: 1px; padding-top: 5px; padding-bottom: 5px; padding-left: 15px; padding-right: 15px; color: #003366; font-weight: bold; font-size: 13px; }
    .stTabs [aria-selected="true"] { background-color: #003366; color: white !important; }

    /* CSS KHUSUS MENU HOME KPI CARDS & BANNER (DENGAN EFEK HOVER) */
    .kpi-container { padding: 10px 0px 20px 0px; }
    .kpi-card { background-color: white; border-radius: 6px; padding: 25px 15px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; border-top: 4px solid transparent; transition: all 0.3s ease; }
    .kpi-card:hover { border-top: 4px solid #F37021; transform: translateY(-4px); box-shadow: 0 8px 16px rgba(0,0,0,0.15); }
    .kpi-header { color: #64748b; font-size: 12px; font-weight: 800; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 12px; transition: color 0.3s; }
    .kpi-card:hover .kpi-header { color: #003366; }
    .kpi-value { color: #0f172a; font-size: 38px; font-weight: 800; margin-bottom: 8px; font-family: 'Arial Black', sans-serif; }
    
    .banner-container { background-color: #003366; border-radius: 6px; padding: 40px 50px; margin-top: 5px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 10px rgba(0,0,0,0.15); }
    .banner-left { border-left: 5px solid #F37021; padding-left: 25px; }
    
    /* Efek Hover Pada Teks Banner */
    .banner-title, .banner-subtitle, .banner-text { transition: transform 0.3s ease, color 0.3s ease, text-shadow 0.3s ease; }
    .banner-title { color: #FBBF24; font-size: 32px; font-weight: 800; letter-spacing: 1px; margin-bottom: 15px; display: inline-block; }
    .banner-container:hover .banner-title { transform: translateX(8px); text-shadow: 2px 2px 5px rgba(0,0,0,0.4); }
    
    .banner-subtitle { color: white; font-size: 18px; font-weight: 700; margin-bottom: 25px; letter-spacing: 0.5px; display: inline-block; }
    .banner-container:hover .banner-subtitle { transform: translateX(8px); color: #e2e8f0; }
    
    .banner-text { color: #cbd5e1; font-size: 14px; margin-bottom: 8px; font-weight: 500; display: inline-block; }
    .banner-container:hover .banner-text { transform: translateX(8px); }

    /* Efek Hover Pada Logo Banner */
    .banner-logo img { height: 70px; filter: brightness(0) invert(1); transition: transform 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275); }
    .banner-logo img:hover { transform: scale(1.15); }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. FUNGSI KONEKSI TUNGGAL (EFFICIENT API)
# ==========================================
@st.cache_resource
def get_gspread_client():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Gagal otentikasi Google Sheets: {e}")
        return None

# ==========================================
# 4. FUNGSI PENARIKAN DATA
# ==========================================
@st.cache_data(ttl=86400, show_spinner=False)
def load_data():
    csv_url_master = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=AIMS_Master"
    try:
        df_master = pd.read_csv(csv_url_master)
        expected_cols = ['TANGGAL', 'BULAN', 'BULAN_WEEK', 'TID', 'LOKASI', 'CABANG', 'WAKTU INSERT', 'GANTUNG_KASET', 'KATEGORI', 'JUMLAH_COMPLAIN', 'TIERING_W>3', 'TIERING_M>3', 'WEEK', 'STATUS MRI', 'TYPE MRI']
        for col in expected_cols:
            if col not in df_master.columns: df_master[col] = np.nan
        df_master = df_master[expected_cols]
        df_master = df_master[df_master['KATEGORI'] != 'Cash Out']
        df_master['JUMLAH_COMPLAIN'] = pd.to_numeric(df_master['JUMLAH_COMPLAIN'], errors='coerce').fillna(0).astype(int)
        df_master['TANGGAL'] = pd.to_datetime(df_master['TANGGAL'], errors='coerce')
        df_master = df_master.dropna(subset=['TANGGAL']) 
        df_master['Periode'] = df_master['TANGGAL'].dt.to_period('M') 
        
        csv_url_asset = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Jml_Kelolaan"
        df_asset = pd.read_csv(csv_url_asset).iloc[:, :2] 
        df_asset.columns = ['CABANG', 'TOTAL ATM REAL']
        df_asset['TOTAL ATM REAL'] = pd.to_numeric(df_asset['TOTAL ATM REAL'], errors='coerce').fillna(0).astype(int)
        return pd.merge(df_master, df_asset, on='CABANG', how='left')
    except Exception as e: 
        print(f"Error load_data: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_mri_data():
    sheet_name_encoded = urllib.parse.quote("Problem MRI 2025/26 Harian")
    csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name_encoded}"
    try:
        df = pd.read_csv(csv_url)
        if len(df.columns) >= 14:
            col_mapping = { df.columns[0]: 'TANGGAL', df.columns[1]: 'TID', df.columns[7]: 'SLM', df.columns[8]: 'CPC', df.columns[12]: 'KATEGORI', df.columns[13]: 'LOKASI' }
            df = df.rename(columns=col_mapping)
            df['TID'] = df['TID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
            df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=True, errors='coerce')
            df = df.dropna(subset=['TANGGAL'])
            df['Periode'] = df['TANGGAL'].dt.to_period('M')
            df['Day'] = df['TANGGAL'].dt.day
            df['Week_Group'] = pd.cut(df['Day'], bins=[0, 7, 14, 21, 31], labels=['W1', 'W2', 'W3', 'W4'])
            df['CPC'] = df['CPC'].astype(str).str.replace('(?i)KEJAR ', '', regex=True).str.upper()
            return df
        return pd.DataFrame()
    except Exception: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_slm_visit_data():
    try:
        client = get_gspread_client()
        if not client: return pd.DataFrame()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("SLM Visit Log")
        data = sheet.get_all_values()
        if not data or len(data) < 2: return pd.DataFrame()
            
        df_raw = pd.DataFrame(data).fillna("")
        df_res = pd.DataFrame()
        df_res['TID'] = df_raw[0].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
        df_res['CABANG'] = df_raw[3].astype(str).str.strip()
        df_res['TGL_VISIT'] = df_raw[5].astype(str).str.strip()
        df_res['ACTION'] = df_raw[6].astype(str).str.strip()
        df_res['TGL_VISIT_DT'] = pd.to_datetime(df_res['TGL_VISIT'], errors='coerce') 
        return df_res
    except Exception: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_elastic_followup_data():
    csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Summary_Monitoring_Cash"
    try:
        df_raw = pd.read_csv(csv_url, header=None)
        start_row, start_col = None, None
        for i in range(len(df_raw)):
            row_vals = df_raw.iloc[i].astype(str).values
            for j in range(len(row_vals) - 1):
                if "STATUS" in row_vals[j] and "Pending" in row_vals[j+1]:
                    start_row, start_col = i, j; break
            if start_row is not None: break
        if start_row is not None and start_col is not None:
            return df_raw.iloc[start_row:start_row+5, start_col:start_col+5].copy().fillna("")
        return pd.DataFrame()
    except Exception: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_complain_followup_data():
    csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Summary_Monitoring_Cash"
    try:
        df_raw = pd.read_csv(csv_url, header=None)
        start_row, start_col = None, None
        for i in range(len(df_raw) - 1): 
            row_vals = df_raw.iloc[i].astype(str).values
            next_row_vals = df_raw.iloc[i+1].astype(str).values
            for j in range(len(row_vals)):
                if "STATUS" in row_vals[j] and "Belum Lengkap" in next_row_vals[j]:
                    start_row, start_col = i, j; break
            if start_row is not None: break
        if start_row is not None and start_col is not None:
            return df_raw.iloc[start_row:start_row+4, start_col:start_col+5].copy().fillna("")
        return pd.DataFrame()
    except Exception: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_logistic_data(sheet_name, range_name):
    csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}&range={range_name}"
    try: return pd.read_csv(csv_url).fillna("")
    except: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False)
def load_worksheet_range(worksheet_name, range_name):
    try:
        client = get_gspread_client()
        if not client: return pd.DataFrame()
        data = client.open_by_key(SPREADSHEET_ID).worksheet(worksheet_name).get(range_name)
        if data and len(data) > 1: return pd.DataFrame(data[1:], columns=data[0])
        elif data and len(data) == 1: return pd.DataFrame(columns=data[0])
        return pd.DataFrame()
    except Exception: return pd.DataFrame()

@st.cache_data(ttl=86400, show_spinner=False) 
def load_dashboard_inputs():
    try:
        client = get_gspread_client()
        if not client: return pd.DataFrame()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet("Data_Input_Dashboard")
        if not sheet.row_values(1):
            sheet.append_row(["Timestamp", "Bulan", "Kategori", "Action_Plan", "Qty_1", "Qty_2", "Qty_3", "Qty_4"])
        return pd.DataFrame(sheet.get_all_records())
    except: return pd.DataFrame()

# ==========================================
# 5. FUNGSI LOGIKA (DATA PROCESSING)
# ==========================================
def get_latest_input(df, bulan, kategori):
    if df.empty: return "", 0, 0, 0, 0
    df_filtered = df[(df['Bulan'] == bulan) & (df['Kategori'] == kategori)]
    if df_filtered.empty: return "", 0, 0, 0, 0
    latest = df_filtered.iloc[-1]
    return latest['Action_Plan'], latest['Qty_1'], latest['Qty_2'], latest['Qty_3'], latest['Qty_4']

def calculate_mom(current, prev):
    if prev == 0 and current == 0: return 0.0, "#475569", "-"
    if prev == 0 and current > 0: return 100.0, "#ef4444", "▲"
    mom = ((current - prev) / prev) * 100
    if mom > 0: return mom, "#ef4444", "▲"  # Bad (Red)
    elif mom < 0: return mom, "#10b981", "▼" # Good (Green)
    else: return mom, "#475569", "-"

def get_tiering_mri(df_sub):
    if df_sub.empty: return 0, 0, 0
    vc = df_sub['TID'].value_counts()
    return (vc == 1).sum(), ((vc >= 2) & (vc <= 3)).sum(), (vc > 3).sum()

def get_tiering_general(df_sub):
    if df_sub.empty: return 0, 0, 0, 0
    vc = df_sub['TID'].value_counts()
    return (vc == 1).sum(), ((vc >= 2) & (vc <= 3)).sum(), (vc > 3).sum(), len(vc)

def get_tiering_complain(df_sub):
    if df_sub.empty: return 0, 0, 0, 0
    vc = df_sub.groupby('TID')['JUMLAH_COMPLAIN'].sum()
    return (vc == 1).sum(), ((vc >= 2) & (vc <= 3)).sum(), (vc > 3).sum(), len(vc)

# ==========================================
# 6. FUNGSI MASTER UI (DRY PRINCIPLE)
# ==========================================
def build_category_dashboard(kat_name, df_kat, m1, m2, m3, m3_str, df_followup=pd.DataFrame(), calc_method='freq'):
    lbl = ["PERGANTIAN PART", "ENVIRONMENT", "REPAIR SPAREPART", "DLL"]
    act_plan, q1, q2, q3, q4 = get_latest_input(df_inputs, m3_str, kat_name)
    
    df_m1 = df_kat[df_kat['Periode'] == m1]
    df_m2 = df_kat[df_kat['Periode'] == m2]
    df_m3 = df_kat[df_kat['Periode'] == m3]
    
    if calc_method == 'complain':
        m1_all = df_m1['JUMLAH_COMPLAIN'].sum()
        m2_all = df_m2['JUMLAH_COMPLAIN'].sum()
        m3_all = df_m3['JUMLAH_COMPLAIN'].sum()
        t1_m1, t23_m1, t3_m1, tot_m1 = get_tiering_complain(df_m1)
        t1_m2, t23_m2, t3_m2, tot_m2 = get_tiering_complain(df_m2)
        t1_m3, t23_m3, t3_m3, tot_m3 = get_tiering_complain(df_m3)
        tot_freq_global = m3_all
    else:
        m1_all = len(df_m1)
        m2_all = len(df_m2)
        m3_all = len(df_m3)
        t1_m1, t23_m1, t3_m1, tot_m1 = get_tiering_general(df_m1)
        t1_m2, t23_m2, t3_m2, tot_m2 = get_tiering_general(df_m2)
        t1_m3, t23_m3, t3_m3, tot_m3 = get_tiering_general(df_m3)
        tot_freq_global = len(df_m3)
        
    tot_tid_global = df_m3['TID'].nunique()
    
    mom_all, color_all, arrow_all = calculate_mom(m3_all, m2_all)
    mom_t1, col_t1, arr_t1 = calculate_mom(t1_m3, t1_m2)
    mom_t23, col_t23, arr_t23 = calculate_mom(t23_m3, t23_m2)
    mom_t3, col_t3, arr_t3 = calculate_mom(t3_m3, t3_m2)
    mom_tot, col_tot, arr_tot = calculate_mom(tot_m3, tot_m2)

    col_left, col_right = st.columns(2, gap="medium")
    
    with col_left:
        html_left = f"""
        <div class="table-title">All {kat_name} Overview</div>
        <table class="custom-table">
            <tr><th>TTL Asset</th><th>{m1.strftime('%B')}</th><th>{m2.strftime('%B')}</th><th>{m3.strftime('%B')}</th><th>Δ MoM</th></tr>
            <tr><td style="background-color: #f2f2f2;"><b>543</b></td><td>{m1_all}</td><td>{m2_all}</td><td>{m3_all}</td><td style="color: {color_all}; font-weight: bold;">{abs(mom_all):.2f}% {arrow_all}</td></tr>
        </table>
        <div class="table-title">Tiering Unit {kat_name}</div>
        <table class="custom-table">
            <tr><th>Tiering</th><th>{m1.strftime('%B')}</th><th>{m2.strftime('%B')}</th><th>{m3.strftime('%B')}</th><th>Δ MoM</th></tr>
            <tr><td style="background-color: #f2f2f2;"><b>1 kali</b></td><td>{t1_m1}</td><td>{t1_m2}</td><td>{t1_m3}</td><td style="color: {col_t1}; font-weight: bold;">{abs(mom_t1):.2f}% {arr_t1}</td></tr>
            <tr><td style="background-color: #f2f2f2;"><b>2-3 kali</b></td><td>{t23_m1}</td><td>{t23_m2}</td><td>{t23_m3}</td><td style="color: {col_t23}; font-weight: bold;">{abs(mom_t23):.2f}% {arr_t23}</td></tr>
            <tr><td style="background-color: #f2f2f2;"><b>> 3 kali</b></td><td style="color: #ef4444; font-weight:bold;">{t3_m1}</td><td style="color: #ef4444; font-weight:bold;">{t3_m2}</td><td style="color: #ef4444; font-weight:bold;">{t3_m3}</td><td style="color: {col_t3}; font-weight: bold;">{abs(mom_t3):.2f}% {arr_t3}</td></tr>
            <tr><td style="background-color: #e2e8f0;"><b>Total Unit</b></td><td style="background-color: #e2e8f0;"><b>{tot_m1}</b></td><td style="background-color: #e2e8f0;"><b>{tot_m2}</b></td><td style="background-color: #e2e8f0;"><b>{tot_m3}</b></td><td style="background-color: #e2e8f0; color: {col_tot}; font-weight: bold;">{abs(mom_tot):.2f}% {arr_tot}</td></tr>
        </table>
        """
        st.markdown(html_left, unsafe_allow_html=True)
        
        if not df_followup.empty:
            table_rows_fu = ""
            if calc_method == 'complain': 
                for i in range(1, len(df_followup)):
                    r = df_followup.iloc[i]
                    table_rows_fu += f'<tr><td style="text-align: left; padding-left: 8px;">{r.iloc[0]}</td><td>{r.iloc[1]}</td><td>{r.iloc[2]}</td><td style="color: #003366; font-weight: bold;">{r.iloc[3]}</td><td style="color: #003366; font-weight: bold;">{r.iloc[4]}</td></tr>'
                st.markdown(f'<div class="table-title">{kat_name} Daily Follow-up</div><table class="custom-table"><tr><th>STATUS</th><th>{m2.strftime("%b")}</th><th>{m3.strftime("%b")}</th><th>% {m2.strftime("%b")}</th><th>% {m3.strftime("%b")}</th></tr>{table_rows_fu}</table>', unsafe_allow_html=True)
            else: 
                for i in range(1, len(df_followup)):
                    r = df_followup.iloc[i]
                    table_rows_fu += f'<tr><td style="text-align: left; padding-left: 8px;">{r.iloc[0]}</td><td style="text-align: left; padding-left: 8px;">{r.iloc[1]}</td><td>{r.iloc[2]}</td><td>{r.iloc[3]}</td><td style="color: #003366; font-weight: bold;">{r.iloc[4]}</td></tr>'
                st.markdown(f'<div class="table-title">{kat_name} Daily Follow-up</div><table class="custom-table"><tr><th>STATUS</th><th>Pending</th><th>Done</th><th>Total</th><th>% TL</th></tr>{table_rows_fu}</table>', unsafe_allow_html=True)

        st.markdown(f'<div class="table-title">Action Plan</div><div class="action-box" style="min-height: 120px;">{act_plan if act_plan else "Belum ada Action Plan."}</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown(f"""
        <div class="table-title">Summary Tindaklanjut {kat_name}</div>
        <table class="custom-table" style="width: 100%;">
            <tr><th>QTY</th><th>ACTION</th></tr>
            <tr><td><b>{q1 if q1 else '-'}</b></td><td>{lbl[0]}</td></tr>
            <tr><td><b>{q2 if q2 else '-'}</b></td><td>{lbl[1]}</td></tr>
            <tr><td><b>{q3 if q3 else '-'}</b></td><td>{lbl[2]}</td></tr>
            <tr><td><b>{q4 if q4 else '-'}</b></td><td>{lbl[3]}</td></tr>
        </table>""", unsafe_allow_html=True)
        
        with st.expander(f"✏️ Update Action Plan & Summary {m3_str} - {kat_name}"):
            with st.form(key=f"form_{kat_name}"):
                new_act = st.text_area("Tulis Action Plan Bulan Selanjutnya:", value=str(act_plan))
                st.markdown(f"**Update QTY Summary Tindaklanjut:**")
                ca, cb = st.columns(2)
                new_q1 = ca.number_input(lbl[0], value=int(q1) if pd.notna(q1) and q1 != "" else 0, step=1)
                new_q2 = cb.number_input(lbl[1], value=int(q2) if pd.notna(q2) and q2 != "" else 0, step=1)
                new_q3 = ca.number_input(lbl[2], value=int(q3) if pd.notna(q3) and q3 != "" else 0, step=1)
                new_q4 = cb.number_input(lbl[3], value=int(q4) if pd.notna(q4) and q4 != "" else 0, step=1)
                
                if st.form_submit_button("🔥 Simpan ke Database"):
                    try:
                        client = get_gspread_client()
                        if client:
                            sheet = client.open_by_key(SPREADSHEET_ID).worksheet("Data_Input_Dashboard")
                            sheet.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), m3_str, kat_name, new_act, new_q1, new_q2, new_q3, new_q4])
                            st.success("Data sukses tertanam! Merefresh...")
                            st.cache_data.clear() 
                            st.rerun()
                    except Exception as e: st.error(f"Error tahe: {e}")
        
        branch_data_list = []
        for branch in df_kat['CABANG'].dropna().unique():
            b_df = df_kat[df_kat['CABANG'] == branch]
            total_atm_real = int(b_df['TOTAL ATM REAL'].iloc[0]) if not b_df.empty and not pd.isna(b_df['TOTAL ATM REAL'].iloc[0]) else 0
            
            if calc_method == 'complain':
                m1_b = b_df[b_df['Periode'] == m1]['JUMLAH_COMPLAIN'].sum()
                m2_b = b_df[b_df['Periode'] == m2]['JUMLAH_COMPLAIN'].sum()
                m3_b = b_df[b_df['Periode'] == m3]['JUMLAH_COMPLAIN'].sum()
            else:
                m1_b = len(b_df[b_df['Periode'] == m1])
                m2_b = len(b_df[b_df['Periode'] == m2])
                m3_b = len(b_df[b_df['Periode'] == m3])
                
            mom_val, color, arrow = calculate_mom(m3_b, m2_b)
            branch_data_list.append({"CABANG": branch, "TOTAL ATM REAL": total_atm_real, m1.strftime('%B'): m1_b, m2.strftime('%B'): m2_b, m3.strftime('%B'): m3_b, "MoM": mom_val, "Color": color, "Arrow": arrow})
        
        df_branches_summary = pd.DataFrame(branch_data_list)
        if not df_branches_summary.empty:
            df_top_5 = df_branches_summary.sort_values(by=m3.strftime('%B'), ascending=False).head(5)
            y_limit = max(df_top_5[m2.strftime('%B')].max(), df_top_5[m3.strftime('%B')].max()) * 1.2
            if y_limit == 0: y_limit = 10
            
            st.markdown(f"<div class='table-title'>Grafik {kat_name} Cabang Tertinggi</div>", unsafe_allow_html=True)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_top_5['CABANG'], y=df_top_5[m2.strftime('%B')], mode='lines+text', name=m2.strftime('%B'), text=df_top_5[m2.strftime('%B')], textposition="top center", textfont=dict(color="#3b82f6", size=13, family="Arial Black"), line=dict(color='#3b82f6', width=3), cliponaxis=False))
            fig.add_trace(go.Scatter(x=df_top_5['CABANG'], y=df_top_5[m3.strftime('%B')], mode='lines+text', name=m3.strftime('%B'), text=df_top_5[m3.strftime('%B')], textposition="top center", textfont=dict(color="#ef4444", size=13, family="Arial Black"), line=dict(color='#ef4444', width=3), cliponaxis=False))
            fig.update_layout(title=f"{kat_name} Problem Statistics Overview", title_font=dict(size=12, color='gray'), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(l=10, r=10, t=35, b=15), xaxis=dict(showgrid=True, gridcolor='#eaeaea', tickfont=dict(size=11)), yaxis=dict(showgrid=True, gridcolor='#eaeaea', tickfont=dict(size=11), range=[0, y_limit]), plot_bgcolor='white', height=240)
            st.plotly_chart(fig, use_container_width=True)
            
            table_rows_top5 = ""
            for _, row in df_top_5.iterrows():
                branch_name = row['CABANG']
                df_drill = df_m3[df_m3['CABANG'] == branch_name]
                drill_html = ""
                if not df_drill.empty:
                    if calc_method == 'complain':
                        br_freq = df_drill['JUMLAH_COMPLAIN'].sum()
                        top_tids = df_drill.groupby('TID')['JUMLAH_COMPLAIN'].sum().reset_index().sort_values(by='JUMLAH_COMPLAIN', ascending=False).head(5)
                        top_tids.columns = ['TID', 'Jumlah']
                    else:
                        br_freq = len(df_drill)
                        top_tids = df_drill['TID'].value_counts().reset_index()
                        top_tids.columns = ['TID', 'Jumlah']
                        top_tids = top_tids.head(5)
                        
                    br_tid = df_drill['TID'].nunique()
                    pct_freq = (br_freq / tot_freq_global * 100) if tot_freq_global > 0 else 0
                    pct_tid = (br_tid / tot_tid_global * 100) if tot_tid_global > 0 else 0
                    
                    drill_html += f"<div style='margin-top: 6px; font-weight: normal; font-size: 10px; background-color: #f1f5f9; padding: 5px; border-radius: 4px; border: 1px solid #cbd5e1;'><div style='margin-bottom: 6px; padding-bottom: 4px; border-bottom: 1px solid #94a3b8; color: #334155;'><div style='display: flex; justify-content: space-between;'><span>Freq % Global:</span> <b>{pct_freq:.1f}%</b></div><div style='display: flex; justify-content: space-between;'><span>Unit % Global:</span> <b>{pct_tid:.1f}%</b></div></div>"
                    
                    lokasi_dict = df_drill.set_index('TID')['LOKASI'].to_dict()
                    for _, t_row in top_tids.iterrows():
                        tid = t_row['TID']
                        lok_raw = str(lokasi_dict.get(tid, ''))
                        lok = (lok_raw[:22] + '..') if len(lok_raw) > 22 else lok_raw
                        jml = t_row['Jumlah']
                        drill_html += f"<div style='border-bottom: 1px dashed #cbd5e1; padding: 3px 0; display: flex; justify-content: space-between; align-items: center;'><span style='text-align: left;'><b style='color:#0f172a;'>{tid}</b> <span style='color:#475569;'>{lok}</span></span> <span style='color: #ef4444; font-weight:bold;'>{jml}</span></div>"
                    drill_html += "</div>"
                
                table_rows_top5 += f"""<tr>
                    <td style="vertical-align: top; padding-top: 8px;">{row['TOTAL ATM REAL']}</td>
                    <td style="text-align: left; padding-left: 10px; vertical-align: top; padding-top: 8px; width: 38%;">
                        <details><summary style="cursor: pointer; color: #003366; font-weight: bold; outline: none;">{branch_name} ▾</summary>{drill_html}</details>
                    </td>
                    <td style="vertical-align: top; padding-top: 8px;">{row[m1.strftime('%B')]}</td>
                    <td style="vertical-align: top; padding-top: 8px;">{row[m2.strftime('%B')]}</td>
                    <td style="vertical-align: top; padding-top: 8px;">{row[m3.strftime('%B')]}</td>
                    <td style="color: {row['Color']}; font-weight: bold; vertical-align: top; padding-top: 8px;">{abs(row['MoM']):.1f}% {row['Arrow']}</td>
                </tr>"""
            st.markdown(f'<table class="custom-table"><tr><th style="width: 12%;">TOTAL ATM</th><th style="width: 38%;">CABANG</th><th style="width: 10%;">{m1.strftime("%B")}</th><th style="width: 10%;">{m2.strftime("%B")}</th><th style="width: 10%;">{m3.strftime("%B")}</th><th style="width: 20%;">Δ MoM</th></tr>{table_rows_top5}</table>', unsafe_allow_html=True)

# ==========================================
# 7. LOAD SEMUA DATA
# ==========================================
df_master_enriched = load_data()
df_mri_master = load_mri_data()
df_slm_visit = load_slm_visit_data() 
df_elastic_fu = load_elastic_followup_data()
df_complain_fu = load_complain_followup_data()
df_log_cassette = load_logistic_data("Stock_Cassette", "A1:F10")
df_log_thermal = load_logistic_data("Stock_Thermal", "A1:D10")
df_log_sdm = load_logistic_data("Stock_SDM", "A1:Q10")
df_inputs = load_dashboard_inputs() 

# ==========================================
# 8. RENDERING DASHBOARD UTAMA
# ==========================================
if df_master_enriched.empty:
    st.warning("Data kosong. Tidak ada yang bisa ditampilkan, David.")
else:
    available_periods = sorted(df_master_enriched['Periode'].unique())
    period_strings = [p.strftime('%B %Y') for p in available_periods]
    
    # RENDER HEADER KORPORAT + RUNNING TEXT
    current_date_header = datetime.now().strftime("%A, %d %B %Y").upper()
    url_header_logo = "https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png"
    
    st.markdown(f"""
    <div class="corporate-header">
        <h2>
            <img src="{url_header_logo}" style="height: 28px; filter: brightness(0) invert(1); margin-right: 15px; vertical-align: middle;"> 
            MONTHLY ATM PERFORMANCE REVIEW
        </h2>
        <div class="ticker-wrapper">
            <div class="ticker-text">
                <span class="blink-dot"></span>
                SYSTEM: SECURE & OPTIMAL &nbsp; | &nbsp; DATA LOADED &nbsp; | &nbsp; SERVER TIME: {current_date_header} &nbsp; | &nbsp; BANK BRI MONITORING ACTIVE &nbsp; | &nbsp; <span style="color: #FBBF24; font-weight: 800; letter-spacing: 1px;">PT KELOLA JASA ARTHA</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    selected_period_str = st.selectbox("Pilih Bulan", period_strings, index=len(period_strings)-1, label_visibility="collapsed")

    selected_idx = period_strings.index(selected_period_str)
    m3 = available_periods[selected_idx]
    m2 = available_periods[selected_idx - 1] if selected_idx >= 1 else m3
    m1 = available_periods[selected_idx - 2] if selected_idx >= 2 else m2
    m3_str = m3.strftime('%B %Y')

    tab_home, tab_mri, tab_elastic, tab_complain, tab_df, tab_out, tab_logistic = st.tabs([
        "Home", "⭐ MRI PROJECT", "Elastic", "Complain", "DF Repeat", "OUT Flm", "Logistic"
    ])
    
    # ----------------------------------------
    # TAB HOME 
    # ----------------------------------------
    with tab_home:
        current_date_str = datetime.now().strftime("%A, %d %B %Y")
        
        df_m3_all = df_master_enriched[df_master_enriched['Periode'] == m3]
        df_m2_all = df_master_enriched[df_master_enriched['Periode'] == m2]

        el_m3 = len(df_m3_all[df_m3_all['KATEGORI'] == 'Elastic'])
        el_m2 = len(df_m2_all[df_m2_all['KATEGORI'] == 'Elastic'])
        el_mom, el_col, el_arr = calculate_mom(el_m3, el_m2)

        cp_m3 = df_m3_all[df_m3_all['KATEGORI'] == 'Complain']['JUMLAH_COMPLAIN'].sum()
        cp_m2 = df_m2_all[df_m2_all['KATEGORI'] == 'Complain']['JUMLAH_COMPLAIN'].sum()
        cp_mom, cp_col, cp_arr = calculate_mom(cp_m3, cp_m2)

        df_m3_cnt = len(df_m3_all[df_m3_all['KATEGORI'].str.contains('DF Repeat', case=False, na=False)])
        df_m2_cnt = len(df_m2_all[df_m2_all['KATEGORI'].str.contains('DF Repeat', case=False, na=False)])
        df_mom, df_col, df_arr = calculate_mom(df_m3_cnt, df_m2_cnt)

        out_m3 = len(df_m3_all[df_m3_all['KATEGORI'].str.contains('OUT Flm', case=False, na=False)])
        out_m2 = len(df_m2_all[df_m2_all['KATEGORI'].str.contains('OUT Flm', case=False, na=False)])
        out_mom, out_col, out_arr = calculate_mom(out_m3, out_m2)

        st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-header">ELASTIC PROBLEM</div>
                <div class="kpi-value">{el_m3}</div>
                <div style="color: {el_col}; font-size: 13px; font-weight: 700;">{el_arr} {abs(el_mom):.1f}% vs Prev MTD</div>
            </div>""", unsafe_allow_html=True)
            
        with col2:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-header">COMPLAIN PROBLEM</div>
                <div class="kpi-value">{cp_m3}</div>
                <div style="color: {cp_col}; font-size: 13px; font-weight: 700;">{cp_arr} {abs(cp_mom):.1f}% vs Prev MTD</div>
            </div>""", unsafe_allow_html=True)
            
        with col3:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-header">DF REPEAT PROBLEM</div>
                <div class="kpi-value">{df_m3_cnt}</div>
                <div style="color: {df_col}; font-size: 13px; font-weight: 700;">{df_arr} {abs(df_mom):.1f}% vs Prev MTD</div>
            </div>""", unsafe_allow_html=True)
            
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-header">OUT FLM PROBLEM</div>
                <div class="kpi-value">{out_m3}</div>
                <div style="color: {out_col}; font-size: 13px; font-weight: 700;">{out_arr} {abs(out_mom):.1f}% vs Prev MTD</div>
            </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        url_logo_bri_banner = "https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png"
        
        st.markdown(f"""
        <div class="banner-container">
            <div class="banner-left">
                <div class="banner-title">MONTHLY PERFORMANCE REVIEW</div><br>
                <div class="banner-subtitle">ATM MONITORING DIVISION</div><br>
                <div class="banner-text">Presenter : Command Center BRI</div><br>
                <div class="banner-text">{current_date_str}</div>
            </div>
            <div class="banner-logo">
                <img src="{url_logo_bri_banner}" alt="Logo BRI">
            </div>
        </div>
        """, unsafe_allow_html=True)

    with tab_mri:
        if not df_mri_master.empty:
            df_m3 = df_mri_master[df_mri_master['Periode'] == m3]
            df_m2 = df_mri_master[df_mri_master['Periode'] == m2]
            
            comp_m3 = df_m3[df_m3['KATEGORI'].astype(str).str.contains('Complain', case=False, na=False)]
            comp_m2 = df_m2[df_m2['KATEGORI'].astype(str).str.contains('Complain', case=False, na=False)]
            df_rep_m3 = df_m3[df_m3['KATEGORI'].astype(str).str.contains('Df Repeat', case=False, na=False)]
            df_rep_m2 = df_m2[df_m2['KATEGORI'].astype(str).str.contains('Df Repeat', case=False, na=False)]
            
            col_mri_left, col_mri_right = st.columns(2, gap="medium")
            
            with col_mri_left:
                st.markdown(f"""
                <div class="table-title">Summary Problem TID MRI</div>
                <table class="custom-table" style="width:60%;"><tr><th>Σ ATM</th><th>Complain</th><th>DF Repeat</th></tr><tr><td style="background-color: #f2f2f2; font-weight: bold;">34</td><td>{len(comp_m3)}</td><td>{len(df_rep_m3)}</td></tr></table>
                """, unsafe_allow_html=True)
                
                c_w1 = comp_m3[comp_m3['Week_Group'] == 'W1']
                c_w2 = comp_m3[comp_m3['Week_Group'] == 'W2']
                c_w3 = comp_m3[comp_m3['Week_Group'] == 'W3']
                c_w4 = comp_m3[comp_m3['Week_Group'] == 'W4']
                
                st.markdown(f"""
                <div class="table-title">Jumlah Complain</div>
                <table class="custom-table"><tr><th>CRO</th><th>Category</th><th>TOTAL ATM</th><th>{m2.strftime('%b')}(prev)</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {m3.strftime('%b')}</th></tr><tr><td>KEJAR</td><td>MRI</td><td>34</td><td>{len(comp_m2)}</td><td>{len(c_w1) if len(c_w1) else ''}</td><td>{len(c_w2) if len(c_w2) else ''}</td><td>{len(c_w3) if len(c_w3) else ''}</td><td>{len(c_w4) if len(c_w4) else ''}</td><td style="font-weight:bold; color:#ef4444;">{len(comp_m3)}</td></tr></table>
                """, unsafe_allow_html=True)
                
                cp1, cp23, cp3 = get_tiering_mri(comp_m2)
                cw1_1, cw1_23, cw1_3 = get_tiering_mri(c_w1)
                cw2_1, cw2_23, cw2_3 = get_tiering_mri(c_w2)
                cw3_1, cw3_23, cw3_3 = get_tiering_mri(c_w3)
                cw4_1, cw4_23, cw4_3 = get_tiering_mri(c_w4)
                ct1, ct23, ct3 = get_tiering_mri(comp_m3)
                
                st.markdown(f"""
                <div class="table-title">Tiering Complain by TID</div>
                <table class="custom-table">
                    <tr><th>CRO</th><th>Category</th><th>TIERING</th><th>{m2.strftime('%b')}(prev)</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {m3.strftime('%b')}</th></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>1 kali</td><td>{cp1 if cp1 else ''}</td><td>{cw1_1 if cw1_1 else ''}</td><td>{cw2_1 if cw2_1 else ''}</td><td>{cw3_1 if cw3_1 else ''}</td><td>{cw4_1 if cw4_1 else ''}</td><td style="font-weight:bold;">{ct1 if ct1 else ''}</td></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>2-3 kali</td><td>{cp23 if cp23 else ''}</td><td>{cw1_23 if cw1_23 else ''}</td><td>{cw2_23 if cw2_23 else ''}</td><td>{cw3_23 if cw3_23 else ''}</td><td>{cw4_23 if cw4_23 else ''}</td><td style="font-weight:bold;">{ct23 if ct23 else ''}</td></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>> 3 kali</td><td>{cp3 if cp3 else ''}</td><td>{cw1_3 if cw1_3 else ''}</td><td>{cw2_3 if cw2_3 else ''}</td><td>{cw3_3 if cw3_3 else ''}</td><td>{cw4_3 if cw4_3 else ''}</td><td style="font-weight:bold; color:#ef4444;">{ct3 if ct3 else ''}</td></tr>
                </table>
                """, unsafe_allow_html=True)
                
                c1, c2 = st.columns([4, 1])
                with c1: st.markdown("<div class='inline-title'>Top Complain Problem Terminal IDs</div>", unsafe_allow_html=True)
                with c2: sort_comp = st.selectbox("SORTby :", [f"Σ {m3.strftime('%b')}", "W1", "W2", "W3", "W4"], key="sort_comp", label_visibility="collapsed")
                
                base_comp = comp_m3 if sort_comp == f"Σ {m3.strftime('%b')}" else c_w1 if sort_comp == "W1" else c_w2 if sort_comp == "W2" else c_w3 if sort_comp == "W3" else c_w4
                
                top_tids_comp = base_comp['TID'].value_counts().head(5).index.tolist()
                html_top_comp = "<table class='custom-table'><tr><th>NO</th><th>TID</th><th>Location</th><th>Branch</th><th>SLM</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th></tr>"
                if not top_tids_comp: html_top_comp += "<tr><td colspan='9'>Data tidak ditemukan untuk periode ini.</td></tr>"
                else:
                    for i, tid in enumerate(top_tids_comp):
                        tid_df = comp_m3[comp_m3['TID'] == tid]
                        loc = str(tid_df['LOKASI'].iloc[0])[:18] + ".." if len(str(tid_df['LOKASI'].iloc[0])) > 18 else str(tid_df['LOKASI'].iloc[0])
                        br, slm = tid_df['CPC'].iloc[0], tid_df['SLM'].iloc[0]
                        w1_cnt, w2_cnt, w3_cnt, w4_cnt = len(tid_df[tid_df['Week_Group'] == 'W1']), len(tid_df[tid_df['Week_Group'] == 'W2']), len(tid_df[tid_df['Week_Group'] == 'W3']), len(tid_df[tid_df['Week_Group'] == 'W4'])
                        
                        slm_history_comp = "<div style='margin-top: 6px; font-weight: normal; font-size: 10px; background-color: #f1f5f9; padding: 6px; border-radius: 4px; border: 1px solid #cbd5e1; width: 220px;'>"
                        prob_dates_str = ", ".join(tid_df['TANGGAL'].dropna().dt.strftime('%d/%m/%Y').unique().tolist()) if not tid_df['TANGGAL'].empty else "-"
                        slm_history_comp += f"<div style='color: #ef4444; font-weight: bold; margin-bottom: 5px; font-size: 11px;'>Tgl Prob: {prob_dates_str}</div>"
                        
                        if not df_slm_visit.empty:
                            df_tid_slm_comp = df_slm_visit[df_slm_visit['TID'] == tid].sort_values(by='TGL_VISIT_DT', ascending=False)
                            if not df_tid_slm_comp.empty:
                                for _, r in df_tid_slm_comp.iterrows():
                                    tgl_visit = r['TGL_VISIT_DT'].strftime('%d/%m/%Y') if pd.notna(r['TGL_VISIT_DT']) else str(r['TGL_VISIT']).strip()[:10]
                                    act = str(r['ACTION']).strip()
                                    if not act or act.lower() == 'nan': act = 'Tidak ada deskripsi'
                                    slm_history_comp += f"<div style='border-top: 1px dashed #cbd5e1; padding: 4px 0; display: flex; flex-direction: column; align-items: flex-start;'><span style='color:#003366; font-weight:bold;'>Visit: {tgl_visit}</span><span style='color:#475569; text-align: left; line-height: 1.2; margin-top:2px;'>{act}</span></div>"
                            else: slm_history_comp += "<div style='color:#94a3b8; padding-top: 4px; border-top: 1px dashed #cbd5e1;'>Belum ada history visit SLM.</div>"
                        else: slm_history_comp += "<div style='color:#94a3b8; padding-top: 4px; border-top: 1px dashed #cbd5e1;'>Data SLM kosong.</div>"

                        tid_cell_comp = f"<td style='text-align: left; vertical-align: top; min-width: 140px;'><details><summary style='cursor: pointer; color: #003366; font-weight: bold; outline: none;'>{tid} ▾</summary>{slm_history_comp}</details></td>"
                        html_top_comp += f"<tr><td style='vertical-align: top;'>{i+1}</td>{tid_cell_comp}<td style='text-align:left; vertical-align: top;'>{loc}</td><td style='vertical-align: top;'>{br}</td><td style='vertical-align: top;'>{slm}</td><td style='vertical-align: top;'>{w1_cnt if w1_cnt else ''}</td><td style='vertical-align: top;'>{w2_cnt if w2_cnt else ''}</td><td style='vertical-align: top;'>{w3_cnt if w3_cnt else ''}</td><td style='vertical-align: top;'>{w4_cnt if w4_cnt else ''}</td></tr>"
                st.markdown(html_top_comp + "</table>", unsafe_allow_html=True)

            with col_mri_right:
                st.markdown(f"""
                <div class="table-title">Summary Pengisian Data MRI</div>
                <table class="custom-table" style="width:80%;"><tr><th>Σ ATM</th><th>Pagi</th><th>Siang</th><th>Malam</th></tr><tr><td style="background-color: #f2f2f2; font-weight: bold;">34</td><td></td><td></td><td></td></tr></table>
                """, unsafe_allow_html=True)
                
                d_w1 = df_rep_m3[df_rep_m3['Week_Group'] == 'W1']
                d_w2 = df_rep_m3[df_rep_m3['Week_Group'] == 'W2']
                d_w3 = df_rep_m3[df_rep_m3['Week_Group'] == 'W3']
                d_w4 = df_rep_m3[df_rep_m3['Week_Group'] == 'W4']
                
                st.markdown(f"""
                <div class="table-title">Jumlah DF</div>
                <table class="custom-table"><tr><th>CRO</th><th>Category</th><th>TOTAL ATM</th><th>{m2.strftime('%b')}(prev)</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {m3.strftime('%b')}</th></tr><tr><td>KEJAR</td><td>MRI</td><td>34</td><td>{len(df_rep_m2)}</td><td>{len(d_w1) if len(d_w1) else ''}</td><td>{len(d_w2) if len(d_w2) else ''}</td><td>{len(d_w3) if len(d_w3) else ''}</td><td>{len(d_w4) if len(d_w4) else ''}</td><td style="font-weight:bold; color:#ef4444;">{len(df_rep_m3)}</td></tr></table>
                """, unsafe_allow_html=True)
                
                dp1, dp23, dp3 = get_tiering_mri(df_rep_m2)
                dw1_1, dw1_23, dw1_3 = get_tiering_mri(d_w1)
                dw2_1, dw2_23, dw2_3 = get_tiering_mri(d_w2)
                dw3_1, dw3_23, dw3_3 = get_tiering_mri(d_w3)
                dw4_1, dw4_23, dw4_3 = get_tiering_mri(d_w4)
                dt1, dt23, dt3 = get_tiering_mri(df_rep_m3)
                
                st.markdown(f"""
                <div class="table-title">Tiering DF Repeat by TID</div>
                <table class="custom-table">
                    <tr><th>CRO</th><th>Category</th><th>TIERING</th><th>{m2.strftime('%b')}(prev)</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {m3.strftime('%b')}</th></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>1 kali</td><td>{dp1 if dp1 else ''}</td><td>{dw1_1 if dw1_1 else ''}</td><td>{dw2_1 if dw2_1 else ''}</td><td>{dw3_1 if dw3_1 else ''}</td><td>{dw4_1 if dw4_1 else ''}</td><td style="font-weight:bold;">{dt1 if dt1 else ''}</td></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>2-3 kali</td><td>{dp23 if dp23 else ''}</td><td>{dw1_23 if dw1_23 else ''}</td><td>{dw2_23 if dw2_23 else ''}</td><td>{dw3_23 if dw3_23 else ''}</td><td>{dw4_23 if dw4_23 else ''}</td><td style="font-weight:bold;">{dt23 if dt23 else ''}</td></tr>
                    <tr><td>KEJAR</td><td>MRI</td><td>> 3 kali</td><td>{dp3 if dp3 else ''}</td><td>{dw1_3 if dw1_3 else ''}</td><td>{dw2_3 if dw2_3 else ''}</td><td>{dw3_3 if dw3_3 else ''}</td><td>{dw4_3 if dw4_3 else ''}</td><td style="font-weight:bold; color:#ef4444;">{dt3 if dt3 else ''}</td></tr>
                </table>
                """, unsafe_allow_html=True)
                
                d1, d2 = st.columns([4, 1])
                with d1: st.markdown("<div class='inline-title'>Top DF Problem Terminal IDs</div>", unsafe_allow_html=True)
                with d2: sort_df = st.selectbox("SORTby :", [f"Σ {m3.strftime('%b')}", "W1", "W2", "W3", "W4"], key="sort_df", label_visibility="collapsed")
                
                base_df = df_rep_m3 if sort_df == f"Σ {m3.strftime('%b')}" else d_w1 if sort_df == "W1" else d_w2 if sort_df == "W2" else d_w3 if sort_df == "W3" else d_w4
                
                top_tids_df = base_df['TID'].value_counts().head(5).index.tolist()
                html_top_df = "<table class='custom-table'><tr><th>NO</th><th>TID</th><th>Location</th><th>Branch</th><th>SLM</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th></tr>"
                if not top_tids_df: html_top_df += "<tr><td colspan='9'>Data tidak ditemukan untuk periode ini.</td></tr>"
                else:
                    for i, tid in enumerate(top_tids_df):
                        tid_df = df_rep_m3[df_rep_m3['TID'] == tid]
                        loc = str(tid_df['LOKASI'].iloc[0])[:18] + ".." if len(str(tid_df['LOKASI'].iloc[0])) > 18 else str(tid_df['LOKASI'].iloc[0])
                        br, slm = tid_df['CPC'].iloc[0], tid_df['SLM'].iloc[0]
                        w1_cnt, w2_cnt, w3_cnt, w4_cnt = len(tid_df[tid_df['Week_Group'] == 'W1']), len(tid_df[tid_df['Week_Group'] == 'W2']), len(tid_df[tid_df['Week_Group'] == 'W3']), len(tid_df[tid_df['Week_Group'] == 'W4'])

                        slm_history_df = "<div style='margin-top: 6px; font-weight: normal; font-size: 10px; background-color: #f1f5f9; padding: 6px; border-radius: 4px; border: 1px solid #cbd5e1; width: 220px;'>"
                        prob_dates_str = ", ".join(tid_df['TANGGAL'].dropna().dt.strftime('%d/%m/%Y').unique().tolist()) if not tid_df['TANGGAL'].empty else "-"
                        slm_history_df += f"<div style='color: #ef4444; font-weight: bold; margin-bottom: 5px; font-size: 11px;'>Tgl Prob: {prob_dates_str}</div>"
                        
                        if not df_slm_visit.empty:
                            df_tid_slm_df = df_slm_visit[df_slm_visit['TID'] == tid].sort_values(by='TGL_VISIT_DT', ascending=False)
                            if not df_tid_slm_df.empty:
                                for _, r in df_tid_slm_df.iterrows():
                                    tgl_visit = r['TGL_VISIT_DT'].strftime('%d/%m/%Y') if pd.notna(r['TGL_VISIT_DT']) else str(r['TGL_VISIT']).strip()[:10]
                                    act = str(r['ACTION']).strip()
                                    if not act or act.lower() == 'nan': act = 'Tidak ada deskripsi'
                                    slm_history_df += f"<div style='border-top: 1px dashed #cbd5e1; padding: 4px 0; display: flex; flex-direction: column; align-items: flex-start;'><span style='color:#003366; font-weight:bold;'>Visit: {tgl_visit}</span><span style='color:#475569; text-align: left; line-height: 1.2; margin-top:2px;'>{act}</span></div>"
                            else: slm_history_df += "<div style='color:#94a3b8; padding-top: 4px; border-top: 1px dashed #cbd5e1;'>Belum ada history visit SLM.</div>"
                        else: slm_history_df += "<div style='color:#94a3b8; padding-top: 4px; border-top: 1px dashed #cbd5e1;'>Data SLM kosong.</div>"

                        tid_cell_df = f"<td style='text-align: left; vertical-align: top; min-width: 140px;'><details><summary style='cursor: pointer; color: #003366; font-weight: bold; outline: none;'>{tid} ▾</summary>{slm_history_df}</details></td>"
                        html_top_df += f"<tr><td style='vertical-align: top;'>{i+1}</td>{tid_cell_df}<td style='text-align:left; vertical-align: top;'>{loc}</td><td style='vertical-align: top;'>{br}</td><td style='vertical-align: top;'>{slm}</td><td style='vertical-align: top;'>{w1_cnt if w1_cnt else ''}</td><td style='vertical-align: top;'>{w2_cnt if w2_cnt else ''}</td><td style='vertical-align: top;'>{w3_cnt if w3_cnt else ''}</td><td style='vertical-align: top;'>{w4_cnt if w4_cnt else ''}</td></tr>"
                st.markdown(html_top_df + "</table>", unsafe_allow_html=True)
        else: st.warning("Data MRI Project kosong atau gagal ditarik.")

    # ==========================================
    # TAB ELASTIC, COMPLAIN, DF REPEAT, OUT FLM (IMPLEMENTASI DRY CODE)
    # ==========================================
    with tab_elastic:
        df_kat = df_master_enriched[df_master_enriched['KATEGORI'] == 'Elastic']
        build_category_dashboard("Elastic", df_kat, m1, m2, m3, m3_str, df_followup=df_elastic_fu, calc_method='freq')

    with tab_complain:
        df_kat = df_master_enriched[df_master_enriched['KATEGORI'] == 'Complain']
        build_category_dashboard("Complain", df_kat, m1, m2, m3, m3_str, df_followup=df_complain_fu, calc_method='complain')

    with tab_df:
        df_kat = df_master_enriched[df_master_enriched['KATEGORI'].str.contains('DF Repeat', case=False, na=False)]
        build_category_dashboard("DF Repeat", df_kat, m1, m2, m3, m3_str, calc_method='freq')

    with tab_out:
        df_kat = df_master_enriched[df_master_enriched['KATEGORI'].str.contains('OUT Flm', case=False, na=False)]
        build_category_dashboard("OUT Flm", df_kat, m1, m2, m3, m3_str, calc_method='freq')

    # ==========================================
    # TAB LOGISTIC
    # ==========================================
    with tab_logistic:
        st.markdown("<div class='table-title' style='margin-bottom: 15px;'>Manajemen Logistik & Resource SDM</div>", unsafe_allow_html=True)
        sub_cassete, sub_thermal, sub_sparepart, sub_sdm, sub_pm = st.tabs(["Stock Cassete", "Stock Thermal", "Stock Sparepart", "Jumlah Teknisi atau SDM", "Preventive Maintenance (PM)"])
        
        with sub_cassete:
            if not df_log_cassette.empty:
                html_cassette = "<div class='table-title'>Data Stock Cassette per Cabang</div><table class='custom-table' style='width: auto; min-width: 60%;'><tr>"
                for col in df_log_cassette.columns: html_cassette += f"<th>{col}</th>"
                html_cassette += "</tr>"
                for _, row in df_log_cassette.iterrows():
                    html_cassette += "<tr>"
                    for i, val in enumerate(row): html_cassette += f"<td style='text-align: left; padding-left: 10px;'><b>{val}</b></td>" if i == 0 else f"<td>{val}</td>"
                    html_cassette += "</tr>"
                st.markdown(html_cassette + "</table>", unsafe_allow_html=True)
            else: st.warning("Data Stock Cassette masih kosong atau gagal ditarik.")
            
        with sub_thermal:
            if not df_log_thermal.empty:
                html_thermal = "<div class='table-title'>Data Stock Thermal per Cabang</div><table class='custom-table' style='width: auto; min-width: 50%;'><tr>"
                for col in df_log_thermal.columns: html_thermal += f"<th>{col}</th>"
                html_thermal += "</tr>"
                for _, row in df_log_thermal.iterrows():
                    html_thermal += "<tr>"
                    for i, val in enumerate(row): html_thermal += f"<td style='text-align: left; padding-left: 10px;'><b>{val}</b></td>" if i == 0 else f"<td>{val}</td>"
                    html_thermal += "</tr>"
                st.markdown(html_thermal + "</table>", unsafe_allow_html=True)
            else: st.warning("Data Stock Thermal masih kosong atau gagal ditarik.")
            
        with sub_sparepart: 
            st.markdown("<div class='table-title'>Inventory Sparepart Monitoring</div>", unsafe_allow_html=True)
            m_hys, m_win, m_ncr = st.tabs(["MESIN HYOSUNG", "MESIN WINCOR", "MESIN NCR"])
            
            for tab_obj, r_name in zip([m_hys, m_win, m_ncr], ["A1:W10", "A12:W21", "A23:Q32"]):
                with tab_obj:
                    df_sp = load_worksheet_range("Stock_Sparepart", r_name)
                    if not df_sp.empty:
                        html_sp = "<table class='custom-table'><tr>" + "".join([f"<th>{c}</th>" for c in df_sp.columns]) + "</tr>"
                        for _, r in df_sp.iterrows(): html_sp += "<tr>" + "".join([f"<td>{v}</td>" for v in r]) + "</tr>"
                        st.markdown(html_sp + "</table>", unsafe_allow_html=True)
                    else: st.warning(f"Data Kosong / Gagal ditarik dari Range {r_name}.")

        with sub_sdm: 
            if not df_log_sdm.empty:
                html_sdm = "<div class='table-title'>Data Resource SDM / Teknisi</div><table class='custom-table' style='width: auto; min-width: 60%;'><tr>"
                for col in df_log_sdm.columns: html_sdm += f"<th>{col}</th>"
                html_sdm += "</tr>"
                for _, row in df_log_sdm.iterrows():
                    html_sdm += "<tr>"
                    for i, val in enumerate(row): html_sdm += f"<td style='text-align: left; padding-left: 10px;'><b>{val}</b></td>" if i == 0 else f"<td>{val}</td>"
                    html_sdm += "</tr>"
                st.markdown(html_sdm + "</table>", unsafe_allow_html=True)
            else: st.warning("Data SDM masih kosong atau gagal ditarik.")
            
        with sub_pm: 
            st.markdown("<div class='table-title'>Data Preventive Maintenance</div>", unsafe_allow_html=True)
            pm_mesin, pm_cassette = st.tabs(["PM Mesin", "PM Cassette"])
            
            for tab_obj, r_name, lbl in zip([pm_mesin, pm_cassette], ["A1:E3", "A6:E9"], ["Mesin", "Cassette"]):
                with tab_obj:
                    df_pm = load_worksheet_range("Preventive_Maintenance", r_name)
                    if not df_pm.empty:
                        html_pm = "<table class='custom-table' style='width: auto; min-width: 50%;'><tr>" + "".join([f"<th>{c}</th>" for c in df_pm.columns]) + "</tr>"
                        for _, r in df_pm.iterrows(): html_pm += "<tr>" + "".join([f"<td>{v if pd.notna(v) else ''}</td>" for v in r]) + "</tr>"
                        st.markdown(html_pm + "</table>", unsafe_allow_html=True)
                    else: st.warning(f"Data PM {lbl} Kosong / Gagal ditarik.")