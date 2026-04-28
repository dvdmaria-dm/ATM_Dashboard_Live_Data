import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os

# ==========================================
# 1. KONFIGURASI PAGE & CSS ELEGANT TABS
# ==========================================
st.set_page_config(page_title="ATM Performance Monitoring", layout="wide", initial_sidebar_state="collapsed")

@st.cache_resource
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    except Exception:
        creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    return gspread.authorize(creds)

client = get_gspread_client()
SHEET_ID = "1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340"

st.markdown("""
    <style>
        .block-container {
            padding-top: 1.5rem !important; 
            padding-bottom: 0.5rem !important;
            padding-left: 2rem !important; 
            padding-right: 2rem !important; 
            max-width: 100% !important;
            background-color: #F8FAFC; 
        }
        header {visibility: hidden;}
        footer {visibility: hidden;}
        
        .table-scroll {
            max-height: 155px; 
            overflow-y: auto;
            overflow-x: auto;
            border-radius: 6px;
            border: 1px solid #E2E8F0;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05); 
            margin-bottom: 12px;
            background-color: #FFFFFF;
            -ms-overflow-style: none;  
            scrollbar-width: none;  
        }
        .table-scroll::-webkit-scrollbar { display: none; }
        
        .dash-table {
            width: 100%;
            border-collapse: separate; border-spacing: 0;
            font-family: 'Inter', 'Segoe UI', sans-serif;
            font-size: 10.5px; text-align: center;
        }
        .dash-table th {
            background-color: #00529C; color: #FFFFFF;
            padding: 5px 4px; font-weight: 600; text-transform: uppercase;
            font-size: 9px; border-bottom: 2px solid #003A70;
            position: sticky; top: 0; z-index: 10;
        }
        .dash-table td { padding: 4px 3px; border-bottom: 1px solid #F1F5F9; color: #334155; }
        .dash-table tr:last-child td { border-bottom: none; }
        .dash-table tr:nth-child(even) { background-color: #F8FAFC; }
        .dash-table tr:hover { background-color: #E2E8F0; }

        .log-table {
            width: 100%; border-collapse: separate; border-spacing: 0;
            font-family: 'Inter', 'Segoe UI', sans-serif;
            font-size: 10.5px; text-align: center; table-layout: fixed; 
        }
        .log-table th {
            background-color: #00529C; color: #FFFFFF;
            padding: 5px 4px; font-weight: 600; text-transform: uppercase;
            font-size: 9px; border-bottom: 2px solid #003A70;
            position: sticky; top: 0; z-index: 10;
            word-wrap: break-word; word-break: break-word; white-space: normal; vertical-align: middle;
        }
        .log-table td {
            padding: 4px 3px; border-bottom: 1px solid #F1F5F9; color: #334155;
            word-wrap: break-word; word-break: break-word; white-space: normal; vertical-align: middle;
        }
        .log-table tr:last-child td { border-bottom: none; }
        .log-table tr:nth-child(even) { background-color: #F8FAFC; }
        .log-table tr:hover { background-color: #E2E8F0; }
        
        .section-title { font-size: 12px; font-weight: 700; color: #00529C; margin-bottom: 4px; margin-top: 0px; border-left: 4px solid #F37021; padding-left: 6px; }

        div[data-baseweb="select"] > div {
            font-size: 13px !important; font-weight: 600 !important; padding-top: 4px !important;
            padding-bottom: 4px !important; min-height: 32px !important; border-radius: 6px !important; color: #00529C !important;
        }
        
        div[data-testid="stRadio"] > div[role="radiogroup"], 
        div.row-widget.stRadio > div[role="radiogroup"] { 
            display: flex !important; flex-direction: row !important; gap: 4px !important; 
            align-items: flex-end !important; border-bottom: 2px solid #E2E8F0 !important; padding-bottom: 0px !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label > div:first-child,
        div.row-widget.stRadio > div[role="radiogroup"] label > div:first-child { display: none !important; }

        div[data-testid="stRadio"] > div[role="radiogroup"] label,
        div.row-widget.stRadio > div[role="radiogroup"] label {
            background-color: #F1F5F9 !important; padding: 10px 22px !important; border-radius: 8px 8px 0px 0px !important; 
            cursor: pointer !important; border: none !important; border-bottom: 4px solid transparent !important;
            transition: all 0.2s ease-in-out !important; margin-bottom: 0px !important; box-shadow: none !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label p,
        div.row-widget.stRadio > div[role="radiogroup"] label p { color: #00529C !important; font-weight: 600 !important; font-size: 12px !important; margin: 0px !important; }

        div[data-testid="stRadio"] > div[role="radiogroup"] label:hover,
        div.row-widget.stRadio > div[role="radiogroup"] label:hover { background-color: #E2E8F0 !important; }
        div[data-testid="stRadio"] > div[role="radiogroup"] label:hover p,
        div.row-widget.stRadio > div[role="radiogroup"] label:hover p { color: #003A70 !important; }

        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked),
        div[data-testid="stRadio"] > div[role="radiogroup"] label[data-checked="true"] {
            background-color: #00529C !important; border-bottom: 4px solid #F37021 !important; box-shadow: 0 -2px 5px rgba(0,0,0,0.1) !important;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) p,
        div[data-testid="stRadio"] > div[role="radiogroup"] label[data-checked="true"] p { color: #FFFFFF !important; }
        
        .stTextArea textarea { font-size: 11px !important; border-left: 3px solid #00529C !important; border-radius: 4px !important; padding: 5px !important; }
        .main-header {
            background: linear-gradient(90deg, #00529C 0%, #003A70 100%); padding: 8px 15px; border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; margin-top: -10px; 
        }
        
        div[data-testid="stFormSubmitButton"] button {
            background-color: transparent !important; color: #CBD5E1 !important; font-weight: 600 !important;
            border: none !important; box-shadow: none !important; float: right; margin-top: -18px;
            padding: 0px 5px !important; font-size: 10px !important; transition: color 0.3s ease; z-index: 99;
        }
        div[data-testid="stFormSubmitButton"] button:hover { color: #F37021 !important; background-color: transparent !important; }
        
        /* Ticker Natural 120s */
        .info-ticker-container {
            flex-grow: 1; overflow: hidden; white-space: nowrap; max-width: 60%; margin-left: 20px;
            -webkit-mask-image: linear-gradient(to right, transparent 0%, black 5%, black 95%, transparent 100%);
            mask-image: linear-gradient(to right, transparent 0%, black 5%, black 95%, transparent 100%);
        }
        .info-ticker-text {
            display: inline-block;
            color: #F8FAFC; font-weight: 500; font-size: 11.5px; opacity: 0.95; letter-spacing: 0.5px;
            animation: ticker-slide 120s linear infinite;
        }
        @keyframes ticker-slide {
            0% { transform: translateX(100vw); }
            50% { transform: translateX(-100vw); } 
            100% { transform: translateX(-100vw); } 
        }

        .home-card {
            background-color: #FFFFFF; border-radius: 8px; padding: 18px 15px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            border-top: 4px solid #00529C; transition: transform 0.2s, box-shadow 0.2s; text-align: center;
        }
        .home-card:hover { transform: translateY(-5px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); border-top: 4px solid #F37021; }
        .card-title { font-size: 13px; font-weight: 700; color: #64748B; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }
        .card-value { font-size: 32px; font-weight: 800; color: #003A70; margin-bottom: 5px; line-height: 1.2; }
        .card-delta-down { font-size: 11.5px; font-weight: 600; color: #10B981; } 
        .card-delta-up { font-size: 11.5px; font-weight: 600; color: #EF4444; } 
        .card-delta-neutral { font-size: 11.5px; font-weight: 600; color: #94A3B8; }
        
        .hero-section {
            background-color: #002244; border-radius: 10px; padding: 35px 45px; margin-top: 25px;
            color: white; background-image: linear-gradient(135deg, #00529C 0%, #001f3f 100%);
            box-shadow: 0 10px 20px -5px rgba(0, 0, 0, 0.25); 
            display: flex; align-items: center; justify-content: space-between; flex-direction: row;
        }
        .hero-left {
            display: flex; flex-direction: column; align-items: flex-start; text-align: left;
            border-left: 5px solid #F37021; padding-left: 25px; 
        }
        .hero-right {
            display: flex; align-items: center; justify-content: center; padding-right: 10px;
        }
        .hero-title { font-size: 34px; font-weight: 800; color: #FBBF24; letter-spacing: 1px; margin-bottom: 8px; text-transform: uppercase; text-shadow: 1px 1px 3px rgba(0,0,0,0.3); line-height: 1.1; }
        .hero-subtitle { font-size: 16px; font-weight: 700; margin-bottom: 20px; color: #FFFFFF; letter-spacing: 0.5px; text-transform: uppercase; opacity: 0.9; }
        .hero-text { font-size: 14px; font-weight: 500; color: #94A3B8; margin-bottom: 4px; }
        
        .bri-logo-watermark {
            width: 240px;
            opacity: 0.9; 
            filter: brightness(0) invert(1) drop-shadow(0px 4px 6px rgba(0,0,0,0.3)); 
            transition: all 0.3s ease;
        }
        .hero-section:hover .bri-logo-watermark {
            transform: scale(1.03); opacity: 1;
        }

        @keyframes heartbeat {
            0% { opacity: 1; transform: scale(1); box-shadow: 0 0 8px #22C55E; }
            50% { opacity: 0.4; transform: scale(0.85); box-shadow: 0 0 2px #22C55E; }
            100% { opacity: 1; transform: scale(1); box-shadow: 0 0 8px #22C55E; }
        }
        .status-dot {
            height: 10px;
            width: 10px;
            background-color: #22C55E;
            border-radius: 50%;
            display: inline-block;
            margin-right: 8px;
            animation: heartbeat 2s infinite ease-in-out;
            vertical-align: middle;
            margin-bottom: 2px;
        }

        /* RESPONSIVE MOBILE OPTIMIZATION */
        @media screen and (max-width: 768px) {
            .main-header { flex-direction: column; align-items: flex-start; gap: 8px; height: auto; padding-bottom: 12px; }
            .info-ticker-container { max-width: 100%; margin-left: 0; width: 100%; }
            
            .hero-section { flex-direction: column; padding: 20px; text-align: center; margin-top: 15px; }
            .hero-left { border-left: none; border-bottom: 3px solid #F37021; padding-left: 0; padding-bottom: 15px; margin-bottom: 15px; align-items: center; width: 100%; }
            .hero-right { padding-right: 0; }
            .bri-logo-watermark { width: 150px; }
            
            div[data-testid="stRadio"] > div[role="radiogroup"], 
            div.row-widget.stRadio > div[role="radiogroup"] { 
                flex-wrap: nowrap !important; 
                overflow-x: auto !important; 
                -webkit-overflow-scrolling: touch; 
                padding-bottom: 5px !important;
            }
            div[data-testid="stRadio"] > div[role="radiogroup"]::-webkit-scrollbar { height: 4px; display: block; }
            div[data-testid="stRadio"] > div[role="radiogroup"]::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 4px; }
            
            .table-scroll { overflow-x: auto !important; }
            .dash-table { min-width: 700px; }
            .log-table { min-width: 800px; }
            
            .home-card { margin-bottom: 10px; }
            
            .header-title-text { font-size: 14px !important; padding-left: 8px !important; margin-left: 8px !important; }
        }
    </style>
""", unsafe_allow_html=True)

def fmt_vis(val, is_visible=True):
    if not is_visible: return ""
    if pd.isna(val): return ""
    try:
        if float(val) == 0: return ""
    except: pass
    if isinstance(val, (int, float)): return str(int(val) if val == int(val) else val)
    return str(val)

@st.cache_data(ttl=3600)
def load_data_gspread(worksheet_name, range_name=None):
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet(worksheet_name)
        if range_name: data = ws.get(range_name)
        else: data = ws.get_all_values()
            
        if not data or len(data) < 2: return pd.DataFrame()
        
        df = pd.DataFrame(data[1:], columns=data[0])
        df.columns = df.columns.astype(str).str.strip().str.upper()
        df = df.loc[:, ~df.columns.duplicated()] 
        df = df.loc[:, df.columns != ''] 
        
        if df.empty: return df
        if 'TID' in df.columns:
            df['TID'] = df['TID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
        if worksheet_name == "SLM Visit Log" and 'TGL VISIT SLM' in df.columns:
            df['TGL_DT'] = pd.to_datetime(df['TGL VISIT SLM'], errors='coerce')
            df = df.dropna(subset=['TGL_DT']) 
            df = df.sort_values(by=['TID', 'TGL_DT'], ascending=[True, False])
            df = df.drop_duplicates(subset=['TID'], keep='first')
            df['TGL VISIT SLM'] = df['TGL_DT'].dt.strftime('%d/%m/%Y')
            
        return df
    except Exception as e:
        st.error(f"Gagal Load {worksheet_name}: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def apply_fallback_logic(df_fresh):
    backup_file = "backup_AIMS_Master.csv"
    kategori_krusial = ["ELASTIC", "COMPLAIN", "DF REPEAT", "OUT FLM"]
    warnings_list = []

    if df_fresh.empty or 'KATEGORI' not in df_fresh.columns:
        if os.path.exists(backup_file):
            warnings_list.append("⚠️ Peringatan Krusial: Tarikan dari Sheet AIMS_Master kosong/gagal. Menggunakan 100% Salinan Backup Lokal.")
            try:
                df_backup = pd.read_csv(backup_file, dtype=str)
                return df_backup, warnings_list
            except Exception as e:
                warnings_list.append(f"⚠️ Gagal membaca salinan backup: {e}")
                return df_fresh, warnings_list
        else:
            return df_fresh, ["⚠️ Peringatan: Data di Google Sheets kosong dan file backup lokal belum tersedia."]

    df_fresh_check = df_fresh.copy()
    df_fresh_check['KAT_CEK'] = df_fresh_check['KATEGORI'].astype(str).str.strip().str.upper()
    df_final = df_fresh.copy()
    
    kategori_hilang = []
    for cat in kategori_krusial:
        if df_fresh_check[df_fresh_check['KAT_CEK'] == cat].empty:
            kategori_hilang.append(cat)

    if kategori_hilang and os.path.exists(backup_file):
        try:
            df_backup = pd.read_csv(backup_file, dtype=str)
            if 'KATEGORI' in df_backup.columns:
                df_backup['KAT_CEK'] = df_backup['KATEGORI'].astype(str).str.strip().str.upper()
                
                for cat in kategori_hilang:
                    df_cat_backup = df_backup[df_backup['KAT_CEK'] == cat]
                    if not df_cat_backup.empty:
                        df_cat_backup_clean = df_cat_backup.drop(columns=['KAT_CEK'], errors='ignore')
                        df_final = pd.concat([df_final, df_cat_backup_clean], ignore_index=True)
                        warnings_list.append(f"⚠️ Peringatan: Sumber Kategori '{cat}' terdeteksi ERROR. Menampilkan data '{cat}' dari Salinan Backup Lokal.")
        except Exception:
            pass

    if 'KAT_CEK' in df_final.columns:
        df_final = df_final.drop(columns=['KAT_CEK'])
        
    try:
        df_final.to_csv(backup_file, index=False)
    except Exception:
        pass

    return df_final, warnings_list

df_master_fresh = load_data_gspread("AIMS_Master")
df_master, master_warnings = apply_fallback_logic(df_master_fresh)

current_date_full = datetime.now().strftime("%A, %d %B %Y")
current_date_header = current_date_full.upper()

st.markdown(f"<div class='main-header'><div style='display: flex; align-items: center; white-space: nowrap;'><img src='https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png' style='height: 28px; filter: brightness(0) invert(1);' alt='Logo BRI'><span class='header-title-text' style='margin-left: 15px; padding-left: 15px; border-left: 2px solid #F37021; color: #FFFFFF; font-weight: 700; font-size: 18px; letter-spacing: 2px; text-shadow: 1px 1px 3px rgba(0,0,0,0.3);'>WEEKLY ATM PERFORMANCE REVIEW</span></div><div class='info-ticker-container'><div class='info-ticker-text'><span class='status-dot'></span>SYSTEM: SECURE & OPTIMAL &nbsp; | &nbsp; DATA LOADED &nbsp; | &nbsp; SERVER TIME: {current_date_header} &nbsp; | &nbsp; BANK BRI MONITORING ACTIVE &nbsp; | &nbsp; <span style='color: #FBBF24; font-weight: 800; letter-spacing: 1px;'>PT KELOLA JASA ARTHA</span></div></div></div>", unsafe_allow_html=True)

for msg in master_warnings:
    st.warning(msg)

col_nav, col_space_top, col_filter1, col_filter2 = st.columns([6.4, 0.4, 1.7, 1.5])

with col_nav:
    menu_list = ["Home", "⭐ MRI PROJECT", "Elastic", "Complain", "DF Repeat", "OUT Flm", "Logistic"]
    menu_pilihan = st.radio("Nav", menu_list, horizontal=True, label_visibility="collapsed", index=0)

with col_filter1:
    list_months = ["January 2026", "February 2026", "March 2026", "April 2026", "May 2026", "June 2026"]
    selected_month_full = st.selectbox("Bulan", list_months, index=3, label_visibility="collapsed")
    selected_month = selected_month_full.split()[0].upper() 
    idx = list_months.index(selected_month_full)
    prev_month = list_months[idx - 1].split()[0].upper() if idx > 0 else list_months[0].split()[0].upper()

with col_filter2:
    selected_week = st.selectbox("Minggu", ["W1", "W2", "W3", "W4"], index=2, label_visibility="collapsed")

show_w1 = selected_week in ['W1', 'W2', 'W3', 'W4']
show_w2 = selected_week in ['W2', 'W3', 'W4']
show_w3 = selected_week in ['W3', 'W4']
show_w4 = selected_week == 'W4'

active_weeks = []
if show_w1: active_weeks.append('W1')
if show_w2: active_weeks.append('W2')
if show_w3: active_weeks.append('W3')
if show_w4: active_weeks.append('W4')

st.markdown("<hr style='margin-top: 15px; margin-bottom: 12px; border: 0; height: 0px;'>", unsafe_allow_html=True)

kategori_valid = ["Elastic", "Complain", "DF Repeat", "OUT Flm"]

# ==========================================
# 4. LOGIKA HALAMAN HOME
# ==========================================
if menu_pilihan == "Home":
    def get_mtd_data(kategori):
        if df_master.empty or 'KATEGORI' not in df_master.columns: return 0, 0
        df_cat = df_master[df_master['KATEGORI'].astype(str).str.upper() == kategori.upper()].copy()
        if df_cat.empty: return 0, 0
        
        df_cat['VAL_METRIC'] = pd.to_numeric(df_cat['JUMLAH_COMPLAIN'], errors='coerce').fillna(0) if (kategori.upper() == 'COMPLAIN' and 'JUMLAH_COMPLAIN' in df_cat.columns) else 1
        if 'WEEK' in df_cat.columns: df_cat['WEEK_CLN'] = df_cat['WEEK'].astype(str).str.strip().str.upper()
        else: df_cat['WEEK_CLN'] = ""
        
        df_curr = df_cat[(df_cat['BULAN'].astype(str).str.upper() == selected_month) & (df_cat['WEEK_CLN'].isin(active_weeks))]
        df_prev = df_cat[(df_cat['BULAN'].astype(str).str.upper() == prev_month) & (df_cat['WEEK_CLN'].isin(active_weeks))]
        
        return int(df_curr['VAL_METRIC'].sum()), int(df_prev['VAL_METRIC'].sum())

    kpi_categories = ["Elastic", "Complain", "DF Repeat", "OUT Flm"]
    kpi_cols = st.columns(4)
    
    for i, cat in enumerate(kpi_categories):
        val_curr, val_prev = get_mtd_data(cat)
        delta_html = ""
        if val_prev > 0:
            diff = val_curr - val_prev
            pct = (diff / val_prev) * 100
            if diff < 0: delta_html = f"<div class='card-delta-down'>▼ {abs(pct):.1f}% vs Prev MTD</div>"
            elif diff > 0: delta_html = f"<div class='card-delta-up'>▲ {abs(pct):.1f}% vs Prev MTD</div>"
            else: delta_html = f"<div class='card-delta-neutral'>▬ No Change vs Prev MTD</div>"
        else: delta_html = f"<div class='card-delta-neutral'>N/A vs Prev MTD</div>"

        with kpi_cols[i]:
            st.markdown(f"<div class='home-card'><div class='card-title'>{cat.upper()} PROBLEM</div><div class='card-value'>{val_curr}</div>{delta_html}</div>", unsafe_allow_html=True)

    st.markdown(f"<div class='hero-section'><div class='hero-left'><div class='hero-title'>WEEKLY PERFORMANCE REVIEW</div><div class='hero-subtitle'>ATM MONITORING DIVISION</div><div class='hero-text'>Presenter : Command Center BRI</div><div class='hero-text'>{current_date_full}</div></div><div class='hero-right'><img src='https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png' class='bri-logo-watermark' alt='Logo BRI'></div></div>", unsafe_allow_html=True)

# ==========================================
# 5. LOGIKA HALAMAN ⭐ MRI PROJECT
# ==========================================
elif menu_pilihan == "⭐ MRI PROJECT":
    df_mri = load_data_gspread("Problem MRI 2025/26 Harian")
    df_slm = load_data_gspread("SLM Visit Log")
    
    if df_mri.empty:
        st.warning("Bah! Data MRI Project kosong. Periksa worksheet 'Problem MRI 2025/26 Harian'!")
    else:
        col_cat3 = "KATEGORI_PROBLEM3"
        
        def calculate_mri_metrics(category_name):
            df_filter = df_mri[df_mri[col_cat3].astype(str).str.upper() == category_name.upper()].copy()
            if 'WEEK' in df_filter.columns: df_filter['WEEK_CLN'] = df_filter['WEEK'].astype(str).str.strip().str.upper()
            else: df_filter['WEEK_CLN'] = ""
            
            v_prev = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == prev_month) & (df_filter['WEEK_CLN'].isin(['W1', 'W2', 'W3', 'W4']))])
            v_w1 = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W1')])
            v_w2 = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W2')])
            v_w3 = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W3')])
            v_w4 = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W4')])
            
            v_total = len(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'].isin(active_weeks))])

            def get_tiers(sub_df):
                if sub_df.empty: return {"1 kali": 0, "2-3 kali": 0, "> 3 kali": 0}
                counts = sub_df.groupby('TID').size()
                return {"1 kali": (counts == 1).sum(), "2-3 kali": ((counts >= 2) & (counts <= 3)).sum(), "> 3 kali": (counts > 3).sum()}

            t_prev = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == prev_month) & (df_filter['WEEK_CLN'].isin(['W1', 'W2', 'W3', 'W4']))])
            t_w1 = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W1')])
            t_w2 = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W2')])
            t_w3 = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W3')])
            t_w4 = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'] == 'W4')])
            
            t_total = get_tiers(df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'].isin(active_weeks))])

            df_curr = df_filter[(df_filter['BULAN'].astype(str).str.upper() == selected_month) & (df_filter['WEEK_CLN'].isin(active_weeks))].copy()
            if not df_curr.empty:
                pivot_t = df_curr.pivot_table(index='TID', columns='WEEK_CLN', aggfunc='size', fill_value=0)
                for w in ['W1', 'W2', 'W3', 'W4']:
                    if w not in pivot_t.columns: pivot_t[w] = 0
                
                df_info = df_curr.drop_duplicates('TID').set_index('TID')
                pivot_t['Location'] = pivot_t.index.map(df_info.get('LOKASI', pd.Series(dtype=str))).fillna('-')
                pivot_t['Branch'] = pivot_t.index.map(df_info.get('CPC', pd.Series(dtype=str))).fillna('-').astype(str).str.upper().str.replace('KEJAR', '', regex=False).str.strip()
                pivot_t['SLM_Vendor'] = pivot_t.index.map(df_info.get('SLM', pd.Series(dtype=str))).fillna('-')
                
                sort_col = selected_week if selected_week in pivot_t.columns else 'W1'
                pivot_t = pivot_t[pivot_t[sort_col] > 0]
                pivot_t = pivot_t.sort_values(by=sort_col, ascending=False).head(5)
            else: pivot_t = pd.DataFrame()

            return {
                "summary": [34, v_prev, v_w1, v_w2, v_w3, v_w4, v_total],
                "tiers": {
                    "1 kali": [t_prev["1 kali"], t_w1["1 kali"], t_w2["1 kali"], t_w3["1 kali"], t_w4["1 kali"], t_total["1 kali"]],
                    "2-3 kali": [t_prev["2-3 kali"], t_w1["2-3 kali"], t_w2["2-3 kali"], t_w3["2-3 kali"], t_w4["2-3 kali"], t_total["2-3 kali"]],
                    "> 3 kali": [t_prev["> 3 kali"], t_w1["> 3 kali"], t_w2["> 3 kali"], t_w3["> 3 kali"], t_w4["> 3 kali"], t_total["> 3 kali"]]
                },
                "top_tid": pivot_t
            }

        mri_complain = calculate_mri_metrics("Complain")
        mri_dfrepeat = calculate_mri_metrics("Df Repeat")

        col_mri_left, col_mri_spacer, col_mri_right = st.columns([10, 0.5, 10])
        prev_lbl, curr_lbl = f"{prev_month[:3].capitalize()}(prev)", f"Σ {selected_month[:3].capitalize()}"

        with col_mri_left:
            st.markdown(f"<div class='section-title'>JUMLAH COMPLAIN</div>", unsafe_allow_html=True)
            s = mri_complain["summary"]
            html_sum = f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width: 18%;'>TOTAL ATM</th><th style='width: 16%;'>{prev_lbl}</th><th style='width: 12%;'>W1</th><th style='width: 12%;'>W2</th><th style='width: 12%;'>W3</th><th style='width: 12%;'>W4</th><th style='width: 18%;'>{curr_lbl}</th></tr><tr><td style='font-weight:700;'>{s[0]}</td><td>{fmt_vis(s[1])}</td><td>{fmt_vis(s[2], show_w1)}</td><td>{fmt_vis(s[3], show_w2)}</td><td>{fmt_vis(s[4], show_w3)}</td><td>{fmt_vis(s[5], show_w4)}</td><td style='font-weight:700;'>{fmt_vis(s[6])}</td></tr></table></div>"
            st.markdown(html_sum, unsafe_allow_html=True)
            
            st.markdown(f"<div class='section-title'>Tiering by TID (Complain)</div>", unsafe_allow_html=True)
            t = mri_complain["tiers"]
            html_tier_rows = "".join([f"<tr><td style='font-weight:600;'>{label}</td><td>{fmt_vis(t[label][0])}</td><td>{fmt_vis(t[label][1], show_w1)}</td><td>{fmt_vis(t[label][2], show_w2)}</td><td>{fmt_vis(t[label][3], show_w3)}</td><td>{fmt_vis(t[label][4], show_w4)}</td><td style='font-weight:700;'>{fmt_vis(t[label][5])}</td></tr>" for label in ["1 kali", "2-3 kali", "> 3 kali"]])
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width: 18%;'>Tiering</th><th style='width: 16%;'>{prev_lbl}</th><th style='width: 12%;'>W1</th><th style='width: 12%;'>W2</th><th style='width: 12%;'>W3</th><th style='width: 12%;'>W4</th><th style='width: 18%;'>{curr_lbl}</th></tr>{html_tier_rows}</table></div>", unsafe_allow_html=True)

            st.markdown(f"<div class='section-title'>Top Complain Problem Terminal IDs</div>", unsafe_allow_html=True)
            df_top_c = mri_complain["top_tid"]
            html_top_c = "".join([f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; max-width:130px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{row['Location']}'>{row['Location']}</td><td>{row['Branch']}</td><td>{row['SLM_Vendor']}</td><td>{fmt_vis(row['W1'], show_w1)}</td><td>{fmt_vis(row['W2'], show_w2)}</td><td>{fmt_vis(row['W3'], show_w3)}</td><td>{fmt_vis(row['W4'], show_w4)}</td></tr>" for i, (tid, row) in enumerate(df_top_c.iterrows(), 1)]) if not df_top_c.empty else f"<tr><td colspan='9' style='padding: 15px; font-weight: bold; color: #10B981; text-align: center;'>✅ Data Kosong - Nihil Problem Complain di {selected_week}</td></tr>"
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th>Branch</th><th>SLM</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th></tr>{html_top_c}</table></div>", unsafe_allow_html=True)

            st.markdown(f"<div class='section-title'>Follow-up Teknisi (Complain)</div>", unsafe_allow_html=True)
            html_fup_c = "".join([f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; max-width:130px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{df_top_c.loc[tid, 'Location']}'>{df_top_c.loc[tid, 'Location']}</td><td style='text-align:left;'>{df_slm[df_slm['TID'] == tid].iloc[0].get('TGL VISIT SLM', '-') if not df_slm[df_slm['TID'] == tid].empty else '-'}</td><td style='text-align:left; max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{df_slm[df_slm['TID'] == tid].iloc[0].get('ACTION', 'Belum ada log') if not df_slm[df_slm['TID'] == tid].empty else 'Belum ada log'}'>{df_slm[df_slm['TID'] == tid].iloc[0].get('ACTION', 'Belum ada log') if not df_slm[df_slm['TID'] == tid].empty else 'Belum ada log'}</td></tr>" for i, tid in enumerate(df_top_c.index, 1)]) if not df_top_c.empty else f"<tr><td colspan='5' style='padding: 15px; font-weight: bold; color: #10B981; text-align: center;'>✅ Data Kosong - Tidak ada Follow-up Complain di {selected_week}</td></tr>"
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th style='text-align:left; width:15%;'>TGL VISIT</th><th style='text-align:left;'>Action</th></tr>{html_fup_c}</table></div>", unsafe_allow_html=True)

        with col_mri_right:
            st.markdown(f"<div class='section-title'>JUMLAH DF REPEAT</div>", unsafe_allow_html=True)
            s = mri_dfrepeat["summary"]
            html_sum = f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width: 18%;'>TOTAL ATM</th><th style='width: 16%;'>{prev_lbl}</th><th style='width: 12%;'>W1</th><th style='width: 12%;'>W2</th><th style='width: 12%;'>W3</th><th style='width: 12%;'>W4</th><th style='width: 18%;'>{curr_lbl}</th></tr><tr><td style='font-weight:700;'>{s[0]}</td><td>{fmt_vis(s[1])}</td><td>{fmt_vis(s[2], show_w1)}</td><td>{fmt_vis(s[3], show_w2)}</td><td>{fmt_vis(s[4], show_w3)}</td><td>{fmt_vis(s[5], show_w4)}</td><td style='font-weight:700;'>{fmt_vis(s[6])}</td></tr></table></div>"
            st.markdown(html_sum, unsafe_allow_html=True)

            st.markdown(f"<div class='section-title'>Tiering by TID (Df Repeat)</div>", unsafe_allow_html=True)
            t = mri_dfrepeat["tiers"]
            html_tier_rows = "".join([f"<tr><td style='font-weight:600;'>{label}</td><td>{fmt_vis(t[label][0])}</td><td>{fmt_vis(t[label][1], show_w1)}</td><td>{fmt_vis(t[label][2], show_w2)}</td><td>{fmt_vis(t[label][3], show_w3)}</td><td>{fmt_vis(t[label][4], show_w4)}</td><td style='font-weight:700;'>{fmt_vis(t[label][5])}</td></tr>" for label in ["1 kali", "2-3 kali", "> 3 kali"]])
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width: 18%;'>Tiering</th><th style='width: 16%;'>{prev_lbl}</th><th style='width: 12%;'>W1</th><th style='width: 12%;'>W2</th><th style='width: 12%;'>W3</th><th style='width: 12%;'>W4</th><th style='width: 18%;'>{curr_lbl}</th></tr>{html_tier_rows}</table></div>", unsafe_allow_html=True)

            st.markdown(f"<div class='section-title'>Top Df Repeat Problem Terminal IDs</div>", unsafe_allow_html=True)
            df_top_d = mri_dfrepeat["top_tid"]
            html_top_d = "".join([f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; max-width:130px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{row['Location']}'>{row['Location']}</td><td>{row['Branch']}</td><td>{row['SLM_Vendor']}</td><td>{fmt_vis(row['W1'], show_w1)}</td><td>{fmt_vis(row['W2'], show_w2)}</td><td>{fmt_vis(row['W3'], show_w3)}</td><td>{fmt_vis(row['W4'], show_w4)}</td></tr>" for i, (tid, row) in enumerate(df_top_d.iterrows(), 1)]) if not df_top_d.empty else f"<tr><td colspan='9' style='padding: 15px; font-weight: bold; color: #10B981; text-align: center;'>✅ Data Kosong - Nihil Problem Df Repeat di {selected_week}</td></tr>"
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th>Branch</th><th>SLM</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th></tr>{html_top_d}</table></div>", unsafe_allow_html=True)

            st.markdown(f"<div class='section-title'>Follow-up Teknisi (Df Repeat)</div>", unsafe_allow_html=True)
            html_fup_d = "".join([f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; max-width:130px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{df_top_d.loc[tid, 'Location']}'>{df_top_d.loc[tid, 'Location']}</td><td style='text-align:left;'>{df_slm[df_slm['TID'] == tid].iloc[0].get('TGL VISIT SLM', '-') if not df_slm[df_slm['TID'] == tid].empty else '-'}</td><td style='text-align:left; max-width:150px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;' title='{df_slm[df_slm['TID'] == tid].iloc[0].get('ACTION', 'Belum ada log') if not df_slm[df_slm['TID'] == tid].empty else 'Belum ada log'}'>{df_slm[df_slm['TID'] == tid].iloc[0].get('ACTION', 'Belum ada log') if not df_slm[df_slm['TID'] == tid].empty else 'Belum ada log'}</td></tr>" for i, tid in enumerate(df_top_d.index, 1)]) if not df_top_d.empty else f"<tr><td colspan='5' style='padding: 15px; font-weight: bold; color: #10B981; text-align: center;'>✅ Data Kosong - Tidak ada Follow-up Df Repeat di {selected_week}</td></tr>"
            st.markdown(f"<div class='table-scroll' style='max-height:unset;'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th style='text-align:left; width:15%;'>TGL VISIT</th><th style='text-align:left;'>Action</th></tr>{html_fup_d}</table></div>", unsafe_allow_html=True)

# ==========================================
# 6. LOGIKA PANDAS DINAMIS (MASTER MONITORING)
# ==========================================
elif menu_pilihan in kategori_valid:
    df_slm = load_data_gspread("SLM Visit Log")
    df_kelolaan = load_data_gspread("Jml_Kelolaan", "A1:B10")
    df_analisa = load_data_gspread("Analisa_dan_Perbaikan") 

    total_atm = 0
    dict_kelolaan = {}
    if not df_kelolaan.empty and 'CABANG' in df_kelolaan.columns and 'TTL ATM' in df_kelolaan.columns:
        df_kelolaan['TTL ATM'] = pd.to_numeric(df_kelolaan['TTL ATM'], errors='coerce').fillna(0)
        total_atm = int(df_kelolaan['TTL ATM'].sum())
        for _, row in df_kelolaan.iterrows():
            if pd.notnull(row['CABANG']):
                dict_kelolaan[str(row['CABANG']).strip().upper()] = str(int(row['TTL ATM']))

    current_cat = menu_pilihan.upper()
    val_prev = val_w1 = val_w2 = val_w3 = val_w4 = val_total = val_avg = 0
    tier_prev = tier_w1 = tier_w2 = tier_w3 = tier_w4 = tier_total = {"1 kali": 0, "2-3 kali": 0, "> 3 kali": 0}
    df_top_tid = pd.DataFrame()
    df_top_branch = pd.DataFrame()
    
    text_analisa_val = ""
    if not df_analisa.empty:
        temp_df = df_analisa[df_analisa['KATEGORI'].astype(str).str.upper() == current_cat]
        if not temp_df.empty:
            col_idx = 2 if 'ANALISA_SOLUSI' in temp_df.columns else (1 if len(temp_df.columns) > 1 else 0)
            val = str(temp_df.iloc[0, col_idx])
            text_analisa_val = val if val.lower() != 'nan' else ""
            
    if current_cat == 'ELASTIC': df_followup_view = load_data_gspread("Summary_Monitoring_Cash", "U3:Y7")
    elif current_cat == 'COMPLAIN': df_followup_view = load_data_gspread("Summary_Monitoring_Cash", "U17:Y20")
    else: df_followup_view = pd.DataFrame() 
    
    if not df_master.empty and 'KATEGORI' in df_master.columns:
        df_cat = df_master[df_master['KATEGORI'].astype(str).str.upper() == current_cat].copy()
        df_cat['VAL_METRIC'] = pd.to_numeric(df_cat['JUMLAH_COMPLAIN'], errors='coerce').fillna(0) if (current_cat == 'COMPLAIN' and 'JUMLAH_COMPLAIN' in df_cat.columns) else 1
        
        if 'WEEK' in df_cat.columns: df_cat['WEEK_CLN'] = df_cat['WEEK'].astype(str).str.strip().str.upper()
        else: df_cat['WEEK_CLN'] = ""
            
        df_prev_m = df_cat[(df_cat['BULAN'].astype(str).str.upper() == prev_month) & (df_cat['WEEK_CLN'].isin(['W1', 'W2', 'W3', 'W4']))]
        val_prev = df_prev_m['VAL_METRIC'].sum()
        prev_counts_tid = df_prev_m.groupby('TID')['VAL_METRIC'].sum()
        prev_counts_branch = df_prev_m.groupby('CABANG')['VAL_METRIC'].sum()
        
        df_curr_m = df_cat[(df_cat['BULAN'].astype(str).str.upper() == selected_month) & (df_cat['WEEK_CLN'].isin(active_weeks))]
        val_total = df_curr_m['VAL_METRIC'].sum() 
        
        val_w1 = df_curr_m[df_curr_m['WEEK_CLN'] == 'W1']['VAL_METRIC'].sum()
        val_w2 = df_curr_m[df_curr_m['WEEK_CLN'] == 'W2']['VAL_METRIC'].sum()
        val_w3 = df_curr_m[df_curr_m['WEEK_CLN'] == 'W3']['VAL_METRIC'].sum()
        val_w4 = df_curr_m[df_curr_m['WEEK_CLN'] == 'W4']['VAL_METRIC'].sum()

        def calculate_tiers(df_p):
            if df_p.empty or 'TID' not in df_p.columns: return {"1 kali": 0, "2-3 kali": 0, "> 3 kali": 0}
            c = df_p.groupby('TID')['VAL_METRIC'].sum()
            return {"1 kali": (c == 1).sum(), "2-3 kali": ((c >= 2) & (c <= 3)).sum(), "> 3 kali": (c > 3).sum()}
        
        tier_prev, tier_total = calculate_tiers(df_prev_m), calculate_tiers(df_curr_m)
        tier_w1, tier_w2 = calculate_tiers(df_curr_m[df_curr_m['WEEK_CLN'] == 'W1']), calculate_tiers(df_curr_m[df_curr_m['WEEK_CLN'] == 'W2'])
        tier_w3, tier_w4 = calculate_tiers(df_curr_m[df_curr_m['WEEK_CLN'] == 'W3']), calculate_tiers(df_curr_m[df_curr_m['WEEK_CLN'] == 'W4'])

        pivot_tid = df_curr_m.pivot_table(index='TID', columns='WEEK_CLN', values='VAL_METRIC', aggfunc='sum', fill_value=0)
        for w in ['W1', 'W2', 'W3', 'W4']:
            if w not in pivot_tid.columns: pivot_tid[w] = 0
        pivot_tid = pivot_tid[['W1', 'W2', 'W3', 'W4']]
        pivot_tid['TOTAL'] = pivot_tid.sum(axis=1)
        
        loc_map, branch_map = df_master.drop_duplicates('TID').set_index('TID')['LOKASI'], df_master.drop_duplicates('TID').set_index('TID')['CABANG']
        df_top_tid = pivot_tid.copy()
        df_top_tid['PREV'], df_top_tid['Location'], df_top_tid['Branch'] = df_top_tid.index.map(prev_counts_tid).fillna(0).astype(int), df_top_tid.index.map(loc_map).fillna('-'), df_top_tid.index.map(branch_map).fillna('-')
        df_top_tid = df_top_tid.sort_values(by=[(selected_week if selected_week in df_top_tid.columns else 'TOTAL'), 'TOTAL'], ascending=[False, False])

        pivot_branch = df_curr_m.pivot_table(index='CABANG', columns='WEEK_CLN', values='VAL_METRIC', aggfunc='sum', fill_value=0)
        for w in ['W1', 'W2', 'W3', 'W4']:
            if w not in pivot_branch.columns: pivot_branch[w] = 0
        pivot_branch['TOTAL'], pivot_branch['PREV'] = pivot_branch.sum(axis=1), pivot_branch.index.map(prev_counts_branch).fillna(0).astype(int)
        df_top_branch = pivot_branch.sort_values(by=[(selected_week if selected_week in pivot_branch.columns else 'TOTAL'), 'TOTAL'], ascending=[False, False])

    val_avg, tgt_tot_val = (int(val_total / len(active_weeks)) if active_weeks and val_total > 0 else 0), int(val_prev * 0.8)
    tgt_weekly_val = int(tgt_tot_val / 4) if val_prev > 0 else 0

    col_left, col_spacer, col_right = st.columns([10, 0.2, 10])

    with col_left:
        st.markdown(f"<div class='section-title'>All {menu_pilihan} Overview</div>", unsafe_allow_html=True)
        val_w1_s, val_w2_s, val_w3_s, val_w4_s = fmt_vis(val_w1, show_w1), fmt_vis(val_w2, show_w2), fmt_vis(val_w3, show_w3), fmt_vis(val_w4, show_w4)
        tgt_w1, tgt_w2, tgt_w3, tgt_w4 = (str(tgt_weekly_val) if val_w1_s != "" else ""), (str(tgt_weekly_val) if val_w2_s != "" else ""), (str(tgt_weekly_val) if val_w3_s != "" else ""), (str(tgt_weekly_val) if val_w4_s != "" else "")
        active_w_count = sum([1 for w in [tgt_w1, tgt_w2, tgt_w3, tgt_w4] if w != ""])
        t_tot, t_avg = (str(tgt_weekly_val * active_w_count) if active_w_count > 0 else ""), (str(tgt_weekly_val) if active_w_count > 0 else "")

        html_t1 = f"<div class='table-scroll' style='max-height: unset; overflow: visible; box-shadow: none; border: none; margin-bottom: 12px;'><table class='dash-table' style='border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);'><tr><th style='width: 14%;'>Total ATM</th><th style='width: 12%;'>{prev_month[:3].capitalize()} (prev)</th><th style='width: 10%;'>W1</th><th style='width: 10%;'>W2</th><th style='width: 10%;'>W3</th><th style='width: 10%;'>W4</th><th style='width: 17%;'>Σ {selected_month[:3].capitalize()}</th><th style='width: 17%;'>Avg/Week</th></tr><tr><td>{total_atm}</td><td>{fmt_vis(val_prev, True)}</td><td>{val_w1_s}</td><td>{val_w2_s}</td><td>{val_w3_s}</td><td>{val_w4_s}</td><td>{fmt_vis(val_total, True)}</td><td>{fmt_vis(val_avg, True)}</td></tr><tr style='background-color:#FFFFFF;'><td colspan='2' style='font-weight:700; color:#00529C;'>TARGET PENURUNAN 20%</td><td>{tgt_w1}</td><td>{tgt_w2}</td><td>{tgt_w3}</td><td>{tgt_w4}</td><td>{t_tot}</td><td>{t_avg}</td></tr></table></div>"
        st.markdown(html_t1, unsafe_allow_html=True)

        st.markdown(f"<div class='section-title'>{menu_pilihan} TID Risk Tiers</div>", unsafe_allow_html=True)
        html_t2 = f"<div class='table-scroll' style='max-height: unset; overflow: visible; box-shadow: none; border: none; margin-bottom: 12px;'><table class='dash-table' style='border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);'><tr><th style='width: 20%;'>Tiering</th><th style='width: 12%;'>{prev_month[:3].capitalize()} (prev)</th><th style='width: 13%;'>W1</th><th style='width: 13%;'>W2</th><th style='width: 13%;'>W3</th><th style='width: 13%;'>W4</th><th style='width: 16%;'>Σ {selected_month[:3].capitalize()}</th></tr><tr><td style='font-weight: 600;'>1 kali</td><td>{fmt_vis(tier_prev['1 kali'], True)}</td><td>{fmt_vis(tier_w1['1 kali'], show_w1)}</td><td>{fmt_vis(tier_w2['1 kali'], show_w2)}</td><td>{fmt_vis(tier_w3['1 kali'], show_w3)}</td><td>{fmt_vis(tier_w4['1 kali'], show_w4)}</td><td>{fmt_vis(tier_total['1 kali'], True)}</td></tr><tr><td style='font-weight: 600;'>2-3 kali</td><td>{fmt_vis(tier_prev['2-3 kali'], True)}</td><td>{fmt_vis(tier_w1['2-3 kali'], show_w1)}</td><td>{fmt_vis(tier_w2['2-3 kali'], show_w2)}</td><td>{fmt_vis(tier_w3['2-3 kali'], show_w3)}</td><td>{fmt_vis(tier_w4['2-3 kali'], show_w4)}</td><td>{fmt_vis(tier_total['2-3 kali'], True)}</td></tr><tr><td style='font-weight: 600;'>> 3 kali</td><td>{fmt_vis(tier_prev['> 3 kali'], True)}</td><td>{fmt_vis(tier_w1['> 3 kali'], show_w1)}</td><td>{fmt_vis(tier_w2['> 3 kali'], show_w2)}</td><td>{fmt_vis(tier_w3['> 3 kali'], show_w3)}</td><td>{fmt_vis(tier_w4['> 3 kali'], show_w4)}</td><td>{fmt_vis(tier_total['> 3 kali'], True)}</td></tr></table></div>"
        st.markdown(html_t2, unsafe_allow_html=True)

        if current_cat in ['ELASTIC', 'COMPLAIN']:
            st.markdown(f"<div class='section-title'>{menu_pilihan} Daily Follow-up</div>", unsafe_allow_html=True)
            html_t3_rows, table_structure = "", ""
            if current_cat == 'ELASTIC':
                if not df_followup_view.empty:
                    html_t3_rows = "".join([f"<tr><td style='text-align:left;'>{str(row.iloc[0])}</td><td style='text-align:center;'>{fmt_vis(row.iloc[1], True)}</td><td style='text-align:center;'>{fmt_vis(row.iloc[2], True)}</td><td style='font-weight:600; text-align:center;'>{fmt_vis(row.iloc[3], True)}</td><td style='color:#00529C; font-weight:600; text-align:center;'>{fmt_vis(row.iloc[4], True) if len(row) > 4 else ''}</td></tr>" for _, row in df_followup_view.iterrows()])
                else: html_t3_rows = "<tr><td colspan='5'>Data Kosong</td></tr>"
                table_structure = f"<tr><th style='text-align: left; width: 50%;'>STATUS</th><th>Pending</th><th>Done</th><th>Total</th><th>% TL</th></tr>{html_t3_rows}"
            elif current_cat == 'COMPLAIN':
                if not df_followup_view.empty:
                    total_rows = len(df_followup_view)
                    for i, (_, row) in enumerate(df_followup_view.iterrows()):
                        if i == 0: html_t3_rows += f"<tr><td style='text-align:left;'>{str(row.iloc[0])}</td><td style='text-align:center; vertical-align:middle;'>{fmt_vis(row.iloc[1], True)}</td><td style='text-align:center; vertical-align:middle;'>{fmt_vis(row.iloc[2], True)}</td><td rowspan='{total_rows}' style='color:#00529C; font-weight:600; text-align:center; vertical-align:middle; border-left: 1px solid #E2E8F0;'>{fmt_vis(row.iloc[3], True) if len(row) > 3 else ''}</td><td rowspan='{total_rows}' style='color:#00529C; font-weight:600; text-align:center; vertical-align:middle; border-left: 1px solid #E2E8F0;'>{fmt_vis(row.iloc[4], True) if len(row) > 4 else ''}</td></tr>"
                        else: html_t3_rows += f"<tr><td style='text-align:left;'>{str(row.iloc[0])}</td><td style='text-align:center; vertical-align:middle;'>{fmt_vis(row.iloc[1], True)}</td><td style='text-align:center; vertical-align:middle;'>{fmt_vis(row.iloc[2], True)}</td></tr>"
                    html_t3_rows += "<tr><td colspan='5' style='text-align:left; font-size: 9.5px; font-weight: 500; color:#64748B; background-color: #F8FAFC;'>Source : Worksheet Complain</td></tr>"
                else: html_t3_rows = "<tr><td colspan='5'>Data Kosong</td></tr>"
                table_structure = f"<tr><th style='text-align: left; width: 36%;'>STATUS</th><th style='text-align: center;'>{prev_month[:3].capitalize()}</th><th style='text-align: center;'>{selected_month[:3].capitalize()}</th><th style='text-align: center;'>% {prev_month[:3].capitalize()}</th><th style='text-align: center;'>% {selected_month[:3].capitalize()}</th></tr>{html_t3_rows}"
            st.markdown(f"<div class='table-scroll' style='max-height: unset; overflow: visible; box-shadow: none; border: none; margin-bottom: 12px;'><table class='dash-table' style='border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);'>{table_structure}</table></div>", unsafe_allow_html=True)

        st.markdown(f"<div class='section-title'>Analisa & Solusi Perbaikan ({menu_pilihan})</div>", unsafe_allow_html=True)
        with st.form(key=f'form_analisa_{current_cat}', clear_on_submit=False):
            input_text = st.text_area("Input", value=text_analisa_val, label_visibility="collapsed", height=160)
            if st.form_submit_button("Save"):
                if input_text.strip() == "": st.warning("Tahe! Jangan kirim data kosong, David!")
                else:
                    try:
                        sh = client.open_by_key(SHEET_ID)
                        sheet = sh.worksheet("Analisa_dan_Perbaikan")
                        sheet.insert_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), current_cat, input_text], index=2)
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e: st.error(f"Gagal koneksi API! Error: {e}")

    with col_spacer: st.empty()
    with col_right:
        st.markdown(f"<div class='section-title'>Top {menu_pilihan} Problem Terminal IDs</div>", unsafe_allow_html=True)
        html_t4_rows = "".join([f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:130px;' title='{row['Location']}'>{row['Location']}</td><td style='text-align:left;'>{row['Branch']}</td><td>{fmt_vis(row['PREV'], True)}</td><td>{fmt_vis(row['W1'], show_w1)}</td><td>{fmt_vis(row['W2'], show_w2)}</td><td>{fmt_vis(row['W3'], show_w3)}</td><td>{fmt_vis(row['W4'], show_w4)}</td><td style='font-weight:600;'>{fmt_vis(row['TOTAL'], True)}</td></tr>" for i, (tid, row) in enumerate(df_top_tid.iterrows(), 1)]) if not df_top_tid.empty else "<tr><td colspan='10'>No Data Available</td></tr>"
        st.markdown(f"<div class='table-scroll'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th style='text-align:left;'>Branch</th><th>{prev_month[:3].capitalize()}</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {selected_month[:3].capitalize()}</th></tr>{html_t4_rows}</table></div>", unsafe_allow_html=True)

        st.markdown(f"<div class='section-title'>{menu_pilihan} Issue Resolution Follow-up</div>", unsafe_allow_html=True)
        html_t5_rows = ""
        if not df_top_tid.empty:
            for i, tid in enumerate(df_top_tid.index, 1):
                slm_row = df_slm[df_slm['TID'] == tid]
                visit_date, action = (slm_row.iloc[0].get('TGL VISIT SLM', '-'), slm_row.iloc[0].get('ACTION', '-')) if not slm_row.empty else ("-", "Belum ada log visit teknisi terbaru")
                html_t5_rows += f"<tr><td>{i}</td><td style='font-weight:600;'>{tid}</td><td style='text-align:left; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:130px;' title='{df_top_tid.loc[tid, 'Location']}'>{df_top_tid.loc[tid, 'Location']}</td><td style='text-align:left;'>{visit_date}</td><td style='text-align:left; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:150px;' title='{action}'>{action}</td></tr>"
        else: html_t5_rows = "<tr><td colspan='5'>No Data</td></tr>"
        st.markdown(f"<div class='table-scroll'><table class='dash-table'><tr><th style='width:5%;'>NO</th><th style='width:10%;'>TID</th><th style='text-align:left; width:28%;'>Location</th><th style='text-align:left; width:15%;'>Visit Date</th><th style='text-align:left;'>Action</th></tr>{html_t5_rows}</table></div>", unsafe_allow_html=True)

        st.markdown(f"<div class='section-title'>Top {menu_pilihan} Problem Branches</div>", unsafe_allow_html=True)
        col_curr, col_prev = (('W4', 'W3') if selected_week == 'W4' else (('W3', 'W2') if selected_week == 'W3' else (('W2', 'W1') if selected_week == 'W2' else ('W1', 'PREV'))))
        df_chart = df_top_branch.head(5)
        if not df_chart.empty:
            fig = go.Figure()
            y_c, y_p = df_chart[col_curr].tolist(), df_chart[col_prev].tolist()
            l_p = prev_month[:3].capitalize() if col_prev == 'PREV' else col_prev
            fig.add_trace(go.Scatter(x=df_chart.index.tolist(), y=y_p, mode='lines+markers+text', name=l_p, text=[str(int(v)) if v > 0 else "" for v in y_p], textposition='top center', line=dict(color='#64748B', width=2, shape='spline')))
            fig.add_trace(go.Scatter(x=df_chart.index.tolist(), y=y_c, mode='lines+markers+text', name=col_curr, text=[str(int(v)) if v > 0 else "" for v in y_c], textposition='top center', line=dict(color='#F37021', width=3, shape='spline')))
            fig.update_layout(margin=dict(l=10, r=10, t=25, b=5), legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, range=[0, max(max(y_p), max(y_c), 1) * 1.4]), plot_bgcolor='rgba(0,0,0,0)', height=110)
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        
        html_t6_rows = "".join([f"<tr><td>{i}</td><td>{dict_kelolaan.get(str(branch).strip().upper(), '')}</td><td style='text-align:left; font-weight:600;'>{branch}</td><td>{fmt_vis(row['PREV'], True)}</td><td>{fmt_vis(row['W1'], show_w1)}</td><td>{fmt_vis(row['W2'], show_w2)}</td><td>{fmt_vis(row['W3'], show_w3)}</td><td>{fmt_vis(row['W4'], show_w4)}</td><td style='font-weight:600;'>{fmt_vis(row['TOTAL'], True)}</td></tr>" for i, (branch, row) in enumerate(df_top_branch.iterrows(), 1)]) if not df_top_branch.empty else "<tr><td colspan='9'>No Data Available</td></tr>"
        st.markdown(f"<div class='table-scroll' style='margin-top:-5px;'><table class='dash-table'><tr><th style='width:5%;'>No</th><th style='width:12%;'>TID Count</th><th style='text-align:left; width:22%;'>Branch_Name</th><th>{prev_month[:3].capitalize()}</th><th>W1</th><th>W2</th><th>W3</th><th>W4</th><th>Σ {selected_month[:3].capitalize()}</th></tr>{html_t6_rows}</table></div>", unsafe_allow_html=True)

# ==========================================
# 7. LOGIKA HALAMAN LOGISTIC
# ==========================================
elif menu_pilihan == "Logistic":
    @st.cache_data(ttl=3600, show_spinner=False)
    def load_logistic_block(worksheet_name, block_index):
        try:
            sh = client.open_by_key(SHEET_ID)
            ws = sh.worksheet(worksheet_name)
            data = ws.get_all_values()
            if not data: return pd.DataFrame()
            blocks, current_block = [], []
            for row in data:
                is_empty = len(row) < 2 or (str(row[0]).strip() == "" and str(row[1]).strip() == "")
                if is_empty:
                    if current_block: blocks.append(current_block); current_block = []
                else: current_block.append(row)
            if current_block: blocks.append(current_block)
            if block_index >= len(blocks): return pd.DataFrame()
            target_data = blocks[block_index]
            if len(target_data) < 2: return pd.DataFrame() 
            headers = [str(h).strip().upper() for h in target_data[0]] 
            max_cols, cleaned_data = len(headers), []
            for row in target_data[1:]:
                if len(row) > max_cols: row = row[:max_cols]
                elif len(row) < max_cols: row.extend([""] * (max_cols - len(row)))
                cleaned_data.append(row)
            df = pd.DataFrame(cleaned_data, columns=headers)
            df = df.loc[:, df.columns != '']; df = df.loc[:, ~df.columns.duplicated()]
            return df
        except Exception as e: st.error(f"Gagal Load Block Logistik {worksheet_name}: {e}"); return pd.DataFrame()

    def render_logistic_table(df, title):
        if df.empty: st.warning(f"Bah! Data {title} kosong!"); return
        st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
        
        th_html = "".join([f"<th style='text-align:left;'>{col}</th>" if col.upper() in ["KANWIL", "KANTOR LAYANAN"] else f"<th>{col}</th>" for col in df.columns])
        
        # PERBAIKAN MARIA: Mengeluarkan kondisional if-else dari dalam f-string untuk mencegah syntax error
        rows_html = "".join([
            "<tr>" + "".join([
                f"<td style='text-align:left; white-space:nowrap;'>{str(row[col]) if pd.notna(row[col]) and str(row[col]).lower() != 'nan' else '-'}</td>" if col.upper() in ["KANWIL", "KANTOR LAYANAN"] 
                else f"<td style='text-align:center;'>{str(row[col]) if pd.notna(row[col]) and str(row[col]).lower() != 'nan' else '-'}</td>"
                for col in df.columns
            ]) + "</tr>" 
            for _, row in df.iterrows()
        ])
        
        st.markdown(f"<div class='table-scroll' style='max-height: 400px;'><table class='log-table'><thead><tr>{th_html}</tr></thead><tbody>{rows_html}</tbody></table></div>", unsafe_allow_html=True)

    tab_log_utama = st.radio("MainTab", ["Stock Sparepart", "Stock Cassette", "Preventive Maintenance (PM)"], horizontal=True, label_visibility="collapsed")
    st.markdown("<br>", unsafe_allow_html=True)

    if tab_log_utama == "Stock Sparepart":
        sub_tab_spare = st.radio("SubSpare", ["Hyosung", "Wincor", "NCR"], horizontal=True, label_visibility="collapsed")
        if sub_tab_spare == "Hyosung": render_logistic_table(load_logistic_block("Stock_Sparepart", 0), "Data Stock Sparepart Hyosung")
        elif sub_tab_spare == "Wincor": render_logistic_table(load_logistic_block("Stock_Sparepart", 1), "Data Stock Sparepart Wincor")
        elif sub_tab_spare == "NCR": render_logistic_table(load_logistic_block("Stock_Sparepart", 2), "Data Stock Sparepart NCR")

    elif tab_log_utama == "Stock Cassette":
        sub_tab_cass = st.radio("SubCass", ["Stock Kaset", "Monitoring Kaset Rusak"], horizontal=True, label_visibility="collapsed")
        col_tabel, col_sisa = st.columns([8.5, 1.5]) 
        with col_tabel:
            if sub_tab_cass == "Stock Kaset":
                df_log = load_logistic_block("Stock_Cassete_Weekly", 0)
                df_log.columns = [str(c).replace(" GOOD ", "<br>GOOD ") for c in df_log.columns]
                render_logistic_table(df_log, "Inventory Stock Kaset (Weekly)")
            elif sub_tab_cass == "Monitoring Kaset Rusak":
                df_log = load_logistic_block("Stock_Cassete_Weekly", 1)
                render_logistic_table(df_log, "Laporan Monitoring Kaset Rusak")

    elif tab_log_utama == "Preventive Maintenance (PM)":
        sub_tab_pm = st.radio("SubPM", ["PM Triwulan", "PM Cassette"], horizontal=True, label_visibility="collapsed")
        col_tabel, col_sisa = st.columns([6, 4])
        with col_tabel:
            if sub_tab_pm == "PM Triwulan": render_logistic_table(load_logistic_block("Preventive_Maintenance", 0), "Jadwal PM Triwulan ATM")
            elif sub_tab_pm == "PM Cassette": render_logistic_table(load_logistic_block("Preventive_Maintenance", 1), "Monitoring PM Cassette")
else: st.info("Pilih salah satu Kategori Monitoring di Menu Navigasi atas.")
