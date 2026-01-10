import streamlit as st
import pandas as pd
import plotly.express as px 
import gspread
import os
from datetime import datetime
import html 

# =========================================================================
# 1. KONFIGURASI HALAMAN & TURBO CACHE SETUP
# =========================================================================

try:
    st.set_page_config(
        layout='wide',
        page_title="ATM Performance Dashboard",
        page_icon="📊", 
        initial_sidebar_state="collapsed"
    )
except:
    st.set_page_config(layout='wide', page_title="ATM Performance Dashboard", page_icon="📊", initial_sidebar_state="collapsed")

# --- INJECT CSS AGAR TAMPILAN FULL & BERSIH ---
st.markdown("""
    <style>
        .block-container {padding-top: 1rem; padding-bottom: 0rem; padding-left: 1rem; padding-right: 1rem;}
        header {visibility: hidden;} 
        footer {visibility: hidden;} 
        .stApp {background-color: #F8FAFC;} 
    </style>
""", unsafe_allow_html=True)


# =========================================================================
# 2. KONEKSI DATA GOOGLE SHEETS (SMART CLOUD & LOCAL - VERSI ANTI NYASAR)
# =========================================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_MAIN = 'AIMS_Master' 
SHEET_SLM = 'SLM Visit Log'
SHEET_MRI = 'Data_Form' 
SHEET_MONITORING = 'Summary Monitoring Cash'
SHEET_SP = 'Sparepart&kaset' 

# --- JURUS KUNCI LOKASI FILE (SUPAYA TIDAK NYASAR DI LOCALHOST) ---
current_dir = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(current_dir, "credentials.json")

gc = None

try:
    # --- PRIORITAS 1: CEK FILE LOKAL (DENGAN PATH LENGKAP) ---
    if os.path.exists(JSON_FILE):
        gc = gspread.service_account(filename=JSON_FILE)
    
    # --- PRIORITAS 2: CEK CLOUD SECRETS ---
    elif 'gcp_service_account' in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
    
    else:
        # Jika file benar-benar tidak ada di folder script
        st.error(f"⚠️ FATAL: File 'credentials.json' TIDAK ADA di folder: {current_dir}")

except Exception as e:
    st.error(f"⚠️ KONEKSI ERROR: {e}")


# =========================================================================
# 3. FUNGSI LOAD DATA (DENGAN LOGIKA FORMATTING KETAT)
# =========================================================================
@st.cache_data(ttl=14400, show_spinner=False)
def load_data():
    # File Backup Lokal
    backup_file = 'DATA_MASTER_ATM.xlsx'
    
    # --- FUNGSI FORMATTING ---
    def clean_and_format(df_in):
        if df_in.empty: return df_in
        df_in.columns = df_in.columns.str.strip().str.upper()
        
        if 'TANGGAL' in df_in.columns: 
            df_in['TANGGAL_OBJ'] = pd.to_datetime(df_in['TANGGAL'], errors='coerce')
            df_in['CALC_MONTH'] = df_in['TANGGAL_OBJ'].dt.strftime('%B')
        else:
            df_in['CALC_MONTH'] = None

        if 'BULAN' in df_in.columns:
            df_in['BULAN'] = df_in['BULAN'].astype(str).str.strip().str.capitalize()
            df_in['BULAN_EN'] = df_in['CALC_MONTH'].fillna(df_in['BULAN'])
        else:
            df_in['BULAN_EN'] = df_in['CALC_MONTH']
            
        if 'TANGGAL_OBJ' in df_in.columns:
            df_in['TANGGAL'] = df_in['TANGGAL_OBJ']
            df_in.drop(columns=['TANGGAL_OBJ', 'CALC_MONTH'], inplace=True)

        if 'WAKTU INSERT' in df_in.columns: df_in['WAKTU_INSERT'] = pd.to_datetime(df_in['WAKTU INSERT'], errors='coerce')
        
        # --- PERBAIKAN KOLOM COMPLAIN (PASTIKAN ANGKA) ---
        if 'JUMLAH_COMPLAIN' in df_in.columns:
            # Paksa jadi numeric, yang error jadi NaN lalu diisi 0
            df_in['JUMLAH_COMPLAIN'] = pd.to_numeric(df_in['JUMLAH_COMPLAIN'].astype(str).str.replace('-', '0'), errors='coerce').fillna(0).astype(int)
        
        if 'WEEK' not in df_in.columns and 'BULAN_WEEK' in df_in.columns: df_in['WEEK'] = df_in['BULAN_WEEK']
        
        return df_in

    # Variabel Status Koneksi
    source_status = "UNKNOWN"

    try:
        # --- PERCOBAAN A: ONLINE (GOOGLE SHEETS) ---
        if gc is None: raise Exception("No Connection") 
        
        sh = gc.open_by_url(SHEET_URL)
        
        # 1. LOAD MASTER
        ws = sh.worksheet(SHEET_MAIN)
        all_vals = ws.get_all_values()
        df = pd.DataFrame(all_vals[1:], columns=all_vals[0]) if all_vals else pd.DataFrame()
        df = clean_and_format(df) 

        # 2. LOAD SLM
        df_slm = pd.DataFrame()
        try:
            ws_slm = sh.worksheet(SHEET_SLM)
            vals_slm = ws_slm.get_all_values()
            if len(vals_slm) > 1:
                df_slm = pd.DataFrame(vals_slm[1:], columns=vals_slm[0])
                col_tgl = next((c for c in df_slm.columns if 'VISIT' in c.upper() or 'TANGGAL' in c.upper()), None)
                if col_tgl:
                    df_slm['TGL_VISIT'] = pd.to_datetime(df_slm[col_tgl], errors='coerce')
                    df_slm['BULAN_EN'] = df_slm['TGL_VISIT'].dt.strftime('%B')
                col_tid_slm = next((c for c in df_slm.columns if 'TID' in c.upper()), None)
                if col_tid_slm:
                    df_slm['TID'] = df_slm[col_tid_slm].astype(str).str.strip()
                    df_slm.rename(columns={col_tid_slm: 'TID'}, inplace=True)
        except: pass
        if 'BULAN_EN' not in df_slm.columns: df_slm['BULAN_EN'] = ''
        if 'TID' not in df_slm.columns: df_slm['TID'] = ''

        # 3. LOAD MRI
        df_mri_ops = pd.DataFrame()
        try:
            ws_mri = sh.worksheet(SHEET_MRI)
            vals_mri = ws_mri.get_all_values()
            if len(vals_mri) > 0: df_mri_ops = pd.DataFrame(vals_mri[1:], columns=vals_mri[0])
        except: pass
        
        # 4. LOAD MONITORING
        df_mon = pd.DataFrame()
        try:
            ws_mon = sh.worksheet(SHEET_MONITORING)
            vals_mon = ws_mon.get_all_values()
            if len(vals_mon) > 0: df_mon = pd.DataFrame(vals_mon)
        except: pass

        # 5. LOAD SPAREPART
        df_sp_raw = pd.DataFrame()
        try:
            ws_sp = sh.worksheet(SHEET_SP)
            vals_sp = ws_sp.get_all_values()
            if len(vals_sp) > 0: df_sp_raw = pd.DataFrame(vals_sp)
        except: pass

        source_status = "ONLINE 🟢"
        return df, df_slm, df_mri_ops, df_mon, df_sp_raw, source_status

    except Exception as e:
        # --- PERCOBAAN B: OFFLINE (LOCAL EXCEL BACKUP) ---
        if os.path.exists(backup_file):
            try:
                # Load Master
                df = pd.read_excel(backup_file, sheet_name=SHEET_MAIN, dtype=str)
                df = clean_and_format(df) 
                
                # Load SLM
                df_slm = pd.read_excel(backup_file, sheet_name=SHEET_SLM, dtype=str)
                col_tgl = next((c for c in df_slm.columns if 'VISIT' in c.upper() or 'TANGGAL' in c.upper()), None)
                if col_tgl:
                    df_slm['TGL_VISIT'] = pd.to_datetime(df_slm[col_tgl], errors='coerce')
                    df_slm['BULAN_EN'] = df_slm['TGL_VISIT'].dt.strftime('%B')
                col_tid_slm = next((c for c in df_slm.columns if 'TID' in c.upper()), None)
                if col_tid_slm:
                    df_slm['TID'] = df_slm[col_tid_slm].astype(str).str.strip()
                    df_slm.rename(columns={col_tid_slm: 'TID'}, inplace=True)
                if 'BULAN_EN' not in df_slm.columns: df_slm['BULAN_EN'] = ''
                if 'TID' not in df_slm.columns: df_slm['TID'] = ''

                try: df_mri_ops = pd.read_excel(backup_file, sheet_name=SHEET_MRI, dtype=str)
                except: df_mri_ops = pd.DataFrame()

                try: df_mon = pd.read_excel(backup_file, sheet_name=SHEET_MONITORING, header=None, dtype=str)
                except: df_mon = pd.DataFrame()
                
                try: df_sp_raw = pd.read_excel(backup_file, sheet_name=SHEET_SP, header=None, dtype=str)
                except: df_sp_raw = pd.DataFrame()

                source_status = "OFFLINE 🟠"
                return df, df_slm, df_mri_ops, df_mon, df_sp_raw, source_status
            except:
                pass

        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), "ERROR 🔴"

# --- HELPER FUNCTIONS (GLOBAL) ---
def get_prev_month_full_en(curr_month_en):
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    try: idx = months.index(curr_month_en); return months[idx - 1] if idx > 0 else months[11]
    except: return None

def clean_zeros(df_in):
    return df_in.astype(str).replace(['0', '0.0', '0.00', 'nan', 'None'], '')

# --- EKSEKUSI LOAD DATA ---
df, df_slm, df_mri_ops, df_mon, df_sp_raw, connection_status = load_data()

# Validasi Data Utama
if df.empty:
    st.warning("⚠️ Data AIMS_Master kosong atau gagal dimuat. Cek koneksi internet atau nama Sheet.")


# =========================================================================
# 4. LOGIKA HALAMAN
# =========================================================================

if 'app_mode' not in st.session_state:
    st.session_state['app_mode'] = 'cover'

# --- A. TAMPILAN HALAMAN PEMBUKA (LANDING PAGE - ULTRA SLIM & ELEGANT) ---
if st.session_state['app_mode'] == 'cover':
    # CSS KHUSUS HALAMAN COVER
    st.markdown("""
        <style>
            /* 1. Ubah Background jadi Biru Gelap Command Center */
            [data-testid="stAppViewContainer"] {
                background-color: #00172E; 
                background-image: linear-gradient(180deg, #00172E 0%, #000F1F 100%);
                color: #FFFFFF;
            }
            
            /* 2. Sembunyikan Header Bawaan */
            [data-testid="stHeader"] { visibility: hidden; }
            
            /* 3. Styling Judul Utama (Kuning Underline) */
            .cover-title {
                font-family: 'Helvetica', sans-serif;
                font-size: 38px;
                font-weight: 700;
                color: #FFC107; /* Kuning Emas */
                text-transform: uppercase;
                border-bottom: 3px solid #FFC107;
                display: inline-block;
                padding-bottom: 5px;
                margin-bottom: 20px;
                letter-spacing: 1px;
            }
            
            /* 4. Sub-Title & Info */
            .cover-subtitle {
                font-size: 18px;
                font-weight: 400;
                color: #FFFFFF;
                margin-bottom: 40px;
                letter-spacing: 1.5px;
            }
            .cover-info {
                font-size: 14px;
                color: #E2E8F0;
                margin-bottom: 8px;
                font-family: 'Inter', sans-serif;
            }
            
            /* 5. Styling Tombol Navigasi (Kotak Outline - Lebih Slim) */
            div.stButton > button {
                border-radius: 4px;
                border: 1px solid #0EA5E9; /* Biru Muda */
                background-color: rgba(14, 165, 233, 0.05); /* Transparan */
                color: #0EA5E9;
                text-align: left;
                padding-left: 15px;
                font-weight: 600;
                height: 45px; /* Sedikit lebih pendek lagi */
                text-transform: uppercase;
                transition: all 0.3s ease;
                font-size: 12px; /* Font diperhalus */
            }
            div.stButton > button:hover {
                border-color: #FFC107;
                color: #FFC107;
                background-color: rgba(255, 193, 7, 0.1);
            }
            
            /* 6. HACK: Tombol Pertama (MRI) jadi Orange Solid */
            [data-testid="column"]:nth-of-type(4) .stButton:nth-of-type(1) button {
                background-color: #D97706; /* Orange */
                color: white;
                border: none;
                font-weight: 800;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
            }
            [data-testid="column"]:nth-of-type(4) .stButton:nth-of-type(1) button:hover {
                background-color: #B45309;
                color: white;
                transform: translateY(-1px);
            }
        </style>
    """, unsafe_allow_html=True)

    # --- LAYOUT DENGAN SPACER YANG LEBIH LUAS ---
    # Rasio Baru: Tombol (1.8) jauh lebih kecil dari sebelumnya (2.5)
    # Spacer kiri-kanan diperlebar (1.2) supaya konten makin di tengah
    c_space_l, c1, c_gap, c2, c_space_r = st.columns([1.2, 4.0, 0.4, 1.8, 1.2])
    
    with c1:
        st.markdown('<div style="height: 60px;"></div>', unsafe_allow_html=True) # Spacer Atas
        
        # --- [NEW] LOGO DISPLAY LOGIC ---
        logo_file = "Logo Command Center.png"
        if os.path.exists(logo_file):
            st.image(logo_file, width=150)
            st.markdown('<div style="height: 15px;"></div>', unsafe_allow_html=True) # Spacer kecil setelah logo
        else:
            # Fallback jika gambar belum diupload user
            st.warning(f"⚠️ Logo not found: {logo_file}")

        # JUDUL WEEKLY
        st.markdown('<div class="cover-title">WEEKLY PERFORMANCE REVIEW</div>', unsafe_allow_html=True)
        st.markdown('<div class="cover-subtitle">ATM MONITORING DIVISION</div>', unsafe_allow_html=True)
        
        # Info Presenter & Tanggal
        st.markdown('<div style="height: 30px;"></div>', unsafe_allow_html=True)
        st.markdown('<div class="cover-info">Presenter : <b>Command Center BRI</b></div>', unsafe_allow_html=True)
        
        # Tanggal Otomatis
        curr_date = datetime.now().strftime("%A, %d %B %Y")
        st.markdown(f'<div class="cover-info">{curr_date}</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div style="height: 120px;"></div>', unsafe_allow_html=True) # Spacer agar sejajar visual
        
        # --- TOMBOL NAVIGASI MENU (ULTRA COMPACT) ---
        
        # 1. MRI PROJECT (Orange)
        if st.button("⭐ PROJECT MRI", use_container_width=True):
            st.session_state['nav_cat'] = 'MRI Project'
            st.session_state['app_mode'] = 'main'
            st.rerun()
            
        # 2. ELASTIC
        if st.button("01 | ELASTIC PROBLEM", use_container_width=True):
            st.session_state['nav_cat'] = 'Elastic'
            st.session_state['app_mode'] = 'main'
            st.rerun()

        # 3. COMPLAIN
        if st.button("02 | COMPLAIN HANDLING", use_container_width=True):
            st.session_state['nav_cat'] = 'Complain'
            st.session_state['app_mode'] = 'main'
            st.rerun()

        # 4. DF REPEAT
        if st.button("03 | DF REPEATED ISSUE", use_container_width=True):
            st.session_state['nav_cat'] = 'DF Repeat'
            st.session_state['app_mode'] = 'main'
            st.rerun()
            
        # 5. OUT FLM
        if st.button("04 | OUT FLM TRACKING", use_container_width=True):
            st.session_state['nav_cat'] = 'OUT Flm'
            st.session_state['app_mode'] = 'main'
            st.rerun()

        # 6. SPAREPART
        if st.button("05 | SPAREPART & KASET", use_container_width=True):
            st.session_state['nav_cat'] = 'SparePart & Kaset'
            st.session_state['app_mode'] = 'main'
            st.rerun()

# --- B. TAMPILAN DASHBOARD UTAMA ---
elif st.session_state['app_mode'] == 'main':
    
    # --- A. LOGIKA DATA HEADER ---
    try:
        h_mon = st.session_state.get('w_mon', df['BULAN_EN'].unique().tolist()[-1] if not df.empty and 'BULAN_EN' in df.columns else '')
        h_week = st.session_state.get('w_week', 'All Week')
        h_cat = st.session_state.get('nav_cat', 'MRI Project') 

        def safe_text(s):
            if pd.isna(s) or s == "": return "N/A"
            return html.escape(str(s)).replace("'", "").replace('"', "")
        
        # --- FIX: LOGIKA HITUNG HEADER AGAR KONSISTEN ---
        def get_val_safe(dframe, cat_name):
            if dframe.empty: return 0
            try:
                # JIKA KATEGORI ADALAH COMPLAIN, WAJIB SUM KOLOM J
                if cat_name == 'Complain':
                    if 'JUMLAH_COMPLAIN' in dframe.columns:
                        return int(dframe['JUMLAH_COMPLAIN'].fillna(0).sum())
                    else: return 0
                
                # JIKA KATEGORI LAIN, HITUNG JUMLAH BARIS
                return len(dframe)
            except:
                return 0

        cat_label = h_cat.upper()
        total_armada = 611 
        
        if h_cat == 'MRI Project':
            col_status = next((c for c in df.columns if 'STATUS' in c and 'MRI' in c), 'STATUS MRI')
            df_raw = df[df[col_status] == 'TID MRI'].copy() if col_status in df.columns else pd.DataFrame()
            df_target = df_raw[df_raw['KATEGORI'].isin(['Complain', 'DF Repeat'])].copy()
            cat_label = "PROJECT MRI"
        elif h_cat == 'SparePart & Kaset':
            df_target = pd.DataFrame(); cat_label = "SPAREPART"
        else:
            df_target = df[df['KATEGORI'] == h_cat].copy()

        # --- INISIALISASI LIST UPDATES DENGAN SIGNATURE MESSAGE (URUTAN 0) ---
        updates = [f"<span style='font-family: monospace; color: #64748B;'>&gt;_ SYSTEM_ORIGIN:</span> <span style='color: #1E293B; font-weight: 800; letter-spacing: 0.5px;'>COMMAND CENTER LT 3 GEDUNG BRI</span>"]
        
        if not df_target.empty:
            months_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
            try: idx_m = months_list.index(h_mon); h_prev_mon = months_list[idx_m - 1] if idx_m > 0 else months_list[11]
            except: h_prev_mon = ""

            df_curr_m = df_target[df_target['BULAN_EN'] == h_mon]
            df_prev_m = df_target[df_target['BULAN_EN'] == h_prev_mon]

            is_weekly_mode = (h_week != 'All Week')
            scope_label = h_week if is_weekly_mode else "MONTHLY"
            
            if is_weekly_mode:
                try: w_num = int(h_week.replace('W','')); prev_w_str = f"W{w_num-1}" if w_num > 1 else ""
                except: prev_w_str = ""
                df_scope_curr = df_curr_m[df_curr_m['WEEK'] == h_week]
                df_scope_prev = df_curr_m[df_curr_m['WEEK'] == prev_w_str] if prev_w_str else pd.DataFrame()
            else:
                df_scope_curr = df_curr_m; df_scope_prev = df_prev_m

            # MONTHLY
            val_m_curr = get_val_safe(df_curr_m, h_cat)
            val_m_prev = get_val_safe(df_prev_m, h_cat)
            diff_m = val_m_curr - val_m_prev
            pct_m = (diff_m / val_m_prev * 100) if val_m_prev > 0 else 100.0 if val_m_curr > 0 else 0.0
            icon_m = "🔺" if diff_m > 0 else "🔻"; color_m = "#DC2626" if diff_m > 0 else "#16A34A"
            updates.append(f"<span style='color: #64748B;'>[MONTHLY] Total {cat_label}: <b>{val_m_curr}</b> Tiket (<span style='color: {color_m}; font-weight: 800;'>{icon_m} {diff_m} / {pct_m:.1f}%</span> vs {h_prev_mon})</span>")

            # SUMMARY
            val_s_curr = get_val_safe(df_scope_curr, h_cat)
            val_s_prev = get_val_safe(df_scope_prev, h_cat)
            diff_s = val_s_curr - val_s_prev
            pct_s = (diff_s / val_s_prev * 100) if val_s_prev > 0 else 100.0 if val_s_curr > 0 else 0.0
            diff_str = f"+{diff_s}" if diff_s > 0 else str(diff_s)
            pct_str = f"+{pct_s:.1f}%" if pct_s > 0 else f"{pct_s:.1f}%"
            color_s = "#DC2626" if diff_s > 0 else "#16A34A"
            updates.append(f"<span style='color: #64748B;'>[{scope_label}] Kategori {cat_label}: <b>{val_s_curr}</b> Tiket. Selisih: <span style='color: {color_s}; font-weight: 800;'>{diff_str} ({pct_str})</span> vs periode lalu.</span>")

            # RECURRING
            if is_weekly_mode and not df_scope_curr.empty and not df_scope_prev.empty and 'TID' in df_scope_curr.columns:
                tids_now = set(df_scope_curr['TID']); tids_bef = set(df_scope_prev['TID'])
                rec_tids = tids_now.intersection(tids_bef)
                cnt_rec = len(rec_tids)
                if cnt_rec > 0:
                    top_rec_list = [safe_text(x) for x in list(rec_tids)[:3]]
                    top_rec_str = ", ".join(top_rec_list)
                    updates.append(f"<span style='color: #64748B;'>[RECURRING] Waspada! Ada <span style='color: #F59E0B; font-weight: 800;'>{cnt_rec} Unit</span> Masalah Berulang dari {prev_w_str} ke {h_week}. (Contoh: {top_rec_str}...)</span>")

            # BRANCH TREND
            if 'CABANG' in df_scope_curr.columns:
                def agg_branch(df_in):
                    if h_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_in.columns:
                        return df_in.groupby('CABANG')['JUMLAH_COMPLAIN'].sum()
                    return df_in['CABANG'].value_counts()

                vc_c_curr = agg_branch(df_scope_curr)
                vc_c_prev = agg_branch(df_scope_prev)
                
                all_cabs = set(vc_c_curr.index).union(set(vc_c_prev.index))
                c_diffs = []
                for c in all_cabs:
                    v1 = vc_c_curr.get(c,0); v0 = vc_c_prev.get(c,0); d = v1 - v0
                    p = (d/v0*100) if v0>0 else 100.0 if v1>0 else 0.0
                    c_diffs.append({'CAB': safe_text(c), 'DIFF': d, 'PCT': p, 'VAL': int(v1)})
                
                df_cd = pd.DataFrame(c_diffs)
                if not df_cd.empty:
                    worst_c = df_cd.sort_values('DIFF', ascending=False).iloc[0]
                    if worst_c['DIFF'] > 0: 
                        updates.append(f"<span style='color: #64748B;'>[BRANCH RISE] Cabang <b>{worst_c['CAB']}</b> ({cat_label}) NAIK <span style='color: #DC2626; font-weight: 800;'>+{int(worst_c['DIFF'])}</span> Tiket (+{worst_c['PCT']:.0f}%) Total: {worst_c['VAL']}.</span>")
                    best_c = df_cd.sort_values('DIFF', ascending=True).iloc[0]
                    if best_c['DIFF'] < 0: 
                        updates.append(f"<span style='color: #64748B;'>[BRANCH DROP] Cabang <b>{best_c['CAB']}</b> ({cat_label}) TURUN <span style='color: #16A34A; font-weight: 800;'>{int(best_c['DIFF'])}</span> Tiket ({best_c['PCT']:.0f}%) Total: {best_c['VAL']}.</span>")

            # TID TREND
            if 'TID' in df_scope_curr.columns:
                def agg_tid(df_in):
                    if h_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_in.columns:
                        return df_in.groupby('TID')['JUMLAH_COMPLAIN'].sum()
                    return df_in['TID'].value_counts()

                vc_t_curr = agg_tid(df_scope_curr)
                vc_t_prev = agg_tid(df_scope_prev)
                
                def get_loc_info(tid_target):
                    try:
                        row = df_target[df_target['TID'] == tid_target].iloc[0]
                        return f"{safe_text(row.get('LOKASI',''))} ({safe_text(row.get('CABANG',''))})"
                    except: return "Lokasi N/A"

                all_tids = set(vc_t_curr.index).union(set(vc_t_prev.index))
                t_diffs = []
                for t in all_tids:
                    v1 = vc_t_curr.get(t,0); v0 = vc_t_prev.get(t,0); d = v1 - v0
                    p = (d/v0*100) if v0>0 else 100.0 if v1>0 else 0.0
                    t_diffs.append({'TID': safe_text(t), 'DIFF': d, 'PCT': p, 'VAL': int(v1)})

                df_td = pd.DataFrame(t_diffs)
                if not df_td.empty:
                    worst_t = df_td.sort_values('DIFF', ascending=False).iloc[0]
                    if worst_t['DIFF'] > 0: 
                        loc_info = get_loc_info(worst_t['TID'])
                        updates.append(f"<span style='color: #64748B;'>[TID RISE] Unit <b>{worst_t['TID']}</b> [{loc_info}] NAIK <span style='color: #DC2626; font-weight: 800;'>+{int(worst_t['DIFF'])}</span> Problem (+{worst_t['PCT']:.0f}%) Total: {worst_t['VAL']}x.</span>")
                    best_t = df_td.sort_values('DIFF', ascending=True).iloc[0]
                    if best_t['DIFF'] < 0: 
                        loc_info = get_loc_info(best_t['TID'])
                        updates.append(f"<span style='color: #64748B;'>[TID DROP] Unit <b>{best_t['TID']}</b> [{loc_info}] TURUN <span style='color: #16A34A; font-weight: 800;'>{int(best_t['DIFF'])}</span> Problem ({best_t['PCT']:.0f}%) Total: {best_t['VAL']}x.</span>")

        # UPDATE GLOBAL ASSET (Ditaruh di akhir)
        updates.append(f"<span style='color: #64748B;'>🌍 GLOBAL ASSETS: <span style='color: #1E293B; font-weight: 800;'>{total_armada}</span> Units Active</span>")

        msg_count = len(updates)
        TIME_SHOW = 8.0; TIME_GAP = 12.0; CYCLE_TIME = TIME_SHOW + TIME_GAP
        TOTAL_DURATION = max(msg_count * CYCLE_TIME, 1)
        PCT_VISIBLE = (TIME_SHOW / TOTAL_DURATION) * 100
        
        fade_html = ""
        for i, item in enumerate(updates):
            delay = i * CYCLE_TIME
            fade_html += f'<div class="whisper-item" style="animation-delay: {delay}s; animation-duration: {TOTAL_DURATION}s;">{item}</div>'

    except Exception as e:
        PCT_VISIBLE = 5.0
        TOTAL_DURATION = 10
        fade_html = f'<div class="whisper-item" style="color:orange;">⚠️ System Syncing... (Check Data Format)</div>'


    # --- B. RENDER LAYOUT HEADER ---
    head_c1, head_c2, head_c3 = st.columns([2.5, 7.0, 2.5])

    with head_c1:
        st.markdown("""
        <div style="line-height: 1.1;">
            <div class="main-title" style="font-size: 30px !important;">ATM WEEKLY PERFORMANCE</div>
            <div class="sub-title" style="font-size: 10px !important; color: #64748B; margin-top: -2px;">PT KELOLA JASA ARTA</div>
        </div>
        """, unsafe_allow_html=True)

    with head_c2:
        css_style = f"""
        <style>
            @keyframes strictSequence {{
                0% {{ opacity: 0; transform: translateY(5px); }} 
                1% {{ opacity: 1; transform: translateY(0px); }} 
                {PCT_VISIBLE:.2f}% {{ opacity: 1; transform: translateY(0px); }} 
                {PCT_VISIBLE + 1.0:.2f}% {{ opacity: 0; transform: translateY(-5px); }} 
                100% {{ opacity: 0; transform: translateY(-5px); }}
            }}
            
            .whisper-container {{ 
                position: relative; 
                height: 35px; 
                margin-top: 5px; 
                display: flex; 
                justify-content: center; 
                align-items: center; 
                width: 100%; 
                overflow: hidden; 
            }}
            .whisper-item {{ 
                position: absolute; 
                width: 100%; 
                opacity: 0; 
                font-family: 'Inter', sans-serif; 
                font-size: 13px; 
                text-align: center; 
                animation-name: strictSequence; 
                animation-timing-function: linear; 
                animation-iteration-count: infinite;
                top: 0; 
                white-space: nowrap;
            }}
        </style>
        """
        st.markdown(css_style + f'<div class="whisper-container">{fade_html}</div>', unsafe_allow_html=True)

    with head_c3:
        # LOGIKA INDIKATOR STATUS DI HEADER
        curr_date = datetime.now().strftime("%d %B %Y")
        
        if "ONLINE" in connection_status:
            status_bg = "#16A34A" # Hijau
            status_text = "ONLINE"
            status_icon = "☁️"
        else:
            status_bg = "#F59E0B" # Orange
            status_text = "OFFLINE"
            status_icon = "📂"

        st.markdown(f"""
        <div style="display: flex; flex-direction: column; align-items: flex-end; width: 100%; margin-right: -10px;">
            <div style="display: flex; gap: 6px; align-items: center; margin-bottom: 2px;">
                 <div style="background-color: {status_bg}; color: white; font-size: 9px; padding: 2px 8px; border-radius: 4px; font-weight: 800; letter-spacing: 0.5px;">
                    {status_icon} {status_text}
                 </div>
                 <div class="date-pill" style="font-size: 10px !important; padding: 2px 8px;">📅 {curr_date}</div>
            </div>
            <div style="font-size: 10px; font-weight: 700; color: #16A34A;">
                LIVE <span id="clock_ticks">--:--:--</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.components.v1.html(
            """
            <script>
                function updateClock() {
                    const now = new Date();
                    const timeString = now.toLocaleTimeString('en-GB', {
                        hour: '2-digit', 
                        minute: '2-digit', 
                        second: '2-digit'
                    });
                    const target = window.parent.document.getElementById('clock_ticks');
                    if (target) {
                        target.innerText = timeString;
                    }
                }
                setInterval(updateClock, 1000);
                updateClock(); 
            </script>
            """,
            height=0,
            width=0
        )

    # --- THEME ENGINE ---
    if 'theme_mode' not in st.session_state:
        st.session_state.theme_mode = False

    use_exec_mode = st.session_state.theme_mode

    if use_exec_mode:
        primary_color = "#0F172A"; secondary_color = "#1E3A8A"; header_bg = "#F8FAFC"; text_color = "#1E293B"
        info_box_bg = "#EFF6FF"; info_box_border = "#1E3A8A"; chart_colors = ['#0F172A', '#94A3B8']
    else:
        primary_color = "#00529C"; secondary_color = "#00386B"; header_bg = "#FFFFFF"; text_color = "#1E293B"
        info_box_bg = "#EFF6FF"; info_box_border = "#60A5FA"; chart_colors = ['#00529C', '#60A5FA']

    st.markdown(f"""
    <style>
    .top-header-bar {{ position: fixed; top: 0; left: 0; width: 100%; height: 8px; background: linear-gradient(90deg, {primary_color} 0%, {secondary_color} 100%); z-index: 99999; }}
    .stDeployButton, [data-testid="stHeader"], [data-testid="stToolbar"] {{ display: none !important; }}
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; color: {text_color}; }}
    [data-testid="stAppViewContainer"] {{ background-color: #F8FAFC; }}
    .block-container {{ padding-top: 1.5rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; padding-bottom: 1rem !important; max-width: 100%; }}
    .element-container, .stMarkdown {{ margin-bottom: -2px !important; }}
    [data-testid="column"] {{ gap: 0px !important; }}
    .main-title {{ font-size: 20px; font-weight: 800; color: {primary_color}; letter-spacing: -0.5px; margin: 0; line-height: 1.2; }}
    .section-header {{ background: {primary_color}; color: #fff; padding: 5px 10px; border-radius: 4px 4px 0 0; font-size: 11px; font-weight: 700; text-transform: uppercase; }}
    .highlight-value {{ font-size: 17px; color: {primary_color}; font-weight: 800; }}
    [data-testid="stDataFrame"] th {{ background-color: {header_bg} !important; color: #334155 !important; border-bottom: 2px solid #E2E8F0 !important; text-align: center !important; }}
    [data-testid="stDataFrame"] td {{ color: {text_color} !important; }}
    div[role="radiogroup"] label[data-checked="true"] {{ background: {primary_color} !important; color: #fff !important; border-color: {primary_color} !important; }}
    button[data-baseweb="tab"][aria-selected="true"] {{ background-color: #F1F5F9 !important; border-bottom-color: {primary_color} !important; color: {primary_color} !important; }}
    div[data-baseweb="notification"] {{ background-color: {info_box_bg} !important; border-left: 4px solid {info_box_border} !important; color: {text_color} !important; }}
    div[data-baseweb="select"] > div {{ background-color: #FFFFFF !important; border-radius: 8px !important; border: 1px solid #E2E8F0 !important; height: 38px !important; overflow: visible !important; }}
    div[data-baseweb="select"] {{ margin-top: 6px !important; }}
    div[role="radiogroup"] {{ justify-content: flex-end; margin-bottom: 8px; gap: 6px; }}
    div[role="radiogroup"] label {{ background: #fff; border: 1px solid #CBD5E1; border-radius: 50px; padding: 3px 14px; font-size: 11px; font-weight: 600; color: #475569; margin-top: 6px !important; }}
    .table-card {{ background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 12px 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 8px; }}
    .table-title {{ color: {primary_color}; font-size: 13px; font-weight: 700; margin-bottom: 18px !important; border-left: 4px solid {primary_color}; padding-left: 8px; line-height: 1; display: block; }}
    [data-testid="stPlotlyChart"] {{ border: 1px solid #E2E8F0; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); background-color: #FFFFFF; }}
    [data-testid="column"]:nth-of-type(3) {{ display: flex; flex-direction: column; align-items: flex-end !important; }}
    </style>
    <div class="top-header-bar"></div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <style>
    div[role="radiogroup"] {
        margin-top: -70px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <style>
    div[data-testid="column"]:nth-of-type(2) {
        margin-top: -50px !important;
    }
    div[data-baseweb="select"],
    div[data-testid="stCheckbox"],
    div[data-testid="stToggle"] {
        margin-top: -50px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # --- NAVIGASI & FILTER ---
    nav_col, filter_col = st.columns([2.5, 1.8], gap="medium")

    with nav_col:
        st.markdown("""<style>div[role="radiogroup"] { justify-content: flex-start !important; flex-wrap: nowrap !important; width: 100% !important; } div[role="radiogroup"] label { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)
        menu_items = ['MRI Project', 'Elastic', 'Complain', 'DF Repeat', 'OUT Flm', 'SparePart & Kaset']
        sel_cat = st.radio("Navigasi:", menu_items, index=0, horizontal=True, label_visibility="collapsed", key="nav_cat")

    # --- MEMORY STATE ---
    months_en = df['BULAN_EN'].unique().tolist() if not df.empty and 'BULAN_EN' in df.columns else []
    default_mon = months_en[-1] if months_en else None
    if 'p_mon' not in st.session_state: st.session_state.p_mon = default_mon
    if 'p_week' not in st.session_state: st.session_state.p_week = 'All Week'
    if 'p_trend' not in st.session_state: st.session_state.p_trend = 'W1 vs W2'

    def save_mon(): st.session_state.p_mon = st.session_state.w_mon
    def save_week(): st.session_state.p_week = st.session_state.w_week
    def save_trend(): st.session_state.p_trend = st.session_state.w_trend

    sel_mon = ""; prev_mon = ""; curr_mon_short = ""; prev_mon_short = ""; sort_week = "All Week"; comp_mode = ""
    use_color = False 

    if sel_cat != 'SparePart & Kaset':
        with filter_col:
            f1, f2, f3, f4, f5 = st.columns([1.4, 1.2, 1.0, 0.6, 0.7], gap="small")
            with f1:
                try: cur_ix_mon = months_en.index(st.session_state.p_mon)
                except: cur_ix_mon = 0
                sel_mon = st.selectbox("Periode:", months_en, index=cur_ix_mon, key='w_mon', on_change=save_mon, label_visibility="collapsed")
            with f2:
                opts_week = ['All Week', 'W1', 'W2', 'W3', 'W4']
                try: cur_ix_week = opts_week.index(st.session_state.p_week)
                except: cur_ix_week = 0
                sort_week = st.selectbox("Week:", opts_week, index=cur_ix_week, key='w_week', on_change=save_week, label_visibility="collapsed")
            with f3:
                opts_trend = ['W1 vs W2', 'W2 vs W3', 'W3 vs W4']
                try: cur_ix_trend = opts_trend.index(st.session_state.p_trend)
                except: cur_ix_trend = 0
                comp_mode = st.selectbox("Tren:", opts_trend, index=cur_ix_trend, key='w_trend', on_change=save_trend, label_visibility="collapsed")
            with f4:
                use_color = st.toggle("🎨", key=f"color_btn_{sel_cat}", help="Indikator Warna")
            with f5:
                exec_toggle = st.toggle("🌙", value=st.session_state.theme_mode, key=f"theme_switch_{sel_cat}", help="Executive Mode")
                if exec_toggle != st.session_state.theme_mode:
                    st.session_state.theme_mode = exec_toggle
                    st.rerun()

        if not sel_mon: sel_mon = st.session_state.p_mon
        if not sort_week: sort_week = st.session_state.p_week
        prev_mon = get_prev_month_full_en(sel_mon)
        curr_mon_short = sel_mon[:3] if sel_mon else ""
        prev_mon_short = prev_mon[:3] if prev_mon else "Prev"
    else:
        st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)

    # --- LOGIKA DATA PROCESSING ---
    df_curr = pd.DataFrame(); df_prev = pd.DataFrame(); total_ticket = 0; avg_ticket = 0

    if sel_cat == 'MRI Project':
        col_status = next((c for c in df.columns if 'STATUS' in c and 'MRI' in c), 'STATUS MRI')
        df_raw_mri = df[(df['BULAN_EN'] == sel_mon) & (df[col_status] == 'TID MRI')].copy() if col_status in df.columns else pd.DataFrame()
        df_curr = df_raw_mri[df_raw_mri['KATEGORI'].isin(['Complain', 'DF Repeat'])].copy()
        
        df_prev_raw = df[(df['BULAN_EN'] == prev_mon) & (df[col_status] == 'TID MRI')].copy() if prev_mon and col_status in df.columns else pd.DataFrame()
        df_prev = df_prev_raw[df_prev_raw['KATEGORI'].isin(['Complain', 'DF Repeat'])].copy()

    elif sel_cat == 'SparePart & Kaset': pass
    else:
        df_curr = df[(df['BULAN_EN'] == sel_mon) & (df['KATEGORI'] == sel_cat)].copy()
        if prev_mon: df_prev = df[(df['BULAN_EN'] == prev_mon) & (df['KATEGORI'] == sel_cat)].copy()

    if sel_cat != 'SparePart & Kaset' and sort_week != 'All Week':
        week_map = {'W1': 1, 'W2': 2, 'W3': 3, 'W4': 4}
        limit_num = week_map.get(sort_week, 4)
        if not df_curr.empty and 'WEEK' in df_curr.columns:
            df_curr['TEMP_W_NUM'] = df_curr['WEEK'].map(week_map).fillna(0)
            df_curr = df_curr[df_curr['TEMP_W_NUM'] <= limit_num].copy()
            df_curr.drop(columns=['TEMP_W_NUM'], inplace=True)

    # --- HITUNG TOTAL (Revisi: Jika Complain, Sum Kolom J) ---
    if sel_cat == 'Complain':
        # Pastikan kolom ada dan numerik (safety)
        if 'JUMLAH_COMPLAIN' in df_curr.columns:
            total_ticket = int(df_curr['JUMLAH_COMPLAIN'].fillna(0).sum())
        else: total_ticket = 0
    else:
        total_ticket = len(df_curr)

    avg_ticket = total_ticket / 4 if sel_cat != 'SparePart & Kaset' else 0

    # --- MICRO METRICS SECTION ---
    if sel_cat != 'SparePart & Kaset':
        if sel_cat == 'MRI Project':
            col_status = next((c for c in df.columns if 'STATUS' in c and 'MRI' in c), 'STATUS MRI')
            df_raw_mri = df[(df['BULAN_EN'] == sel_mon) & (df[col_status] == 'TID MRI')].copy() if col_status in df.columns else pd.DataFrame()
            df_met = df_raw_mri[df_raw_mri['KATEGORI'].isin(['Complain', 'DF Repeat'])].copy()
            
            df_prev_raw_mri = df[(df['BULAN_EN'] == prev_mon) & (df[col_status] == 'TID MRI')].copy() if prev_mon and col_status in df.columns else pd.DataFrame()
            df_prev_met = df_prev_raw_mri[df_prev_raw_mri['KATEGORI'].isin(['Complain', 'DF Repeat'])].copy()
        else:
            df_met = df[(df['BULAN_EN'] == sel_mon) & (df['KATEGORI'] == sel_cat)].copy()
            df_prev_met = df[(df['BULAN_EN'] == prev_mon) & (df['KATEGORI'] == sel_cat)].copy() if prev_mon else pd.DataFrame()

        # FIX METRICS: GUNAKAN SUM UNTUK COMPLAIN
        if sel_cat == 'Complain':
            total_t = int(df_met['JUMLAH_COMPLAIN'].sum()) if 'JUMLAH_COMPLAIN' in df_met.columns else 0
            prev_t = int(df_prev_met['JUMLAH_COMPLAIN'].sum()) if 'JUMLAH_COMPLAIN' in df_prev_met.columns else 0
        else:
            total_t = len(df_met)
            prev_t = len(df_prev_met)

        avg_t = total_t / 4
        diff_t = total_t - prev_t
        t_icon = "▲" if diff_t > 0 else ("▼" if diff_t < 0 else "•")
        t_color = "#DC2626" if diff_t > 0 else ("#16A34A" if diff_t < 0 else "#64748B")

        pill_bg = "#1E293B" if use_exec_mode else "#FFFFFF"
        pill_text = "#F8FAFC" if use_exec_mode else "#1E293B"
        pill_border = "#334155" if use_exec_mode else "#E2E8F0"
        label_color = "#94A3B8" if use_exec_mode else "#64748B"

        metric_html = f"""
        <div style="display: flex; gap: 12px; margin-top: -10px; margin-bottom: 12px; padding-left: 2px;">
            <div style="display: flex; align-items: center; background: {pill_bg}; border: 1px solid {pill_border}; padding: 4px 14px; border-radius: 50px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <span style="font-size: 11px; font-weight: 700; color: {label_color}; margin-right: 8px; text-transform: uppercase;">TOTAL {sel_cat}</span>
                <span style="font-size: 16px; font-weight: 800; color: {pill_text};">{total_t}</span>
                <span style="font-size: 11px; font-weight: 800; color: {t_color}; margin-left: 8px;">{t_icon} {abs(diff_t)}</span>
            </div>
            <div style="display: flex; align-items: center; background: {pill_bg}; border: 1px solid {pill_border}; padding: 4px 14px; border-radius: 50px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                <span style="font-size: 11px; font-weight: 700; color: {label_color}; margin-right: 8px; text-transform: uppercase;">AVG WEEKLY</span>
                <span style="font-size: 16px; font-weight: 800; color: {pill_text};">{avg_t:.1f}</span>
            </div>
        </div>
        """
        st.markdown(metric_html, unsafe_allow_html=True)

    # --- MAIN CONTENT RENDERING ---
    st.markdown("""
        <style>
            [data-testid="stDataFrame"], .stDataFrame {
                margin-top: 6px !important;
            }
            .section-header {
                margin-top: 5px !important;
            }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom: 5px;'></div>", unsafe_allow_html=True) 

    # --- UNIVERSAL STYLING FUNCTION (FIXED BUG) ---
    def get_styled_dataframe(df_in):
        # 1. Create Base Styler
        styler = df_in.style

        # 2. Logic Warna Merah/Hijau (Jika toggle ON)
        if use_color:
            def style_logic(row):
                color_bad = 'color: #B91C1C; font-weight: 700;' 
                color_good = 'color: #15803D; font-weight: 700;' 
                styles = [''] * len(row)
                
                def get_val(val):
                    try: return float(val) if val != "" else 0
                    except: return 0
                
                col_names = row.index.tolist()
                chain = [('W2', 'W1'), ('W3', 'W2'), ('W4', 'W3')]
                
                for curr_col, prev_col in chain:
                    if curr_col in col_names and prev_col in col_names:
                        try:
                            curr_idx = col_names.index(curr_col)
                            prev_idx = col_names.index(prev_col)
                            curr_val = get_val(row[curr_idx])
                            prev_val = get_val(row[prev_idx])
                            
                            if curr_val > prev_val: styles[curr_idx] = color_bad
                            elif curr_val < prev_val: styles[curr_idx] = color_good
                        except: pass
                return styles
            
            try:
                styler = styler.apply(style_logic, axis=1)
            except: pass

        # 3. Logic Warna Kolom (Dec & Jan) - UNIVERSAL (Always On)
        # Prev Month (Dec) -> Very subtle Grey
        if prev_mon_short in df_in.columns:
            styler = styler.map(lambda x: 'background-color: #F9FAFB; color: #444;', subset=[prev_mon_short])
        
        # Current Total (Jan) -> Very subtle Blue + Bold
        col_total_curr = f'Σ {curr_mon_short}'
        if col_total_curr in df_in.columns:
            styler = styler.map(lambda x: 'background-color: #F0F9FF; color: #000; font-weight: 600;', subset=[col_total_curr])

        return styler

    if sel_cat == 'SparePart & Kaset':
        st.markdown("""<style>[data-testid="stDataFrame"] th { font-size: 10px !important; background-color: #F8FAFC !important; }[data-testid="stDataFrame"] td { font-size: 10px !important; }</style>""", unsafe_allow_html=True)
        
        def make_unique_df(subset_data):
            try:
                raw_h = [str(x).strip() if str(x).strip() != "" else "Info" for x in subset_data.iloc[0]]
                final_h = []
                counts = {}
                for h in raw_h:
                    if h in counts:
                        counts[h] += 1
                        final_h.append(f"{h}_{counts[h]}")
                    else:
                        counts[h] = 0
                        final_h.append(h)
                return pd.DataFrame(subset_data.values[1:], columns=final_h)
            except: return pd.DataFrame()

        subset_kaset = df_sp_raw.iloc[11:22, 0:12]
        manual_headers = ["CABANG", "JML TID", "NOV GOOD CURRENT", "NOV GOOD REJECT", "W1 DEC GOOD CURRENT", "W1 DEC GOOD REJECT", "W2 DEC GOOD REJECT", "W2 DEC GOOD CURRENT", "W3 DEC GOOD CURRENT", "W3 DEC GOOD REJECT", "W4 DEC GOOD CURRENT", "W4 DEC GOOD REJECT"]
        df_kaset_final = pd.DataFrame(subset_kaset.values[1:], columns=manual_headers)
        df_kaset_final = df_kaset_final[(df_kaset_final['CABANG'].str.strip() != "") & (df_kaset_final['CABANG'].notna()) & (df_kaset_final['CABANG'].str.upper() != "CABANG")]

        tab1, tab2, tab3 = st.tabs(["🛠️ Stock Sparepart", "📼 Stock Kaset", "⚠️ Monitoring & PM"])
        
        with tab1:
            st.markdown('<div class="section-header">🛠️ Ketersediaan SparePart</div>', unsafe_allow_html=True)
            df_sp_clean = make_unique_df(df_sp_raw.iloc[0:10, 0:22])
            st.dataframe(df_sp_clean, use_container_width=True, hide_index=True)

        with tab2:
            st.markdown('<div class="section-header">📼 Ketersediaan Kaset</div>', unsafe_allow_html=True)
            for col in df_kaset_final.columns:
                if "CABANG" not in col:
                    try:
                        n = pd.to_numeric(df_kaset_final[col].astype(str).str.replace('%',''), errors='coerce')
                        df_kaset_final[col] = n.apply(lambda x: f"{x:.0%}" if (pd.notnull(x) and x <= 1.5) else (f"{x:.0f}" if pd.notnull(x) else ""))
                    except: pass
            st.dataframe(df_kaset_final, use_container_width=True, hide_index=True)

        with tab3:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown('<div class="section-header">⚠️ Rekap Kaset Rusak</div>', unsafe_allow_html=True)
                df_rsk = make_unique_df(df_sp_raw.iloc[23:27, 0:6])
                st.dataframe(df_rsk, use_container_width=True, hide_index=True)
                
            with c2:
                st.markdown('<div class="section-header">🧹 PM Kaset</div>', unsafe_allow_html=True)
                df_pm = make_unique_df(df_sp_raw.iloc[31:39, 0:7])
                st.dataframe(df_pm, use_container_width=True, hide_index=True)


    elif sel_cat == 'MRI Project':
        col_left, col_right = st.columns(2, gap="medium")
        df_mri_comp = df_curr[df_curr['KATEGORI'] == 'Complain'].copy()
        df_mri_df   = df_curr[df_curr['KATEGORI'] == 'DF Repeat'].copy()
        df_prev_comp = df_prev[df_prev['KATEGORI'] == 'Complain'].copy() if not df_prev.empty else pd.DataFrame()
        df_prev_df   = df_prev[df_prev['KATEGORI'] == 'DF Repeat'].copy() if not df_prev.empty else pd.DataFrame()
        total_atm_mri = 34 

        # --- FUNGSI KHUSUS UNTUK MEMBEDAKAN CARA HITUNG TIER MRI ---
        def calc_mri_tiers_fixed(dframe, category_type):
            if dframe.empty: return 0, 0, 0
            
            # Jika Complain: SUM kolom JUMLAH_COMPLAIN
            if category_type == 'Complain' and 'JUMLAH_COMPLAIN' in dframe.columns:
                counts = dframe.groupby('TID')['JUMLAH_COMPLAIN'].sum()
            
            # Jika DF Repeat (atau lainnya): HITUNG FREKUENSI TID (Baris)
            else:
                counts = dframe['TID'].value_counts()
                
            return (counts == 1).sum(), ((counts >= 2) & (counts <= 3)).sum(), (counts > 3).sum()

        with col_left:
            st.markdown(f'<div class="section-header">🔴 Summary Problem TID MRI</div>', unsafe_allow_html=True)
            sum_data = {"TOTAL ATM": [total_atm_mri], "Complain": [len(df_mri_comp)], "DF": [len(df_mri_df)]}
            st.dataframe(clean_zeros(pd.DataFrame(sum_data)), use_container_width=True, hide_index=True)
            
            # 1. JML COMPLAIN (Color)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">📊 JML Complain</div>', unsafe_allow_html=True)
            jml_data = { "TOTAL ATM": [total_atm_mri], f"{prev_mon_short}": [len(df_prev_comp)], "W1": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W1'])], "W2": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W2'])], "W3": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W3'])], "W4": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W4'])], f"Σ {curr_mon_short}": [len(df_mri_comp)] }
            st.dataframe(get_styled_dataframe(clean_zeros(pd.DataFrame(jml_data))), use_container_width=True, hide_index=True)
            
            # 2. TIERING COMPLAIN (Pakai Logic 'Complain' -> SUM + ADD TOTAL ROW)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">⚠️ Tiering Complain</div>', unsafe_allow_html=True)
            p_t = calc_mri_tiers_fixed(df_prev_comp, 'Complain'); c_t = calc_mri_tiers_fixed(df_mri_comp, 'Complain')
            def get_w_risk_mri(df_target, w, cat_type): return calc_mri_tiers_fixed(df_target[df_target['WEEK'] == w], cat_type)
            
            w1_t = get_w_risk_mri(df_mri_comp, 'W1', 'Complain'); w2_t = get_w_risk_mri(df_mri_comp, 'W2', 'Complain')
            w3_t = get_w_risk_mri(df_mri_comp, 'W3', 'Complain'); w4_t = get_w_risk_mri(df_mri_comp, 'W4', 'Complain')
            
            # Create Tier Data
            col_tot = f'Σ {curr_mon_short}'
            tier_data_mri = { 'TIERING': ['1 kali', '2-3 kali', '> 3 kali'], f'{prev_mon_short}': [p_t[0], p_t[1], p_t[2]], 'W1': [w1_t[0], w1_t[1], w1_t[2]], 'W2': [w2_t[0], w2_t[1], w2_t[2]], 'W3': [w3_t[0], w3_t[1], w3_t[2]], 'W4': [w4_t[0], w4_t[1], w4_t[2]], col_tot: [c_t[0], c_t[1], c_t[2]] }
            df_tier_mri = pd.DataFrame(tier_data_mri)
            
            # Add TOTAL UNIT Row
            total_row_mri = {
                'TIERING': 'TOTAL UNIT',
                f'{prev_mon_short}': df_tier_mri[f'{prev_mon_short}'].sum(),
                'W1': df_tier_mri['W1'].sum(),
                'W2': df_tier_mri['W2'].sum(),
                'W3': df_tier_mri['W3'].sum(),
                'W4': df_tier_mri['W4'].sum(),
                col_tot: df_tier_mri[col_tot].sum()
            }
            df_tier_mri = pd.concat([df_tier_mri, pd.DataFrame([total_row_mri])], ignore_index=True)
            
            # Styling for Total Row
            def highlight_total_mri(x):
                df1 = pd.DataFrame('', index=x.index, columns=x.columns)
                try: df1.iloc[3, :] = 'font-weight: 800; background-color: rgba(128, 128, 128, 0.1); border-top: 2px solid #94A3B8;'
                except: pass
                return df1

            st.dataframe(get_styled_dataframe(clean_zeros(df_tier_mri)).apply(highlight_total_mri, axis=None), use_container_width=True, hide_index=True)
            
            # 3. TOP TID COMPLAIN (MODIFIKASI: SCROLLABLE & SORT BY DROPDOWN)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">🔥 Top Complain Problem Terminal IDs</div>', unsafe_allow_html=True)
            if not df_mri_comp.empty or not df_prev_comp.empty:
                # A. Pivot Current Data (W1-W4)
                if not df_mri_comp.empty:
                    if 'JUMLAH_COMPLAIN' in df_mri_comp.columns:
                        piv = df_mri_comp.pivot_table(index=['TID','LOKASI','CABANG','TYPE MRI'], columns='WEEK', values='JUMLAH_COMPLAIN', aggfunc='sum', fill_value=0).reset_index()
                    else:
                        piv = df_mri_comp.pivot_table(index=['TID','LOKASI','CABANG','TYPE MRI'], columns='WEEK', aggfunc='size', fill_value=0).reset_index()
                else:
                    piv = pd.DataFrame(columns=['TID','LOKASI','CABANG','TYPE MRI'])

                # B. Prepare Previous Data (Dec)
                if not df_prev_comp.empty:
                    if 'JUMLAH_COMPLAIN' in df_prev_comp.columns:
                        prev_grp = df_prev_comp.groupby('TID')['JUMLAH_COMPLAIN'].sum().reset_index()
                    else:
                        prev_grp = df_prev_comp['TID'].value_counts().reset_index()
                        prev_grp.columns = ['TID', 'JUMLAH_COMPLAIN']
                    prev_grp.rename(columns={'JUMLAH_COMPLAIN': prev_mon_short}, inplace=True)
                    
                    # Merge Prev to Curr
                    piv = pd.merge(piv, prev_grp[['TID', prev_mon_short]], on='TID', how='outer').fillna(0)
                    
                    # Fill Metadata for rows that only exist in Prev
                    if 'LOKASI' in df_prev_comp.columns:
                        lookup_loc = df_prev_comp.set_index('TID')['LOKASI'].to_dict()
                        piv['LOKASI'] = piv.apply(lambda r: lookup_loc.get(r['TID'], '') if pd.isna(r['LOKASI']) or r['LOKASI'] == 0 else r['LOKASI'], axis=1)
                    if 'CABANG' in df_prev_comp.columns:
                        lookup_cab = df_prev_comp.set_index('TID')['CABANG'].to_dict()
                        piv['CABANG'] = piv.apply(lambda r: lookup_cab.get(r['TID'], '') if pd.isna(r['CABANG']) or r['CABANG'] == 0 else r['CABANG'], axis=1)
                else:
                    piv[prev_mon_short] = 0

                # Ensure Columns Exist
                for w in ['W1','W2','W3','W4']: 
                    if w not in piv.columns: piv[w] = 0
                
                # C. Calculate Total Current (Σ Jan)
                col_total_curr = f'Σ {curr_mon_short}'
                piv[col_total_curr] = piv[['W1','W2','W3','W4']].sum(axis=1)
                
                # D. Sorting Dynamic based on Week Dropdown
                sort_col = col_total_curr # Default Sort
                if sort_week != 'All Week' and sort_week in piv.columns:
                    sort_col = sort_week
                
                # Sort descending based on selected criterion
                piv = piv.sort_values(sort_col, ascending=False).reset_index(drop=True)
                
                # E. Column Ordering
                cols_show = ['TID', 'LOKASI', 'CABANG', 'TYPE MRI', prev_mon_short, 'W1', 'W2', 'W3', 'W4', col_total_curr]
                cols_final = [c for c in cols_show if c in piv.columns]
                
                # F. Display
                df_disp = clean_zeros(piv[cols_final])
                
                # Convert numbers to int before display
                num_cols = [prev_mon_short, 'W1', 'W2', 'W3', 'W4', col_total_curr]
                for c in num_cols:
                    if c in df_disp.columns:
                        df_disp[c] = pd.to_numeric(df_disp[c]).fillna(0).astype(int).astype(str).replace('0','')
                
                # G. APPLY SPECIAL STYLING (Column Backgrounds)
                final_styler = get_styled_dataframe(df_disp)

                # HEIGHT DISET 200px Biar Scrollable
                event_mri_c = st.dataframe(final_styler, height=200, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
                if len(event_mri_c.selection.rows) > 0:
                    idx = event_mri_c.selection.rows[0]; sel_tid = str(piv.iloc[idx]['TID']); sel_loc = piv.iloc[idx]['LOKASI']
                    time_str = "N/A"
                    tid_problems = df_mri_comp[df_mri_comp['TID'].astype(str) == sel_tid]
                    if not tid_problems.empty and 'TANGGAL' in tid_problems.columns:
                        last_date = tid_problems['TANGGAL'].max()
                        if pd.notnull(last_date): days = (datetime.now() - last_date).days; time_str = "Hari ini" if days == 0 else ("Kemarin" if days == 1 else f"{days} hari lalu")
                    st.info(f"📋 **History TID: {sel_tid}** ({sel_loc}) • **Last Problem:** {time_str}")
                    if not df_slm.empty:
                        slm_det = df_slm[(df_slm['TID'] == sel_tid) & (df_slm['BULAN_EN'] == sel_mon)].copy()
                        if not slm_det.empty:
                            slm_det = slm_det.sort_values('TGL_VISIT', ascending=False).head(2); slm_det['TGL_VISIT'] = slm_det['TGL_VISIT'].dt.strftime('%d-%b-%Y')
                            col_act = next((c for c in slm_det.columns if 'ACTION' in c.upper() or 'KETERANGAN' in c.upper()), None)
                            if col_act: st.dataframe(slm_det[['TGL_VISIT', col_act]], hide_index=True)
                        else: st.caption(f"No Visit Data for {sel_tid}")

        with col_right:
            st.markdown(f'<div class="section-header">Summary Pengisian Data MRI</div>', unsafe_allow_html=True)
            col_visit = next((c for c in df_mri_ops.columns if 'Range' in c or 'Waktu' in c), None)
            if col_visit:
                pagi = df_mri_ops[col_visit].str.contains('Pagi', case=False, na=False).sum(); siang = df_mri_ops[col_visit].str.contains('Siang', case=False, na=False).sum(); malam = df_mri_ops[col_visit].str.contains('Malam', case=False, na=False).sum()
            else: pagi, siang, malam = 0, 0, 0
            visit_data = {"TOTAL ATM": [total_atm_mri], "Pagi": [pagi], "Siang": [siang], "Malam": [malam]}
            st.dataframe(clean_zeros(pd.DataFrame(visit_data)), use_container_width=True, hide_index=True)
            
            # 4. JML DF (Color)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">🔵 JML DF Repeat</div>', unsafe_allow_html=True)
            jml_df_data = { "TOTAL ATM": [total_atm_mri], f"{prev_mon_short}": [len(df_prev_df)], "W1": [len(df_mri_df[df_mri_df['WEEK'] == 'W1'])], "W2": [len(df_mri_df[df_mri_df['WEEK'] == 'W2'])], "W3": [len(df_mri_df[df_mri_df['WEEK'] == 'W3'])], "W4": [len(df_mri_df[df_mri_df['WEEK'] == 'W4'])], f"Σ {curr_mon_short}": [len(df_mri_df)] }
            st.dataframe(get_styled_dataframe(clean_zeros(pd.DataFrame(jml_df_data))), use_container_width=True, hide_index=True)
            
            # 5. TIERING DF (Pakai Logic 'DF' -> COUNT BARIS + ADD TOTAL ROW)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">⚠️ Tiering DF Repeat</div>', unsafe_allow_html=True)
            
            p_t_df = calc_mri_tiers_fixed(df_prev_df, 'DF'); c_t_df = calc_mri_tiers_fixed(df_mri_df, 'DF')
            
            w1_t_d = get_w_risk_mri(df_mri_df, 'W1', 'DF'); w2_t_d = get_w_risk_mri(df_mri_df, 'W2', 'DF')
            w3_t_d = get_w_risk_mri(df_mri_df, 'W3', 'DF'); w4_t_d = get_w_risk_mri(df_mri_df, 'W4', 'DF')
            
            # Create Tier Data
            col_tot = f'Σ {curr_mon_short}'
            tier_data_df = { 'TIERING': ['1 kali', '2-3 kali', '> 3 kali'], f'{prev_mon_short}': [p_t_df[0], p_t_df[1], p_t_df[2]], 'W1': [w1_t_d[0], w1_t_d[1], w1_t_d[2]], 'W2': [w2_t_d[0], w2_t_d[1], w2_t_d[2]], 'W3': [w3_t_d[0], w3_t_d[1], w3_t_d[2]], 'W4': [w4_t_d[0], w4_t_d[1], w4_t_d[2]], col_tot: [c_t_df[0], c_t_df[1], c_t_df[2]] }
            df_tier_df = pd.DataFrame(tier_data_df)

            # Add TOTAL UNIT Row
            total_row_df = {
                'TIERING': 'TOTAL UNIT',
                f'{prev_mon_short}': df_tier_df[f'{prev_mon_short}'].sum(),
                'W1': df_tier_df['W1'].sum(),
                'W2': df_tier_df['W2'].sum(),
                'W3': df_tier_df['W3'].sum(),
                'W4': df_tier_df['W4'].sum(),
                col_tot: df_tier_df[col_tot].sum()
            }
            df_tier_df = pd.concat([df_tier_df, pd.DataFrame([total_row_df])], ignore_index=True)

            st.dataframe(get_styled_dataframe(clean_zeros(df_tier_df)).apply(highlight_total_mri, axis=None), use_container_width=True, hide_index=True)
            
            # 6. TOP TID DF (MODIFIKASI: SCROLLABLE & SORT BY DROPDOWN)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">🔥 Top DF Problem Terminal IDs</div>', unsafe_allow_html=True)
            if not df_mri_df.empty or not df_prev_df.empty:
                # A. Pivot Current (Count/Size)
                if not df_mri_df.empty:
                    piv_df = df_mri_df.pivot_table(index=['TID','LOKASI','CABANG','TYPE MRI'], columns='WEEK', aggfunc='size', fill_value=0).reset_index()
                else:
                    piv_df = pd.DataFrame(columns=['TID','LOKASI','CABANG','TYPE MRI'])

                # B. Prepare Previous (Dec)
                if not df_prev_df.empty:
                    prev_grp_df = df_prev_df['TID'].value_counts().reset_index()
                    prev_grp_df.columns = ['TID', prev_mon_short]
                    
                    # Merge
                    piv_df = pd.merge(piv_df, prev_grp_df[['TID', prev_mon_short]], on='TID', how='outer').fillna(0)
                    
                    # Fill Metadata
                    if 'LOKASI' in df_prev_df.columns:
                        lookup_loc_df = df_prev_df.set_index('TID')['LOKASI'].to_dict()
                        piv_df['LOKASI'] = piv_df.apply(lambda r: lookup_loc_df.get(r['TID'], '') if pd.isna(r['LOKASI']) or r['LOKASI'] == 0 else r['LOKASI'], axis=1)
                    if 'CABANG' in df_prev_df.columns:
                        lookup_cab_df = df_prev_df.set_index('TID')['CABANG'].to_dict()
                        piv_df['CABANG'] = piv_df.apply(lambda r: lookup_cab_df.get(r['TID'], '') if pd.isna(r['CABANG']) or r['CABANG'] == 0 else r['CABANG'], axis=1)
                else:
                    piv_df[prev_mon_short] = 0

                # Ensure Columns
                for w in ['W1','W2','W3','W4']:
                    if w not in piv_df.columns: piv_df[w] = 0

                # C. Calculate Total Current (Σ Jan)
                col_total_curr_df = f'Σ {curr_mon_short}'
                piv_df[col_total_curr_df] = piv_df[['W1','W2','W3','W4']].sum(axis=1)

                # D. Sorting Dynamic
                sort_col_df = col_total_curr_df # Default
                if sort_week != 'All Week' and sort_week in piv_df.columns:
                    sort_col_df = sort_week
                
                # Sort descending
                piv_df = piv_df.sort_values(sort_col_df, ascending=False).reset_index(drop=True)
                
                # E. Column Ordering
                cols_show_df = ['TID', 'LOKASI', 'CABANG', 'TYPE MRI', prev_mon_short, 'W1', 'W2', 'W3', 'W4', col_total_curr_df]
                cols_final_df = [c for c in cols_show_df if c in piv_df.columns]
                
                # F. Display
                df_disp_df = clean_zeros(piv_df[cols_final_df])
                
                # Convert numbers to int
                num_cols_df = [prev_mon_short, 'W1', 'W2', 'W3', 'W4', col_total_curr_df]
                for c in num_cols_df:
                    if c in df_disp_df.columns:
                        df_disp_df[c] = pd.to_numeric(df_disp_df[c]).fillna(0).astype(int).astype(str).replace('0','')

                # G. APPLY SPECIAL STYLING
                final_styler_df = get_styled_dataframe(df_disp_df)

                # HEIGHT DISET 200px Biar Scrollable
                event_mri_d = st.dataframe(final_styler_df, height=200, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
                if len(event_mri_d.selection.rows) > 0:
                    idx = event_mri_d.selection.rows[0]; sel_tid = str(piv_df.iloc[idx]['TID']); sel_loc = piv_df.iloc[idx]['LOKASI']
                    time_str = "N/A"
                    tid_problems = df_mri_df[df_mri_df['TID'].astype(str) == sel_tid]
                    
                    if not tid_problems.empty:
                        col_time = 'WAKTU_INSERT' if 'WAKTU_INSERT' in tid_problems.columns else 'TANGGAL'
                        if col_time in tid_problems.columns:
                            last_time = tid_problems[col_time].max()
                            if pd.notnull(last_time):
                                diff = datetime.now() - last_time; days = diff.days
                                if days > 0: time_str = f"{days} hari lalu"
                                else: hrs = int(diff.seconds // 3600); time_str = f"{hrs} jam lalu" if hrs > 0 else "Baru saja"
                    st.info(f"📋 **History TID: {sel_tid}** ({sel_loc}) • **Last Problem:** {time_str}")
                    if not df_slm.empty:
                        slm_det = df_slm[(df_slm['TID'] == sel_tid) & (df_slm['BULAN_EN'] == sel_mon)].copy()
                        if not slm_det.empty:
                            slm_det = slm_det.sort_values('TGL_VISIT', ascending=False).head(2); slm_det['TGL_VISIT'] = slm_det['TGL_VISIT'].dt.strftime('%d-%b-%Y')
                            col_act = next((c for c in slm_det.columns if 'ACTION' in c.upper() or 'KETERANGAN' in c.upper()), None)
                            if col_act: st.dataframe(slm_det[['TGL_VISIT', col_act]], hide_index=True)
                        else: st.caption(f"No Visit Data for {sel_tid}")
    
    else:
        col_left, col_right = st.columns(2, gap="medium")
        
        with col_left:
            # 1. OVERVIEW SUMMARY
            st.markdown(f'<div class="section-header">📊 {sel_cat} Overview Summary</div>', unsafe_allow_html=True)
            
            # --- FIX: LOGIKA HITUNG SUMMARY STANDARD (COMPLAIN WAJIB SUM) ---
            def get_val_std(dframe):
                if dframe.empty: return 0
                
                # KHUSUS COMPLAIN: SUM KOLOM JUMLAH_COMPLAIN
                if sel_cat == 'Complain':
                    if 'JUMLAH_COMPLAIN' in dframe.columns:
                        # Safety: Konversi ke numeric dulu sebelum sum
                        return int(pd.to_numeric(dframe['JUMLAH_COMPLAIN'], errors='coerce').fillna(0).sum())
                    return 0

                # KATEGORI LAIN: COUNT BARIS
                return len(dframe)

            val_total_atm = 543 
            val_prev = get_val_std(df_prev)
            weeks = ['W1', 'W2', 'W3', 'W4']
            w_vals = {}; curr_total = 0
            for w in weeks:
                val = get_val_std(df_curr[df_curr['WEEK'] == w])
                w_vals[w] = val; curr_total += val
                
            avg_val = curr_total / 4
            prob_val = (curr_total / val_total_atm * 100) if val_total_atm > 0 else 0

            overview_data = { 
                'TOTAL ATM': [str(val_total_atm)], f'{prev_mon_short} (Prev)': [val_prev], 
                'W1': [w_vals['W1']], 'W2': [w_vals['W2']], 'W3': [w_vals['W3']], 'W4': [w_vals['W4']], 
                f'Σ {curr_mon_short}': [curr_total], 'AVG': [f"{avg_val:.1f}"], 'PROB %': [f"{prob_val:.2f}%"] 
            }
            st.dataframe(get_styled_dataframe(clean_zeros(pd.DataFrame(overview_data))), use_container_width=True, hide_index=True)
            
            # 2. RISK TIERS ANALYSIS
            st.markdown(f'<div class="section-header" style="margin-top: 15px;">⚠️ Risk Tiers Analysis</div>', unsafe_allow_html=True)
            def safe_risk_calc(dframe):
                if dframe.empty: return [0, 0, 0]
                # FIX: Pastikan Complain hitung SUM per TID
                if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in dframe.columns: 
                    tid_counts = dframe.groupby('TID')['JUMLAH_COMPLAIN'].sum()
                else: 
                    tid_counts = dframe['TID'].value_counts()
                    
                return [tid_counts[tid_counts == 1].count(), tid_counts[(tid_counts >= 2) & (tid_counts <= 3)].count(), tid_counts[tid_counts > 3].count()]

            p_t = safe_risk_calc(df_prev); c_t = safe_risk_calc(df_curr)
            w1_t, w2_t, w3_t, w4_t = safe_risk_calc(df_curr[df_curr['WEEK'] == 'W1']), safe_risk_calc(df_curr[df_curr['WEEK'] == 'W2']), safe_risk_calc(df_curr[df_curr['WEEK'] == 'W3']), safe_risk_calc(df_curr[df_curr['WEEK'] == 'W4'])
            
            tier_data = { 
                'TIERING': ['1x Kali', '2-3x Kali', '>3x Kali', 'TOTAL UNIT'], 
                f'{prev_mon_short}': [p_t[0], p_t[1], p_t[2], sum(p_t)], 
                'W1': [w1_t[0], w1_t[1], w1_t[2], sum(w1_t)], 'W2': [w2_t[0], w2_t[1], w2_t[2], sum(w2_t)],
                'W3': [w3_t[0], w3_t[1], w3_t[2], sum(w3_t)], 'W4': [w4_t[0], w4_t[1], w4_t[2], sum(w4_t)],
                f'Σ {curr_mon_short}': [c_t[0], c_t[1], c_t[2], sum(c_t)]
            }
            df_tiers = pd.DataFrame(tier_data)
            def highlight_total_row(x):
                df1 = pd.DataFrame('', index=x.index, columns=x.columns)
                try: df1.iloc[3, :] = 'font-weight: 800; background-color: rgba(128, 128, 128, 0.1); border-top: 2px solid #94A3B8;'
                except: pass
                return df1
            
            base_obj = get_styled_dataframe(clean_zeros(df_tiers))
            try: st.dataframe(base_obj.apply(highlight_total_row, axis=None), use_container_width=True, hide_index=True)
            except: st.dataframe(base_obj, use_container_width=True, hide_index=True)

            # 3. FOLLOW UP / TOP LOCATION
            if sel_cat in ['Elastic', 'Complain']:
                st.markdown(f'<div class="section-header" style="margin-top: 15px;">🛠️ Follow-up Status</div>', unsafe_allow_html=True)
                def get_monitor_slice(r_start, r_end, c_start_idx=20, c_end_idx=25):
                    if not df_mon.empty and df_mon.shape[0] >= r_end and df_mon.shape[1] >= c_end_idx:
                        subset = df_mon.iloc[r_start:r_end, c_start_idx:c_end_idx]; headers = subset.iloc[0].astype(str).tolist(); seen = {}; final_cols = []
                        for col in headers:
                            col = col.strip(); cnt = seen.get(col, 0) + 1 if col in seen else 0; seen[col] = cnt; final_cols.append(f"{col}_{cnt}" if col and cnt>0 else col)
                        subset.columns = final_cols; return subset[1:]
                    return pd.DataFrame()
                
                df_fu = get_monitor_slice(2, 7) if sel_cat == 'Elastic' else get_monitor_slice(16, 20)
                if not df_fu.empty: st.dataframe(clean_zeros(df_fu), use_container_width=True, hide_index=True)
                else: st.caption("Data Follow-up belum tersedia.")
            else:
                st.markdown(f'<div class="section-header" style="margin-top: 15px;">📍 Top Impacted Locations</div>', unsafe_allow_html=True)
                if not df_curr.empty and 'LOKASI' in df_curr.columns:
                    loc_counts = df_curr['LOKASI'].value_counts().reset_index(); loc_counts.columns = ['LOKASI', 'FREQ']; top_locs = loc_counts.head(50) 
                    st.dataframe(top_locs, height=200, column_config={ "LOKASI": st.column_config.TextColumn("Lokasi", width="medium"), "FREQ": st.column_config.ProgressColumn("Frekuensi", format="%d", min_value=0, max_value=int(top_locs['FREQ'].max()) if not top_locs.empty else 10, width="small") }, use_container_width=True, hide_index=True)
            
            # --- ANALISA & CATATAN ---
            st.markdown(f'<div class="section-header" style="margin-top: 15px; margin-bottom: 5px !important;">📝 Analisa & Catatan</div>', unsafe_allow_html=True)
            input_height = 90 if sel_cat == 'Elastic' else (100 if sel_cat == 'Complain' else 80)
            current_analysis_text = ""
            if not df_curr.empty:
                if 'ANALISA' in df_curr.columns: current_analysis_text = df_curr['ANALISA'].iloc[0]
                elif 'KETERANGAN' in df_curr.columns: current_analysis_text = df_curr['KETERANGAN'].iloc[0]
            st.markdown("""<style>div[data-testid="stTextArea"] > label {display: none !important;} div[data-testid="stTextArea"] {margin-top: 0px !important;}</style>""", unsafe_allow_html=True)
            st.text_area("Analisa Sheet:", value=str(current_analysis_text), height=input_height, label_visibility="collapsed", placeholder="Ketik analisa di sini...", key=f"analisa_box_{sel_cat}")
        
        with col_right:
            # 1. TOP CRITICAL TIDS (SCROLLABLE ALL DATA)
            st.markdown(f'<div class="section-header">🔥 Critical TIDs (Scroll for More)</div>', unsafe_allow_html=True)
            if 'TID' in df_curr.columns:
                def agg_tid_piv(df_in):
                    if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_in.columns: return df_in.pivot_table(index=['TID', 'LOKASI', 'CABANG'], columns='WEEK', values='JUMLAH_COMPLAIN', aggfunc='sum', fill_value=0).reset_index()
                    return df_in.pivot_table(index=['TID', 'LOKASI', 'CABANG'], columns='WEEK', aggfunc='size', fill_value=0).reset_index()

                pivot_tid = agg_tid_piv(df_curr)
                for w in weeks: 
                    if w not in pivot_tid.columns: pivot_tid[w] = 0
                
                if not df_prev.empty:
                    if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_prev.columns: prev_counts = df_prev.groupby('TID')['JUMLAH_COMPLAIN'].sum().reset_index()
                    else: prev_counts = df_prev['TID'].value_counts().reset_index()
                    prev_counts.columns = ['TID', prev_mon_short] 
                else: prev_counts = pd.DataFrame(columns=['TID', prev_mon_short])
                    
                merged = pd.merge(pivot_tid, prev_counts, on='TID', how='left').fillna(0)
                col_total = f'Σ {curr_mon_short}'; merged[col_total] = merged[weeks].sum(axis=1)
                sort_col = col_total if sort_week == 'All Week' else sort_week
                
                top_all_df = merged.sort_values(sort_col, ascending=False).reset_index(drop=True)
                
                cols_to_convert = [prev_mon_short] + weeks + [col_total]
                for c in cols_to_convert:
                    if c in top_all_df.columns: top_all_df[c] = top_all_df[c].astype(int).astype(str)
                    
                display_cols = ['TID', 'LOKASI', 'CABANG'] + cols_to_convert
                col_config = {
                    "TID": st.column_config.TextColumn("TID", width="small"), 
                    "LOKASI": st.column_config.TextColumn("LOKASI", width="medium"), 
                    "CABANG": st.column_config.TextColumn("CABANG", width="small"), 
                    prev_mon_short: st.column_config.TextColumn(prev_mon_short, width="small"), 
                    "W1": st.column_config.TextColumn("W1", width="small"), 
                    "W2": st.column_config.TextColumn("W2", width="small"),
                    "W3": st.column_config.TextColumn("W3", width="small"), 
                    "W4": st.column_config.TextColumn("W4", width="small"), 
                    col_total: st.column_config.TextColumn(col_total, width="small")
                }
                # USE UNIVERSAL STYLER HERE TOO
                final_styler_tids = get_styled_dataframe(clean_zeros(top_all_df[display_cols]))

                event = st.dataframe(final_styler_tids, height=220, column_config=col_config, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
                if len(event.selection.rows) > 0:
                    selected_idx = event.selection.rows[0]; selected_tid = str(top_all_df.iloc[selected_idx]['TID']); selected_loc = top_all_df.iloc[selected_idx]['LOKASI']
                    time_str = "N/A"; tid_problems = df_curr[df_curr['TID'].astype(str) == selected_tid]
                    if not tid_problems.empty:
                        now = datetime.now()
                        if sel_cat in ['Elastic', 'Complain'] and 'TANGGAL' in tid_problems.columns:
                            last_prob_date = tid_problems['TANGGAL'].max()
                            if pd.notnull(last_prob_date): days = (now - last_prob_date).days; time_str = "Hari ini" if days == 0 else ("Kemarin" if days == 1 else f"{days} hari lalu")
                        elif sel_cat in ['DF Repeat', 'OUT Flm'] and 'WAKTU_INSERT' in tid_problems.columns:
                            last_prob_time = tid_problems['WAKTU_INSERT'].max()
                            if pd.notnull(last_prob_time): diff = now - last_prob_time; days = diff.days; hrs = int(diff.seconds // 3600); time_str = f"{days} hari lalu" if days > 0 else (f"{hrs} jam lalu" if hrs > 0 else "Baru saja")
                    st.info(f"📋 **History TID: {selected_tid}** ({selected_loc}) • **Last Problem:** {time_str}")
                    if not df_slm.empty and 'BULAN_EN' in df_slm.columns:
                        slm_detail = df_slm[(df_slm['TID'] == selected_tid) & (df_slm['BULAN_EN'] == sel_mon)].copy()
                        if not slm_detail.empty:
                            slm_detail = slm_detail.sort_values('TGL_VISIT', ascending=False).head(2); slm_detail['TGL_VISIT'] = slm_detail['TGL_VISIT'].dt.strftime('%d-%b-%Y')
                            col_action = next((c for c in slm_detail.columns if 'ACTION' in c.upper() or 'KETERANGAN' in c.upper()), None)
                            if col_action: st.dataframe(slm_detail[['TGL_VISIT', col_action]], use_container_width=True, hide_index=True)
                            else: st.dataframe(slm_detail, use_container_width=True, hide_index=True)
                        else: st.caption("Belum ada data kunjungan bulan ini.")

            # 2. BRANCH TREND VISUALIZATION
            st.markdown(f'<div class="section-header" style="margin-top: 10px; margin-bottom: 0px !important;">📈 Branch Trend Visualization</div>', unsafe_allow_html=True) 
            if 'CABANG' in df_curr.columns:
                def agg_branch_piv(df_in):
                    if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_in.columns: return df_in.pivot_table(index='CABANG', columns='WEEK', values='JUMLAH_COMPLAIN', aggfunc='sum', fill_value=0).reset_index()
                    return df_in.pivot_table(index='CABANG', columns='WEEK', aggfunc='size', fill_value=0).reset_index()

                p_cab = agg_branch_piv(df_curr)
                for w in weeks: 
                    if w not in p_cab.columns: p_cab[w] = 0
                
                if not df_prev.empty and 'CABANG' in df_prev.columns:
                    if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_prev.columns: branch_prev = df_prev.groupby('CABANG')['JUMLAH_COMPLAIN'].sum().reset_index()
                    else: branch_prev = df_prev['CABANG'].value_counts().reset_index()
                    branch_prev.columns = ['CABANG', prev_mon_short] 
                else: branch_prev = pd.DataFrame(columns=['CABANG', prev_mon_short])
                
                merged_cab = pd.merge(p_cab, branch_prev, on='CABANG', how='left').fillna(0)
                col_total_cab = f'Σ {curr_mon_short}'; merged_cab[col_total_cab] = merged_cab[weeks].sum(axis=1)
                
                top_5_cab_chart = merged_cab.sort_values(col_total_cab, ascending=False).head(5)
                top_all_cab_table = merged_cab.sort_values(col_total_cab, ascending=False)
                
                week_pair = comp_mode.split(' vs ')
                df_melt = top_5_cab_chart[['CABANG'] + week_pair].melt(id_vars='CABANG', var_name='Week', value_name='Total')
                
                # --- CHART STYLING (ALWAYS WHITE / CLEAN) ---
                # User Request: Background putih agar tidak norak
                chart_bg_color = "#FFFFFF"
                c_text = "#1E293B"
                c_grid = "#E2E8F0"
                current_chart_pal = ['#0F172A', '#60A5FA'] # Navy & Light Blue
                
                fig = px.line(df_melt, x='CABANG', y='Total', color='Week', markers=True, text='Total', color_discrete_sequence=current_chart_pal)
                
                # UPDATE LAYOUT: ALWAYS WHITE BACKGROUND
                fig.update_layout(
                    height=180, 
                    margin=dict(l=10, r=0, t=35, b=10), 
                    paper_bgcolor=chart_bg_color, # FORCE WHITE
                    plot_bgcolor=chart_bg_color,  # FORCE WHITE
                    font=dict(family="Inter", size=11, color=c_text), 
                    xaxis=dict(showgrid=True, gridcolor=c_grid, title=None, tickfont=dict(color=c_text)),
                    yaxis=dict(showgrid=True, gridcolor=c_grid, title=None, zeroline=False, showticklabels=False), 
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, title=None, font=dict(color=c_text)), 
                    hovermode="x unified"
                )
                fig.update_traces(
                    mode='lines+markers+text', 
                    line=dict(width=2.5), 
                    marker=dict(size=7, symbol='circle', line=dict(width=1.5, color=c_text)),
                    textposition="top center", 
                    textfont=dict(size=12, color=c_text, family="Inter", weight="bold"), 
                    cliponaxis=False
                )
                
                # INJECT CSS TO MATCH WHITE CONTAINER
                st.markdown(f"""<style>[data-testid="stPlotlyChart"] {{ background-color: {chart_bg_color} !important; border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); width: 100% !important; overflow: hidden !important; margin-top: -10px !important;}} iframe[title="streamlit_plotly_events.plotly_chart"] {{width: 100% !important;}}</style>""", unsafe_allow_html=True)
                st.plotly_chart(fig, use_container_width=True)
                
                final_cols_cab = [prev_mon_short] + weeks + [col_total_cab]
                top_cab_str = top_all_cab_table.copy()
                for c in final_cols_cab: 
                    if c in top_cab_str.columns: top_cab_str[c] = top_cab_str[c].astype(int).astype(str)
                cols_to_show = ['CABANG'] + [c for c in final_cols_cab if c in top_cab_str.columns]
                
                st.dataframe(get_styled_dataframe(clean_zeros(top_cab_str[cols_to_show])), height=200, use_container_width=True, hide_index=True)
