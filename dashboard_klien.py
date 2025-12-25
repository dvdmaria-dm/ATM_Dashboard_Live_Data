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

# PENTING: File 'image_11.png' HARUS ADA di folder yang sama!
try:
    st.set_page_config(
        layout='wide',
        page_title="ATM Performance Dashboard",
        page_icon="image_11.png", 
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
        .stApp {background-color: #F8FAFC;} /* Background Default Terang */
    </style>
""", unsafe_allow_html=True)


# --- 2. KONEKSI DATA GOOGLE SHEETS (SMART CLOUD & LOCAL) ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_MAIN = 'AIMS_Master' 
SHEET_SLM = 'SLM Visit Log'
SHEET_MRI = 'Data_Form' 
SHEET_MONITORING = 'Summary Monitoring Cash'
SHEET_SP = 'Sparepart&kaset' 
JSON_FILE = "credentials.json" 

# Logika Koneksi Pintar: Cek Secrets dulu (Cloud), baru File JSON (Lokal)
gc = None
try:
    if 'gcp_service_account' in st.secrets:
        # Koneksi via Streamlit Cloud Secrets
        creds_dict = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
    elif os.path.exists(JSON_FILE):
        # Koneksi via File Lokal
        gc = gspread.service_account(filename=JSON_FILE)
except Exception as e:
    # Jika gagal koneksi (misal internet mati), biarkan gc None agar masuk mode Offline di load_data
    pass


# --- 3. FUNGSI LOAD DATA (HYBRID: ONLINE + OFFLINE BACKUP) ---
@st.cache_data(ttl=14400, show_spinner=False)
def load_data():
    # File Backup Lokal
    backup_file = 'DATA_MASTER_ATM.xlsx'
    
    # --- FUNGSI FORMATTING REVISI (LEBIH PINTAR & HORMAT DATA) ---
    def clean_and_format(df_in):
        if df_in.empty: return df_in
        # Bersihkan Header
        df_in.columns = df_in.columns.str.strip().str.upper()
        
        # 1. Format Tanggal (Coba baca, tapi jangan panik kalau gagal)
        if 'TANGGAL' in df_in.columns: 
            # Coba konversi tanggal, kalau gagal jadi NaT
            df_in['TANGGAL_OBJ'] = pd.to_datetime(df_in['TANGGAL'], errors='coerce')
            
            # Coba ambil nama bulan dari tanggal yang berhasil dibaca
            df_in['CALC_MONTH'] = df_in['TANGGAL_OBJ'].dt.strftime('%B')
        else:
            df_in['CALC_MONTH'] = None

        # 2. LOGIKA PENENTUAN BULAN (FIXED)
        # Prioritas 1: Ambil dari hasil hitungan tanggal (CALC_MONTH)
        # Prioritas 2: Jika hasil hitungan 'nan', AMBIL DARI KOLOM 'BULAN' EXCEL (Fallback)
        if 'BULAN' in df_in.columns:
            # Pastikan kolom BULAN Excel bersih
            df_in['BULAN'] = df_in['BULAN'].astype(str).str.strip().str.capitalize()
            # Isi BULAN_EN. Jika CALC_MONTH kosong/nan, pakai isi kolom BULAN
            df_in['BULAN_EN'] = df_in['CALC_MONTH'].fillna(df_in['BULAN'])
        else:
            # Kalau tidak ada kolom BULAN di Excel, terpaksa pakai hasil hitungan
            df_in['BULAN_EN'] = df_in['CALC_MONTH']
            
        # Kembalikan kolom TANGGAL asli ke objek datetime (untuk sorting week dll)
        if 'TANGGAL_OBJ' in df_in.columns:
            df_in['TANGGAL'] = df_in['TANGGAL_OBJ']
            df_in.drop(columns=['TANGGAL_OBJ', 'CALC_MONTH'], inplace=True)

        if 'WAKTU INSERT' in df_in.columns: df_in['WAKTU_INSERT'] = pd.to_datetime(df_in['WAKTU INSERT'], errors='coerce')
        
        # Format Angka Complain
        if 'JUMLAH_COMPLAIN' in df_in.columns:
            df_in['JUMLAH_COMPLAIN'] = pd.to_numeric(df_in['JUMLAH_COMPLAIN'].astype(str).str.replace('-', '0'), errors='coerce').fillna(0).astype(int)
        
        # Fallback Week
        if 'WEEK' not in df_in.columns and 'BULAN_WEEK' in df_in.columns: df_in['WEEK'] = df_in['BULAN_WEEK']
        
        return df_in

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

        return df, df_slm, df_mri_ops, df_mon, df_sp_raw

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

                return df, df_slm, df_mri_ops, df_mon, df_sp_raw
            except:
                pass

        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# --- HELPER FUNCTIONS (GLOBAL) ---
def calculate_risk_tiers(df_target):
    if df_target.empty: return 0, 0, 0
    # Jika Complain, hitung berdasarkan SUM. Jika lain, hitung Frequency.
    if 'JUMLAH_COMPLAIN' in df_target.columns and df_target['JUMLAH_COMPLAIN'].sum() > len(df_target):
         counts = df_target.groupby('TID')['JUMLAH_COMPLAIN'].sum()
    else:
         counts = df_target['TID'].value_counts() if 'TID' in df_target.columns else []
         
    return (counts == 1).sum(), ((counts >= 2) & (counts <= 3)).sum(), (counts > 3).sum()

def get_prev_month_full_en(curr_month_en):
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    try: idx = months.index(curr_month_en); return months[idx - 1] if idx > 0 else months[11]
    except: return None

def clean_zeros(df_in):
    return df_in.astype(str).replace(['0', '0.0', '0.00', 'nan', 'None'], '')

# --- EKSEKUSI LOAD DATA ---
df, df_slm, df_mri_ops, df_mon, df_sp_raw = load_data()

# Validasi Data Utama
if df.empty:
    st.warning("⚠️ Data AIMS_Master kosong atau gagal dimuat. Cek koneksi internet atau nama Sheet.")


# =========================================================================
# 4. LOGIKA HALAMAN (INISIALISASI DULU BARU DICEK)
# =========================================================================

# --- LANGKAH 1: INISIALISASI (WAJIB PALING ATAS SEBELUM IF) ---
if 'app_mode' not in st.session_state:
    st.session_state['app_mode'] = 'cover'

# --- LANGKAH 2: CEK KONDISI HALAMAN ---

# --- A. TAMPILAN HALAMAN PEMBUKA (LANDING PAGE) ---
if st.session_state['app_mode'] == 'cover':
    # 1. CSS Styling (Murni CSS, Tanpa f-string)
    st.markdown("""
        <style>
            .cover-container {
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                height: 75vh;
                text-align: center;
            }
            .big-title {
                font-family: 'Helvetica', sans-serif;
                font-weight: 900;
                font-size: 60px;
                background: -webkit-linear-gradient(#0F172A, #2563EB);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin: 0;
            }
            .subtitle {
                font-size: 20px;
                color: #64748B;
                letter-spacing: 2px;
                font-weight: 500;
                margin-bottom: 40px;
            }
            .presenter-box {
                margin-top: 50px;
                padding-top: 20px;
                border-top: 1px solid #E2E8F0;
                width: 450px;
                display: flex;
                flex-direction: column;
                align-items: center;
            }
        </style>
    """, unsafe_allow_html=True)

    # 2. Konten HTML (Gunakan penggabungan string biasa agar aman dari f-string error)
    hari_ini = datetime.now().strftime("%d %B %Y")
    
    html_markup = '<div class="cover-container">'
    html_markup += '<div style="font-size: 80px; margin-bottom: 10px;">🏦</div>'
    html_markup += '<div class="big-title">ATM PERFORMANCE</div>'
    html_markup += '<div class="big-title" style="font-size: 40px; margin-top: -15px;">MONITORING SYSTEM</div>'
    html_markup += '<p class="subtitle">EXECUTIVE SUMMARY & STRATEGIC INSIGHTS</p>'
    html_markup += '<div class="presenter-box">'
    html_markup += '<p style="margin-bottom: 5px; font-weight: bold; font-size: 12px; letter-spacing: 1px; color: #94A3B8;">PRESENTED BY</p>'
    html_markup += '<p style="font-size: 24px; font-weight: 800; color: #1E293B; margin: 0;">DAVID JAMES SIMANJUNTAK</p>'
    html_markup += '<p style="font-size: 14px; color: #64748B; margin-top: 5px;">DATA ANALYST & STRATEGY</p>'
    html_markup += '<p style="font-size: 12px; color: #94A3B8; margin-top: 25px;">' + hari_ini + '</p>'
    html_markup += '</div></div>'
    
    st.markdown(html_markup, unsafe_allow_html=True)

    # 3. Tombol Launch
    col1, col2, col3 = st.columns([5, 2, 5])
    with col2:
        if st.button("🚀 LAUNCH DASHBOARD", use_container_width=True, type="primary"):
            st.session_state['app_mode'] = 'main'
            st.rerun()



# --- B. TAMPILAN DASHBOARD UTAMA (SCRIPT ASLI ABANG ADA DI SINI) ---
elif st.session_state['app_mode'] == 'main':
    
    # =========================================================================
    # 5. HEADER SECTION & THEME ENGINE (FIXED: ANIMASI JALAN & DATA COMPLAIN AMAN)
    # =========================================================================
    
    # --- A. LOGIKA DATA HEADER (GENERATOR PESAN) ---
    try:
        # 1. AMBIL STATE
        h_mon = st.session_state.get('w_mon', df['BULAN_EN'].unique().tolist()[-1] if not df.empty and 'BULAN_EN' in df.columns else '')
        h_week = st.session_state.get('w_week', 'All Week')
        h_cat = st.session_state.get('nav_cat', 'MRI Project') 

        # 2. FILTER DATA & FUNGSI PEMBERSIH
        def safe_text(s):
            if pd.isna(s) or s == "": return "N/A"
            return html.escape(str(s)).replace("'", "").replace('"', "")
        
        # FUNGSI PINTAR: HITUNG COMPLAIN (SUM) vs LAINNYA (COUNT)
        def get_val_safe(dframe, cat_name):
            if dframe.empty: return 0
            try:
                if cat_name == 'Complain' and 'JUMLAH_COMPLAIN' in dframe.columns:
                    # Pastikan fillna(0) agar tidak error kalau ada cell kosong
                    return int(dframe['JUMLAH_COMPLAIN'].fillna(0).sum())
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

        updates = []
        
        if not df_target.empty:
            # Persiapan Data Waktu
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

            # --- GENERATE PESAN (GUNAKAN get_val_safe) ---
            
            # 1. MONTHLY
            val_m_curr = get_val_safe(df_curr_m, h_cat)
            val_m_prev = get_val_safe(df_prev_m, h_cat)
            diff_m = val_m_curr - val_m_prev
            pct_m = (diff_m / val_m_prev * 100) if val_m_prev > 0 else 100.0 if val_m_curr > 0 else 0.0
            icon_m = "🔺" if diff_m > 0 else "🔻"; color_m = "#DC2626" if diff_m > 0 else "#16A34A"
            updates.append(f"<span style='color: #64748B;'>[MONTHLY] Total {cat_label}: <b>{val_m_curr}</b> Tiket (<span style='color: {color_m}; font-weight: 800;'>{icon_m} {diff_m} / {pct_m:.1f}%</span> vs {h_prev_mon})</span>")

            # 2. SUMMARY
            val_s_curr = get_val_safe(df_scope_curr, h_cat)
            val_s_prev = get_val_safe(df_scope_prev, h_cat)
            diff_s = val_s_curr - val_s_prev
            pct_s = (diff_s / val_s_prev * 100) if val_s_prev > 0 else 100.0 if val_s_curr > 0 else 0.0
            diff_str = f"+{diff_s}" if diff_s > 0 else str(diff_s)
            pct_str = f"+{pct_s:.1f}%" if pct_s > 0 else f"{pct_s:.1f}%"
            color_s = "#DC2626" if diff_s > 0 else "#16A34A"
            updates.append(f"<span style='color: #64748B;'>[{scope_label}] Kategori {cat_label}: <b>{val_s_curr}</b> Tiket. Selisih: <span style='color: {color_s}; font-weight: 800;'>{diff_str} ({pct_str})</span> vs periode lalu.</span>")

            # 3. RECURRING (Tetap hitung Unit/TID Unik)
            if is_weekly_mode and not df_scope_curr.empty and not df_scope_prev.empty and 'TID' in df_scope_curr.columns:
                tids_now = set(df_scope_curr['TID']); tids_bef = set(df_scope_prev['TID'])
                rec_tids = tids_now.intersection(tids_bef)
                cnt_rec = len(rec_tids)
                if cnt_rec > 0:
                    top_rec_list = [safe_text(x) for x in list(rec_tids)[:3]]
                    top_rec_str = ", ".join(top_rec_list)
                    updates.append(f"<span style='color: #64748B;'>[RECURRING] Waspada! Ada <span style='color: #F59E0B; font-weight: 800;'>{cnt_rec} Unit</span> Masalah Berulang dari {prev_w_str} ke {h_week}. (Contoh: {top_rec_str}...)</span>")

            # 4 & 5. BRANCH TREND (Aggregasi Sum Complain)
            if 'CABANG' in df_scope_curr.columns:
                # Fungsi Aggregator
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

            # 6 & 7. TID TREND
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

        # UPDATE GLOBAL
        updates.append(f"<span style='color: #64748B;'>🌍 GLOBAL ASSETS: <span style='color: #1E293B; font-weight: 800;'>{total_armada}</span> Units Active</span>")

        # --- LOGIKA MATEMATIKA CSS (FIXED) ---
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


    # --- B. RENDER LAYOUT HEADER (CSS) ---
    head_c1, head_c2, head_c3 = st.columns([2.5, 7.0, 2.5])

    with head_c1:
        st.markdown("""
        <div style="line-height: 1.1;">
            <div class="main-title" style="font-size: 30px !important;">MONITORING PERFORMANCE</div>
            <div class="sub-title" style="font-size: 14px !important; margin-top: -2px; color: #64748B;">PT KELOLA JASA ARTA</div>
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
        # 1. Tampilkan Tanggal Statis
        curr_date = datetime.now().strftime("%d %B %Y")
        
        # 2. Wadah Jam (Span dengan ID khusus) & FIX Mentok Kanan
        st.markdown(f"""
        <div style="display: flex; flex-direction: column; align-items: flex-end; width: 100%; margin-right: -10px;">
            <div class="date-pill" style="font-size: 10px !important; padding: 2px 8px; margin-bottom: 2px;">📅 {curr_date}</div>
            <div style="font-size: 10px; font-weight: 700; color: #16A34A;">
                ONLINE <span id="clock_ticks">--:--:--</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 3. Script JavaScript untuk Detik Berjalan (Live Ticker)
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
                    // Cari elemen dengan ID 'clock_ticks' di parent frame
                    const target = window.parent.document.getElementById('clock_ticks');
                    if (target) {
                        target.innerText = timeString;
                    }
                }
                // Jalankan setiap 1000ms (1 detik)
                setInterval(updateClock, 1000);
                updateClock(); // Jalankan langsung biar gak nunggu 1 detik
            </script>
            """,
            height=0, # Sembunyikan iframe script ini
            width=0
        )


    # C. LOGIKA WARNA & THEME ENGINE
    # Kunci agar variabel use_exec_mode selalu ada (Anti-Error)
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

    # --- FIX MERAPATKAN NAVIGASI ---
    st.markdown("""
    <style>
    div[role="radiogroup"] {
        margin-top: -70px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # --- FIX dropdown + toggle ikut naik ---
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

    # --- 6. NAVIGASI & FILTER ---
    nav_col, filter_col = st.columns([2.5, 1.8], gap="medium")

    with nav_col:
        st.markdown("""<style>div[role="radiogroup"] { justify-content: flex-start !important; flex-wrap: nowrap !important; width: 100% !important; } div[role="radiogroup"] label { white-space: nowrap !important; }</style>""", unsafe_allow_html=True)
        menu_items = ['MRI Project', 'Elastic', 'Complain', 'DF Repeat', 'OUT Flm', 'SparePart & Kaset']
        # Tambahkan key="nav_cat" agar Header di atas bisa membacanya
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
            # f1=Periode, f2=Week, f3=Trend, f4=Warna, f5=Mode
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
                # TOMBOL SWITCH MODE ADA DI SINI SEKARANG
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

    # --- LOGIKA DATA PROCESSING (FIXED: COMPLAIN DIHITUNG SUM) ---
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
    if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in df_curr.columns:
        total_ticket = int(df_curr['JUMLAH_COMPLAIN'].sum())
    else:
        total_ticket = len(df_curr)

    avg_ticket = total_ticket / 4 if sel_cat != 'SparePart & Kaset' else 0

    # --- MICRO METRICS SECTION (FIXED: COMPLAIN DIHITUNG SUM) ---
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

        # --- HITUNG METRIK (Revisi: Jika Complain, Sum Kolom J) ---
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

        # Style
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

    # --- 7. MAIN CONTENT RENDERING ---

    # INI KODE PENDORONG TABEL AGAR TIDAK MENIMPA JUDUL
    st.markdown("""
        <style>
            /* Mencari semua elemen tabel dan mendorongnya turun */
            [data-testid="stDataFrame"], .stDataFrame {
                margin-top: 6px !important;
            }
            
            /* Jika kau menggunakan judul manual (st.markdown) di atas tabel, 
               kode ini juga akan memberikan jarak pada judul tersebut */
            .section-header {
                margin-top: 5px !important;
            }
        </style>
    """, unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom: 5px;'></div>", unsafe_allow_html=True) 

    # 🔥 LOGIKA PEWARNAAN W2 vs W1 (Chain Comparison) 🔥
    def apply_corporate_style(df_in):
        if not use_color: return df_in 

        def style_logic(row):
            color_bad = 'color: #B91C1C; font-weight: 700;' 
            color_good = 'color: #15803D; font-weight: 700;' 
            styles = [''] * len(row)
            
            def get_val(val):
                try: return float(val) if val != "" else 0
                except: return 0
            
            col_names = row.index.tolist()
            # Logika: Bandingkan Current vs Previous dalam urutan W1->W2->W3->W4
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
            return df_in.style.apply(style_logic, axis=1)
        except:
            return df_in
    # =========================================================================
    # 1. LAYOUT KHUSUS: Sparepart&kaset (NAMA SUDAH DIPERBAIKI)
    # =========================================================================
    # Perhatikan: string di bawah ini sekarang persis 'Sparepart&kaset'
    if sel_cat == 'Sparepart&kaset':
        st.markdown("""<style>[data-testid="stDataFrame"] th { font-size: 10px !important; padding: 4px 6px !important; white-space: normal !important; vertical-align: top !important; line-height: 1.2 !important; height: auto !important; background-color: #F8FAFC !important; }[data-testid="stDataFrame"] td { font-size: 10px !important; padding: 3px 6px !important; white-space: nowrap !important; }</style>""", unsafe_allow_html=True)
        
        # Helper Data Custom (X1:AI10)
        def get_custom_data(r_header_idx, r_data_stop_idx, c_start_idx, c_end_idx):
            try:
                if not df_sp_raw.empty:
                    # Cek ketersediaan kolom
                    if df_sp_raw.shape[1] < c_start_idx:
                        return pd.DataFrame(columns=[f"Error: Sheet 'Sparepart&kaset' terbaca, tapi kolom kurang (Cuma {df_sp_raw.shape[1]})"])

                    # 1. AMBIL HEADER
                    header_row = df_sp_raw.iloc[r_header_idx, c_start_idx:c_end_idx]
                    raw_headers = header_row.astype(str).str.strip().tolist()
                    
                    # 2. BERSIHKAN HEADER
                    final_headers = []
                    seen_counts = {}
                    for col in raw_headers:
                        if col.lower() in ['nan', 'none', '']: col = "Info"
                        if col in seen_counts:
                            seen_counts[col] += 1
                            final_headers.append(f"{col}_{seen_counts[col]}")
                        else:
                            seen_counts[col] = 0
                            final_headers.append(col)
                    
                    # 3. AMBIL DATA
                    data_values = df_sp_raw.iloc[r_header_idx+1 : r_data_stop_idx, c_start_idx:c_end_idx].values
                    
                    # 4. BUAT TABEL
                    new_df = pd.DataFrame(data_values, columns=final_headers)
                    return clean_zeros(new_df)
            except Exception as e:
                return pd.DataFrame(columns=[f"Error Script: {str(e)}"])
            return pd.DataFrame()

        tab1, tab2, tab3 = st.tabs(["🛠️ Stock Sparepart", "📼 Stock Kaset", "⚠️ Monitoring & PM"])
        
        with tab1: 
            st.markdown(f'<div class="section-header">🛠️ Ketersediaan SparePart</div>', unsafe_allow_html=True)
            # Sparepart asumsi masih di A1:V10 (Aman)
            st.dataframe(get_custom_data(0, 10, 0, 22), use_container_width=True, hide_index=True)
            
        with tab2: 
            st.markdown(f'<div class="section-header">📼 Ketersediaan Kaset (X1:AI10)</div>', unsafe_allow_html=True)
            
            # --- UJI COBA KOORDINAT X1:AI10 ---
            # Row 0-10, Kolom 23 (X) - 35 (AI)
            
            df_kaset = get_custom_data(0, 10, 23, 35) 
            
            if not df_kaset.empty and "Error" not in df_kaset.columns[0]:
                for col in df_kaset.columns:
                    if "CABANG" in col.upper(): continue
                    try:
                        clean_val = df_kaset[col].astype(str).str.replace('%', '').str.strip()
                        s_numeric = pd.to_numeric(clean_val, errors='coerce')
                        
                        df_kaset[col] = s_numeric.apply(
                            lambda x: f"{x:.0%}" if (pd.notnull(x) and x <= 1.5) else (f"{x:.0f}" if pd.notnull(x) else "")
                        )
                    except: pass
                st.dataframe(df_kaset, use_container_width=True, hide_index=True)
            else:
                st.warning("⚠️ Data Kosong / Salah Sheet")
                st.write("Coba cek: Apakah nama sheet di `load_data` bagian atas script sudah diganti jadi `Sparepart&kaset`?")

        with tab3:
            c1, c2 = st.columns(2)
            with c1: st.markdown(f'<div class="section-header">⚠️ Rekap Kaset Rusak</div>', unsafe_allow_html=True); st.dataframe(get_custom_data(24, 29, 0, 6), use_container_width=True, hide_index=True)
            with c2: st.markdown(f'<div class="section-header">🧹 PM Kaset</div>', unsafe_allow_html=True); st.dataframe(get_custom_data(31, 38, 0, 7), use_container_width=True, hide_index=True)
  
   
    # =========================================================================
    # 2. LAYOUT KHUSUS: MRI PROJECT (V61.46: FIX VARIABLE NAME TYPO)
    # =========================================================================
    elif sel_cat == 'MRI Project':
        col_left, col_right = st.columns(2, gap="medium")
        df_mri_comp = df_curr[df_curr['KATEGORI'] == 'Complain'].copy()
        df_mri_df   = df_curr[df_curr['KATEGORI'] == 'DF Repeat'].copy()
        df_prev_comp = df_prev[df_prev['KATEGORI'] == 'Complain'].copy() if not df_prev.empty else pd.DataFrame()
        df_prev_df   = df_prev[df_prev['KATEGORI'] == 'DF Repeat'].copy() if not df_prev.empty else pd.DataFrame()
        total_atm_mri = 34 # REVISI: JUMLAH ATM JADI 34 SESUAI REQUEST ABANG

        with col_left:
            st.markdown(f'<div class="section-header">🔴 Summary Problem TID MRI</div>', unsafe_allow_html=True)
            sum_data = {"TOTAL ATM": [total_atm_mri], "Complain": [len(df_mri_comp)], "DF": [len(df_mri_df)]}
            st.dataframe(clean_zeros(pd.DataFrame(sum_data)), use_container_width=True, hide_index=True)
            
            # 1. JML COMPLAIN (Color)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">📊 JML Complain</div>', unsafe_allow_html=True)
            jml_data = { "TOTAL ATM": [total_atm_mri], f"{prev_mon_short}": [len(df_prev_comp)], "W1": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W1'])], "W2": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W2'])], "W3": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W3'])], "W4": [len(df_mri_comp[df_mri_comp['WEEK'] == 'W4'])], f"Σ {curr_mon_short}": [len(df_mri_comp)] }
            st.dataframe(apply_corporate_style(clean_zeros(pd.DataFrame(jml_data))), use_container_width=True, hide_index=True)
            
            # 2. TIERING COMPLAIN (Color)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">⚠️ Tiering Complain</div>', unsafe_allow_html=True)
            p_t = calculate_risk_tiers(df_prev_comp); c_t = calculate_risk_tiers(df_mri_comp)
            def get_w_risk_mri(df_target, w): return calculate_risk_tiers(df_target[df_target['WEEK'] == w])
            w1_t = get_w_risk_mri(df_mri_comp, 'W1'); w2_t = get_w_risk_mri(df_mri_comp, 'W2'); w3_t = get_w_risk_mri(df_mri_comp, 'W3'); w4_t = get_w_risk_mri(df_mri_comp, 'W4')
            tier_data_mri = { 'TIERING': ['1 kali', '2-3 kali', '> 3 kali'], f'{prev_mon_short}': [p_t[0], p_t[1], p_t[2]], 'W1': [w1_t[0], w1_t[1], w1_t[2]], 'W2': [w2_t[0], w2_t[1], w2_t[2]], 'W3': [w3_t[0], w3_t[1], w3_t[2]], 'W4': [w4_t[0], w4_t[1], w4_t[2]], f'Σ {curr_mon_short}': [c_t[0], c_t[1], c_t[2]] }
            st.dataframe(apply_corporate_style(clean_zeros(pd.DataFrame(tier_data_mri))), use_container_width=True, hide_index=True)
            
            # 3. TOP TID COMPLAIN (Color + Fix Logic)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">🔥 Top Complain Problem Terminal IDs</div>', unsafe_allow_html=True)
            if not df_mri_comp.empty:
                piv = df_mri_comp.pivot_table(index=['TID','LOKASI','CABANG','TYPE MRI'], columns='WEEK', aggfunc='size', fill_value=0).reset_index()
                for w in ['W1','W2','W3','W4']: 
                    if w not in piv.columns: piv[w] = 0
                piv['Total'] = piv[['W1','W2','W3','W4']].sum(axis=1)
                sort_col_mri = sort_week if sort_week != 'All Week' else 'Total'
                piv = piv.sort_values(sort_col_mri, ascending=False).head(8).reset_index(drop=True)
                cols_show = ['TID', 'LOKASI', 'CABANG', 'TYPE MRI', 'W1', 'W2', 'W3', 'W4']
                cols_final = [c for c in cols_show if c in piv.columns]
                
                df_disp = clean_zeros(piv[cols_final])
                event_mri_c = st.dataframe(apply_corporate_style(df_disp), use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
                if len(event_mri_c.selection.rows) > 0:
                    idx = event_mri_c.selection.rows[0]; sel_tid = str(piv.iloc[idx]['TID']); sel_loc = piv.iloc[idx]['LOKASI']
                    time_str = "N/A"
                    # FIXED VARIABLE NAME
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
            st.dataframe(apply_corporate_style(clean_zeros(pd.DataFrame(jml_df_data))), use_container_width=True, hide_index=True)
            
            # 5. TIERING DF (Color)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">⚠️ Tiering DF Repeat</div>', unsafe_allow_html=True)
            p_t_df = calculate_risk_tiers(df_prev_df); c_t_df = calculate_risk_tiers(df_mri_df)
            w1_t_d = get_w_risk_mri(df_mri_df, 'W1'); w2_t_d = get_w_risk_mri(df_mri_df, 'W2'); w3_t_d = get_w_risk_mri(df_mri_df, 'W3'); w4_t_d = get_w_risk_mri(df_mri_df, 'W4')
            tier_data_df = { 'TIERING': ['1 kali', '2-3 kali', '> 3 kali'], f'{prev_mon_short}': [p_t_df[0], p_t_df[1], p_t_df[2]], 'W1': [w1_t_d[0], w1_t_d[1], w1_t_d[2]], 'W2': [w2_t_d[0], w2_t_d[1], w2_t_d[2]], 'W3': [w3_t_d[0], w3_t_d[1], w3_t_d[2]], 'W4': [w4_t_d[0], w4_t_d[1], w4_t_d[2]], f'Σ {curr_mon_short}': [c_t_df[0], c_t_df[1], c_t_df[2]] }
            st.dataframe(apply_corporate_style(clean_zeros(pd.DataFrame(tier_data_df))), use_container_width=True, hide_index=True)
            
            # 6. TOP TID DF (Color + Fix Logic)
            st.markdown(f'<div class="section-header" style="margin-top:15px;">🔥 Top DF Problem Terminal IDs</div>', unsafe_allow_html=True)
            if not df_mri_df.empty:
                piv_df = df_mri_df.pivot_table(index=['TID','LOKASI','CABANG','TYPE MRI'], columns='WEEK', aggfunc='size', fill_value=0).reset_index()
                for w in ['W1','W2','W3','W4']: 
                    if w not in piv_df.columns: piv_df[w] = 0
                piv_df['Total'] = piv_df[['W1','W2','W3','W4']].sum(axis=1)
                sort_col_mri = sort_week if sort_week != 'All Week' else 'Total'
                piv_df = piv_df.sort_values(sort_col_mri, ascending=False).head(4).reset_index(drop=True)
                cols_show_df = ['TID', 'LOKASI', 'CABANG', 'TYPE MRI', 'W1', 'W2', 'W3', 'W4']
                cols_final_df = [c for c in cols_show_df if c in piv_df.columns]
                
                df_disp_df = clean_zeros(piv_df[cols_final_df])
                event_mri_d = st.dataframe(apply_corporate_style(df_disp_df), use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
                if len(event_mri_d.selection.rows) > 0:
                    idx = event_mri_d.selection.rows[0]; sel_tid = str(piv_df.iloc[idx]['TID']); sel_loc = piv_df.iloc[idx]['LOKASI']
                    time_str = "N/A"
                    # FIXED VARIABLE NAME
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
    # =========================================================================
    # 3. LAYOUT STANDARD (ELASTIC, COMPLAIN, OUT FLM, DF REPEAT) - FINAL FIX
    # =========================================================================
    else:
        col_left, col_right = st.columns(2, gap="medium")
        
        # --- KOLOM KIRI ---
        with col_left:
            # 1. OVERVIEW SUMMARY
            st.markdown(f'<div class="section-header">📊 {sel_cat} Overview Summary</div>', unsafe_allow_html=True)
            
            def get_val_std(dframe):
                if dframe.empty: return 0
                if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in dframe.columns:
                    return int(dframe['JUMLAH_COMPLAIN'].fillna(0).sum())
                return len(dframe)

            val_total_atm = 611 
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
            st.dataframe(apply_corporate_style(clean_zeros(pd.DataFrame(overview_data))), use_container_width=True, hide_index=True)
            
            # 2. RISK TIERS ANALYSIS
            st.markdown(f'<div class="section-header" style="margin-top: 15px;">⚠️ Risk Tiers Analysis</div>', unsafe_allow_html=True)
            def safe_risk_calc(dframe):
                if dframe.empty: return [0, 0, 0]
                if sel_cat == 'Complain' and 'JUMLAH_COMPLAIN' in dframe.columns: tid_counts = dframe.groupby('TID')['JUMLAH_COMPLAIN'].sum()
                else: tid_counts = dframe['TID'].value_counts()
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
            
            base_obj = apply_corporate_style(clean_zeros(df_tiers))
            try: st.dataframe(base_obj.style.apply(highlight_total_row, axis=None), use_container_width=True, hide_index=True)
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
                    loc_counts = df_curr['LOKASI'].value_counts().reset_index(); loc_counts.columns = ['LOKASI', 'FREQ']; top_locs = loc_counts.head(50) # Scrollable
                    st.dataframe(top_locs, height=200, column_config={ "LOKASI": st.column_config.TextColumn("Lokasi", width="medium"), "FREQ": st.column_config.ProgressColumn("Frekuensi", format="%d", min_value=0, max_value=int(top_locs['FREQ'].max()) if not top_locs.empty else 10, width="small") }, use_container_width=True, hide_index=True)
                else: st.info("Data Lokasi tidak cukup.")
            
            # --- ANALISA & CATATAN ---
            st.markdown(f'<div class="section-header" style="margin-top: 15px; margin-bottom: 5px !important;">📝 Analisa & Catatan</div>', unsafe_allow_html=True)
            input_height = 150 if sel_cat == 'Elastic' else (100 if sel_cat == 'Complain' else 80)
            current_analysis_text = ""
            if not df_curr.empty:
                if 'ANALISA' in df_curr.columns: current_analysis_text = df_curr['ANALISA'].iloc[0]
                elif 'KETERANGAN' in df_curr.columns: current_analysis_text = df_curr['KETERANGAN'].iloc[0]
            st.markdown("""<style>div[data-testid="stTextArea"] > label {display: none !important;} div[data-testid="stTextArea"] {margin-top: 0px !important;}</style>""", unsafe_allow_html=True)
            st.text_area("Analisa Sheet:", value=str(current_analysis_text), height=input_height, label_visibility="collapsed", placeholder="Ketik analisa di sini...", key=f"analisa_box_{sel_cat}")
        
        # --- KOLOM KANAN ---
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
                
                # AMBIL SEMUA DATA (HAPUS .head(5))
                top_all_df = merged.sort_values(sort_col, ascending=False).reset_index(drop=True)
                
                cols_to_convert = [prev_mon_short] + weeks + [col_total]
                for c in cols_to_convert:
                    if c in top_all_df.columns: top_all_df[c] = top_all_df[c].astype(int).astype(str)
                    
                display_cols = ['TID', 'LOKASI', 'CABANG'] + cols_to_convert
                col_config = {
                    "TID": st.column_config.TextColumn("TID", width="small"), 
                    "LOKASI": st.column_config.TextColumn("LOKASI", width="medium"), # Fixed Scroll
                    "CABANG": st.column_config.TextColumn("CABANG", width="small"), # Fixed Scroll
                    prev_mon_short: st.column_config.TextColumn(prev_mon_short, width="small"), 
                    "W1": st.column_config.TextColumn("W1", width="small"), 
                    "W2": st.column_config.TextColumn("W2", width="small"),
                    "W3": st.column_config.TextColumn("W3", width="small"), 
                    "W4": st.column_config.TextColumn("W4", width="small"), 
                    col_total: st.column_config.TextColumn(col_total, width="small")
                }
                # HEIGHT DISET 220px AGAR SCROLLABLE
                event = st.dataframe(apply_corporate_style(clean_zeros(top_all_df[display_cols])), height=220, column_config=col_config, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
                
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

            # 2. BRANCH TREND VISUALIZATION (CHART: TOP 5, TABLE: ALL SCROLLABLE)
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
                
                # DATA UNTUK GRAFIK (TETAP TOP 5 AGAR TIDAK KUSUT)
                top_5_cab_chart = merged_cab.sort_values(col_total_cab, ascending=False).head(5)
                
                # DATA UNTUK TABEL (ALL DATA - SCROLLABLE)
                top_all_cab_table = merged_cab.sort_values(col_total_cab, ascending=False)
                
                week_pair = comp_mode.split(' vs ')
                df_melt = top_5_cab_chart[['CABANG'] + week_pair].melt(id_vars='CABANG', var_name='Week', value_name='Total')
                
                if use_exec_mode: chart_bg_color = "#1E293B"; c_text = "#F8FAFC"; c_grid = "rgba(255, 255, 255, 0.1)"; current_chart_pal = ['#F59E0B', '#38BDF8'] 
                else: chart_bg_color = "#FFFFFF"; c_text = "#1E293B"; c_grid = "#E2E8F0"; current_chart_pal = ['#0F172A', '#60A5FA'] 
                
                fig = px.line(df_melt, x='CABANG', y='Total', color='Week', markers=True, text='Total', color_discrete_sequence=current_chart_pal)
                fig.update_layout(height=180, margin=dict(l=10, r=0, t=35, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="Inter", size=11, color=c_text), xaxis=dict(showgrid=True, gridcolor=c_grid, title=None, tickfont=dict(color=c_text)),
                    yaxis=dict(showgrid=True, gridcolor=c_grid, title=None, zeroline=False, showticklabels=False), 
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, title=None, font=dict(color=c_text)), hovermode="x unified")
                fig.update_traces(mode='lines+markers+text', line=dict(width=2.5), marker=dict(size=7, symbol='circle', line=dict(width=1.5, color=c_text)),
                    textposition="top center", textfont=dict(size=12, color=c_text, family="Inter", weight="bold"), cliponaxis=False)
                
                st.markdown(f"""<style>[data-testid="stPlotlyChart"] {{ background-color: {chart_bg_color} !important; border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); width: 100% !important; overflow: hidden !important; margin-top: -10px !important;}} iframe[title="streamlit_plotly_events.plotly_chart"] {{width: 100% !important;}}</style>""", unsafe_allow_html=True)
                st.plotly_chart(fig, use_container_width=True)
                
                final_cols_cab = [prev_mon_short] + weeks + [col_total_cab]
                top_cab_str = top_all_cab_table.copy()
                for c in final_cols_cab: 
                    if c in top_cab_str.columns: top_cab_str[c] = top_cab_str[c].astype(int).astype(str)
                cols_to_show = ['CABANG'] + [c for c in final_cols_cab if c in top_cab_str.columns]
                
                # TABEL SCROLLABLE (HEIGHT 200px)

                st.dataframe(apply_corporate_style(clean_zeros(top_cab_str[cols_to_show])), height=200, use_container_width=True, hide_index=True)





















