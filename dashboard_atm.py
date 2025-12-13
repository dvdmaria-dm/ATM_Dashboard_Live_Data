# ATM Executive Dashboard - V24 FINAL (LIVE GSheets Connected)
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import gspread # Library yang kita pakai
# from gspread_dataframe import get_dataframe 
import json
from io import StringIO 

# --- 0. KONFIGURASI URL DATA MUTLAK (INJEKSI) ---
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit?gid=98670277#gid=98670277"

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Dashboard Executive ATM", layout="wide")

# --- 2. CSS STYLING ---
st.markdown("""
<style>
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    h1 { font-size: 24px !important; margin-bottom: 5px; padding-bottom: 0px; }
    .stDataFrame { font-size: 13px; }
    div[data-testid="stDataFrame"] div[class*="stDataFrame"] { font-size: 12px !important; }
    .streamlit-expanderHeader { font-weight: bold; color: #E74C3C; padding: 8px 0px 8px 0px; }
    .st-emotion-cache-1ftnylj { background-color: #f7f7f7; border-radius: 5px; margin-bottom: 5px; }
</style>
""", unsafe_allow_html=True)

# --- 3. DATA POPULASI MANUAL (Hardcode Tetap) ---
POPULASI_MANUAL = {
    "BANDUNG": 58, "BEKASI": 78, "DENPASAR": 86, "JAKARTA": 147, "JOGJA": 20,
    "KUPANG": 15, "MAKASSAR": 89, "MEDAN": 47, "PARE-PARE": 71
}
TOTAL_ASSET_GLOBAL = sum(POPULASI_MANUAL.values())

# --- 4. FUNGSI LOAD DATA BARU (GOOGLE SHEETS - FIX KONVERSI) ---
@st.cache_data(ttl=300) 
def load_data_gsheets(gsheet_url):
    
    # 1. KONEKSI KE GOOGLE DRIVE MENGGUNAKAN SERVICE ACCOUNT JSON
    try:
      import json
import streamlit as st
from gspread_pandas import spread

# Ambil data Secrets dari Streamlit Cloud, bukan dari file lokal.
# Kunci gspread_service_account adalah nama header TOML yang kita masukkan.
secrets = st.secrets["gspread_service_account"]
secrets_dict = dict(secrets)

# Gunakan data secrets yang sudah di-load sebagai dict
gc = gspread.service_account_from_dict(secrets_dict
    except Exception as e:
        st.error(f"⚠️ GAGAL KONEKSI GOOGLE API. Pastikan file 'credentials.json' ada di folder DASHBOARD_ATM. Error: {e}")
        return pd.DataFrame(), pd.DataFrame() 
        
    # 2. BUKA GOOGLE SHEET BERDASARKAN URL
    try:
        spreadsheet = gc.open_by_url(gsheet_url)
    except Exception as e:
        st.error(f"⚠️ GAGAL MEMBUKA GOOGLE SHEET. Pastikan URL sudah benar dan SERVICE ACCOUNT sudah diberi akses 'Viewer'. Error: {e}")
        return pd.DataFrame(), pd.DataFrame()
        
    # 3. AMBIL DATA DARI SHEET 'AIMS_Master'
    try:
        ws_aims = spreadsheet.worksheet("AIMS_Master")
        aims_data = ws_aims.get_all_values()
        
        # Tentukan Header Row (Cari baris yang mengandung 'TANGGAL')
        header_row_idx = None
        for idx, row in enumerate(aims_data):
            if 'TANGGAL' in [str(x).strip().upper() for x in row]:
                header_row_idx = idx
                break
                
        if header_row_idx is None:
            st.error("Header 'TANGGAL' tidak ditemukan di AIMS_Master.")
            return pd.DataFrame(), pd.DataFrame()
            
        df = pd.DataFrame(aims_data[header_row_idx+1:], columns=aims_data[header_row_idx])
        
    except Exception as e:
        st.error(f"Gagal memproses sheet AIMS_Master. Pastikan nama sheet benar. Error: {e}")
        return pd.DataFrame(), pd.DataFrame()
        
    # 4. AMBIL DATA DARI SHEET 'SLM Visit Log'
    try:
        ws_slm = spreadsheet.worksheet("SLM Visit Log")
        slm_data = ws_slm.get_all_values()
        
        # Ambil header baris ke-2 (index 1) dari data Google Sheet
        header_slm = slm_data[1] 
        # Ambil data mulai baris ke-3 (index 2)
        df_slm_raw = pd.DataFrame(slm_data[2:], columns=header_slm)

        # MAPPING KOLOM FIX (Kolom A, F, G)
        COL_TID = header_slm[0]
        COL_TGL = header_slm[5]
        COL_ACTION = header_slm[6]
        
        df_slm = df_slm_raw[[COL_TID, COL_TGL, COL_ACTION]].copy()
        df_slm.columns = ['TID', 'TGL Visit SLM', 'Action']
        
        df_slm['TID'] = df_slm['TID'].astype(str).str.strip()
        df_slm['TGL Visit SLM'] = pd.to_datetime(df_slm['TGL Visit SLM'], errors='coerce')
        df_slm.dropna(subset=['TID', 'TGL Visit SLM'], inplace=True)
        
    except Exception as e:
        st.warning(f"Gagal memproses sheet SLM Visit Log. Fitur detail SLM dimatikan. Error: {e}")
        df_slm = pd.DataFrame(columns=['TID', 'TGL Visit SLM', 'Action'])

    # 5. PEMBUATAN KOLOM DAN NORMALISASI (SAMA)
    df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], errors='coerce')
    
    # Konversi kolom-kolom penting menjadi tipe numerik yang benar
    for col in ['JUMLAH_COMPLAIN']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Hapus baris dengan Tanggal atau TID kosong setelah konversi
    df.dropna(subset=['TANGGAL', 'TID'], inplace=True)
    
    cols = [c.upper() for c in df.columns]
    if 'MASALAH' in cols: df.rename(columns={'MASALAH': 'KATEGORI'}, inplace=True)
    if 'JUMLAH_COMPLAIN' not in df.columns or df['JUMLAH_COMPLAIN'].sum() == 0: 
        df['JUMLAH_COMPLAIN'] = 1 
    
    df['MONTH_NUM'] = df['TANGGAL'].dt.month
    df['BULAN'] = df['TANGGAL'].dt.strftime('%B')
    df['WEEK'] = 'W' + df['TANGGAL'].dt.day.apply(lambda x: str((x-1)//7 + 1))
    df['CABANG'] = df['CABANG'].astype(str).str.upper().str.strip()
    
    return df, df_slm

# --- LOAD DATA UTAMA ---
df, df_slm = load_data_gsheets(GOOGLE_SHEET_URL)

if df.empty:
    st.error("Dashboard tidak dapat menampilkan data. Cek koneksi API dan URL Google Sheet.")
    st.stop()
    
# --- FUNGSI UTAMA UNTUK MENGHITUNG TIKET/COMPLAIN (SAMA) ---
def get_count_or_sum(df_source, category):
    if df_source.empty: return 0
    if category.upper() == 'COMPLAIN':
        return df_source['JUMLAH_COMPLAIN'].sum()
    else:
        return len(df_source)

# --- 5. TAMPILAN DASHBOARD ---

st.title("📊 ATM Executive Dashboard")

# === AREA KONTROL KANAN ATAS (BULAN & KATEGORI) ===
col_filter_left, col_filter_right = st.columns([1, 1])

with col_filter_left:
    # 1. FILTER KATEGORI
    kategori_list = ["Elastic", "Complain", "DF Repeat", "OUT Flm", "Cash Out"]
    selected_kategori = st.radio("Pilih Kategori:", kategori_list, horizontal=True)

with col_filter_right:
    # 2. FILTER BULAN (Dinamis)
    
    available_months_num = df['MONTH_NUM'].dropna().unique()
    available_months_num.sort()
    
    month_map = df[['MONTH_NUM', 'BULAN']].drop_duplicates().set_index('MONTH_NUM')['BULAN'].to_dict()
    sorted_month_names = [month_map[m] for m in available_months_num]
    
    default_ix = len(sorted_month_names) - 1 if len(sorted_month_names) > 0 else 0
    
    selected_month_name = st.selectbox("Pilih Bulan Analisis:", sorted_month_names, index=default_ix)
    
    current_month_num = [k for k, v in month_map.items() if v == selected_month_name][0]
    previous_month_num = current_month_num - 1
    previous_month_name = month_map.get(previous_month_num, "N/A")
    
    label_current = selected_month_name[:3].upper() 
    label_previous = previous_month_name[:3].upper() 

# --- PENYESUAIAN DATA FRAME (DINAMIS) ---
df_cat = df[df['KATEGORI'].astype(str).str.contains(selected_kategori, case=False, na=False)]
df_current = df_cat[df_cat['MONTH_NUM'] == current_month_num]
df_previous = df_cat[df_cat['MONTH_NUM'] == previous_month_num]

df_dec = df_current 
df_nov = df_previous 


# LAYOUT SIMETRIS 50:50
col_left, col_right = st.columns(2)

# =========================================================
# KIRI: GLOBAL + BREAKDOWN
# =========================================================
with col_left:
    st.subheader(f"🌍 All {selected_kategori} Overview (Month: {label_current})")
    
    # 1. TABEL GLOBAL
    w_labels = ['W1', 'W2', 'W3', 'W4']
    weeks_passed = df_current['WEEK'].nunique()
    if weeks_passed == 0: weeks_passed = 1 
    
    freq_dec = get_count_or_sum(df_current, selected_kategori)
    freq_nov = get_count_or_sum(df_previous, selected_kategori)
    
    tid_dec = df_current['TID'].nunique()
    tid_nov = df_previous['TID'].nunique()
    
    weekly_data = {}
    for w in w_labels:
        data_w = df_current[df_current['WEEK'] == w]
        weekly_data[w] = {
            'freq': get_count_or_sum(data_w, selected_kategori), 
            'tid': data_w['TID'].nunique()
        }

    gl_data = {
        'Metric': ['Global Ticket (Freq)', 'Global Unique TID'],
        'Total ATM': [TOTAL_ASSET_GLOBAL, TOTAL_ASSET_GLOBAL],
        f'{label_previous}(prev)': [freq_nov if freq_nov > 0 else np.nan, tid_nov if tid_nov > 0 else np.nan],
        'W1': [weekly_data['W1']['freq'] if weekly_data['W1']['freq'] > 0 else np.nan, weekly_data['W1']['tid'] if weekly_data['W1']['tid'] > 0 else np.nan],
        'W2': [weekly_data['W2']['freq'] if weekly_data['W2']['freq'] > 0 else np.nan, weekly_data['W2']['tid'] if weekly_data['W2']['tid'] > 0 else np.nan],
        'W3': [weekly_data['W3']['freq'] if weekly_data['W3']['freq'] > 0 else np.nan, weekly_data['W3']['tid'] if weekly_data['W3']['tid'] > 0 else np.nan],
        'W4': [weekly_data['W4']['freq'] if weekly_data['W4']['freq'] > 0 else np.nan, weekly_data['W4']['tid'] if weekly_data['W4']['tid'] > 0 else np.nan],
        f'Σ {label_current}': [freq_dec, tid_dec],
        'Avg/Week': [round(freq_dec/weeks_passed, 1), round(tid_dec/weeks_passed, 1)],
        'Prob %': [
            f"{round((freq_dec/TOTAL_ASSET_GLOBAL)*100, 1)}%", 
            f"{round((tid_dec/TOTAL_ASSET_GLOBAL)*100, 1)}%"
        ]
    }
    
    df_global = pd.DataFrame(gl_data).set_index('Metric')
    df_global.columns = [f'{col}' if col not in ['Σ Dec', 'Nov(prev)', 'Prob %'] else col.replace('Dec', label_current).replace('Nov', label_previous) for col in df_global.columns]
    st.dataframe(df_global, use_container_width=True)
    
    # 2. TABEL BREAKDOWN (SHOW/HIDE)
    with st.expander(f"📂 Klik untuk Lihat Rincian Per Cabang ({label_current})", expanded=False):
        
        semua_cabang = pd.Series(POPULASI_MANUAL.keys()).str.upper().tolist()
        
        if selected_kategori.upper() == 'COMPLAIN':
            piv = df_current.pivot_table(index='CABANG', columns='WEEK', values='JUMLAH_COMPLAIN', aggfunc='sum', fill_value=0)
            nov_counts = df_previous.groupby('CABANG')['JUMLAH_COMPLAIN'].sum()
        else:
            piv = df_current.pivot_table(index='CABANG', columns='WEEK', values='TID', aggfunc='count', fill_value=0)
            nov_counts = df_previous.groupby('CABANG').size()


        piv = piv.reindex(semua_cabang, fill_value=0)
        for w in w_labels:
            if w not in piv.columns: piv[w] = 0
            
        df_br = pd.DataFrame(index=piv.index)
        df_br['Total ATM'] = df_br.index.map(POPULASI_MANUAL).fillna(0).astype(int)
        
        df_br[f'{label_previous}'] = df_br.index.map(nov_counts).fillna(0)
        df_br['W1'] = piv['W1']
        df_br['W2'] = piv['W2']
        df_br['W3'] = piv['W3']
        df_br['W4'] = piv['W4']
        
        df_br[f'Σ {label_current}'] = piv.sum(axis=1)
        df_br['Avg/Week'] = (df_br[f'Σ {label_current}'] / weeks_passed).round(1)
        
        df_br['Prob %'] = df_br.apply(lambda x: f"{round((x[f'Σ {label_current}']/x['Total ATM'])*100, 1)}%" if x['Total ATM'] > 0 else np.nan, axis=1)
        
        cols_to_clean = [f'{label_previous}', 'W1', 'W2', 'W3', 'W4', f'Σ {label_current}', 'Avg/Week']
        df_br[cols_to_clean] = df_br[cols_to_clean].replace(0, np.nan)
        df_br = df_br.fillna(np.nan).sort_values(by=f'Σ {label_current}', ascending=False, na_position='last')
        
        st.dataframe(df_br, use_container_width=True)

# =========================================================
# KANAN: TOP TID + DETAIL SLM INTERAKTIF + GRAFIK
# =========================================================
with col_right:
    st.subheader(f"🔥 Top 5 {selected_kategori} Unit Problem ({label_current})")
    
    # --- 1. DROPDOWN SORTIR FINAL ---
    col_sort_week, _ = st.columns([1, 2])
    with col_sort_week:
        sort_options = [f'Σ {label_current}', 'W1', 'W2', 'W3', 'W4', f'{label_previous}(prev)']
        sort_options_clean = [opt for opt in sort_options if opt in df_current['WEEK'].unique() or opt.startswith('Σ') or opt.endswith('(prev)')]
        
        sel_sort_col = st.selectbox("Urutkan berdasarkan:", sort_options_clean, index=0, label_visibility="collapsed")
        
        sort_by_col_name = sel_sort_col.replace(f'(prev)', '') 
        if sort_by_col_name == f'{label_previous}':
             sort_by_col_name = f'{label_previous}'


    if not df_current.empty:
        
        # 2. PERSIAPAN DATA TOP TID
        if selected_kategori.upper() == 'COMPLAIN':
            p_tid = df_current.pivot_table(index=['TID', 'LOKASI', 'CABANG'], columns='WEEK', values='JUMLAH_COMPLAIN', aggfunc='sum', fill_value=0)
            nov_agg = df_previous.groupby('TID')['JUMLAH_COMPLAIN'].sum()
        else:
            p_tid = df_current.pivot_table(index=['TID', 'LOKASI', 'CABANG'], columns='WEEK', values='KATEGORI', aggfunc='count', fill_value=0)
            nov_agg = df_previous.groupby('TID').size()

        d_tid = p_tid.reset_index()
        for w in w_labels:
            if w not in d_tid.columns: d_tid[w] = 0
            
        d_tid[f'{label_previous}'] = d_tid['TID'].map(nov_agg).fillna(0)
        d_tid[f'Σ {label_current}'] = d_tid[w_labels].sum(axis=1)
        
        # --- LOGIC SORTING FINAL ---
        d_tid_sort = d_tid.copy()
        
        if sort_by_col_name == f'{label_previous}':
             d_tid_sort['SORT_COL'] = d_tid_sort[f'{label_previous}'].astype(int)
        elif sort_by_col_name in w_labels:
             d_tid_sort['SORT_COL'] = d_tid_sort[sort_by_col_name].astype(int)
        else: 
             d_tid_sort['SORT_COL'] = d_tid_sort[f'Σ {label_current}'].astype(int)
        
        d_tid_sort = d_tid_sort.sort_values(by='SORT_COL', ascending=False, na_position='last')
        top_5_data_raw = d_tid_sort.head(5)
            
        # --- HAPUS DESIMAL DENGAN UBAH KE INT UNTUK TAMPILAN ---
        d_tid_int = d_tid.copy()
        cols_to_clean_tid = [f'{label_previous}', 'W1', 'W2', 'W3', 'W4', f'Σ {label_current}']
        d_tid_int[cols_to_clean_tid] = d_tid_int[cols_to_clean_tid].replace(0, np.nan)

        # 3. ITERASI EXPANDER (Detail SLM)
        for index, row in top_5_data_raw.iterrows():
            tid = row['TID']
            
            row_display = d_tid_int[d_tid_int['TID'] == tid].iloc[0]
            
            current_sum = row_display[f'Σ {label_current}']
            nov_val = row_display[f'{label_previous}']
            
            header_title = f"TID: **{tid}** | **{int(current_sum) if not pd.isna(current_sum) else 0}x** ({label_current}) | {row_display['LOKASI']} ({row_display['CABANG']})"
            
            w1_val = int(row_display['W1']) if not pd.isna(row_display['W1']) else '-'
            w2_val = int(row_display['W2']) if not pd.isna(row_display['W2']) else '-'
            w3_val = int(row_display['W3']) if not pd.isna(row_display['W3']) else '-'
            w4_val = int(row_display['W4']) if not pd.isna(row_display['W4']) else '-'
            nov_val_fmt = int(nov_val) if not pd.isna(nov_val) else '-'
            
            subheader_info = (
                f"{label_previous}: {nov_val_fmt} | W1: {w1_val} | W2: {w2_val} | W3: {w3_val} | W4: {w4_val}"
            )
            
            with st.expander(header_title):
                st.caption(subheader_info)
                st.markdown("---")
                
                # Cek Detail SLM
                if not df_slm.empty and 'TID' in df_slm.columns:
                    slm_history = df_slm[df_slm['TID'].astype(str) == str(tid)].sort_values(by='TGL Visit SLM', ascending=False)
                    
                    if slm_history.empty:
                        st.info(f"Tidak ada catatan SLM di log untuk TID {tid}.")
                    else:
                        st.markdown("<p style='font-weight: bold; font-size: 14px; margin-bottom: 5px;'>Riwayat Kunjungan SLM:</p>", unsafe_allow_html=True)
                        slm_history['TANGGAL'] = slm_history['TGL Visit SLM'].dt.strftime('%Y-%m-%d')
                        slm_display = slm_history[['TANGGAL', 'Action']].copy()
                        slm_display.columns = ['Tanggal Visit', 'Action di Lapangan']
                        
                        st.dataframe(slm_display, hide_index=True, use_container_width=True)
                else:
                    st.info("Fitur Detail SLM tidak aktif.")

        # 3. GRAFIK TREND HARIAN
        st.markdown("<div style='margin-top: 20px'></div>", unsafe_allow_html=True) 
        st.subheader(f"📈 Tren Harian (Ticket Volume - {label_current})")
        
        if selected_kategori.upper() == 'COMPLAIN':
            daily_trend = df_current.groupby(df_current['TANGGAL'].dt.date)['JUMLAH_COMPLAIN'].sum().reset_index(name='JUMLAH')
        else:
            daily_trend = df_current.groupby(df_current['TANGGAL'].dt.date).size().reset_index(name='JUMLAH')
            
        daily_trend.columns = ['Tanggal', 'Jumlah']
        
        fig_line = px.line(daily_trend, x='Tanggal', y='Jumlah', markers=True, 
                           title=f"Tiket Harian: {selected_kategori}", 
                           labels={'Jumlah': 'Jumlah Problem', 'Tanggal': 'Tanggal'},
                           text='Jumlah',
                           height=200) 
                           
        fig_line.update_traces(line_color='#E74C3C', line_width=3, textposition='top center')
        fig_line.update_layout(margin=dict(l=20, r=20, t=30, b=20), uniformtext_minsize=8, uniformtext_mode='hide')
        
        st.plotly_chart(fig_line, use_container_width=True)
    else:

        st.info(f"Tidak ada data {selected_kategori} di bulan {selected_month_name}.")
