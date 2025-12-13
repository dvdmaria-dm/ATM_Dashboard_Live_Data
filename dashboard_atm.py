import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# CSS Styling
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    .dataframe {font-size: 13px !important;}
    th {background-color: #262730 !important; color: white !important;}
    .top-card {
        background-color: #1E1E1E; 
        border-left: 4px solid #FF4B4B;
        padding: 10px; 
        margin-bottom: 8px; 
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA (V49 ENGINE) ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_NAME = 'AIMS_Master'

try:
    if "gcp_service_account" not in st.secrets:
        st.error("Secrets not found.")
        st.stop()
    
    creds = st.secrets["gcp_service_account"]
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
    st.error(f"Connection Error: {e}")
    sys.exit()

@st.cache_data(ttl=600)
def load_data():
    try:
        sh = gc.open_by_url(SHEET_URL)
        ws = sh.worksheet(SHEET_NAME)
        all_values = ws.get_all_values()
        
        if not all_values: return pd.DataFrame()

        headers = all_values[0]
        rows = all_values[1:]
        df = pd.DataFrame(rows, columns=headers)
        
        # CLEANING
        df = df.loc[:, df.columns != '']
        df.columns = df.columns.str.strip().str.upper()

        if 'TANGGAL' in df.columns:
            df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=True, errors='coerce')
        
        # LOGIKA KHUSUS KOLOM COMPLAIN
        if 'JUMLAH_COMPLAIN' in df.columns:
             # Bersihkan karakter aneh, ganti '-' jadi 0
             df['JUMLAH_COMPLAIN'] = pd.to_numeric(df['JUMLAH_COMPLAIN'].astype(str).str.replace('-', '0'), errors='coerce').fillna(0).astype(int)
        else:
             df['JUMLAH_COMPLAIN'] = 0

        # Normalisasi Kolom WEEK
        if 'WEEK' not in df.columns and 'BULAN_WEEK' in df.columns:
            df['WEEK'] = df['BULAN_WEEK']
            
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- 3. LOGIKA MATRIX EXECUTIVE (V53 - FIXED LOGIC) ---
def build_executive_summary(df_curr, kategori_pilih):
    """
    Membangun tabel W1, W2, W3, W4.
    LOGIKA PERBAIKAN:
    - Jika 'Complain': SUM kolom JUMLAH_COMPLAIN
    - Jika Lainnya: COUNT Baris (TID)
    """
    weeks = ['W1', 'W2', 'W3', 'W4']
    
    # Tentukan Mode Hitung
    is_complain_mode = 'Complain' in kategori_pilih
    
    # 1. Hitung Baris "Global Ticket (Freq)"
    row_ticket = {}
    total_ticket = 0
    
    for w in weeks:
        df_week = df_curr[df_curr['WEEK'] == w] if 'WEEK' in df_curr.columns else pd.DataFrame()
        
        if not df_week.empty:
            if is_complain_mode:
                # Jika Complain -> SUM
                val = df_week['JUMLAH_COMPLAIN'].sum()
            else:
                # Jika Lainnya -> COUNT BARIS
                val = len(df_week)
        else:
            val = 0
            
        row_ticket[w] = val
        total_ticket += val
    
    row_ticket['TOTAL'] = total_ticket
    row_ticket['AVG/WEEK'] = round(total_ticket / 4, 1)

    # 2. Hitung Baris "Global Unique TID" (Sama untuk semua kategori)
    row_tid = {}
    total_tid_set = set()
    for w in weeks:
        tids = df_curr[df_curr['WEEK'] == w]['TID'].unique() if 'WEEK' in df_curr.columns and 'TID' in df_curr.columns else []
        count = len(tids)
        row_tid[w] = count
        total_tid_set.update(tids)
    
    row_tid['TOTAL'] = len(total_tid_set)
    row_tid['AVG/WEEK'] = round(len(total_tid_set) / 4, 1)

    matrix_df = pd.DataFrame([row_ticket, row_tid], index=['Global Ticket (Freq)', 'Global Unique TID'])
    
    cols_order = ['W1', 'W2', 'W3', 'W4', 'TOTAL', 'AVG/WEEK']
    for c in cols_order:
        if c not in matrix_df.columns: matrix_df[c] = 0
            
    return matrix_df[cols_order]

# --- 4. UI DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data belum tersedia.")
else:
    # JUDUL
    st.markdown("### 🇮🇩 ATM Executive Dashboard")
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        if 'KATEGORI' in df.columns:
            cats = sorted(df['KATEGORI'].dropna().unique().tolist())
            sel_cat = st.radio("Pilih Kategori:", cats, index=0, horizontal=True)
        else:
            sel_cat = "Semua"
    with col_f2:
        if 'BULAN' in df.columns:
            months = df['BULAN'].dropna().unique().tolist()
            sel_mon = st.selectbox("Pilih Bulan Analisis:", months, index=len(months)-1 if months else 0)
        else:
            sel_mon = "Semua"

    # FILTER UTAMA
    df_main = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_main.columns:
        df_main = df_main[df_main['KATEGORI'] == sel_cat]
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]

    st.markdown("---")
    
    # CEK MODE HITUNG
    is_complain_mode = 'Complain' in sel_cat

    # === LAYOUT UTAMA (50:50) ===
    col_left, col_right = st.columns(2)

    # ---------------------------------------------------------
    # KOLOM KIRI (MATRIX & BREAKDOWN)
    # ---------------------------------------------------------
    with col_left:
        st.subheader(f"🌏 {sel_cat} Overview (Month: {sel_mon})")
        
        # 1. MATRIX TABLE (LOGIKA BARU)
        matrix_result = build_executive_summary(df_main, sel_cat)
        st.dataframe(
            matrix_result.style.highlight_max(axis=1, color='#262730').format("{:,.0f}"), 
            use_container_width=True
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 2. BREAKDOWN CABANG (Show Week)
        with st.expander(f"📂 Klik untuk Lihat Rincian Per Cabang ({sel_mon})"):
            if 'CABANG' in df_main.columns:
                # Logika Agregasi Cabang
                if is_complain_mode:
                    # Sum Jumlah Complain
                    agg_col = 'JUMLAH_COMPLAIN'
                else:
                    # Count Baris (Kita buat kolom dummy hitung)
                    df_main['COUNT_TEMP'] = 1
                    agg_col = 'COUNT_TEMP'
                
                # Kita group by Cabang dan Week agar week muncul
                cabang_pivot = df_main.groupby(['CABANG', 'WEEK'])[agg_col].sum().reset_index()
                cabang_pivot = cabang_pivot.sort_values(agg_col, ascending=False)
                cabang_pivot.columns = ['CABANG', 'WEEK', 'TOTAL TIKET']
                
                st.dataframe(cabang_pivot, use_container_width=True, hide_index=True)
            else:
                st.info("Data Cabang tidak tersedia.")

    # ---------------------------------------------------------
    # KOLOM KANAN (TOP 5 & TREND)
    # ---------------------------------------------------------
    with col_right:
        # 3. TOP 5 UNIT PROBLEM (Show Week)
        st.subheader(f"🔥 Top 5 {sel_cat} Unit Problem ({sel_mon})")
        
        if 'TID' in df_main.columns and 'LOKASI' in df_main.columns:
            # Tentukan kolom nilai
            metric_field = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID' # Kalau count TID, nanti di-count
            
            if is_complain_mode:
                top5 = df_main.groupby(['TID', 'LOKASI', 'WEEK'])[metric_field].sum().reset_index()
                top5 = top5.sort_values(metric_field, ascending=False).head(5)
                val_col = metric_field
            else:
                # Count Mode
                top5 = df_main.groupby(['TID', 'LOKASI', 'WEEK']).size().reset_index(name='TOTAL_FREQ')
                top5 = top5.sort_values('TOTAL_FREQ', ascending=False).head(5)
                val_col = 'TOTAL_FREQ'
            
            if not top5.empty:
                for idx, row in top5.iterrows():
                    st.markdown(f"""
                    <div class="top-card">
                        <div style="display:flex; justify-content:space-between;">
                            <span style="font-weight:bold; font-size:16px; color:#FFF;">
                                TID: {row['TID']} <span style="font-size:12px; color:#AAA; margin-left:5px;">({row['WEEK']})</span>
                            </span>
                            <span style="color:#FF4B4B; font-weight:bold;">{row[val_col]}x</span>
                        </div>
                        <div style="font-size:12px; color:#AAA; margin-top:4px;">{row['LOKASI']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("Tidak ada data Top 5.")

        # 4. GRAFIK TREN HARIAN (Per Tanggal + Label Angka)
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"📈 Tren Harian (Ticket Volume - {sel_mon})")
        
        if 'TANGGAL' in df_main.columns:
            # Group by Tanggal
            if is_complain_mode:
                daily = df_main.groupby('TANGGAL')['JUMLAH_COMPLAIN'].sum().reset_index()
                y_val = 'JUMLAH_COMPLAIN'
            else:
                daily = df_main.groupby('TANGGAL').size().reset_index(name='TOTAL_FREQ')
                y_val = 'TOTAL_FREQ'
            
            if not daily.empty:
                # Pastikan urut tanggal
                daily = daily.sort_values('TANGGAL')
                
                fig = px.line(daily, x='TANGGAL', y=y_val, markers=True, text=y_val, template="plotly_dark")
                fig.update_traces(
                    line_color='#FF4B4B', 
                    line_width=3,
                    textposition="top center" # Munculkan angka di atas titik
                )
                fig.update_layout(
                    xaxis_title=None, 
                    yaxis_title="Volume", 
                    height=300,
                    margin=dict(l=0, r=0, t=20, b=0),
                    # FORMAT TANGGAL HARIAN
                    xaxis=dict(
                        tickformat="%d %b", # Format: 01 Dec
                        dtick="D1" # Paksa interval harian (1 Hari)
                    )
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Data harian kosong.")
