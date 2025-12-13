import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN (EKSEKUTIF DARK MODE) ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# CSS untuk memiripkan tampilan dengan Screenshot (Card Hitam & Teks)
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 2rem;}
    
    /* Styling Tabel Matrix agar mirip screenshot */
    .dataframe {font-size: 14px !important;}
    th {background-color: #262730 !important; color: white !important;}
    
    /* Styling Card Top 5 */
    .top-card {
        background-color: #1E1E1E; 
        border-left: 4px solid #FF4B4B;
        padding: 10px; 
        margin-bottom: 8px; 
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA (ENGINE V49 - STABIL) ---
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
        
        if 'JUMLAH_COMPLAIN' in df.columns:
             df['JUMLAH_COMPLAIN'] = pd.to_numeric(df['JUMLAH_COMPLAIN'].replace('-', '0'), errors='coerce').fillna(0).astype(int)
        else:
             df['JUMLAH_COMPLAIN'] = 1 
             
        # Pastikan kolom WEEK/BULAN_WEEK ada (sesuai screenshot kolom S)
        # Kita normalisasi nama kolom minggu
        if 'WEEK' not in df.columns and 'BULAN_WEEK' in df.columns:
            df['WEEK'] = df['BULAN_WEEK'] # Rename logic
            
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- 3. LOGIKA BUILD TABLE MATRIX ---
def build_executive_summary(df_curr):
    """
    Fungsi ini membangun tabel W1, W2, W3, W4 persis seperti screenshot.
    Baris 1: Global Ticket (Sum Jumlah Complain)
    Baris 2: Global Unique TID (Count Unique TID)
    """
    # Siapkan Kolom Mingguan
    weeks = ['W1', 'W2', 'W3', 'W4']
    
    # 1. Hitung Baris "Global Ticket (Freq)"
    row_ticket = {}
    total_ticket = 0
    for w in weeks:
        # Filter data minggu ini
        val = df_curr[df_curr['WEEK'] == w]['JUMLAH_COMPLAIN'].sum() if 'WEEK' in df_curr.columns else 0
        row_ticket[w] = val
        total_ticket += val
    
    # Tambah kolom Total & Avg
    row_ticket['TOTAL'] = total_ticket
    row_ticket['AVG/WEEK'] = round(total_ticket / 4, 1)

    # 2. Hitung Baris "Global Unique TID"
    row_tid = {}
    total_tid_set = set()
    for w in weeks:
        tids = df_curr[df_curr['WEEK'] == w]['TID'].unique() if 'WEEK' in df_curr.columns and 'TID' in df_curr.columns else []
        count = len(tids)
        row_tid[w] = count
        total_tid_set.update(tids)
    
    row_tid['TOTAL'] = len(total_tid_set)
    row_tid['AVG/WEEK'] = round(len(total_tid_set) / 4, 1)

    # Buat DataFrame Matrix
    matrix_df = pd.DataFrame([row_ticket, row_tid], index=['Global Ticket (Freq)', 'Global Unique TID'])
    
    # Reorder kolom biar rapi
    cols_order = ['W1', 'W2', 'W3', 'W4', 'TOTAL', 'AVG/WEEK']
    # Pastikan kolom ada (jika data minggu kosong, isi 0)
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
    
    # FILTER AREA
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        # Filter Kategori (Radio Horizontal)
        if 'KATEGORI' in df.columns:
            cats = sorted(df['KATEGORI'].dropna().unique().tolist())
            sel_cat = st.radio("Pilih Kategori:", cats, index=0, horizontal=True)
        else:
            sel_cat = "Semua"
            
    with col_f2:
        # Filter Bulan
        if 'BULAN' in df.columns:
            months = df['BULAN'].dropna().unique().tolist()
            # Default pilih bulan terakhir
            sel_mon = st.selectbox("Pilih Bulan Analisis:", months, index=len(months)-1 if months else 0)
        else:
            sel_mon = "Semua"

    # TERAPKAN FILTER
    df_main = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_main.columns:
        df_main = df_main[df_main['KATEGORI'] == sel_cat]
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]

    st.markdown("---")

    # === LAYOUT UTAMA (50:50) ===
    col_left, col_right = st.columns(2)

    # ---------------------------------------------------------
    # KOLOM KIRI (THE MATRIX TABLE & BREAKDOWN)
    # ---------------------------------------------------------
    with col_left:
        st.subheader(f"🌏 {sel_cat} Overview (Month: {sel_mon})")
        
        # 1. MATRIX TABLE (PERSIS SCREENSHOT)
        matrix_result = build_executive_summary(df_main)
        # Tampilkan tabel dengan highlight max value
        st.dataframe(
            matrix_result.style.highlight_max(axis=1, color='#262730'), 
            use_container_width=True
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 2. BREAKDOWN CABANG (EXPANDER)
        # Kita buat seperti screenshot: "Klik untuk Lihat Rincian Per Cabang"
        with st.expander(f"📂 Klik untuk Lihat Rincian Per Cabang ({sel_mon})"):
            if 'CABANG' in df_main.columns:
                # Pivot per Cabang
                cabang_pivot = df_main.groupby('CABANG')['JUMLAH_COMPLAIN'].sum().reset_index()
                cabang_pivot = cabang_pivot.sort_values('JUMLAH_COMPLAIN', ascending=False)
                cabang_pivot.columns = ['CABANG', 'TOTAL TIKET']
                st.dataframe(cabang_pivot, use_container_width=True, hide_index=True)
            else:
                st.info("Data Cabang tidak tersedia.")

    # ---------------------------------------------------------
    # KOLOM KANAN (TOP 5 & TREND)
    # ---------------------------------------------------------
    with col_right:
        # 3. TOP 5 UNIT PROBLEM
        st.subheader(f"🔥 Top 5 {sel_cat} Unit Problem ({sel_mon})")
        
        if 'TID' in df_main.columns and 'LOKASI' in df_main.columns:
            top5 = df_main.groupby(['TID', 'LOKASI'])['JUMLAH_COMPLAIN'].sum().reset_index()
            top5 = top5.sort_values('JUMLAH_COMPLAIN', ascending=False).head(5)
            
            if not top5.empty:
                for idx, row in top5.iterrows():
                    # Layout Card Custom
                    st.markdown(f"""
                    <div class="top-card">
                        <div style="display:flex; justify-content:space-between;">
                            <span style="font-weight:bold; font-size:16px; color:#FFF;">TID: {row['TID']}</span>
                            <span style="color:#FF4B4B; font-weight:bold;">{row['JUMLAH_COMPLAIN']}x</span>
                        </div>
                        <div style="font-size:12px; color:#AAA; margin-top:4px;">{row['LOKASI']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("Tidak ada data Top 5.")

        # 4. GRAFIK TREN HARIAN (DI BAWAH TOP 5)
        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader(f"📈 Tren Harian (Ticket Volume - {sel_mon})")
        
        if 'TANGGAL' in df_main.columns:
            daily = df_main.groupby('TANGGAL')['JUMLAH_COMPLAIN'].sum().reset_index()
            if not daily.empty:
                fig = px.line(daily, x='TANGGAL', y='JUMLAH_COMPLAIN', markers=True, template="plotly_dark")
                fig.update_traces(line_color='#FF4B4B', line_width=3)
                fig.update_layout(
                    xaxis_title=None, 
                    yaxis_title="Tiket", 
                    height=250,
                    margin=dict(l=0, r=0, t=10, b=0)
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Data harian kosong.")
