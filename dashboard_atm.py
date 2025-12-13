import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN (LAYOUT WIDE) ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# CSS Custom untuk merapikan margin dan font
st.markdown("""
<style>
    .block-container {padding-top: 1rem; padding-bottom: 1rem;}
    [data-testid="stMetricValue"] {font-size: 20px; color: #4ea8de;}
    .stExpander {border: 1px solid #444; border-radius: 5px;}
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA (ENGINE V49 - SUDAH STABIL) ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_NAME = 'AIMS_Master'

try:
    if "gcp_service_account" not in st.secrets:
        st.error("Setup Error: Secrets not found.")
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
             
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- 3. UI DASHBOARD (LAYOUT LOCALHOST REPLICA) ---
df = load_data()

if df.empty:
    st.warning("Data belum tersedia. Pastikan koneksi aman.")
else:
    # JUDUL
    st.markdown("### 🇮🇩 ATM Executive Dashboard")
    
    # --- FILTER AREA ---
    col_filter1, col_filter2 = st.columns([2, 1])
    
    with col_filter1:
        st.caption("Pilih Kategori:")
        if 'KATEGORI' in df.columns:
            kategori_list = sorted(df['KATEGORI'].dropna().unique().tolist())
            pilih_kategori = st.radio("Kategori", kategori_list, index=0, horizontal=True, label_visibility="collapsed")
        else:
            pilih_kategori = "Semua"

    with col_filter2:
        st.caption("Pilih Bulan Analisis:")
        if 'BULAN' in df.columns:
            bulan_list = df['BULAN'].dropna().unique().tolist()
            pilih_bulan = st.selectbox("Bulan", bulan_list, index=len(bulan_list)-1 if bulan_list else 0, label_visibility="collapsed")
        else:
            pilih_bulan = "Semua"

    # --- FILTERING ---
    df_filtered = df.copy()
    if 'KATEGORI' in df.columns and pilih_kategori:
        df_filtered = df_filtered[df_filtered['KATEGORI'] == pilih_kategori]
    if 'BULAN' in df.columns and pilih_bulan:
        df_filtered = df_filtered[df_filtered['BULAN'] == pilih_bulan]

    st.markdown("---")

    # --- MAIN LAYOUT (50:50 SPLIT) ---
    col_left, col_right = st.columns(2) # Simetris

    # ==========================
    # KOLOM KIRI (TABLES)
    # ==========================
    with col_left:
        st.subheader(f"🌍 {pilih_kategori} Overview (Month: {pilih_bulan})")
        
        # Metric Bar Kecil
        m1, m2 = st.columns(2)
        total_ticket = df_filtered['JUMLAH_COMPLAIN'].sum() if 'JUMLAH_COMPLAIN' in df.columns else len(df_filtered)
        unique_tid = df_filtered['TID'].nunique() if 'TID' in df.columns else 0
        m1.metric("Global Ticket (Freq)", f"{total_ticket:,}")
        m2.metric("Global Unique TID", f"{unique_tid:,}")

        # 1. TABEL GLOBAL (PIVOT CABANG)
        if 'CABANG' in df_filtered.columns:
            # Grouping sederhana per cabang
            global_table = df_filtered.groupby('CABANG')['JUMLAH_COMPLAIN'].sum().reset_index()
            global_table = global_table.rename(columns={'JUMLAH_COMPLAIN': 'TOTAL TIKET'})
            global_table = global_table.sort_values('TOTAL TIKET', ascending=False)
            
            # Tampilkan Tabel Utama
            st.dataframe(global_table, use_container_width=True, hide_index=True)
        else:
            st.info("Kolom CABANG tidak ditemukan.")

        # 2. BREAKDOWN SHOW/HIDE (EXPANDER)
        st.markdown("<br>", unsafe_allow_html=True) # Spasi
        with st.expander("📂 Klik untuk Lihat Rincian Data (Raw Data)"):
            st.write("Detail Data Tiket:")
            # Tampilkan kolom-kolom penting saja agar rapi
            cols_to_show = [c for c in ['TANGGAL', 'TID', 'LOKASI', 'JUMLAH_COMPLAIN', 'KETERANGAN'] if c in df_filtered.columns]
            if cols_to_show:
                st.dataframe(df_filtered[cols_to_show], use_container_width=True, hide_index=True)
            else:
                st.dataframe(df_filtered, use_container_width=True)

    # ==========================
    # KOLOM KANAN (TOP 5 & CHART)
    # ==========================
    with col_right:
        # 1. TOP 5 PROBLEM
        st.subheader(f"🔥 Top 5 {pilih_kategori} Unit Problem ({pilih_bulan})")
        
        if 'TID' in df_filtered.columns and 'LOKASI' in df_filtered.columns:
            # Cari Top 5
            top_problem = df_filtered.groupby(['TID', 'LOKASI'])['JUMLAH_COMPLAIN'].sum().reset_index()
            top_problem = top_problem.sort_values('JUMLAH_COMPLAIN', ascending=False).head(5)
            
            # Tampilan Custom Card
            if not top_problem.empty:
                for index, row in top_problem.iterrows():
                    tid = row['TID']
                    loc = row['LOKASI']
                    count = row['JUMLAH_COMPLAIN']
                    st.markdown(f"""
                    <div style="
                        background-color: #1E1E1E; 
                        padding: 10px 15px; 
                        border-radius: 8px; 
                        margin-bottom: 8px; 
                        border: 1px solid #333;
                        display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <span style="color: #FFF; font-weight: bold;">TID: {tid}</span><br>
                            <span style="color: #AAA; font-size: 12px;">{loc}</span>
                        </div>
                        <div style="color: #FF4B4B; font-weight: bold; font-size: 16px;">
                            {count}x
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("Tidak ada data Top 5.")

        # 2. GRAFIK TREN HARIAN (DI BAWAH TOP 5)
        st.markdown("<br>", unsafe_allow_html=True) # Spasi
        st.subheader(f"📈 Tren Harian (Ticket Volume - {pilih_bulan})")
        
        if 'TANGGAL' in df_filtered.columns:
            daily_trend = df_filtered.groupby('TANGGAL')['JUMLAH_COMPLAIN'].sum().reset_index()
            if not daily_trend.empty:
                fig = px.line(daily_trend, x='TANGGAL', y='JUMLAH_COMPLAIN', markers=True, template="plotly_dark")
                fig.update_traces(line_color='#FF4B4B', line_width=3)
                fig.update_layout(
                    xaxis_title="Tanggal", 
                    yaxis_title="Jumlah Tiket",
                    height=300, # Tinggi disesuaikan agar pas di kanan
                    margin=dict(l=20, r=20, t=30, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Data harian tidak tersedia.")
