import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN (LAYOUT MEWAH) ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# Styling CSS Khusus agar mirip tampilan "Dark Executive"
st.markdown("""
<style>
    [data-testid="stMetricValue"] {
        font-size: 24px;
        color: #4ea8de;
    }
    .big-font {
        font-size:20px !important;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA (ENGINE V49 - JANGAN DIUBAH LAGI) ---
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
        
        # CLEANING & MAPPING KOLOM
        df = df.loc[:, df.columns != '']
        df.columns = df.columns.str.strip().str.upper()

        # Pastikan kolom-kolom krusial ada (berdasarkan screenshot Sheetmu)
        # TANGGAL, KATEGORI, BULAN, WEEK, JUMLAH_COMPLAIN, TID, LOKASI, CABANG
        
        if 'TANGGAL' in df.columns:
            df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=True, errors='coerce')
        
        # Konversi angka
        if 'JUMLAH_COMPLAIN' in df.columns:
             df['JUMLAH_COMPLAIN'] = pd.to_numeric(df['JUMLAH_COMPLAIN'].replace('-', '0'), errors='coerce').fillna(0).astype(int)
        else:
             df['JUMLAH_COMPLAIN'] = 1 # Fallback jika kolom tidak ada
             
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- 3. UI DASHBOARD EKSEKUTIF ---
df = load_data()

if df.empty:
    st.warning("Data belum tersedia. Pastikan koneksi aman.")
else:
    # Header Dashboard
    st.title("🇮🇩 ATM Executive Dashboard")
    
    # --- BAGIAN FILTER ATAS (Mirip Screenshot Localhost) ---
    col_filter1, col_filter2 = st.columns([2, 1])
    
    with col_filter1:
        st.subheader("Pilih Kategori:")
        # Ambil unik kategori dari data (Elastic, Complain, dll)
        if 'KATEGORI' in df.columns:
            kategori_list = sorted(df['KATEGORI'].dropna().unique().tolist())
            # Radio button horizontal
            pilih_kategori = st.radio("Kategori", kategori_list, index=0, horizontal=True, label_visibility="collapsed")
        else:
            pilih_kategori = "Semua"
            st.info("Kolom 'KATEGORI' tidak ditemukan di Excel.")

    with col_filter2:
        st.subheader("Pilih Bulan:")
        if 'BULAN' in df.columns:
            bulan_list = df['BULAN'].dropna().unique().tolist()
            # Coba urutkan bulan jika formatnya dikenali, kalau tidak alphabet
            pilih_bulan = st.selectbox("Bulan", bulan_list, index=len(bulan_list)-1 if bulan_list else 0, label_visibility="collapsed")
        else:
            pilih_bulan = "Semua"

    # --- FILTERING LOGIC ---
    df_filtered = df.copy()
    
    # Filter 1: Kategori
    if 'KATEGORI' in df.columns and pilih_kategori:
        df_filtered = df_filtered[df_filtered['KATEGORI'] == pilih_kategori]
        
    # Filter 2: Bulan
    if 'BULAN' in df.columns and pilih_bulan:
        df_filtered = df_filtered[df_filtered['BULAN'] == pilih_bulan]

    st.markdown("---")

    # --- BAGIAN UTAMA (KIRI: TABEL, KANAN: LIST) ---
    col_main1, col_main2 = st.columns([3, 2])

    with col_main1:
        st.subheader(f"🌏 {pilih_kategori} Overview (Month: {pilih_bulan})")
        
        # Metric Cards Sederhana di atas tabel
        m1, m2 = st.columns(2)
        total_ticket = df_filtered['JUMLAH_COMPLAIN'].sum() if 'JUMLAH_COMPLAIN' in df.columns else len(df_filtered)
        unique_tid = df_filtered['TID'].nunique() if 'TID' in df.columns else 0
        
        m1.metric("Global Ticket (Freq)", f"{total_ticket:,}")
        m2.metric("Global Unique TID", f"{unique_tid:,}")
        
        # TABEL OVERVIEW PER MINGGU (WEEK)
        # Kita pivot datanya: Baris = Cabang, Kolom = Week, Isi = Jumlah Complain
        if 'WEEK' in df_filtered.columns and 'CABANG' in df_filtered.columns:
            try:
                # Pivot Table
                pivot_week = df_filtered.pivot_table(
                    index='CABANG', 
                    columns='WEEK', 
                    values='JUMLAH_COMPLAIN', 
                    aggfunc='sum', 
                    fill_value=0
                )
                # Tambahkan kolom Total
                pivot_week['TOTAL'] = pivot_week.sum(axis=1)
                # Sort berdasarkan Total tertinggi
                pivot_week = pivot_week.sort_values('TOTAL', ascending=False)
                
                st.dataframe(pivot_week, use_container_width=True)
            except Exception as e:
                st.info("Data tidak cukup untuk membuat Pivot Table Mingguan.")
        else:
            st.info("Kolom 'WEEK' atau 'CABANG' tidak ditemukan untuk membuat tabel overview.")

    with col_main2:
        st.subheader(f"🔥 Top 5 {pilih_kategori} Unit Problem")
        
        # List Top 5 TID/Lokasi bermasalah
        if 'TID' in df_filtered.columns and 'LOKASI' in df_filtered.columns:
            top_problem = df_filtered.groupby(['TID', 'LOKASI'])['JUMLAH_COMPLAIN'].sum().reset_index()
            top_problem = top_problem.sort_values('JUMLAH_COMPLAIN', ascending=False).head(5)
            
            for index, row in top_problem.iterrows():
                tid = row['TID']
                loc = row['LOKASI']
                count = row['JUMLAH_COMPLAIN']
                # Tampilan Card Custom ala Executive
                st.markdown(f"""
                <div style="background-color: #262730; padding: 10px; border-radius: 5px; margin-bottom: 10px; border-left: 5px solid #ff4b4b;">
                    <b>TID: {tid}</b> | <span style="color:#ff4b4b; font-weight:bold;">{count}x Ticket</span><br>
                    <small>{loc}</small>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Kolom TID/LOKASI tidak lengkap.")

    # --- BAGIAN BAWAH (CHART) ---
    st.subheader(f"📈 Tren Harian (Ticket Volume - {pilih_bulan})")
    
    if 'TANGGAL' in df_filtered.columns:
        daily_trend = df_filtered.groupby('TANGGAL')['JUMLAH_COMPLAIN'].sum().reset_index()
        if not daily_trend.empty:
            fig = px.line(daily_trend, x='TANGGAL', y='JUMLAH_COMPLAIN', markers=True, template="plotly_dark")
            fig.update_traces(line_color='#ff4b4b', line_width=3)
            fig.update_layout(xaxis_title="Tanggal", yaxis_title="Jumlah Tiket", height=350)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Tidak ada data harian untuk ditampilkan.")
