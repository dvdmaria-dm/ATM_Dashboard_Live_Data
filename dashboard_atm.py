import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN (WIDE & COMPACT) ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# Styling CSS (Ultra Compact untuk Satu Layar Penuh)
st.markdown("""
<style>
    /* 1. MENGECILKAN JARAK ANTAR ELEMEN (PADDING) */
    .block-container {
        padding-top: 1.5rem !important; /* Jarak atas sangat minim */
        padding-bottom: 1rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    
    /* 2. MENGECILKAN HEADER AGAR TIDAK MAKAN TEMPAT */
    h1 { font-size: 1.4rem !important; margin-bottom: 0px !important; padding-bottom: 0px !important;}
    h2 { font-size: 1.2rem !important; margin-bottom: 0px !important; padding-bottom: 0px !important;}
    h3 { font-size: 1.1rem !important; margin-bottom: 5px !important; padding-bottom: 0px !important;}
    
    /* 3. FONT GLOBAL LEBIH KECIL (11px) */
    html, body, [class*="st-emotion-"] { 
        font-size: 11px; 
    }

    /* 4. TABEL ULTRA RINGKAS & RATA TENGAH */
    .dataframe {
        font-size: 10px !important; 
    }
    .dataframe td {
        text-align: center !important; 
        padding: 2px 4px !important; /* Padding sel sangat tipis */
    }
    .dataframe th {
        text-align: center !important;
        padding: 2px 4px !important;
        font-size: 10px !important;
    }
    /* Header Baris (Nama Cabang) Rata Kiri */
    .dataframe tbody th {
        text-align: left !important;
    }
    
    /* 5. MEMBUANG ELEMEN PENGGANGGU */
    #MainMenu, footer {visibility: hidden;}
    .st-emotion-cache-1j8u2d7 {visibility: hidden;} 
    
    /* Warna Header Tabel */
    th {background-color: #262730 !important; color: white !important;}
    thead tr th:first-child {display:none}
    tbody th {display:none}
    
    /* Margin Grafik Nol */
    .js-plotly-plot {margin-bottom: 0px !important;}
    .stPlotlyChart {margin-bottom: 0px !important;}
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA ---
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
        df = df.loc[:, ~df.columns.duplicated()]

        if 'TANGGAL' in df.columns:
            df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=True, errors='coerce')
        
        if 'JUMLAH_COMPLAIN' in df.columns:
             df['JUMLAH_COMPLAIN'] = pd.to_numeric(df['JUMLAH_COMPLAIN'].astype(str).str.replace('-', '0'), errors='coerce').fillna(0).astype(int)
        else:
             df['JUMLAH_COMPLAIN'] = 0

        if 'WEEK' not in df.columns and 'BULAN_WEEK' in df.columns:
            df['WEEK'] = df['BULAN_WEEK']
            
        if 'BULAN' in df.columns:
            df['BULAN'] = df['BULAN'].astype(str).str.strip()
            
        if 'TID' in df.columns:
            df['TID'] = df['TID'].astype(str)
        if 'LOKASI' in df.columns:
            df['LOKASI'] = df['LOKASI'].astype(str)
            
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- 3. LOGIKA MATRIX ---
def build_executive_summary(df_curr, is_complain_mode):
    weeks = ['W1', 'W2', 'W3', 'W4']
    row_ticket = {}
    total_ticket = 0
    for w in weeks:
        df_week = df_curr[df_curr['WEEK'] == w] if 'WEEK' in df_curr.columns else pd.DataFrame()
        val = df_week['JUMLAH_COMPLAIN'].sum() if is_complain_mode and not df_week.empty else len(df_week) if not df_week.empty else 0
        row_ticket[w] = val
        total_ticket += val
    
    row_ticket['TOTAL'] = total_ticket
    row_ticket['AVG/WEEK'] = round(total_ticket / 4, 1)

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
    # Header sangat tipis
    st.markdown("### 🇮🇩 ATM Executive Dashboard")
    
    # FILTER (Dibuat Compact)
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        if 'KATEGORI' in df.columns:
            cats = sorted(df['KATEGORI'].dropna().unique().tolist())
            sel_cat = st.radio("Kategori:", cats, index=0, horizontal=True, label_visibility="collapsed") # Label disembunyikan biar hemat
            st.caption(f"Kategori Terpilih: **{sel_cat}**") # Ganti label dengan caption kecil
        else:
            sel_cat = "Semua"
    with col_f2:
        if 'BULAN' in df.columns:
            months = df['BULAN'].unique().tolist()
            sel_mon = st.selectbox("Bulan:", months, index=len(months)-1 if months else 0, label_visibility="collapsed")
            st.caption(f"Bulan: **{sel_mon}**")
        else:
            sel_mon = "Semua"

    # DATA FILTERING
    df_main = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_main.columns:
        df_main = df_main[df_main['KATEGORI'] == sel_cat]
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]

    is_complain_mode = 'Complain' in sel_cat
    
    st.markdown("---") 

    # =========================================================================
    # BAGIAN 1: GRAFIK TREN (POSISI ATAS - FULL WIDTH - PENDEK)
    # =========================================================================
    # st.subheader hemat ruang dengan markdown bold kecil
    st.markdown(f"**📈 Tren Harian (Ticket Volume - {sel_mon})**")
    
    if 'TANGGAL' in df_main.columns:
        if is_complain_mode:
            daily = df_main.groupby('TANGGAL')['JUMLAH_COMPLAIN'].sum().reset_index()
            y_val = 'JUMLAH_COMPLAIN'
        else:
            daily = df_main.groupby('TANGGAL').size().reset_index(name='TOTAL_FREQ')
            y_val = 'TOTAL_FREQ'
        
        if not daily.empty:
            daily = daily.sort_values('TANGGAL')
            daily['TANGGAL_STR'] = daily['TANGGAL'].dt.strftime('%Y-%m-%d')
            
            fig = px.line(daily, x='TANGGAL_STR', y=y_val, markers=True, text=y_val, template="plotly_dark")
            fig.update_traces(line_color='#FF4B4B', line_width=2, textposition="top center") # Line width dikecilkan sedikit
            
            # Layout Sangat Compact (Tinggi 220px)
            fig.update_layout(
                xaxis_title=None, 
                yaxis_title=None, # Hilangkan label sumbu Y biar hemat
                height=220, # TINGGI DIKURANGI AGAR MUAT 1 LAYAR
                margin=dict(l=10, r=10, t=15, b=10), # Margin tipis
                xaxis=dict(
                    tickangle=0, 
                    type='category' 
                )
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Data harian kosong.")

    # =========================================================================
    # BAGIAN 2: TABEL (SPLIT COLUMN DI BAWAH GRAFIK)
    # =========================================================================
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown(f"**🌏 {sel_cat} Overview**")
        matrix_result = build_executive_summary(df_main, is_complain_mode)
        st.dataframe(matrix_result.style.highlight_max(axis=1, color='#262730').format("{:,.0f}"), use_container_width=True)
        
        # Expander default closed biar hemat, user klik kalau butuh detail
        with st.expander(f"📂 Rincian Cabang"):
            if 'CABANG' in df_main.columns and 'WEEK' in df_main.columns:
                try:
                    val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                    agg_func = 'sum' if is_complain_mode else 'count'
                    pivot_cabang = df_main.pivot_table(index='CABANG', columns='WEEK', values=val_col, aggfunc=agg_func, fill_value=0)
                    desired_cols = ['W1', 'W2', 'W3', 'W4']
                    for c in desired_cols:
                        if c not in pivot_cabang.columns: pivot_cabang[c] = 0
                    pivot_cabang = pivot_cabang[desired_cols]
                    pivot_cabang['TOTAL'] = pivot_cabang.sum(axis=1)
                    pivot_cabang = pivot_cabang.sort_values('TOTAL', ascending=False)
                    st.dataframe(pivot_cabang, use_container_width=True)
                except Exception as e:
                    st.error(f"Error pivot: {e}")

    with col_right:
        st.markdown(f"**🔥 Top 5 Problem Unit**")
        if 'TID' in df_main.columns and 'LOKASI' in df_main.columns and 'WEEK' in df_main.columns:
            try:
                val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                agg_func = 'sum' if is_complain_mode else 'count'
                grouped_df = df_main.groupby(['TID', 'LOKASI', 'WEEK'])[val_col].agg(agg_func).reset_index(name='VAL')
                pivot_top5 = grouped_df.pivot_table(index=['TID', 'LOKASI'], columns='WEEK', values='VAL', aggfunc='sum', fill_value=0)
                desired_cols = ['W1', 'W2', 'W3', 'W4']
                for c in desired_cols:
                    if c not in pivot_top5.columns: pivot_top5[c] = 0
                pivot_top5 = pivot_top5[desired_cols]
                pivot_top5['TOTAL'] = pivot_top5.sum(axis=1)
                top5_final = pivot_top5.sort_values('TOTAL', ascending=False).head(5)
                st.dataframe(top5_final, use_container_width=True)
            except Exception as e:
                 st.error(f"Error Top 5: {e}")
