import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# Styling CSS (Ultra Compact & Forced Centering)
st.markdown("""
<style>
    /* LAYOUTING */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    
    /* HEADER */
    h1 { font-size: 1.5rem !important; margin-bottom: 0.5rem !important;}
    h2 { font-size: 1.2rem !important; margin-bottom: 0px !important;}
    h3 { font-size: 1.1rem !important; margin-bottom: 5px !important;}
    
    /* FONT GLOBAL */
    html, body, [class*="st-emotion-"] { 
        font-size: 11px; 
    }

    /* HAPUS ELEMEN PENGGANGGU */
    #MainMenu, footer, header {visibility: hidden;}
    .st-emotion-cache-1j8u2d7 {visibility: hidden;} 
    
    /* PLOTLY MARGIN */
    .js-plotly-plot {margin-bottom: 0px !important;}
    .stPlotlyChart {margin-bottom: 0px !important;}
    
    /* FORCE CENTER ALIGNMENT (CSS LEVEL) */
    .dataframe {
        text-align: center !important;
    }
    .dataframe td {
        text-align: center !important;
    }
    .dataframe th {
        text-align: center !important;
        background-color: #262730 !important; 
        color: white !important;
    }
    /* Kecuali kolom pertama (index) rata kiri */
    .dataframe tbody th {
        text-align: left !important;
    }
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
            df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=False, errors='coerce')
        
        if 'JUMLAH_COMPLAIN' in df.columns:
             df['JUMLAH_COMPLAIN'] = pd.to_numeric(df['JUMLAH_COMPLAIN'].astype(str).str.replace('-', '0'), errors='coerce').fillna(0).astype(int)
        else:
             df['JUMLAH_COMPLAIN'] = 0

        if 'WEEK' not in df.columns and 'BULAN_WEEK' in df.columns:
            df['WEEK'] = df['BULAN_WEEK']
            
        if 'BULAN' in df.columns:
            df['BULAN'] = df['BULAN'].astype(str).str.strip().str.title()
            
        if 'TID' in df.columns:
            df['TID'] = df['TID'].astype(str)
        if 'LOKASI' in df.columns:
            df['LOKASI'] = df['LOKASI'].astype(str)
            
        return df

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame()

# --- HELPER: GET SHORT MONTH ---
def get_short_month_name(full_month_str):
    if not full_month_str: return ""
    return full_month_str[:3] # Ambil 3 huruf pertama (Dec, Nov, Jan)

def get_prev_month_full(curr_month):
    months = ['January', 'February', 'March', 'April', 'May', 'June', 
              'July', 'August', 'September', 'October', 'November', 'December']
    try:
        curr_idx = months.index(curr_month)
        prev_idx = (curr_idx - 1) if curr_idx > 0 else 11
        return months[prev_idx]
    except:
        return None

# --- HELPER: STYLE DATAFRAME (CENTER & ZERO BLANK) ---
def style_dataframe(df_to_style):
    # 1. Replace 0 with Empty String (Zero Suppression)
    df_clean = df_to_style.replace(0, "")
    
    # 2. Pandas Styler (Python Level Centering)
    return df_clean.style.set_properties(**{
        'text-align': 'center', 
        'white-space': 'nowrap',
        'font-size': '10px'
    }).set_table_styles([
        # Paksa Header Rata Tengah
        dict(selector='th', props=[('text-align', 'center !important'), ('font-size', '10px')]),
        # Paksa Cell Rata Tengah
        dict(selector='td', props=[('text-align', 'center !important')])
    ]).format(lambda x: "{:,.0f}".format(x) if x != "" else "")

# --- 3. LOGIKA MATRIX ---
def build_executive_summary(df_curr, df_prev, is_complain_mode, prev_month_short, curr_month_short):
    # 1. Data Current (W1-W4)
    weeks = ['W1', 'W2', 'W3', 'W4']
    row_ticket = {}
    row_tid = {}
    
    total_ticket_curr = 0
    total_tid_set_curr = set()

    for w in weeks:
        # Ticket
        df_week = df_curr[df_curr['WEEK'] == w] if 'WEEK' in df_curr.columns else pd.DataFrame()
        val = df_week['JUMLAH_COMPLAIN'].sum() if is_complain_mode and not df_week.empty else len(df_week) if not df_week.empty else 0
        row_ticket[w] = val
        total_ticket_curr += val
        
        # TID
        tids = df_curr[df_curr['WEEK'] == w]['TID'].unique() if 'WEEK' in df_curr.columns and 'TID' in df_curr.columns else []
        row_tid[w] = len(tids)
        total_tid_set_curr.update(tids)

    # 2. Data Prev Month (Total)
    val_prev = df_prev['JUMLAH_COMPLAIN'].sum() if is_complain_mode and not df_prev.empty else len(df_prev) if not df_prev.empty else 0
    tid_prev = len(df_prev['TID'].unique()) if 'TID' in df_prev.columns and not df_prev.empty else 0

    # 3. Header Styling
    col_prev = f"{prev_month_short}"
    col_total = f"Σ {curr_month_short.upper()}" # Menambahkan Sigma
    
    row_ticket[col_prev] = val_prev
    row_ticket[col_total] = total_ticket_curr
    
    row_tid[col_prev] = tid_prev
    row_tid[col_total] = len(total_tid_set_curr)

    # 4. Create DF
    matrix_df = pd.DataFrame([row_ticket, row_tid], index=['Global Ticket (Freq)', 'Global Unique TID'])
    
    # 5. Reorder
    cols_order = [col_prev, 'W1', 'W2', 'W3', 'W4', col_total]
    for c in cols_order:
        if c not in matrix_df.columns: matrix_df[c] = 0
        
    return matrix_df[cols_order]

# --- 4. UI DASHBOARD ---
df = load_data()

if df.empty:
    st.warning("Data belum tersedia.")
else:
    st.markdown("### 🇮🇩 ATM Executive Dashboard")
    
    # FILTER
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        if 'KATEGORI' in df.columns:
            cats = sorted(df['KATEGORI'].dropna().unique().tolist())
            sel_cat = st.radio("Kategori:", cats, index=0, horizontal=True, label_visibility="collapsed")
            st.caption(f"Kategori: **{sel_cat}**") 
        else:
            sel_cat = "Semua"
    with col_f2:
        if 'BULAN' in df.columns:
            months = df['BULAN'].unique().tolist()
            sel_mon = st.selectbox("Bulan:", months, index=len(months)-1 if months else 0, label_visibility="collapsed")
            st.caption(f"Bulan: **{sel_mon}**")
        else:
            sel_mon = "Semua"

    # PROCESS DATA
    df_cat = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_cat.columns:
        df_cat = df_cat[df_cat['KATEGORI'] == sel_cat]
        
    # Current Month
    df_main = df_cat.copy()
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]
        
    # Prev Month
    prev_mon_full = get_prev_month_full(sel_mon)
    df_prev = pd.DataFrame()
    if prev_mon_full and 'BULAN' in df_cat.columns:
        df_prev = df_cat[df_cat['BULAN'] == prev_mon_full]

    # Short Names for Columns
    curr_mon_short = get_short_month_name(sel_mon)
    prev_mon_short = get_short_month_name(prev_mon_full) if prev_mon_full else "Prev"

    is_complain_mode = 'Complain' in sel_cat
    
    st.markdown("---") 

    # =========================================================================
    # BAGIAN 1: GRAFIK TREN
    # =========================================================================
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
            daily['TANGGAL_STR'] = daily['TANGGAL'].dt.strftime('%d-%m-%Y')
            
            fig = px.line(daily, x='TANGGAL_STR', y=y_val, markers=True, text=y_val, template="plotly_dark")
            fig.update_traces(line_color='#FF4B4B', line_width=2, textposition="top center")
            
            fig.update_layout(
                xaxis_title=None, 
                yaxis_title=None,
                height=180, 
                margin=dict(l=10, r=10, t=10, b=10),
                xaxis=dict(tickangle=0, type='category')
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Data harian kosong.")

    # =========================================================================
    # BAGIAN 2: TABEL
    # =========================================================================
    col_left, col_right = st.columns(2)
    
    # NAMA KOLOM UTAMA
    col_prev_header = prev_mon_short
    col_curr_total_header = f"Σ {curr_mon_short.upper()}"

    with col_left:
        st.markdown(f"**🌏 {sel_cat} Overview**")
        
        # 1. OVERVIEW MATRIX
        matrix_result = build_executive_summary(df_main, df_prev, is_complain_mode, prev_mon_short, curr_mon_short)
        st.dataframe(style_dataframe(matrix_result), use_container_width=True)
        
        # 2. RINCIAN CABANG
        with st.expander(f"📂 Rincian Cabang"):
            if 'CABANG' in df_main.columns and 'WEEK' in df_main.columns:
                try:
                    val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                    agg_func = 'sum' if is_complain_mode else 'count'
                    
                    # Current
                    pivot_curr = df_main.pivot_table(index='CABANG', columns='WEEK', values=val_col, aggfunc=agg_func, fill_value=0)
                    desired_cols = ['W1', 'W2', 'W3', 'W4']
                    for c in desired_cols:
                        if c not in pivot_curr.columns: pivot_curr[c] = 0
                    pivot_curr = pivot_curr[desired_cols]
                    
                    # Prev
                    prev_grp = df_prev.groupby('CABANG')[val_col].agg(agg_func).reset_index(name=col_prev_header) if not df_prev.empty else pd.DataFrame(columns=['CABANG', col_prev_header])
                    prev_grp = prev_grp.set_index('CABANG')
                    
                    # Merge
                    final_cabang = pivot_curr.join(prev_grp, how='left').fillna(0)
                    final_cabang[col_curr_total_header] = final_cabang[['W1', 'W2', 'W3', 'W4']].sum(axis=1)
                    
                    # Sort & Clean
                    final_cols = [col_prev_header] + desired_cols + [col_curr_total_header]
                    final_cabang = final_cabang[final_cols]
                    final_cabang = final_cabang.sort_values(col_curr_total_header, ascending=False)
                    
                    st.dataframe(style_dataframe(final_cabang), use_container_width=True)
                except Exception as e:
                    st.error(f"Error pivot: {e}")

    with col_right:
        # HEADER DENGAN DROPDOWN SORTING
        c_head1, c_head2 = st.columns([2, 1])
        with c_head1:
             st.markdown(f"**🔥 Top 5 Problem Unit**")
        with c_head2:
             # DROPDOWN SORTING (Req No. 3)
             sort_options = [col_curr_total_header, 'W1', 'W2', 'W3', 'W4']
             sort_by = st.selectbox("Urutkan:", sort_options, index=0, label_visibility="collapsed")
        
        if 'TID' in df_main.columns and 'LOKASI' in df_main.columns and 'WEEK' in df_main.columns:
            try:
                val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                agg_func = 'sum' if is_complain_mode else 'count'
                
                # Current
                grouped_df = df_main.groupby(['TID', 'LOKASI', 'WEEK'])[val_col].agg(agg_func).reset_index(name='VAL')
                pivot_top5 = grouped_df.pivot_table(index=['TID', 'LOKASI'], columns='WEEK', values='VAL', aggfunc='sum', fill_value=0)
                desired_cols = ['W1', 'W2', 'W3', 'W4']
                for c in desired_cols:
                    if c not in pivot_top5.columns: pivot_top5[c] = 0
                pivot_top5 = pivot_top5[desired_cols]
                
                # Prev
                prev_grp_top5 = df_prev.groupby(['TID', 'LOKASI'])[val_col].agg(agg_func).reset_index(name=col_prev_header) if not df_prev.empty else pd.DataFrame(columns=['TID', 'LOKASI', col_prev_header])
                prev_grp_top5 = prev_grp_top5.set_index(['TID', 'LOKASI'])
                
                # Merge
                final_top5 = pivot_top5.join(prev_grp_top5, how='left').fillna(0)
                final_top5[col_curr_total_header] = final_top5[['W1', 'W2', 'W3', 'W4']].sum(axis=1)
                
                # Reorder
                final_cols_top = [col_prev_header] + desired_cols + [col_curr_total_header]
                final_top5 = final_top5[final_cols_top]
                
                # SORTING DINAMIS BERDASARKAN DROPDOWN
                # Sort descending berdasarkan kolom yang dipilih di dropdown
                top5_final = final_top5.sort_values(sort_by, ascending=False).head(5)
                
                st.dataframe(style_dataframe(top5_final), use_container_width=True)
                
            except Exception as e:
                 st.error(f"Error Top 5: {e}")
