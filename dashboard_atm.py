import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="ATM Executive Dashboard", initial_sidebar_state="collapsed")

# Styling CSS (Ultra Compact, Kategori Besar & Rata Tengah Mutlak)
st.markdown("""
<style>
    /* LAYOUTING */
    .block-container {
        padding-top: 1.5rem !important;
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
    
    /* --- CSS PERBAIKAN RATA TENGAH (THE FINAL BOSS) --- */
    /* Target tabel dataframe secara spesifik */
    [data-testid="stDataFrame"] {
        width: 100%;
    }
    
    /* Target Header (th) */
    [data-testid="stDataFrame"] th {
        text-align: center !important;
        background-color: #262730 !important; 
        color: white !important;
        font-size: 11px !important;
    }
    
    /* Target Cell (td) */
    [data-testid="stDataFrame"] td {
        text-align: center !important;
        font-size: 11px !important;
    }
    
    /* Target DIV di dalam Header & Cell (Ini kuncinya!) */
    [data-testid="stDataFrame"] div[data-testid="stVerticalBlock"] > div {
        display: flex;
        justify-content: center;
        align-items: center;
    }

    /* Kecuali Kolom Pertama (Index/Nama Cabang) Rata Kiri */
    [data-testid="stDataFrame"] tbody th, 
    [data-testid="stDataFrame"] tbody th div {
        text-align: left !important;
        justify-content: flex-start !important;
    }

    /* --- STYLING RADIO BUTTON (KATEGORI) AGAR BESAR --- */
    div[role="radiogroup"] > label {
        font-size: 14px !important; /* Huruf lebih besar */
        font-weight: bold !important;
        background-color: #1E1E1E;
        padding: 5px 10px;
        border-radius: 5px;
        border: 1px solid #444;
        margin-right: 5px;
    }
    div[role="radiogroup"] > label:hover {
        border-color: #FF4B4B;
    }
    
    /* Styling Expander Top 5 */
    .streamlit-expanderHeader {
        font-size: 12px !important;
        font-weight: bold !important;
        background-color: #262730 !important;
        color: white !important;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# --- 2. KONEKSI DATA ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit"
SHEET_MAIN = 'AIMS_Master'
SHEET_SLM = 'SLM Visit Log'

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
        
        # --- LOAD DATA MASTER ---
        ws = sh.worksheet(SHEET_MAIN)
        all_values = ws.get_all_values()
        
        if not all_values: return pd.DataFrame(), pd.DataFrame()

        headers = all_values[0]
        rows = all_values[1:]
        df = pd.DataFrame(rows, columns=headers)
        
        # Cleaning Master
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
            df['TID'] = df['TID'].astype(str).str.strip()
        if 'LOKASI' in df.columns:
            df['LOKASI'] = df['LOKASI'].astype(str)

        # --- LOAD DATA SLM VISIT LOG ---
        try:
            ws_slm = sh.worksheet(SHEET_SLM)
            slm_values = ws_slm.get_all_values()
            if slm_values:
                h_slm = slm_values[0]
                r_slm = slm_values[1:]
                df_slm = pd.DataFrame(r_slm, columns=h_slm)
                df_slm.columns = df_slm.columns.str.strip().str.upper()
                if 'TID' in df_slm.columns:
                    df_slm['TID'] = df_slm['TID'].astype(str).str.strip()
                date_cols = [c for c in df_slm.columns if 'TGL' in c or 'VISIT' in c or 'DATE' in c]
                df_slm['TGL_VISIT'] = df_slm[date_cols[0]] if date_cols else "-"
                act_cols = [c for c in df_slm.columns if 'ACTION' in c or 'TINDAKAN' in c or 'KET' in c]
                df_slm['ACTION'] = df_slm[act_cols[0]] if act_cols else "-"
            else:
                df_slm = pd.DataFrame()
        except:
            df_slm = pd.DataFrame(columns=['TID', 'TGL_VISIT', 'ACTION'])
            
        return df, df_slm

    except Exception as e:
        st.error(f"Data Loading Error: {e}")
        return pd.DataFrame(), pd.DataFrame()

# --- HELPER FUNCTIONS ---
def get_short_month_name(full_month_str):
    if not full_month_str: return ""
    return full_month_str[:3]

def get_prev_month_full(curr_month):
    months = ['January', 'February', 'March', 'April', 'May', 'June', 
              'July', 'August', 'September', 'October', 'November', 'December']
    try:
        curr_idx = months.index(curr_month)
        prev_idx = (curr_idx - 1) if curr_idx > 0 else 11
        return months[prev_idx]
    except:
        return None

# --- STYLING ELEGANT FUNCTION (COLOR CODED) ---
def style_elegant(df_to_style, col_prev, col_total):
    """
    Fungsi ini melakukan 3 hal:
    1. Mengganti angka 0 dengan string kosong (Clean look)
    2. Rata tengah (Python level enforcement)
    3. Mewarnai Angka Total: 
       - Merah jika NAIK (Total > Prev)
       - Hijau jika TURUN (Total < Prev)
    """
    # 1. Zero Suppression (Logic di luar Styler)
    # Kita tidak ubah data asli jadi string di awal agar bisa dikomparasi, 
    # tapi kita format tampilan nanti.
    
    def highlight_trend(row):
        # Default style (no color)
        styles = [''] * len(row)
        
        # Cek apakah kolom ada
        if col_prev not in row.index or col_total not in row.index:
            return styles
            
        prev_val = row[col_prev]
        curr_val = row[col_total]
        
        # Pastikan numeric
        try:
            p = float(prev_val)
            c = float(curr_val)
        except:
            return styles

        # Logic Warna (Problem Dashboard: Naik = Merah, Turun = Hijau)
        # Cari index kolom total
        idx_total = row.index.get_loc(col_total)
        
        if c > p:
            styles[idx_total] = 'color: #FF4B4B; font-weight: bold;' # Merah Terang
        elif c < p:
            styles[idx_total] = 'color: #00FF00; font-weight: bold;' # Hijau Neon
            
        return styles

    # Apply Style
    styler = df_to_style.style.apply(highlight_trend, axis=1)
    
    # Format Properties Global
    styler = styler.set_properties(**{
        'text-align': 'center', 
        'vertical-align': 'middle',
        'font-size': '11px'
    })
    
    # Format Header khusus
    styler = styler.set_table_styles([
        {'selector': 'th', 'props': [('text-align', 'center'), ('background-color', '#262730'), ('color', 'white')]},
        {'selector': 'td', 'props': [('text-align', 'center')]}
    ])
    
    # Format Angka (0 jadi blank, ribuan pakai koma)
    styler = styler.format(lambda x: "{:,.0f}".format(x) if (isinstance(x, (int, float)) and x != 0) else "")
    
    return styler

# --- 3. LOGIKA MATRIX ---
def build_executive_summary(df_curr, df_prev, is_complain_mode, prev_month_short, curr_month_short):
    weeks = ['W1', 'W2', 'W3', 'W4']
    row_ticket = {}
    row_tid = {}
    total_ticket_curr = 0
    total_tid_set_curr = set()

    for w in weeks:
        df_week = df_curr[df_curr['WEEK'] == w] if 'WEEK' in df_curr.columns else pd.DataFrame()
        val = df_week['JUMLAH_COMPLAIN'].sum() if is_complain_mode and not df_week.empty else len(df_week) if not df_week.empty else 0
        row_ticket[w] = val
        total_ticket_curr += val
        
        tids = df_curr[df_curr['WEEK'] == w]['TID'].unique() if 'WEEK' in df_curr.columns and 'TID' in df_curr.columns else []
        row_tid[w] = len(tids)
        total_tid_set_curr.update(tids)

    val_prev = df_prev['JUMLAH_COMPLAIN'].sum() if is_complain_mode and not df_prev.empty else len(df_prev) if not df_prev.empty else 0
    tid_prev = len(df_prev['TID'].unique()) if 'TID' in df_prev.columns and not df_prev.empty else 0

    col_prev = f"{prev_month_short}"
    col_total = f"Σ {curr_month_short.upper()}"
    
    row_ticket[col_prev] = val_prev
    row_ticket[col_total] = total_ticket_curr
    row_tid[col_prev] = tid_prev
    row_tid[col_total] = len(total_tid_set_curr)

    matrix_df = pd.DataFrame([row_ticket, row_tid], index=['Global Ticket (Freq)', 'Global Unique TID'])
    cols_order = [col_prev, 'W1', 'W2', 'W3', 'W4', col_total]
    for c in cols_order:
        if c not in matrix_df.columns: matrix_df[c] = 0
    return matrix_df[cols_order], col_prev, col_total

# --- 4. UI DASHBOARD ---
df, df_slm = load_data()

if df.empty:
    st.warning("Data Master belum tersedia.")
else:
    st.markdown("### 🇮🇩 ATM Executive Dashboard")
    
    # FILTER
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        if 'KATEGORI' in df.columns:
            # 1. URUTAN KATEGORI DIKUNCI (Fixed Order)
            fixed_order = ['Elastic', 'Complain', 'DF Repeat', 'OUT Flm', 'Cash Out']
            available_cats = df['KATEGORI'].dropna().unique().tolist()
            # Filter hanya yang ada di data, tapi pertahankan urutan
            final_cats = [c for c in fixed_order if c in available_cats]
            # Tambahkan sisa kategori lain jika ada yang tidak masuk list
            remaining = [c for c in available_cats if c not in final_cats]
            final_cats.extend(remaining)
            
            sel_cat = st.radio("Kategori:", final_cats, index=0, horizontal=True, label_visibility="collapsed")
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

    # DATA FILTERING
    df_cat = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_cat.columns:
        df_cat = df_cat[df_cat['KATEGORI'] == sel_cat]
        
    df_main = df_cat.copy()
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]
        
    prev_mon_full = get_prev_month_full(sel_mon)
    df_prev = pd.DataFrame()
    if prev_mon_full and 'BULAN' in df_cat.columns:
        df_prev = df_cat[df_cat['BULAN'] == prev_mon_full]

    curr_mon_short = get_short_month_name(sel_mon)
    prev_mon_short = get_short_month_name(prev_mon_full) if prev_mon_full else "Prev"
    
    # Define Header Names for styling purpose
    col_prev_head = prev_mon_short
    col_total_head = f"Σ {curr_mon_short.upper()}"
    
    is_complain_mode = 'Complain' in sel_cat
    
    st.markdown("---") 

    # 1. GRAFIK TREN
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
                xaxis_title=None, yaxis_title=None, height=180, 
                margin=dict(l=10, r=10, t=10, b=10), xaxis=dict(tickangle=0, type='category')
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Data harian kosong.")

    # 2. DETAIL & SLM LOG
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown(f"**🌏 {sel_cat} Overview**")
        matrix_result, c_p, c_t = build_executive_summary(df_main, df_prev, is_complain_mode, prev_mon_short, curr_mon_short)
        
        # APPLY ELEGANT STYLING (WARNA & ZERO BLANK)
        st.dataframe(style_elegant(matrix_result, c_p, c_t), use_container_width=True)
        
        with st.expander(f"📂 Rincian Cabang"):
            if 'CABANG' in df_main.columns and 'WEEK' in df_main.columns:
                try:
                    val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                    agg_func = 'sum' if is_complain_mode else 'count'
                    
                    pivot_curr = df_main.pivot_table(index='CABANG', columns='WEEK', values=val_col, aggfunc=agg_func, fill_value=0)
                    desired_cols = ['W1', 'W2', 'W3', 'W4']
                    for c in desired_cols:
                        if c not in pivot_curr.columns: pivot_curr[c] = 0
                    pivot_curr = pivot_curr[desired_cols]
                    
                    prev_grp = df_prev.groupby('CABANG')[val_col].agg(agg_func).reset_index(name=col_prev_head) if not df_prev.empty else pd.DataFrame(columns=['CABANG', col_prev_head])
                    prev_grp = prev_grp.set_index('CABANG')
                    
                    final_cabang = pivot_curr.join(prev_grp, how='left').fillna(0)
                    final_cabang[col_total_head] = final_cabang[['W1', 'W2', 'W3', 'W4']].sum(axis=1)
                    
                    final_cols = [col_prev_head] + desired_cols + [col_total_head]
                    final_cabang = final_cabang[final_cols].sort_values(col_total_head, ascending=False)
                    
                    # APPLY ELEGANT STYLING
                    st.dataframe(style_elegant(final_cabang, col_prev_head, col_total_head), use_container_width=True)
                except Exception as e:
                    st.error(f"Error pivot: {e}")

    with col_right:
        c_head1, c_head2 = st.columns([2, 1])
        with c_head1:
             st.markdown(f"**🔥 Top 5 Problem Unit (Click to Expand)**")
        with c_head2:
             sort_options = [col_total_head, 'W1', 'W2', 'W3', 'W4']
             sort_by = st.selectbox("Urutkan:", sort_options, index=0, label_visibility="collapsed")
        
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
                
                prev_grp_top5 = df_prev.groupby(['TID', 'LOKASI'])[val_col].agg(agg_func).reset_index(name=col_prev_head) if not df_prev.empty else pd.DataFrame(columns=['TID', 'LOKASI', col_prev_head])
                prev_grp_top5 = prev_grp_top5.set_index(['TID', 'LOKASI'])
                
                final_top5 = pivot_top5.join(prev_grp_top5, how='left').fillna(0)
                final_top5[col_total_head] = final_top5[['W1', 'W2', 'W3', 'W4']].sum(axis=1)
                final_cols_top = [col_prev_head] + desired_cols + [col_total_head]
                final_top5 = final_top5[final_cols_top]
                
                # Sort Top 5
                top5_final = final_top5.sort_values(sort_by, ascending=False).head(5)
                
                if top5_final.empty:
                    st.info("Tidak ada data unit problem.")
                else:
                    for idx, row in top5_final.iterrows():
                        tid_val = idx[0]
                        lokasi_val = idx[1]
                        total_val = int(row[col_total_head])
                        curr_mon_code = curr_mon_short.upper()
                        
                        # Label Expander dengan Logic Warna di Icon (Visual Hack dengan Emoji)
                        # Kita tidak bisa warnai teks expander secara dinamis mudah, jadi pakai emoji
                        prev_val_row = row[col_prev_head]
                        trend_emoji = "🔴" if total_val > prev_val_row else "🟢" if total_val < prev_val_row else "⚪"
                        
                        label = f"{trend_emoji} TID: {tid_val} | {total_val}x ({curr_mon_code}) | {lokasi_val}"
                        
                        with st.expander(label):
                            cols_detail = st.columns(5)
                            cols_detail[0].caption(f"{col_prev_head}: {int(prev_val_row)}")
                            cols_detail[1].caption(f"W1: {int(row['W1'])}")
                            cols_detail[2].caption(f"W2: {int(row['W2'])}")
                            cols_detail[3].caption(f"W3: {int(row['W3'])}")
                            cols_detail[4].caption(f"W4: {int(row['W4'])}")
                            
                            st.divider()
                            
                            if not df_slm.empty and 'TID' in df_slm.columns:
                                slm_hist = df_slm[df_slm['TID'] == str(tid_val)]
                                if not slm_hist.empty:
                                    st.markdown("**Riwayat Kunjungan SLM:**")
                                    display_slm = slm_hist[['TGL_VISIT', 'ACTION']].reset_index(drop=True)
                                    # Pakai Styler juga untuk SLM biar rapi
                                    st.dataframe(display_slm.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
                                else:
                                    st.caption("Belum ada data kunjungan SLM di log.")
                            else:
                                st.caption("Data SLM Visit Log tidak tersedia/kosong.")

            except Exception as e:
                 st.error(f"Error Top 5: {e}")
