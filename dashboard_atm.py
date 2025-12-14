import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys
import re
import streamlit.components.v1 as components
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(layout='wide', page_title="ATM Performance Monitoring", initial_sidebar_state="collapsed")

# Styling CSS (The "Monitoring Hub" Theme - V94 Complete)
st.markdown("""
<style>
    /* 1. LAYOUTING */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
    }
    
    /* TYPOGRAPHY */
    h1 { 
        font-size: 2.2rem !important; 
        font-weight: 800 !important; 
        margin-bottom: 0rem !important; 
        margin-top: 0rem !important;
        color: #FFFFFF !important; 
        letter-spacing: 1px;
    }
    h2 { font-size: 1.1rem !important; margin-bottom: 0px !important;}
    
    html, body, [class*="st-emotion-"] { 
        font-size: 10px; 
    }

    #MainMenu, footer, header {visibility: hidden;}
    .st-emotion-cache-1j8u2d7 {visibility: hidden;} 
    
    /* --- 2. TABLE ALIGNMENT --- */
    [data-testid="stDataFrame"] td {
        text-align: right !important;
        padding-top: 2px !important;
        padding-bottom: 2px !important;
    }
    [data-testid="stDataFrame"] thead tr th:first-child,
    [data-testid="stDataFrame"] tbody th {
        text-align: left !important;
    }

    /* --- 3. TOMBOL KATEGORI --- */
    div[role="radiogroup"] > label {
        font-size: 12px !important;
        font-weight: bold !important;
        background-color: #1E1E1E;
        padding: 6px 12px; 
        border-radius: 8px;
        border: 1px solid #444; 
        margin-right: 4px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }
    div[role="radiogroup"] > label:hover {
        border-color: #888; 
        transform: translateY(-1px);
    }
    
    /* --- 4. ANIMASI KEDAP KEDIP (BLINKING) --- */
    @keyframes blink-animation {
      0% { opacity: 1; box-shadow: 0 0 5px #ff0000; transform: scale(1); }
      50% { opacity: 0.8; box-shadow: 0 0 20px #ff0000; transform: scale(1.02); }
      100% { opacity: 1; box-shadow: 0 0 5px #ff0000; transform: scale(1); }
    }
    .blinking-alert {
        animation: blink-animation 1.5s infinite;
        background-color: #FF4B4B;
        color: white;
        padding: 10px;
        border-radius: 8px;
        text-align: center;
        font-weight: bold;
        font-size: 14px;
        margin-bottom: 10px;
        border: 2px solid #ff9999;
    }

    /* --- 5. DESAIN CARD --- */
    [data-testid="stDataFrame"], .stPlotlyChart {
        border: 1px solid #333;
        border-radius: 8px; 
        padding: 5px; 
        background-color: #1a1a1a; 
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }
    .streamlit-expanderHeader {
        font-size: 11px !important;
        font-weight: bold !important;
        background-color: #262730 !important;
        color: white !important;
        border-radius: 5px;
        border: 1px solid #444;
        padding-top: 2px !important;
        padding-bottom: 2px !important;
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
        ws = sh.worksheet(SHEET_MAIN)
        all_values = ws.get_all_values()
        
        if not all_values: return pd.DataFrame(), pd.DataFrame()

        headers = all_values[0]
        rows = all_values[1:]
        df = pd.DataFrame(rows, columns=headers)
        
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
        if 'CABANG' in df.columns:
            df['CABANG'] = df['CABANG'].astype(str)

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

# --- STYLING ELEGANT FUNCTION ---
def style_elegant(df_to_style, col_prev, col_total):
    def highlight_trend(row):
        styles = [''] * len(row)
        if col_prev not in row.index or col_total not in row.index:
            return styles
        prev_val = row[col_prev]
        curr_val = row[col_total]
        try:
            p = float(prev_val)
            c = float(curr_val)
        except:
            return styles
        idx_total = row.index.get_loc(col_total)
        # LOGIKA ADAPTIF: Merah jika > Prev, Hijau jika < Prev
        if c > p:
            styles[idx_total] = 'color: #FF4B4B; font-weight: bold;'
        elif c < p:
            styles[idx_total] = 'color: #00FF00; font-weight: bold;'
        return styles

    styler = df_to_style.style.apply(highlight_trend, axis=1)
    styler = styler.set_properties(**{
        'text-align': 'right', 
        'vertical-align': 'middle', 
        'font-size': '10px'
    })
    styler = styler.set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#262730'), ('color', 'white'), ('font-size', '10px')]},
    ])
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
    # HEADER
    c_title, c_clock = st.columns([3, 1]) 
    with c_title:
        st.markdown("<h1>ATM Performance Monitoring</h1>", unsafe_allow_html=True)
    with c_clock:
        components.html(
            """
            <!DOCTYPE html>
            <html>
            <head>
            <style>
                body { 
                    margin: 0; padding: 0; background-color: transparent; color: #BBBBBB; 
                    font-family: 'Source Sans Pro', sans-serif; text-align: right; 
                    font-size: 13px; font-weight: 600; display: flex; 
                    justify-content: flex-end; align-items: center; height: 100vh;
                }
            </style>
            </head>
            <body>
                <div id="clock"></div>
                <script>
                function updateTime() {
                    var now = new Date();
                    var options = { timeZone: 'Asia/Jakarta', weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false };
                    var formatter = new Intl.DateTimeFormat('id-ID', options);
                    var timeString = formatter.format(now);
                    timeString = timeString.replace("pukul ", "") + " WIB";
                    document.getElementById('clock').innerHTML = timeString;
                }
                setInterval(updateTime, 1000);
                updateTime();
                </script>
            </body>
            </html>
            """, height=40
        )
    
    # FILTER
    all_months = df['BULAN'].unique().tolist() if 'BULAN' in df.columns else []
    default_ix = len(all_months)-1 if all_months else 0
    
    col_f1, col_f2 = st.columns([2, 1])
    with col_f2:
        sel_mon = st.selectbox("Bulan:", all_months, index=default_ix, label_visibility="collapsed")

    prev_mon_full_calc = get_prev_month_full(sel_mon)
    df_mon_curr = df[df['BULAN'] == sel_mon] if 'BULAN' in df.columns else df
    df_mon_prev = df[df['BULAN'] == prev_mon_full_calc] if (prev_mon_full_calc and 'BULAN' in df.columns) else pd.DataFrame()
    
    fixed_order = ['Elastic', 'Complain', 'DF Repeat', 'OUT Flm', 'Cash Out']
    available_cats = df['KATEGORI'].dropna().unique().tolist() if 'KATEGORI' in df.columns else []
    final_cats_raw = [c for c in fixed_order if c in available_cats]
    remaining = [c for c in available_cats if c not in final_cats_raw]
    final_cats_raw.extend(remaining)
    
    cat_labels = []
    cat_map = {} 
    dynamic_css = [] 
    
    for idx, c in enumerate(final_cats_raw):
        df_c_curr = df_mon_curr[df_mon_curr['KATEGORI'] == c]
        if 'Complain' in c: 
            count_curr = df_c_curr['JUMLAH_COMPLAIN'].sum() if 'JUMLAH_COMPLAIN' in df_c_curr.columns else 0
        else: 
            count_curr = len(df_c_curr)
            
        count_prev = 0
        has_prev = False
        if not df_mon_prev.empty:
            df_c_prev = df_mon_prev[df_mon_prev['KATEGORI'] == c]
            if 'Complain' in c:
                count_prev = df_c_prev['JUMLAH_COMPLAIN'].sum() if 'JUMLAH_COMPLAIN' in df_c_prev.columns else 0
            else:
                count_prev = len(df_c_prev)
            has_prev = True
            
        trend_str = ""
        text_color = "#E0E0E0" 
        
        if has_prev:
            if count_prev > 0:
                pct_change = ((count_curr - count_prev) / count_prev) * 100
                if pct_change > 0:
                    trend_str = f"| ▲ +{int(pct_change)}%" 
                    text_color = "#FF4B4B" 
                elif pct_change < 0:
                    trend_str = f"| ▼ {int(pct_change)}%"
                    text_color = "#00FF00" 
                else:
                    trend_str = "| - 0%"
            elif count_curr > 0:
                trend_str = "| ▲ New"
                text_color = "#FF4B4B"
        
        label = f"{c} ({count_curr} {trend_str})"
        cat_labels.append(label)
        cat_map[label] = c
        
        rule = f"""
            div[role="radiogroup"] > label:nth-of-type({idx+1}) {{ color: {text_color} !important; }}
            div[role="radiogroup"] > label:nth-of-type({idx+1}) p {{ color: {text_color} !important; }}
        """
        dynamic_css.append(rule)

    st.markdown(f"<style>{''.join(dynamic_css)}</style>", unsafe_allow_html=True)

    with col_f1:
        sel_cat_label = st.radio("Kategori:", cat_labels, index=0, horizontal=True, label_visibility="collapsed")
        sel_cat = cat_map[sel_cat_label]

    # DATA PROCESSING
    df_cat = df.copy()
    if sel_cat != "Semua" and 'KATEGORI' in df_cat.columns:
        df_cat = df_cat[df_cat['KATEGORI'] == sel_cat]
        
    df_main = df_cat.copy()
    if sel_mon != "Semua" and 'BULAN' in df_main.columns:
        df_main = df_main[df_main['BULAN'] == sel_mon]
        
    df_prev = pd.DataFrame()
    if prev_mon_full_calc and 'BULAN' in df_cat.columns:
        df_prev = df_cat[df_cat['BULAN'] == prev_mon_full_calc]

    curr_mon_short = get_short_month_name(sel_mon)
    prev_mon_short = get_short_month_name(prev_mon_full_calc) if prev_mon_full_calc else "Prev"
    col_prev_head = prev_mon_short
    col_total_head = f"Σ {curr_mon_short.upper()}"
    is_complain_mode = 'Complain' in sel_cat
    
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
            avg_val = daily[y_val].mean()
            
            fig = px.area(daily, x='TANGGAL_STR', y=y_val, markers=True, text=y_val, template="plotly_dark")
            fig.update_traces(
                line_color='#FF4B4B', line_width=3, line_shape='spline', textposition="top center",
                fill='tozeroy', fillcolor='rgba(255, 75, 75, 0.1)'
            )
            fig.add_hline(
                y=avg_val, line_dash="dash", line_color="rgba(255, 255, 255, 0.5)", 
                annotation_text=f"AVG: {avg_val:.1f}", annotation_position="bottom right"
            )
            fig.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                xaxis_title=None, yaxis_title=None, height=170, 
                margin=dict(l=10, r=10, t=10, b=10), 
                xaxis=dict(tickangle=0, type='category', showgrid=False),
                yaxis=dict(showgrid=True, gridcolor='#333')
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Data harian kosong.")

    # 2. DETAIL
    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown(f"**🌏 {sel_cat} Overview**")
        matrix_result, c_p, c_t = build_executive_summary(df_main, df_prev, is_complain_mode, prev_mon_short, curr_mon_short)
        st.dataframe(style_elegant(matrix_result, c_p, c_t), use_container_width=True)
        
        with st.expander(f"📂 Rincian Cabang (Total: {len(df_main['CABANG'].unique())} Unit)", expanded=True):
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
                    st.dataframe(style_elegant(final_cabang, col_prev_head, col_total_head), use_container_width=True)
                except Exception as e:
                    st.error(f"Error pivot: {e}")

    with col_right:
        c_head1, c_head2 = st.columns([2, 1])
        with c_head1:
             st.markdown(f"**🔥 Top 10 Problem Unit (Diagnosis)**")
        with c_head2:
             sort_options = [col_total_head, 'W1', 'W2', 'W3', 'W4']
             sort_by = st.selectbox("Urutkan:", sort_options, index=0, label_visibility="collapsed")
        
        # --- UPDATE V94: MENAMBAHKAN CABANG KE GROUPING ---
        if 'TID' in df_main.columns and 'LOKASI' in df_main.columns and 'CABANG' in df_main.columns and 'WEEK' in df_main.columns:
            try:
                val_col = 'JUMLAH_COMPLAIN' if is_complain_mode else 'TID'
                agg_func = 'sum' if is_complain_mode else 'count'
                
                # GROUP BY TID + LOKASI + CABANG
                grouped_df = df_main.groupby(['TID', 'LOKASI', 'CABANG', 'WEEK'])[val_col].agg(agg_func).reset_index(name='VAL')
                pivot_top5 = grouped_df.pivot_table(index=['TID', 'LOKASI', 'CABANG'], columns='WEEK', values='VAL', aggfunc='sum', fill_value=0)
                
                desired_cols = ['W1', 'W2', 'W3', 'W4']
                for c in desired_cols:
                    if c not in pivot_top5.columns: pivot_top5[c] = 0
                pivot_top5 = pivot_top5[desired_cols]
                
                prev_grp_top5 = df_prev.groupby(['TID', 'LOKASI', 'CABANG'])[val_col].agg(agg_func).reset_index(name=col_prev_head) if not df_prev.empty else pd.DataFrame(columns=['TID', 'LOKASI', 'CABANG', col_prev_head])
                prev_grp_top5 = prev_grp_top5.set_index(['TID', 'LOKASI', 'CABANG'])
                
                final_top5 = pivot_top5.join(prev_grp_top5, how='left').fillna(0)
                final_top5[col_total_head] = final_top5[['W1', 'W2', 'W3', 'W4']].sum(axis=1)
                final_cols_top = [col_prev_head] + desired_cols + [col_total_head]
                final_top5 = final_top5[final_cols_top]
                
                top5_final = final_top5.sort_values(sort_by, ascending=False).head(10)
                
                if sort_by in ['W1', 'W2', 'W3', 'W4']:
                    top5_final = top5_final[top5_final[sort_by] > 0]
                
                if top5_final.empty:
                    st.info(f"Belum ada unit problem yang tercatat di {sort_by}.")
                else:
                    today_dt = pd.Timestamp.now()
                    is_realtime_cat = sel_cat in ['DF Repeat', 'OUT Flm', 'Cash Out']
                    
                    for idx, row in top5_final.iterrows():
                        # Unpack 3 index
                        tid_val = idx[0]
                        lokasi_val = idx[1]
                        cabang_val = idx[2] # CABANG
                        
                        total_val = int(row[col_total_head])
                        curr_mon_code = curr_mon_short.upper()
                        
                        is_sick = False
                        time_str = "⏱️ ?"
                        
                        try:
                            mask_tid = df_main['TID'] == str(tid_val)
                            if mask_tid.any():
                                last_date = df_main[mask_tid]['TANGGAL'].max()
                                if pd.notna(last_date):
                                    days_diff = (today_dt - last_date).days
                                    
                                    if is_realtime_cat:
                                        if days_diff == 0: is_sick = True
                                    else: 
                                        if days_diff <= 1: is_sick = True
                                    
                                    if days_diff == 0:
                                        time_str = "⏱️ Hari ini"
                                    elif days_diff == 1:
                                        time_str = "⏱️ Kemarin"
                                    else:
                                        time_str = f"⏱️ {days_diff} hari lalu"
                                else:
                                    time_str = "⏱️ -"
                            else:
                                time_str = "⏱️ -"
                        except:
                            time_str = "⏱️ ?"
                        
                        prev_val_row = row[col_prev_head]
                        trend_emoji = "🔴" if total_val > prev_val_row else "🟢" if total_val < prev_val_row else "⚪"
                        
                        alert_prefix = "🚨 " if is_sick else ""
                        # LABEL BARU: TAMBAHKAN CABANG
                        label = f"{alert_prefix}{trend_emoji} TID: {tid_val} | {total_val}x ({curr_mon_code}) | {time_str} | {lokasi_val} ({cabang_val})"
                        
                        with st.expander(label):
                            if is_sick:
                                st.markdown(f"""
                                <div class="blinking-alert">
                                    ⚡ UNIT INI SAKIT PARAH (DATA TERBARU) - PERLU FOLLOW UP SEGERA! ⚡
                                </div>
                                """, unsafe_allow_html=True)

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
                                    st.dataframe(display_slm.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
                                else:
                                    st.caption("Belum ada data kunjungan SLM di log.")
                            else:
                                st.caption("Data SLM Visit Log tidak tersedia/kosong.")

            except Exception as e:
                 st.error(f"Error Top 5: {e}")
