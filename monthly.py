import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.graph_objects as go
import base64
import io
import urllib.parse
from PIL import Image
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

# ==========================================
# 0. GLOBAL CONFIGURATION & CONSTANTS
# ==========================================
SPREADSHEET_ID = "1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340"

# ==========================================
# 1. KONFIGURASI HALAMAN UTAMA
# ==========================================
def get_image_base64(filepath):
    try:
        img = Image.open(filepath)
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode()
    except FileNotFoundError:
        return ""

bri_b64 = get_image_base64('logo_BRI.PNG')

st.set_page_config(
    page_title="BRI ATM Performance Review",
    page_icon=Image.open('logo_BRI.PNG') if bri_b64 else "🏦", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# 2. INJEKSI CUSTOM CSS BRUTAL (FULL SCREEN OPTIMIZED)
# ==========================================
st.markdown("""
    <style>
    .block-container { padding-top: 0.2rem !important; padding-bottom: 1rem !important; padding-left: 3rem !important; padding-right: 3rem !important; max-width: 98% !important; }
    [data-testid="stHeader"] { display: none; }
    .main {background-color: #F8F9FA; font-family: 'Segoe UI', Tahoma, sans-serif;}
    
    /* HEADER KORPORAT */
    .corporate-header { background-color: #003366; color: white; padding: 5px 16px; height: 55px; border-radius: 4px; display: flex; align-items: center; margin-bottom: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; position: relative; }
    .corporate-header h2 { margin: 0; color: white !important; font-size: 18px; font-weight: bold; display: flex; align-items: center; white-space: nowrap; z-index: 10; background-color: #003366; padding-right: 25px; }
    
    /* TICKER RUNNING TEXT */
    .ticker-wrapper { flex-grow: 1; overflow: hidden; white-space: nowrap; margin-left: 10px; display: flex; align-items: center; z-index: 1; -webkit-mask-image: linear-gradient(to right, transparent 0%, black 5%, black 95%, transparent 100%); }
    .ticker-text { display: inline-block; font-size: 13px; font-weight: 600; color: #e2e8f0; animation: slideTicker 60s linear infinite; }
    @keyframes slideTicker { 0% { transform: translateX(0); } 100% { transform: translateX(-100%); } }
    
    .blink-dot { display: inline-block; width: 9px; height: 9px; background-color: #10b981; border-radius: 50%; margin-right: 10px; box-shadow: 0 0 6px #10b981; animation: blink 1s infinite alternate; vertical-align: middle; }
    @keyframes blink { 0% { opacity: 1; } 100% { opacity: 0.4; } }

    /* PANGKAS SPASI ANTARA RADAR DAN TABS */
    div[data-testid="stExpander"] { margin-bottom: -15px !important; border: 1px solid #ddd !important; }
    div[data-testid="stVerticalBlock"] > div { margin-top: -8px !important; }

    /* TABLE STYLE */
    .custom-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; background-color: white; border: 1px solid #ddd; }
    .custom-table th { background-color: #003366; color: white; padding: 7px; font-weight: bold; border: 1px solid #ddd; font-size: 12px; text-align: center; }
    .custom-table td { text-align: center; padding: 7px 8px; border: 1px solid #ddd; color: #333; font-size: 12px; }
    
    /* KPI CARDS */
    .kpi-card { background-color: white; border-radius: 6px; padding: 20px 10px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.08); border-top: 4px solid #003366; }
    .kpi-header { color: #64748b; font-size: 11px; font-weight: 800; text-transform: uppercase; margin-bottom: 8px; }
    .kpi-value { color: #0f172a; font-size: 32px; font-weight: 800; font-family: 'Arial Black', sans-serif; }

    /* TABS STYLE */
    .stTabs [data-baseweb="tab-list"] { gap: 5px; }
    .stTabs [data-baseweb="tab"] { height: 38px; background-color: #e9ecef; border-radius: 4px 4px 0 0; padding: 5px 20px; color: #003366; font-weight: bold; }
    .stTabs [aria-selected="true"] { background-color: #003366; color: white !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. FUNGSI LOAD DATA (SEMUA SHEET)
# ==========================================
@st.cache_resource
def get_gspread_client():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        return gspread.authorize(creds)
    except: return None

@st.cache_data(ttl=3600)
def load_all_data():
    csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=AIMS_Master"
    df = pd.read_csv(csv_url)
    df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], errors='coerce')
    df['WAKTU INSERT'] = pd.to_datetime(df['WAKTU INSERT'], errors='coerce')
    df['Periode'] = df['TANGGAL'].dt.to_period('M')
    
    # Load Asset Map
    csv_asset = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=Jml_Kelolaan"
    df_asset = pd.read_csv(csv_asset).iloc[:, :2]
    df_asset.columns = ['CABANG', 'TOTAL ATM REAL']
    return pd.merge(df, df_asset, on='CABANG', how='left')

@st.cache_data(ttl=3600)
def load_mri_data():
    sheet_name = urllib.parse.quote("Problem MRI 2025/26 Harian")
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    df = pd.read_csv(url)
    df.columns = ['TANGGAL','TID','3','4','5','6','7','SLM','CPC','10','11','12','KATEGORI','LOKASI']
    df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], dayfirst=True, errors='coerce')
    df['Periode'] = df['TANGGAL'].dt.to_period('M')
    df['Week_Group'] = pd.cut(df['TANGGAL'].dt.day, bins=[0,7,14,21,31], labels=['W1','W2','W3','W4'])
    return df

@st.cache_data(ttl=3600)
def load_slm_visit():
    client = get_gspread_client()
    if not client: return pd.DataFrame()
    data = client.open_by_key(SPREADSHEET_ID).worksheet("SLM Visit Log").get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df['TGL_VISIT_DT'] = pd.to_datetime(df['Tgl Visit'], errors='coerce')
    return df

# ==========================================
# 4. DASHBOARD BUILDER (LOGIC RE-USABLE)
# ==========================================
def calculate_mom(curr, prev):
    if prev == 0: return (100.0 if curr > 0 else 0.0), "#ef4444", "▲"
    mom = ((curr - prev) / prev) * 100
    return mom, ("#ef4444" if mom > 0 else "#10b981"), ("▲" if mom > 0 else "▼")

def build_category_dashboard(kat, df_all, m3, m2, m1):
    df_kat = df_all[df_all['KATEGORI'].str.contains(kat, na=False, case=False)]
    m3_val = len(df_kat[df_kat['Periode'] == m3])
    m2_val = len(df_kat[df_kat['Periode'] == m2])
    m1_val = len(df_kat[df_kat['Periode'] == m1])
    
    mom, color, arrow = calculate_mom(m3_val, m2_val)
    
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""<div class="table-title">Overview {kat}</div>
        <table class="custom-table">
            <tr><th>Period</th><th>{m1}</th><th>{m2}</th><th>{m3}</th><th>MoM</th></tr>
            <tr><td>Frequency</td><td>{m1_val}</td><td>{m2_val}</td><td>{m3_val}</td><td style="color:{color}; font-weight:bold;">{abs(mom):.1f}% {arrow}</td></tr>
        </table>""", unsafe_allow_html=True)
    
    with c2:
        top_br = df_kat[df_kat['Periode'] == m3]['CABANG'].value_counts().head(5)
        st.markdown(f"<div class='table-title'>Top 5 Cabang {kat}</div>", unsafe_allow_html=True)
        fig = go.Figure(go.Bar(x=top_br.index, y=top_br.values, marker_color='#003366'))
        fig.update_layout(height=200, margin=dict(l=0,r=0,t=20,b=0))
        st.plotly_chart(fig, use_container_width=True)

# ==========================================
# 5. MAIN EXECUTION
# ==========================================
df_master = load_all_data()

if not df_master.empty:
    # --- HEADER ---
    st.markdown(f"""<div class="corporate-header">
        <h2><img src="https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png" style="height:28px; filter:brightness(0) invert(1); margin-right:15px;"> MONITORING ATM COMMAND CENTER</h2>
        <div class="ticker-wrapper"><div class="ticker-text"><span class="blink-dot"></span> STATUS: SECURE & OPTIMAL &nbsp; | &nbsp; LAST REFRESH: {datetime.now().strftime('%H:%M:%S')} &nbsp; | &nbsp; PT KELOLA JASA ARTHA</div></div>
    </div>""", unsafe_allow_html=True)

    # --- RADAR & PERIOD SELECTOR (ONE ROW) ---
    col_rad, col_sel = st.columns([8.5, 1.5])
    
    with col_sel:
        periods = sorted(df_master['Periode'].unique())
        sel_p = st.selectbox("Periode", [p.strftime('%B %Y') for p in periods], index=len(periods)-1, label_visibility="collapsed")
        m3 = periods[[p.strftime('%B %Y') for p in periods].index(sel_p)]
        m2 = periods[max(0, periods.index(m3)-1)]
        m1 = periods[max(0, periods.index(m3)-2)]

    with col_rad:
        with st.expander("📡 RADAR: LATEST INCIDENTS", expanded=False):
            r1, r2, r3, r4 = st.columns(4)
            df_curr = df_master[df_master['Periode'] == m3]
            
            def get_radar(df_sub, lbl, is_h=False):
                col_t = 'WAKTU INSERT' if is_h else 'TANGGAL'
                latest = df_sub.dropna(subset=[col_t]).sort_values(col_t, ascending=False)
                if latest.empty: return "<b>-</b>"
                max_t = latest.iloc[0][col_t]
                cnt = len(latest[latest[col_t] == max_t])
                return f"<div style='border-left:3px solid #F37021; padding-left:10px;'><b style='font-size:12px;'>{lbl}</b><br><small style='color:red;'>{max_t.strftime('%d/%m %H:%M' if is_h else '%d/%m/%y')} ({cnt} Unit)</small></div>"

            r1.markdown(get_radar(df_curr[df_curr['KATEGORI']=='Elastic'], "ELASTIC"), unsafe_allow_html=True)
            r2.markdown(get_radar(df_curr[df_curr['KATEGORI']=='Complain'], "COMPLAIN"), unsafe_allow_html=True)
            r3.markdown(get_radar(df_curr[df_curr['KATEGORI'].str.contains('DF Repeat', na=False)], "DF REPEAT", True), unsafe_allow_html=True)
            r4.markdown(get_radar(df_curr[df_curr['KATEGORI'].str.contains('OUT Flm', na=False)], "OUT FLM", True), unsafe_allow_html=True)

    # --- TABS ---
    t_home, t_mri, t_el, t_cp, t_df, t_out, t_log = st.tabs(["Home", "⭐ MRI", "Elastic", "Complain", "DF Repeat", "OUT Flm", "Logistic"])

    with t_home:
        c1, c2, c3, c4 = st.columns(4)
        for c, k, n in zip([c1,c2,c3,c4],['Elastic','Complain','DF Repeat','OUT Flm'],['ELASTIC','COMPLAIN','DF REPEAT','OUT FLM']):
            v3 = len(df_curr[df_curr['KATEGORI'].str.contains(k, na=False)])
            v2 = len(df_master[(df_master['Periode']==m2) & (df_master['KATEGORI'].str.contains(k, na=False))])
            mom, col, arr = calculate_mom(v3, v2)
            c.markdown(f'<div class="kpi-card"><div class="kpi-header">{n}</div><div class="kpi-value">{v3}</div><div style="color:{col}; font-size:12px; font-weight:700;">{arr} {abs(mom):.1f}%</div></div>', unsafe_allow_html=True)
        
        st.markdown(f"""<div class="banner-container"><div style="border-left:5px solid #F37021; padding-left:20px;"><div class="banner-title">MONTHLY PERFORMANCE</div><div class="banner-subtitle">REPORTER: COMMAND CENTER &nbsp; | &nbsp; {sel_p}</div></div><img src="https://upload.wikimedia.org/wikipedia/commons/9/97/Logo_BRI.png" style="height:60px; filter:brightness(0) invert(1);"></div>""", unsafe_allow_html=True)

    with t_mri:
        df_mri = load_mri_data()
        df_slm = load_slm_visit()
        curr_mri = df_mri[df_mri['Periode'] == m3]
        
        st.markdown("<div class='table-title'>MRI Problem Monitoring (MTD)</div>", unsafe_allow_html=True)
        top_mri = curr_mri['TID'].value_counts().head(5)
        
        html_mri = "<table class='custom-table'><tr><th>TID</th><th>Lokasi</th><th>Branch</th><th>Freq</th><th>Latest SLM Action</th></tr>"
        for tid, freq in top_mri.items():
            row = curr_mri[curr_mri['TID'] == tid].iloc[0]
            slm_act = df_slm[df_slm['TID']==tid].sort_values('TGL_VISIT_DT', ascending=False)
            act_text = slm_act.iloc[0]['Action'] if not slm_act.empty else "No Visit Data"
            html_mri += f"<tr><td><b>{tid}</b></td><td>{row['LOKASI'][:20]}..</td><td>{row['CPC']}</td><td style='color:red; font-weight:bold;'>{freq}</td><td style='text-align:left;'>{act_text[:50]}..</td></tr>"
        st.markdown(html_mri + "</table>", unsafe_allow_html=True)

    with t_el: build_category_dashboard("Elastic", df_master, m3, m2, m1)
    with t_cp: build_category_dashboard("Complain", df_master, m3, m2, m1)
    with t_df: build_category_dashboard("DF Repeat", df_master, m3, m2, m1)
    with t_out: build_category_dashboard("OUT Flm", df_master, m3, m2, m1)

    with t_log:
        st.markdown("<div class='table-title'>Resource & Logistics</div>", unsafe_allow_html=True)
        st.warning("Silakan hubungkan sheet Stock_Cassette & Stock_Thermal untuk sinkronisasi data logistik.")

else:
    st.error("Bah! Data gagal ditarik. Cek Spreadsheet ID kau!")
