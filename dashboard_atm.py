import streamlit as st
import pandas as pd
import plotly.express as px
import gspread
import sys

# =========================================================
# 1. KONFIGURASI HALAMAN
# =========================================================
st.set_page_config(
    layout="wide",
    page_title="ATM Executive Dashboard",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
html, body, [class*="st-emotion-"] { font-size: 12px; }
h1, h2, h3, h4, h5, h6 { font-size: 1.1em !important; }
.block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
.dataframe { font-size: 9px !important; text-align: center; }
.dataframe td { text-align: center !important; }
.dataframe tbody th { text-align: left !important; }
th { background-color: #262730 !important; color: white !important; }
thead tr th:first-child { display:none }
tbody th { display:none }
.js-plotly-plot { margin-bottom: 0px !important; }
.stPlotlyChart { margin-bottom: 0px !important; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# 2. KONFIGURASI GOOGLE SHEET
# =========================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1pApEIA9BEYEojW4a6Fvwykkf-z-UqeQ8u2pmrqQc340/edit?gid=98670277#gid=98670277"
SHEET_NAME = "AIMS_Master"

# =========================================================
# 3. AUTHENTICATION GOOGLE SERVICE ACCOUNT (FIX FINAL)
# =========================================================
try:
    if "gcp_service_account" not in st.secrets:
        st.error("❌ gcp_service_account tidak ditemukan di secrets.")
        st.stop()

    creds = st.secrets["gcp_service_account"]

    creds_dict = {
        "type": creds["type"],
        "project_id": creds["project_id"],
        "private_key_id": creds["private_key_id"],
        "private_key": creds["private_key"].replace("\\n", "\n"),
        "client_email": creds["client_email"],
        "client_id": creds["client_id"],
        "auth_uri": creds["auth_uri"],
        "token_uri": creds["token_uri"],
        "auth_provider_x509_cert_url": creds["auth_provider_x509_cert_url"],
        "client_x509_cert_url": creds["client_x509_cert_url"],
    }

    gc = gspread.service_account_from_dict(creds_dict)

except Exception as e:
    st.error("🚨 AUTHENTICATION ERROR")
    st.exception(e)
    st.stop()

# =========================================================
# 4. LOAD DATA
# =========================================================
@st.cache_data(ttl=600)
def load_data():
    try:
        sh = gc.open_by_url(SHEET_ID)
        ws = sh.worksheet(SHEET_NAME)
        values = ws.get_all_values()

        if not values:
            return pd.DataFrame()

        headers = values[0]
        rows = values[1:]
        df = pd.DataFrame(rows, columns=headers)

        # --- CLEANING ---
        df = df.loc[:, df.columns != ""]
        df.columns = df.columns.str.strip().str.upper()
        df = df.loc[:, ~df.columns.duplicated()]

        if "TANGGAL" in df.columns:
            df["TANGGAL"] = pd.to_datetime(df["TANGGAL"], dayfirst=True, errors="coerce")

        if "JUMLAH_COMPLAIN" in df.columns:
            df["JUMLAH_COMPLAIN"] = (
                df["JUMLAH_COMPLAIN"]
                .astype(str)
                .str.replace("-", "0")
                .astype(float)
                .fillna(0)
                .astype(int)
            )
        else:
            df["JUMLAH_COMPLAIN"] = 0

        if "WEEK" not in df.columns and "BULAN_WEEK" in df.columns:
            df["WEEK"] = df["BULAN_WEEK"]

        if "BULAN" in df.columns:
            df["BULAN"] = df["BULAN"].astype(str).str.strip().str.upper()

        if "TID" in df.columns:
            df["TID"] = df["TID"].astype(str)

        if "LOKASI" in df.columns:
            df["LOKASI"] = df["LOKASI"].astype(str)

        return df

    except Exception as e:
        st.error("🚨 DATA LOADING ERROR")
        st.exception(e)
        return pd.DataFrame()

# =========================================================
# 5. LOGIKA MATRIX
# =========================================================
def build_executive_summary(df_curr, is_complain_mode):
    weeks = ["W1", "W2", "W3", "W4"]

    row_ticket = {}
    total_ticket = 0
    for w in weeks:
        df_w = df_curr[df_curr["WEEK"] == w]
        val = (
            df_w["JUMLAH_COMPLAIN"].sum()
            if is_complain_mode and not df_w.empty
            else len(df_w)
        )
        row_ticket[w] = val
        total_ticket += val

    row_ticket["TOTAL"] = total_ticket
    row_ticket["AVG/WEEK"] = round(total_ticket / 4, 1)

    row_tid = {}
    tid_set = set()
    for w in weeks:
        tids = df_curr[df_curr["WEEK"] == w]["TID"].unique()
        row_tid[w] = len(tids)
        tid_set.update(tids)

    row_tid["TOTAL"] = len(tid_set)
    row_tid["AVG/WEEK"] = round(len(tid_set) / 4, 1)

    df_matrix = pd.DataFrame(
        [row_ticket, row_tid],
        index=["Global Ticket (Freq)", "Global Unique TID"]
    )

    cols = ["W1", "W2", "W3", "W4", "TOTAL", "AVG/WEEK"]
    return df_matrix[cols]

# =========================================================
# 6. UI DASHBOARD
# =========================================================
df = load_data()

if df.empty:
    st.warning("Data belum tersedia.")
    st.stop()

st.markdown("### 🇮🇩 ATM Executive Dashboard")

col_f1, col_f2 = st.columns([2, 1])

with col_f1:
    cats = df["KATEGORI"].dropna().unique().tolist()
    sel_cat = st.radio("Pilih Kategori:", cats, horizontal=True)

with col_f2:
    months = df["BULAN"].unique().tolist()
    sel_mon = st.selectbox("Pilih Bulan:", months, index=len(months) - 1)

df_main = df[(df["KATEGORI"] == sel_cat) & (df["BULAN"] == sel_mon)]
is_complain_mode = "COMPLAIN" in sel_cat.upper()

st.markdown("---")

col_left, col_right = st.columns(2)

with col_left:
    st.subheader("🌏 Overview")
    matrix = build_executive_summary(df_main, is_complain_mode)
    st.dataframe(matrix, use_container_width=True)

with col_right:
    st.subheader("📈 Tren Harian")
    daily = (
        df_main.groupby("TANGGAL")["JUMLAH_COMPLAIN"].sum().reset_index()
        if is_complain_mode
        else df_main.groupby("TANGGAL").size().reset_index(name="TOTAL")
    )

    fig = px.line(
        daily,
        x="TANGGAL",
        y=daily.columns[-1],
        markers=True,
        template="plotly_dark"
    )
    st.plotly_chart(fig, use_container_width=True)



