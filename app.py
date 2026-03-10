import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. INJEKSI CSS GLOBAL (Memaksa Alignment Kiri)
st.markdown("""
    <style>
        /* Memaksa teks di dalam dataframe rata kiri */
        [data-testid="stDataFrame"] div[data-testid="stTable"] th { text-align: left !important; }
        [data-testid="stDataFrame"] div[data-testid="stTable"] td { text-align: left !important; }
        /* Menghilangkan margin default pada markdown agar lebih rapat ke kiri */
        .stMarkdown div { text-align: left !important; }
    </style>
""", unsafe_allow_html=True)

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        # (Logika pembersihan data internal tetap sama)
        return df_main, bln_raw, thn_raw
    except Exception as e:
        return pd.DataFrame(), "N/A", "N/A"

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names)
    
    # Memuat data utama (Asumsi fungsi load_all_data sudah didefinisikan lengkap)
    # df_main, df_history, bln_ref, thn_ref = load_all_data(...) 

    st.title(f"📊 Reliability Analysis")

    # 4. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    # filtered = ... (logika filter data)
    event = st.dataframe(df_main, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # 5. PART REMOVAL DETAIL (SEMUA RATA KIRI)
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = df_main.iloc[selected_idx]
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {row.get('PART NUMBER', 'N/A')}")
        
        # Penggunaan HTML Div untuk alignment absolut di sisi kiri
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            st.markdown(f"<div style='text-align: left;'><p style='margin-bottom: -5px; color: gray; font-size: 14px;'>Description</p><h2 style='margin-top: 0;'>{row.get('DESCRIPTION', 'N/A')}</h2></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div style='text-align: left;'><p style='margin-bottom: -5px; color: gray; font-size: 14px;'>Current Rate</p><h2 style='margin-top: 0;'>{row.get('RATE', 0):.2f}</h2></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div style='text-align: left;'><p style='margin-bottom: -5px; color: gray; font-size: 14px;'>Total Qty Rem</p><h2 style='margin-top: 0;'>{row.get('QTY REM', 0)} EA</h2></div>", unsafe_allow_html=True)

        # 6. TABLE HISTORY (Rata Kiri Terkunci)
        # if not df_history.empty:
        #    st.dataframe(hist_match, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Sistem Error: {e}")
