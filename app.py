import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman & UI Profesional
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

def clean_df(df):
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = [str(col).strip() for col in df.columns]
    return df

# 2. Fungsi Load Data Utama
@st.cache_data
def load_reliability_data(file_name):
    try:
        # Ambil kriteria periode dari sheet utama (A2 & A3)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_ref = str(df_crit.iloc[1, 0]).strip().upper()
        thn_ref = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load History (Sheet yang dikoreksi Bapak)
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist = clean_df(df_hist)
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
        return df_hist, bln_ref, thn_ref
    except Exception as e:
        st.error(f"Gagal memuat data dasar: {e}")
        return pd.DataFrame(), "N/A", "N/A"

# --- LOGIKA UTAMA ---
try:
    FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(FILE_PATH)
    
    # MUAT DATA DASAR & KRITERIA
    df_history, bulan_excel, tahun_excel = load_reliability_data(FILE_PATH)
    
    # LOGIKA NAVIGASI (SIDEBAR) - INI YANG KEMBALI DIMUNCULKAN
    st.sidebar.title("🚀 Navigation")
    all_sheets = xls.sheet_names
    sheet_pilihan = st.sidebar.selectbox("Select Report Sheet:", all_sheets)
    
    # LOGIKA PENGURANGAN 1 BULAN (Contoh: Dec -> Nov)
    months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                  'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    
    m_num = months_map.get(bulan_excel, 12)
    current_dt = datetime(int(tahun_excel) if tahun_excel.isdigit() else 2026, m_num, 1)
    prev_dt = current_dt - timedelta(days=1)
    target_m_num = prev_dt.month
    target_y = prev_dt.year
    target_m_name = [k for k, v in months_map.items() if v == target_m_num][0]

    st.sidebar.markdown("---")
    st.sidebar.info(f"📅 **Excel Period:** {bulan_excel} {tahun_excel}\n\n🔍 **Displaying:** {target_m_name} {target_y}")

    # HEADER DASHBOARD
    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    
    # LOAD DATA SHEET YANG DIPILIH
    df_main = pd.read_excel(FILE_PATH, sheet_name=sheet_pilihan, header=1)
    df_main = clean_df(df_main)

    # SECTION: TOP 10 & CHART
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        df_main['RATE'] = pd.to_numeric(df_main['RATE'], errors='coerce').fillna(0)
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.subheader(f"🏆 Top 10 Removal Rate ({target_m_name})")
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], use_container_width=True, hide_index=True)
        with col_b:
            fig = px.bar(top_10, x='PART NUMBER', y='RATE', color='RATE', color_continuous_scale='Reds')
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # SECTION: EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    
    if search:
        df_main = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

    # Tabel Utama Interaktif
    event = st.dataframe(df_main, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # DETAIL HISTORY (DRILL-DOWN)
    if event.selection.rows:
        row = df_main.iloc[event.selection.rows[0]]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.info(f"### 🛠️ Maintenance History for P/N: {pn_selected}")
        
        if not df_history.empty:
            # Cari kolom P/N yang benar di history
            col_pn_hist = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_history.columns else 'PART NUMBER'
            
            # Filter berdasarkan P/N dan Periode Bulan Sebelumnya
            hist_match = df_history[
                (df_history[col_pn_hist].astype(str).str.strip() == pn_selected) &
                (df_history['DATE'].dt.month == target_m_num) &
                (df_history['DATE'].dt.year == target_y)
            ]
            
            show_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
            available = [c for c in show_cols if c in df_history.columns]
            
            if not hist_match.empty:
                st.table(hist_match[available])
            else:
                st.warning(f"Tidak ada record removal pada bulan {target_m_name} {target_y}.")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
