import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability Airfast DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Kriteria (Bulan dari A2, Tahun dari A3)
@st.cache_data
def load_filter_criteria(file_name):
    try:
        df_criteria = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bulan_str = str(df_criteria.iloc[1, 0]).strip().upper() # Cell A2
        tahun_str = str(df_criteria.iloc[2, 0]).strip().replace('.0', '') # Cell A3
        return bulan_str, tahun_str
    except Exception as e:
        return "N/A", "N/A"

# 3. Fungsi Load Data Utama & History
@st.cache_data
def load_all_data(file_name, sheet_pilihan):
    try:
        df_utama = pd.read_excel(file_name, sheet_name=sheet_pilihan, header=1)
        df_utama = df_utama.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df_utama.columns = [str(col).strip() for col in df_utama.columns]
        
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_utama, df_hist
    except:
        return pd.DataFrame(), pd.DataFrame()

# --- EKSEKUSI DASHBOARD ---
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    
    # Ambil Kriteria dari Excel
    bulan_excel, tahun_excel = load_filter_criteria(file_target)
    
    # Logic Pengurangan 1 Bulan
    months_map = {
        'JANUARY': 1, 'FEBRUARY': 2, 'MARCH': 3, 'APRIL': 4, 'MAY': 5, 'JUNE': 6,
        'JULY': 7, 'AUGUST': 8, 'SEPTEMBER': 9, 'OCTOBER': 10, 'NOVEMBER': 11, 'DECEMBER': 12
    }
    
    current_month_num = months_map.get(bulan_excel)
    
    if current_month_num and tahun_excel.isdigit():
        # Buat objek tanggal berdasarkan input Excel
        current_date = datetime(int(tahun_excel), current_month_num, 1)
        # Kurangi 1 bulan menggunakan timedelta (mundur ke hari terakhir bulan sebelumnya)
        previous_date = current_date - timedelta(days=1)
        
        target_month_num = previous_date.month
        target_year_val = previous_date.year
        # Cari nama bulan untuk tampilan UI
        target_month_name = [k for k, v in months_map.items() if v == target_month_num][0]
    else:
        target_month_num, target_year_val, target_month_name = None, None, "N/A"

    # Sidebar & Data Load
    xls = pd.ExcelFile(file_target)
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet:", ["REMOVAL RATE CALCULATION", "ALERT LEVEL CALCULATION"])
    data_utama, data_history = load_all_data(file_target, sheet_pilihan)

    st.sidebar.info(f"📌 **Excel Status:** {bulan_excel} {tahun_excel}\n\n🔍 **Displaying Data:** {target_month_name} {target_year_val}")

    # Tabel Utama
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        data_utama = data_utama[data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

    event = st.dataframe(data_utama, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # --- DETAIL HISTORY DENGAN FILTER PREVIOUS MONTH ---
    if event.selection.rows:
        idx = event.selection.rows[0]
        row = data_utama.iloc[idx]
        pn = str(row['PART NUMBER']).strip()
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn}")
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.info(f"**Description:**\n{row.get('DESCRIPTION', 'N/A')}")
                st.metric(f"Total Qty Removal ({target_month_name})", f"{row.get('QTY REM', 0)} EA")
            
            with col2:
                st.markdown(f"**📅 Removal Records in {target_month_name} {target_year_val}:**")
                if not data_history.empty and target_month_num:
                    # Filter P/N DAN Filter Bulan Sebelumnya
                    detail = data_history[
                        (data_history['PART NUMBER OFF'].astype(str).str.strip() == pn) &
                        (data_history['DATE'].dt.month == target_month_num) &
                        (data_history['DATE'].dt.year == target_year_val)
                    ]
                    
                    show_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                    valid = [c for c in show_cols if c in detail.columns]
                    
                    if not detail.empty:
                        st.table(detail[valid])
                    else:
                        st.warning(f"Tidak ada removal pada bulan {target_month_name} {target_year_val}.")

except Exception as e:
    st.error(f"Error: {e}")
