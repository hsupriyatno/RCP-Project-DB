import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability Airfast DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Kriteria (Bulan dari A2, Tahun dari A3 di sheet REMOVAL RATE)
@st.cache_data
def load_filter_criteria(file_name):
    try:
        df_criteria = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bulan_str = str(df_criteria.iloc[1, 0]).strip().upper() 
        tahun_str = str(df_criteria.iloc[2, 0]).strip().replace('.0', '') 
        return bulan_str, tahun_str
    except:
        return "DECEMBER", "2025" # Fallback default

# 3. Fungsi Load Data Utama
@st.cache_data
def load_main_table(file_name, sheet_name):
    try:
        # Header tabel rata-rata dimulai baris ke-2 (index 1)
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        # Bersihkan nama kolom agar tidak error 'Series' object has no attribute 'strip'
        df.columns = [str(col).strip() for col in df.columns]
        return df
    except:
        return pd.DataFrame()

# 4. Fungsi Load History
@st.cache_data
def load_history_data(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

# --- EKSEKUSI ---
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # PERBAIKAN: Memunculkan semua sheet kembali di Sidebar
    # Kita ambil semua list sheet yang ada di file Excel Bapak
    all_sheets = xls.sheet_names
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet:", all_sheets)
    
    # Ambil Kriteria & Hitung Bulan Sebelumnya (Logic November)
    bulan_excel, tahun_excel = load_filter_criteria(file_target)
    months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                  'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    
    c_month_num = months_map.get(bulan_excel, 12)
    current_date = datetime(int(tahun_excel), c_month_num, 1)
    prev_date = current_date - timedelta(days=1)
    
    target_month_num = prev_date.month
    target_year_val = prev_date.year
    target_month_name = [k for k, v in months_map.items() if v == target_month_num][0]

    st.sidebar.info(f"📅 **Displaying:** {target_month_name} {target_year_val}")

    # Tampilkan konten berdasarkan sheet yang dipilih
    if sheet_pilihan == "CHART DASHBOARD":
        st.subheader("📈 Visualisasi Dashboard")
        data_calc = load_main_table(file_target, "REMOVAL RATE CALCULATION")
        if not data_calc.empty and 'RATE' in data_calc.columns:
            fig = px.bar(data_calc.head(15), x='PART NUMBER', y='RATE', color='RATE', title="Top Removal Rates")
            st.plotly_chart(fig, use_container_width=True)
    else:
        # Load Tabel untuk sheet lainnya
        df_display = load_main_table(file_target, sheet_pilihan)
        data_history = load_history_data(file_target)

        search = st.text_input(f"🔍 Cari di {sheet_pilihan}:")
        if search:
            df_display = df_display[df_display.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

        event = st.dataframe(df_display, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

        # Detail History (Hanya muncul jika baris diklik)
        if event.selection.rows:
            idx = event.selection.rows[0]
            row = df_display.iloc[idx]
            # Cek apakah kolom PART NUMBER ada untuk menghindari error struktur
            pn_col = 'PART NUMBER' if 'PART NUMBER' in df_display.columns else None
            
            if pn_col:
                pn_val = str(row[pn_col]).strip()
                st.markdown(f"### 🔍 Detailed History: {pn_val}")
                
                # Filter History
                if not data_history.empty and 'PART NUMBER OFF' in data_history.columns:
                    detail = data_history[
                        (data_history['PART NUMBER OFF'].astype(str).str.strip() == pn_val) &
                        (data_history['DATE'].dt.month == target_month_num) &
                        (data_history['DATE'].dt.year == target_year_val)
                    ]
                    if not detail.empty:
                        st.table(detail[['DATE', 'REASON OF REMOVAL', 'REMARK']])
                    else:
                        st.warning(f"Tidak ada data untuk {target_month_name} {target_year_val}")
                else:
                    st.error("Kolom 'PART NUMBER OFF' tidak ditemukan di sheet Replacement.")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")

