import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman & Judul
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
        return "DECEMBER", "2025"

# 3. Fungsi Load Data Utama (Sheet Kalkulasi)
@st.cache_data
def load_calc_data(file_name, sheet_name):
    try:
        # Baca mulai baris ke-2 (Header Excel di baris 2)
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() for col in df.columns]
        # Pastikan kolom RATE adalah angka
        if 'RATE' in df.columns:
            df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame()

# 4. Fungsi Load Database History (Sheet Replacement)
@st.cache_data
def load_history_db(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

# --- LOGIKA UTAMA ---
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    
    # A. Ambil Periode Filter (Logic: Jika Dec 2025 di A2, maka Target = Nov 2025)
    bulan_excel, tahun_excel = load_filter_criteria(file_target)
    months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                  'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    
    current_month_num = months_map.get(bulan_excel, 12)
    current_dt = datetime(int(tahun_excel), current_month_num, 1)
    prev_dt = current_dt - timedelta(days=1)
    
    target_m_num = prev_dt.month
    target_y_val = prev_dt.year
    target_m_name = [k for k, v in months_map.items() if v == target_m_num][0]

    # B. Sidebar: Menampilkan Semua Sheet
    xls = pd.ExcelFile(file_target)
    all_sheets = xls.sheet_names
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Laporan:", all_sheets, index=all_sheets.index("REMOVAL RATE CALCULATION") if "REMOVAL RATE CALCULATION" in all_sheets else 0)
    
    st.sidebar.markdown("---")
    st.sidebar.info(f"📅 **Excel Period:** {bulan_excel} {tahun_excel}\n\n🔍 **Displaying Data:** {target_m_name} {target_y_val}")

    # C. Ambil Data
    df_main = load_calc_data(file_target, sheet_pilihan)
    df_history = load_history_db(file_target)

    # D. Tampilan Top 10 & Chart (Khusus sheet kalkulasi)
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        st.subheader(f"📊 Top 10 Highest Removal Rate - {target_m_name} {target_y_val}")
        
        # Ambil Top 10 berdasarkan RATE tertinggi
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        col_table, col_chart = st.columns([1, 1])
        with col_table:
            st.write("**Table View (Top 10):**")
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], use_container_width=True, hide_index=True)
            
        with col_chart:
            fig = px.bar(top_10, x='PART NUMBER', y='RATE', color='RATE', 
                         hover_data=['DESCRIPTION'], color_continuous_scale='Reds')
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")

    # E. Tabel Utama (Semua Data) & Pencarian
    st.subheader(f"📋 Full Data List: {sheet_pilihan}")
    search = st.text_input("🔍 Cari Part Number atau Nama Komponen:")
    if search:
        df_main = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
    
    # Pilih Baris untuk melihat Detail History
    event = st.dataframe(df_main, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # F. Detail History (Muncul saat baris diklik)
    if event.selection.rows:
        idx = event.selection.rows[0]
        row_terpilih = df_main.iloc[idx]
        
        # Proteksi jika kolom PART NUMBER tidak ada
        if 'PART NUMBER' in row_terpilih:
            pn = str(row_terpilih['PART NUMBER']).strip()
            st.markdown(f"### 🔍 Detailed History for P/N: {pn}")
            
            with st.container(border=True):
                c1, c2 = st.columns([1, 2])
                with c1:
                    st.info(f"**Description:**\n{row_terpilih.get('DESCRIPTION', 'N/A')}")
                    st.metric(f"Removal Rate ({target_m_name})", row_terpilih.get('RATE', 0))
                
                with c2:
                    st.markdown(f"**📅 Removal Records in {target_m_name} {target_y_val}:**")
                    if not df_history.empty and 'PART NUMBER OFF' in df_history.columns:
                        # Filter History berdasarkan P/N dan Bulan Sebelumnya
                        detail = df_history[
                            (df_history['PART NUMBER OFF'].astype(str).str.strip() == pn) &
                            (df_history['DATE'].dt.month == target_m_num) &
                            (df_history['DATE'].dt.year == target_y_val)
                        ]
                        
                        cols_show = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                        valid_cols = [c for c in cols_show if c in detail.columns]
                        
                        if not detail.empty:
                            st.table(detail[valid_cols])
                        else:
                            st.warning(f"Tidak ditemukan catatan penggantian untuk part ini di bulan {target_m_name}.")
                    else:
                        st.error("Sheet 'COMPONENT REPLACEMENT' atau kolom 'PART NUMBER OFF' tidak ditemukan.")
        else:
            st.warning("Baris yang dipilih tidak memiliki kolom 'PART NUMBER'.")

except Exception as e:
    st.error(f"Sistem mengalami kendala teknis: {e}")
