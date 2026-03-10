import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability Airfast DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Kriteria (Bulan dari A2, Tahun dari A3)
@st.cache_data
def load_filter_criteria(file_name):
    try:
        # Membaca hanya 3 baris pertama dari kolom A di sheet target
        df_criteria = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bulan = str(df_criteria.iloc[1, 0]).strip().upper() # Cell A2
        tahun = str(df_criteria.iloc[2, 0]).strip()         # Cell A3
        # Membersihkan format tahun jika terbaca sebagai float (misal 2025.0)
        tahun = tahun.replace('.0', '')
        return bulan, tahun
    except Exception as e:
        st.error(f"Gagal membaca kriteria di A2/A3: {e}")
        return "N/A", "N/A"

# 3. Fungsi Load Data Utama
@st.cache_data
def load_main_data(file_name, sheet_name):
    try:
        # Header tabel dimulai dari baris ke-2 (index 1 di Python)
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() for col in df.columns]
        return df
    except:
        return pd.DataFrame()

# 4. Fungsi Load History Removal
@st.cache_data
def load_history(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

# --- EKSEKUSI DASHBOARD ---
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    
    # Ambil Kriteria Global dari sheet REMOVAL RATE CALCULATION
    target_month_name, target_year = load_filter_criteria(file_target)
    
    # Sidebar
    xls = pd.ExcelFile(file_target)
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Laporan:", xls.sheet_names)
    
    st.sidebar.markdown("---")
    st.sidebar.info(f"📅 **Periode Laporan:**\n{target_month_name} {target_year}")

    # Load Data
    data_utama = load_main_data(file_target, sheet_pilihan)
    data_history = load_history(file_target)

    # Input Pencarian
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        data_utama = data_utama[data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

    # Tabel Utama
    st.markdown(f"### 📊 Data: {sheet_pilihan}")
    event = st.dataframe(data_utama, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # --- LOGIKA DETAIL HISTORY DENGAN FILTER TANGGAL ---
    if event.selection.rows:
        idx = event.selection.rows[0]
        row = data_utama.iloc[idx]
        pn_terpilih = str(row['PART NUMBER']).strip()
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn_terpilih}")
        
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.info(f"**Description:**\n\n{row.get('DESCRIPTION', 'N/A')}")
                st.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")
            
            with col2:
                # Filter Data History berdasarkan P/N dan Periode dari A2/A3
                if not data_history.empty and 'PART NUMBER OFF' in data_history.columns:
                    # 1. Filter P/N
                    detail = data_history[data_history['PART NUMBER OFF'].astype(str).str.strip() == pn_terpilih].copy()
                    
                    # 2. Mapping Nama Bulan ke Angka
                    months_map = {
                        'JANUARY': 1, 'FEBRUARY': 2, 'MARCH': 3, 'APRIL': 4, 'MAY': 5, 'JUNE': 6,
                        'JULY': 7, 'AUGUST': 8, 'SEPTEMBER': 9, 'OCTOBER': 10, 'NOVEMBER': 11, 'DECEMBER': 12
                    }
                    target_month_num = months_map.get(target_month_name)
                    
                    # 3. Filter Tanggal
                    if target_month_num and target_year.isdigit():
                        detail = detail[
                            (detail['DATE'].dt.month == target_month_num) & 
                            (detail['DATE'].dt.year == int(target_year))
                        ]
                        
                        show_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                        cols_available = [c for c in show_cols if c in detail.columns]
                        
                        if not detail.empty:
                            st.table(detail[cols_available])
                        else:
                            st.warning(f"Tidak ditemukan data removal pada {target_month_name} {target_year}.")
                    else:
                        st.error(f"Kriteria filter tidak valid: {target_month_name} / {target_year}")

    # Grafik Top 10
    st.markdown("---")
    if 'PART NUMBER' in data_utama.columns and 'RATE' in data_utama.columns:
        fig = px.bar(data_utama.head(10), x='PART NUMBER', y='RATE', title="Top 10 Part Removal Rate", color='RATE', color_continuous_scale='Reds')
        st.plotly_chart(fig, use_container_width=True)

except Exception as e:
    st.error(f"Sistem mengalami kendala: {e}")
