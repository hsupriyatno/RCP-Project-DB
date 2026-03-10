import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Dasar
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Data Utama (Sheet Kalkulasi)
@st.cache_data
def load_main_data(file_name, sheet_name):
    try:
        # Membaca kriteria Bulan (A2) dan Tahun (A3) secara eksplisit
        # header=None agar A1, A2, A3 terbaca sebagai data mentah
        criteria = pd.read_excel(file_name, sheet_name=sheet_name, header=None, nrows=3, usecols="A")
        bulan_target = str(criteria.iloc[1, 0]).strip().upper() # Cell A2
        tahun_target = str(criteria.iloc[2, 0]).strip()         # Cell A3
        
        # Membaca data tabel utama (Mulai baris ke-2)
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() if "Unnamed" not in str(col) else "" for col in df.columns]
        
        return df, bulan_target, tahun_target
    except Exception as e:
        st.error(f"Error load_main_data: {e}")
        return pd.DataFrame(), "N/A", "N/A"

# 3. Fungsi Load History (Sheet COMPONENT REPLACEMENT)
@st.cache_data
def load_history_data(file_name):
    try:
        # header=0 karena Judul 'PART NUMBER OFF' ada di baris pertama
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        # Bersihkan nama kolom dari spasi atau karakter aneh
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        
        # Konversi kolom DATE menjadi format waktu yang benar
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
        return df_hist
    except Exception as e:
        st.error(f"Error load_history: {e}")
        return pd.DataFrame()

# 4. Alur Dashboard
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Laporan:", xls.sheet_names)
    
    # Load data dengan filter dinamis
    data_utama, target_month, target_year = load_main_data(file_target, sheet_pilihan)
    data_history = load_history_data(file_target)
    
    st.markdown(f"### 📊 DASHBOARD: {sheet_pilihan}")
    st.sidebar.success(f"📅 Filter Periode: {target_month} {target_year}")

    # Pencarian
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        mask = data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_data = data_utama[mask]
    else:
        display_data = data_utama

    # Tampilkan Tabel Utama
    event = st.dataframe(
        display_data, 
        use_container_width=True, 
        hide_index=True,
        on_select="rerun", 
        selection_mode="single-row"
    )

    # --- LOGIKA DETAIL HISTORY ---
    if event.selection.rows:
        index_terpilih = event.selection.rows[0]
        row_data = display_data.iloc[index_terpilih]
        pn_terpilih = str(row_data['PART NUMBER']).strip()
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn_terpilih}")
        
        with st.container(border=True):
            c1, c2 = st.columns([1, 2])
            with c1:
                st.info(f"**Description:**\n\n{row_data.get('DESCRIPTION', 'N/A')}")
                st.metric("Total Qty Rem", f"{row_data.get('QTY REM', 0)} EA")
            
            with c2:
                st.markdown(f"**📅 Removal Records in {target_month} {target_year}:**")
                
                # Filter berdasarkan P/N dan Waktu
                if 'PART NUMBER OFF' in data_history.columns and 'DATE' in data_history.columns:
                    # Filter 1: Part Number
                    detail = data_history[data_history['PART NUMBER OFF'].astype(str).str.strip() == pn_terpilih].copy()
                    
                    # Mapping Bulan
                    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
                    t_month_num = m_map.get(target_month)
                    
                    # Filter 2: Bulan & Tahun
                    if t_month_num and target_year.isdigit():
                        detail = detail[
                            (detail['DATE'].dt.month == t_month_num) & 
                            (detail['DATE'].dt.year == int(target_year))
                        ]
                        
                        show_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                        valid_cols = [c for c in show_cols if c in detail.columns]
                        
                        if not detail.empty:
                            st.table(detail[valid_cols])
                        else:
                            st.warning(f"Tidak ada removal untuk P/N ini pada periode {target_month} {target_year}.")
                    else:
                        st.error(f"Kriteria filter tidak valid (Bulan: {target_month}, Tahun: {target_year})")
                else:
                    st.error("Struktur sheet 'COMPONENT REPLACEMENT' tidak sesuai.")

    # --- GRAFIK ---
    st.markdown("---")
    if 'PART NUMBER' in display_data.columns and 'RATE' in display_data.columns:
        top_10 = display_data.head(10).copy()
        top_10['Label'] = top_10['PART NUMBER'].astype(str) + " (" + top_10['DESCRIPTION'].astype(str) + ")"
        fig = px.bar(top_10, x='Label', y='RATE', color='RATE', color_continuous_scale='Reds')
        st.plotly_chart(fig, use_container_width=True)

except Exception as e:
    st.error(f"Sistem mengalami kendala: {e}")
