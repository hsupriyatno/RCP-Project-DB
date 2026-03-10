import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Data Utama (Termasuk mengambil Bulan & Tahun)
@st.cache_data
def load_data(file_name, sheet_name):
    try:
        # Load data utama (Tabel dimulai baris ke-2)
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() if "Unnamed" not in str(col) else "" for col in df.columns]
        
        # Ambil Kriteria Bulan (A2) dan Tahun (A3) dari Sheet yang sama
        # Karena header=1, maka A2 di Excel menjadi baris 0 di DataFrame awal
        criteria_df = pd.read_excel(file_name, sheet_name=sheet_name, header=None, nrows=3)
        bulan_target = str(criteria_df.iloc[1, 0]).strip() # Cell A2
        tahun_target = str(criteria_df.iloc[2, 0]).strip() # Cell A3
        
        return df, bulan_target, tahun_target
    except Exception as e:
        st.error(f"Gagal memuat sheet {sheet_name}: {e}")
        return pd.DataFrame(), "", ""

# 3. Fungsi Load History (COMPONENT REPLACEMENT)
@st.cache_data
def load_history(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        # Pastikan kolom DATE dalam format datetime agar bisa difilter
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

# 4. Alur Utama
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    
    # Muat Data dan Kriteria Filter
    data_utama, target_month, target_year = load_data(file_target, sheet_pilihan)
    data_history = load_history(file_target)
    
    st.markdown(f"### 📊 REPORT: {sheet_pilihan}")
    st.sidebar.info(f"📅 **Filter Periode:** {target_month} {target_year}")
    
    search = st.text_input("🔍 Cari Part Number / Description:")
    display_data = data_utama[data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)] if search else data_utama

    st.info("💡 Klik baris tabel untuk melihat detail history removal.")
    event = st.dataframe(display_data, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # --- LOGIKA DRILL-DOWN DENGAN FILTER BULAN/TAHUN ---
    if event.selection.rows:
        index_terpilih = event.selection.rows[0]
        row_data = display_data.iloc[index_terpilih]
        pn_terpilih = str(row_data['PART NUMBER']).strip()
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn_terpilih}")
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.info(f"**Description:**\n\n{row_data.get('DESCRIPTION', 'N/A')}")
                st.metric("Total Qty Removal", f"{row_data.get('QTY REM', 0)} EA")
            with col2:
                st.markdown(f"**📅 Records for {target_month} {target_year}:**")
                
                if 'PART NUMBER OFF' in data_history.columns and 'DATE' in data_history.columns:
                    # 1. Filter berdasarkan P/N
                    detail_pn = data_history[data_history['PART NUMBER OFF'].astype(str).strip() == pn_terpilih].copy()
                    
                    # 2. Filter berdasarkan Bulan dan Tahun (Konversi nama bulan ke angka)
                    months_map = {
                        'JANUARY': 1, 'FEBRUARY': 2, 'MARCH': 3, 'APRIL': 4, 'MAY': 5, 'JUNE': 6,
                        'JULY': 7, 'AUGUST': 8, 'SEPTEMBER': 9, 'OCTOBER': 10, 'NOVEMBER': 11, 'DECEMBER': 12
                    }
                    target_month_num = months_map.get(target_month.upper())
                    
                    if target_month_num and target_year.isdigit():
                        detail_filtered = detail_pn[
                            (detail_pn['DATE'].dt.month == target_month_num) & 
                            (detail_pn['DATE'].dt.year == int(target_year))
                        ]
                        
                        cols_to_show = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                        available = [c for c in cols_to_show if c in detail_filtered.columns]
                        
                        if not detail_filtered.empty:
                            st.table(detail_filtered[available])
                        else:
                            st.warning(f"Tidak ada data removal untuk {target_month} {target_year}.")
                else:
                    st.error("Struktur kolom 'DATE' atau 'PART NUMBER OFF' tidak ditemukan.")

    # --- GRAFIK ---
    st.markdown("---")
    if 'PART NUMBER' in display_data.columns and 'RATE' in display_data.columns:
        chart_data = display_data.head(10).copy()
        chart_data['Label'] = chart_data['PART NUMBER'].astype(str) + " - " + chart_data['DESCRIPTION'].astype(str)
        fig = px.bar(chart_data, x='Label', y='RATE', text='RATE', color='RATE', color_continuous_scale='Reds', height=600)
        st.plotly_chart(fig, use_container_width=True)

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
