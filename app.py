import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman & Sidebar
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

def clean_df(df):
    """Membersihkan dataframe dari spasi dan kolom kosong"""
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = [str(col).strip() for col in df.columns]
    return df

# 2. Fungsi Load Data Utama
@st.cache_data
def load_reliability_data(file_name):
    try:
        # Ambil periode laporan (A2 & A3)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_ref = str(df_crit.iloc[1, 0]).strip().upper()
        thn_ref = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load History dari sheet COMPONENT REPLACEMENT
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist = clean_df(df_hist)
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
        return df_hist, bln_ref, thn_ref
    except Exception as e:
        st.error(f"Gagal memuat file: {e}")
        return pd.DataFrame(), "N/A", "N/A"

# --- PROSES DATA ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    df_history, bulan_excel, tahun_excel = load_reliability_data(FILE_PATH)

    # 3. SIDEBAR NAVIGATION
    st.sidebar.title("🚀 Navigation")
    all_sheets = xls.sheet_names
    sheet_pilihan = st.sidebar.selectbox("Select Report Sheet:", all_sheets)
    
    # Logika Pengurangan 1 Bulan untuk Filter
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

    # 4. MAIN DASHBOARD
    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")

    # Load Data Sheet yang dipilih
    df_main = pd.read_excel(FILE_PATH, sheet_name=sheet_pilihan, header=1)
    df_main = clean_df(df_main)

    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        df_main['RATE'] = pd.to_numeric(df_main['RATE'], errors='coerce').fillna(0)
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)

        # --- POSISI CHART PALING ATAS ---
        st.subheader(f"📈 Top 10 Removal Rate Chart ({target_m_name})")
        fig = px.bar(top_10, x='PART NUMBER', y='RATE', color='RATE', 
                     color_continuous_scale='Reds', text_auto='.2f')
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. COMPONENT EXPLORER (DI BAWAH CHART)
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    
    if search:
        df_main = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

    # Tabel interaktif
    selection = st.dataframe(df_main, use_container_width=True, hide_index=True, 
                             on_select="rerun", selection_mode="single-row")

    # 6. DETAIL HISTORY (DRILL-DOWN)
    if selection.selection.rows:
        row_idx = selection.selection.rows[0]
        selected_pn = str(df_main.iloc[row_idx]['PART NUMBER']).strip()
        
        st.info(f"### 🛠️ Maintenance History for P/N: {selected_pn}")
        
        if not df_history.empty:
            col_pn_hist = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_history.columns else 'PART NUMBER'
            
            # Filter History berdasarkan P/N dan Periode
            hist_match = df_history[
                (df_history[col_pn_hist].astype(str).str.strip() == selected_pn) &
                (df_history['DATE'].dt.month == target_m_num) &
                (df_history['DATE'].dt.year == target_y)
            ]
            
            show_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
            available = [c for c in show_cols if c in df_history.columns]
            
            if not hist_match.empty:
                st.table(hist_match[available])
            else:
                st.warning(f"Tidak ditemukan catatan removal untuk {selected_pn} pada periode {target_m_name} {target_y}.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")


