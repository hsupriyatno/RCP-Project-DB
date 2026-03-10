import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. FUNGSI PEMBERSIH & LOAD DATA
def clean_dynamic_columns(df):
    for i in range(len(df)):
        row_values = [str(val).upper() for val in df.iloc[i].values]
        if 'PART NUMBER' in row_values:
            new_cols = df.iloc[i].values
            df = df.iloc[i+1:].copy()
            df.columns = [str(c).strip().upper() for c in new_cols]
            break
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df = df.fillna(0)
    if 'RATE' in df.columns:
        df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
    if 'QTY REM' in df.columns:
        df['QTY REM'] = pd.to_numeric(df['QTY REM'], errors='coerce').fillna(0)
    return df

@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        df_hist = pd.read_excel(file_name, sheet_name=2) 
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        
        date_col = next((c for c in df_hist.columns if 'DATE' in c), None)
        if date_col:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist[date_col], errors='coerce')
            df_hist['DATE_STR'] = df_hist['DATE_DT'].dt.strftime('%d-%m-%Y')
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# 3. LOGIKA PERIODE
def get_period_info(bulan, tahun):
    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    m_num = m_map.get(bulan, 12)
    try:
        y_int = int(float(tahun))
        curr = datetime(y_int, m_num, 1)
        prev = curr - timedelta(days=1)
        p_name = [k for k, v in m_map.items() if v == prev.month][0]
        return prev.month, prev.year, p_name
    except:
        return 1, 2026, "JANUARY"

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names)
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, target_m_name = get_period_info(bln_ref, thn_ref)
    full_period = f"{target_m_name} {target_y}"

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    
    # METRICS
    if not df_main.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Components", f"{len(df_main)} Items")
        c2.metric("Total Removals", f"{int(df_main['QTY REM'].sum())} EA")
        c3.metric("Avg Rate", f"{df_main['RATE'].mean():.2f}")

    st.divider()

    # 4. CHART & TOMBOL DATA TABLE (TOP 10)
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        st.subheader(f"📈 Top 10 Removal Rate ({full_period})")
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + " - " + top_10['DESCRIPTION'].astype(str)
        
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200')
        st.plotly_chart(fig, use_container_width=True)

        # PASTIKAN INI MUNCUL: Tabel di bawah chart
        with st.expander("👉 KLIK DISINI UNTUK LIHAT TABEL DATA TOP 10", expanded=False):
            st.write("Data detail untuk 10 Part Number dengan Rate tertinggi:")
            st.table(top_10[['PART NUMBER', 'DESCRIPTION', 'QTY REM', 'RATE']])

    st.divider()

    # 5. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    # Menggunakan on_select untuk interaksi tabel
    event = st.dataframe(
        filtered, 
        use_container_width=True, 
        hide_index=True, 
        on_select="rerun", 
        selection_mode="single-row"
    )

    # 6. TABEL DETAIL REMOVAL (HISTORY)
    if event.selection.rows:
        idx = event.selection.rows[0]
        row = filtered.iloc[idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ DETAIL REMOVAL: {pn_selected}")
        
        # Tampilkan tabel history jika data tersedia
        if not df_history.empty:
            col_pn_h = next((c for c in df_history.columns if 'PART' in c), None)
            if col_pn_h and 'DATE_DT' in df_history.columns:
                hist_match = df_history[
                    (df_history[col_pn_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE_DT'].dt.month == target_m) & 
                    (df_history['DATE_DT'].dt.year == target_y)
                ].copy()
                
                if not hist_match.empty:
                    cols_to_show = ['DATE_STR', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                    existing = [c for c in cols_to_show if c in hist_match.columns]
                    st.dataframe(hist_match[existing], use_container_width=True, hide_index=True)
                else:
                    st.warning(f"Tidak ada record history untuk {pn_selected} pada periode {full_period}.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")

st.sidebar.markdown("---")
st.sidebar.info("User: HERY SUPRIYATNO")
