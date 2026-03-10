import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. KONFIGURASI HALAMAN & CSS
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    [data-testid="stMetricLabel"] { font-size: 14px !important; }
    [data-testid="stMetricValue"] { font-size: 22px !important; }
    .stMetric { 
        background-color: #ffffff; 
        padding: 10px; 
        border-radius: 10px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
    }
    /* Mematikan interaksi hover berlebih agar stabil di HP */
    [data-testid="stDataFrame"] { pointer-events: auto; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNGSI PEMBERSIH DATA
def clean_dynamic_columns(df):
    for i in range(len(df)):
        row_values = [str(val).upper() for val in df.iloc[i].values]
        if 'PART NUMBER' in row_values:
            new_cols = df.iloc[i].values
            df = df.iloc[i+1:].copy()
            df.columns = [str(c).strip().upper() for c in new_cols]
            break
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).fillna(0)
    if 'RATE' in df.columns:
        df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
    if 'QTY REM' in df.columns:
        df['QTY REM'] = pd.to_numeric(df['QTY REM'], errors='coerce').fillna(0)
    return df

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().split('.')[0]
        
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        
        if 'DATE' in df_hist.columns:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            df_hist['DATE_STR'] = df_hist['DATE_DT'].dt.strftime('%d-%m-%Y')
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# 4. LOGIKA PERIODE
def get_period_info(bulan, tahun):
    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    try:
        curr = datetime(int(tahun), m_map.get(bulan, 1), 1)
        prev = curr - timedelta(days=1)
        p_name = [k for k, v in m_map.items() if v == prev.month][0]
        return prev.month, prev.year, p_name
    except:
        return 1, 2026, "JANUARY"

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    st.sidebar.title("Navigation")
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names)
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, target_m_name = get_period_info(bln_ref, thn_ref)

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")

    # 5. CHART TOP 10
    if 'PART NUMBER' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        st.subheader(f"📈 Top 10 Removal Rate ({target_m_name} {target_y})")
        fig = px.bar(top_10, x='PART NUMBER', y='RATE', text_auto='.2f', hover_data=['DESCRIPTION'])
        fig.update_traces(marker_color='#F2B200')
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 6. COMPONENT EXPLORER (Mode Stabil untuk HP)
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    # PERUBAHAN UTAMA: Gunakan selectbox untuk memilih part agar tabel tidak sensitif sentuhan
    list_pn = filtered['PART NUMBER'].unique().tolist()
    pn_selected = st.selectbox("Pilih Part Number untuk melihat detail:", ["-- Pilih --"] + list_pn)

    # Tabel ditampilkan murni sebagai informasi (tanpa mode seleksi agar tidak 'hilang' saat disentuh)
    st.dataframe(filtered, use_container_width=True, hide_index=True)

    # 7. PART REMOVAL DETAIL
    if pn_selected != "-- Pilih --":
        row = filtered[filtered['PART NUMBER'] == pn_selected].iloc[0]
        
        st.write("---")
        st.subheader(f"🛠️ DETAIL REMOVAL: {pn_selected}")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        c2.metric("Rate", f"{row.get('RATE', 0):.2f}")
        c3.metric("Qty", f"{int(row.get('QTY REM', 0))} EA")

        if not df_history.empty and 'DATE_DT' in df_history.columns:
            hist_match = df_history[
                (df_history.iloc[:, 1].astype(str).str.strip() == str(pn_selected)) & 
                (df_history['DATE_DT'].dt.month == target_m) & 
                (df_history['DATE_DT'].dt.year == target_y)
            ].copy()
            
            if not hist_match.empty:
                st.dataframe(hist_match[['DATE_STR', 'REASON OF REMOVAL', 'TSN', 'TSO']], 
                             use_container_width=True, hide_index=True)
            else:
                st.info("Tidak ada data removal pada periode ini.")

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.markdown("---")
st.sidebar.info("User: Hery Supriyatno")
