Mohon maaf sekali, Pak Hery. Sepertinya saat saya merapikan kode CSS untuk mengecilkan huruf metrik tadi, bagian Summary Table yang ada di dalam expander tidak sengaja tergeser atau terhapus.

Mari kita kembalikan tombol "View Top 10 Summary Table" tersebut tepat di bawah Chart dan di atas Component Explorer.

Berikut adalah kode lengkap dengan CSS (huruf kecil), Judul PART REMOVAL DETAIL, dan Tombol Summary Chart yang sudah kembali:

Python
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi & Custom CSS (Huruf Metrik Diperkecil)
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
    </style>
    """, unsafe_allow_html=True)

# 2. Fungsi Load Data
@st.cache_data
def load_reliability_data(file_name, sheet_name):
    try:
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() for col in df.columns]
        if 'RATE' in df.columns:
            df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
        return df, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), "N/A", "N/A"

@st.cache_data
def load_replacement_history(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

def get_previous_month(bulan, tahun):
    months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                  'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    try:
        m_num = months_map.get(bulan, 12)
        current_dt = datetime(int(tahun), m_num, 1)
        prev_dt = current_dt - timedelta(days=1)
        return prev_dt.month, prev_dt.year, [k for k, v in months_map.items() if v == prev_dt.month][0]
    except:
        return 11, 2025, "NOVEMBER"

# --- MAIN APP ---
try:
    FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(FILE_PATH)
    
    # Sidebar
    st.sidebar.title("Navigation")
    sheet_pilihan = st.sidebar.selectbox("Select Report Sheet:", xls.sheet_names)
    
    df_main, bln_ref, thn_ref = load_reliability_data(FILE_PATH, sheet_pilihan)
    target_m_num, target_y, target_m_name = get_previous_month(bln_ref, thn_ref)
    df_history = load_replacement_history(FILE_PATH)

    # Header
    full_period = f"{target_m_name} {target_y}"
    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Reporting Period: {bln_ref} {thn_ref} | Analysis Data: {full_period}")

    # 4. CHART & SUMMARY TABLE (Kembali dimunculkan)
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        fig = px.bar(top_10, x='PART NUMBER', y='RATE', color='RATE', color_continuous_scale='Reds', text_auto='.4f')
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        # --- TOMBOL SUMMARY YANG TADI HILANG ---
        with st.expander("📊 View Top 10 Summary Table"):
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], use_container_width=True, hide_index=True)

    st.divider()

    # 5. EXPLORER & PART REMOVAL DETAIL
    st.subheader("🔍 Component Explorer")
    search_query = st.text_input("Search by Part Number or Description", placeholder="e.g. 040-21101")
    
    filtered_df = df_main.copy()
    if search_query:
        filtered_df = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)]

    event = st.dataframe(filtered_df, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered_df.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        m2.metric("Current Rate", f"{row.get('RATE', 0):.4f}")
        m3.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        if not df_history.empty:
            col_pn_hist = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_history.columns else 'PART NUMBER'
            hist_match = df_history[(df_history[col_pn_hist].astype(str).str.strip() == pn_selected) & 
                                    (df_history['DATE'].dt.month == target_m_num) & 
                                    (df_history['DATE'].dt.year == target_y)]
            
            if not hist_match.empty:
                st.table(hist_match[['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']])
            else:
                st.warning(f"No removal records found in {full_period} for this part.")

except Exception as e:
    st.error(f"Critical System Error: {e}")
