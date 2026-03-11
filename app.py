import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from dateutil.relativedelta import relativedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. FUNGSI LOAD DATA (DIPERKUAT)
def load_all_data(file_name):
    try:
        # AMBIL REFERENCE PERIOD (A2/A3)
        df_ref = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        ref_month_str = str(df_ref.iloc[1, 0]).strip().upper() 
        ref_year_int = int(float(str(df_ref.iloc[2, 0]).strip())) 
        
        m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                 'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
        ref_month_idx = m_map.get(ref_month_str, 1)

        # LOGIKA N-1 (Mundur 1 Bulan)
        input_date = datetime(ref_year_int, ref_month_idx, 1)
        target_date = input_date - relativedelta(months=1)
        
        c_month_name = target_date.strftime('%B').upper()
        c_year_int = target_date.year
        c_m_idx = target_date.month

        # LOAD TABEL UTAMA
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        h_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                h_idx = i
                break
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df_main.columns = [str(c).strip().upper() for c in df_main.columns]
        
        # Rate Kolom O (Index 14)
        df_main['RATE_1MO'] = pd.to_numeric(df_main.iloc[:, 14], errors='coerce').fillna(0)
        df_main['PN_DESC_CHART'] = df_main['PART NUMBER'].astype(str) + "<br>" + df_main['DESCRIPTION'].astype(str).str[:25]
        
        # LOAD HISTORY
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        date_col = next((c for c in df_hist.columns if 'DATE' in c), None)
        if date_col:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist[date_col], errors='coerce')
            df_hist['DATE_DISPLAY'] = df_hist['DATE_DT'].dt.strftime('%d-%b-%Y')
        
        return df_main, df_hist, ref_month_str, ref_year_int, c_month_name, c_year_int, c_m_idx
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return None, None, None, None, None, None, None

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

# 3. PENGGUNAAN SESSION STATE AGAR DATA STABIL SAAT DI-SCROLL DI HP
if 'data_loaded' not in st.session_state or st.sidebar.button("🔄 Sync with Excel"):
    res = load_all_data(FILE_PATH)
    if res[0] is not None:
        st.session_state.df_main = res[0]
        st.session_state.df_history = res[1]
        st.session_state.p_month = res[2]
        st.session_state.p_year = res[3]
        st.session_state.c_month = res[4]
        st.session_state.c_year = res[5]
        st.session_state.c_m_idx = res[6]
        st.session_state.data_loaded = True
        st.cache_data.clear()

if 'data_loaded' in st.session_state:
    st.title("📊 Reliability Analysis Dashboard")
    
    # Header Info
    c1, c2 = st.columns(2)
    c1.info(f"📁 **Excel Period:** {st.session_state.p_month} {st.session_state.p_year}")
    c2.success(f"⚙️ **Processing For (N-1):** {st.session_state.c_month} {st.session_state.c_year}")

    # 4. CHART TOP 10 (DIBUNGKUS CONTAINER AGAR TIDAK FLICKER)
    with st.container():
        st.subheader(f"📈 Top 10 Removal Rate ({st.session_state.c_month})")
        
        # Sorting Data
        top_10 = st.session_state.df_main.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
        
        # Konfigurasi Chart Plotly agar ramah HP (Responsive)
        fig = px.bar(top_10, x='PN_DESC_CHART', y='RATE_1MO', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.5)
        fig.update_layout(
            xaxis_tickangle=-45, 
            margin=dict(l=20, r=20, t=20, b=120), # Margin bawah lebih luas untuk label PN
            xaxis_title=None,
            hovermode="closest",
            dragmode=False # Mematikan zoom agar tidak sengaja ter-zoom saat scroll di HP
        )
        
        # Tampilkan Chart dengan config statis
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    st.divider()

    # 5. DATA TABLE SUMMARY (Menggunakan st.data_editor agar lebih ringan di HP)
    with st.expander("🔍 Detail List Top 10", expanded=True):
        st.write("Pilih baris untuk melihat detail history:")
        selected_event = st.dataframe(
            top_10[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
            use_container_width=True, 
            hide_index=True,
            on_select="rerun", 
            selection_mode="single-row"
        )

    # 6. DETAIL & HISTORY
    if selected_event.selection.rows:
        row_idx = selected_event.selection.rows[0]
        sel_row = top_10.iloc[row_idx]
        pn_selected = str(sel_row['PART NUMBER']).strip()
        
        pn_col_h = next((c for c in st.session_state.df_history.columns if 'PART' in c), st.session_state.df_history.columns[1])
        
        # Filter History N-1
        hist_match = st.session_state.df_history[
            (st.session_state.df_history[pn_col_h].astype(str).str.strip() == pn_selected) & 
            (st.session_state.df_history['DATE_DT'].dt.month == st.session_state.df_m_idx if 'df_m_idx' in st.session_state else st.session_state.c_m_idx) & 
            (st.session_state.df_history['DATE_DT'].dt.year == st.session_state.c_year)
        ].copy().sort_values(by='DATE_DT', ascending=False)

        st.markdown(f"### 🛠️ Investigation: {pn_selected}")
        
        m1, m2 = st.columns(2)
        m1.metric("Rate", f"{sel_row['RATE_1MO']:.2f}")
        m2.metric("Qty Removed (N-1)", f"{len(hist_match)} EA")
        
        if not hist_match.empty:
            st.table(hist_match[['DATE_DISPLAY', 'REASON OF REMOVAL', 'TSN', 'TSO']]) # Menggunakan st.table agar statis di HP
        else:
            st.info("No removal records found for this period.")

st.sidebar.info(f"User: HERY SUPRIYATNO\nReliability Engineer")
