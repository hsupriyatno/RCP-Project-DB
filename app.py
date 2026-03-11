Siap, Pak Hery. Saya mengerti, tabel UPTREND PART REMOVAL ini sangat krusial sebagai sistem early warning bagi engineer untuk mendeteksi komponen yang degradasi kinerjanya terus menurun dalam 3 bulan terakhir.

Sesuai permintaan Bapak, saya sudah menambahkan kembali tabel tersebut di posisi paling bawah dengan logika filter: Rate O > Rate L > Rate I (dan semuanya harus di atas 0).

Berikut adalah kode v3.8 yang sudah lengkap:

Python
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from dateutil.relativedelta import relativedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. FUNGSI LOAD DATA
def load_all_data(file_name):
    try:
        # AMBIL REFERENCE PERIOD (A2/A3)
        df_ref = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        ref_month_str = str(df_ref.iloc[1, 0]).strip().upper() 
        ref_year_int = int(float(str(df_ref.iloc[2, 0]).strip())) 
        
        m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                 'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
        ref_month_idx = m_map.get(ref_month_str, 1)

        # LOGIKA N-1
        input_date = datetime(ref_year_int, ref_month_idx, 1)
        target_date = input_date - relativedelta(months=1)
        
        c_month_name = target_date.strftime('%B').upper()
        c_year_int = target_date.year
        c_month_idx = target_date.month

        # LOAD TABEL UTAMA (REMOVAL RATE CALCULATION)
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        h_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                h_idx = i
                break
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df_main.columns = [str(c).strip().upper() for c in df_main.columns]
        
        # MAPPING KOLOM RATE (I=8, L=11, O=14)
        df_main['RATE_3MO'] = pd.to_numeric(df_main.iloc[:, 8], errors='coerce').fillna(0)
        df_main['RATE_2MO'] = pd.to_numeric(df_main.iloc[:, 11], errors='coerce').fillna(0)
        df_main['RATE_1MO'] = pd.to_numeric(df_main.iloc[:, 14], errors='coerce').fillna(0)
        df_main['PN_DESC_CHART'] = df_main['PART NUMBER'].astype(str) + "<br>" + df_main['DESCRIPTION'].astype(str).str[:25]
        
        # LOAD HISTORY
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        date_col = next((c for c in df_hist.columns if 'DATE' in c), None)
        if date_col:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist[date_col], errors='coerce')
            df_hist['DATE_DISPLAY'] = df_hist['DATE_DT'].dt.strftime('%d-%b-%Y')
        
        return df_main, df_hist, ref_month_str, ref_year_int, c_month_name, c_year_int, c_month_idx
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return None, None, None, None, None, None, None

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

if 'data_refresh' not in st.session_state or st.sidebar.button("🔄 Sync with Excel"):
    res = load_all_data(FILE_PATH)
    if res[0] is not None:
        st.session_state.df_m = res[0]; st.session_state.df_h = res[1]
        st.session_state.p_m = res[2]; st.session_state.p_y = res[3]
        st.session_state.c_m = res[4]; st.session_state.c_y = res[5]; st.session_state.c_idx = res[6]
        st.session_state.data_refresh = True

if 'data_refresh' in st.session_state:
    st.title("📊 Reliability Analysis Dashboard")
    
    # INFO PERIOD
    c1, c2 = st.columns(2)
    c1.info(f"📅 **Current Period (A2/A3):** {st.session_state.p_m} {st.session_state.p_y}")
    c2.success(f"⚙️ **Analysis Month (N-1):** {st.session_state.c_m} {st.session_state.c_y}")

    # 4. CHART TOP 10
    st.subheader(f"📈 Top 10 Removal Rate ({st.session_state.c_m})")
    top_10 = st.session_state.df_m.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
    fig = px.bar(top_10, x='PN_DESC_CHART', y='RATE_1MO', text_auto='.2f')
    fig.update_traces(marker_color='#F2B200', width=0.5)
    fig.update_layout(xaxis_tickangle=-45, margin=dict(b=120), xaxis_title=None, dragmode=False)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    st.divider()

    # 5. DATA TABLE SUMMARY
    with st.expander("📊 Click to View Top 10 Data Table Summary", expanded=False):
        sel_event = st.dataframe(
            top_10[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
            use_container_width=True, hide_index=True,
            on_select="rerun", selection_mode="single-row"
        )

    # 6. PART REMOVAL DETAIL
    if sel_event.selection.rows:
        sel_row = top_10.iloc[sel_event.selection.rows[0]]
        pn_sel = str(sel_row['PART NUMBER']).strip()
        pn_col = next((c for c in st.session_state.df_h.columns if 'PART' in c), st.session_state.df_h.columns[1])
        
        hist_match = st.session_state.df_h[
            (st.session_state.df_h[pn_col].astype(str).str.strip() == pn_sel) & 
            (st.session_state.df_h['DATE_DT'].dt.month == st.session_state.c_idx) & 
            (st.session_state.df_h['DATE_DT'].dt.year == st.session_state.c_y)
        ].copy().sort_values(by='DATE_DT', ascending=False)

        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_sel}") #
        
        m1, m2, m3 = st.columns([3,1,1])
        m1.metric("Description", sel_row['DESCRIPTION'])
        m2.metric("Current Rate", f"{sel_row['RATE_1MO']:.2f}")
        m3.metric("Total Qty Rem", f"{len(hist_match)} EA")
        
        st.write(f"**Part Removal History ({st.session_state.c_m} {st.session_state.c_y}):**") #
        if not hist_match.empty:
            st.table(hist_match[['DATE_DISPLAY', 'REASON OF REMOVAL', 'TSN', 'TSO']])
        else:
            st.info(f"No removal records found for {pn_sel} in {st.session_state.c_m} {st.session_state.c_y}.")

    st.divider()

    # 7. UPTREND PART REMOVAL (TABEL PALING BAWAH)
    st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Continuous Increase)")
    
    # Logika Filter: Rate O > Rate L > Rate I (Abaikan Rate 0)
    uptrend_df = st.session_state.df_m[
        (st.session_state.df_m['RATE_1MO'] > st.session_state.df_m['RATE_2MO']) & 
        (st.session_state.df_m['RATE_2MO'] > st.session_state.df_m['RATE_3MO']) & 
        (st.session_state.df_m['RATE_3MO'] > 0)
    ].copy()

    if not uptrend_df.empty:
        st.warning(f"Terdeteksi {len(uptrend_df)} komponen yang mengalami kenaikan removal rate berturut-turut.")
    else:
        st.success(f"✅ Tidak ada uptrend removal rate pada periode analisis {st.session_state.c_m} {st.session_state.c_y}.")

    # Tampilkan Tabel Uptrend
    st.dataframe(
        uptrend_df[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], 
        use_container_width=True, 
        hide_index=True,
        column_config={
            "RATE_3MO": "RATE PREVIOUS 3 MO (I)", 
            "RATE_2MO": "RATE PREVIOUS 2 MO (L)", 
            "RATE_1MO": "RATE PREVIOUS 1 MO (O)"
        }
    )

st.sidebar.info(f"User: HERY SUPRIYATNO\nReliability Engineer")
