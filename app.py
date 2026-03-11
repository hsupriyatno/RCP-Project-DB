import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# 1. KONFIGURASI HALAMAN (SUDAH DIPERBAIKI)
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. FUNGSI LOAD DATA
def load_all_data(file_name):
    try:
        # AMBIL PERIOD DARI SEL A2 & A3
        df_ref = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        ref_month_str = str(df_ref.iloc[1, 0]).strip().upper() 
        ref_year_int = int(float(str(df_ref.iloc[2, 0]).strip())) 
        
        m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                 'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
        ref_month_idx = m_map.get(ref_month_str, 1)

        # LOAD TABEL UTAMA
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        h_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                h_idx = i
                break
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df_main.columns = [str(c).strip().upper() for c in df_main.columns]
        
        # Mapping Kolom Rate secara posisi (I=8, L=11, O=14)
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
        
        return df_main, df_hist, ref_month_str, ref_year_int, ref_month_idx
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", 2026, 1

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

# Tombol Sync di Sidebar
if st.sidebar.button("🔄 Sync with Excel"):
    st.cache_data.clear()

df_main, df_history, p_month, p_year, p_m_idx = load_all_data(FILE_PATH)

try:
    if not df_main.empty:
        st.title("📊 Reliability Analysis Dashboard")
        st.info(f"📅 **Current Period (A2/A3):** {p_month} {p_year}")
        st.divider()

        # 4. CHART TOP 10
        st.subheader(f"📈 Top 10 Removal Rate ({p_month})")
        top_10 = df_main.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
        fig = px.bar(top_10, x='PN_DESC_CHART', y='RATE_1MO', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        fig.update_layout(xaxis_tickangle=-45, margin=dict(b=100), xaxis_title=None)
        st.plotly_chart(fig, use_container_width=True)

        # 5. DATA TABLE SUMMARY
        with st.expander("📊 Click to View Data Table Summary", expanded=False):
            event_top10 = st.dataframe(
                top_10[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
                use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row"
            )

        # 6. DETAIL & HISTORY
        if event_top10.selection.rows:
            sel_row = top_10.iloc[event_top10.selection.rows[0]]
            pn_selected = str(sel_row['PART NUMBER']).strip()
            
            pn_col_h = next((c for c in df_history.columns if 'PART' in c), df_history.columns[1])
            
            # Filter History: PN + Bulan (p_m_idx) + Tahun (p_year)
            hist_match = df_history[
                (df_history[pn_col_h].astype(str).str.strip() == pn_selected) & 
                (df_history['DATE_DT'].dt.month == p_m_idx) & 
                (df_history['DATE_DT'].dt.year == p_year)
            ].copy().sort_values(by='DATE_DT', ascending=False)
            
            qty_rem = len(hist_match)

            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            c1.metric("Description", sel_row.get('DESCRIPTION', 'N/A'))
            c2.metric("Rate Period", f"{sel_row['RATE_1MO']:.2f}")
            c3.metric("Qty Removed", f"{qty_rem} EA")
            
            st.write(f"**History on {p_month} {p_year}:**")
            if not hist_match.empty:
                st.dataframe(hist_match[['DATE_DISPLAY', 'REASON OF REMOVAL', 'TSN', 'TSO']], use_container_width=True, hide_index=True)
            else:
                st.info(f"Tidak ada data pelepasan untuk {pn_selected} pada {p_month} {p_year}.")
        
        st.divider()

        # 7. UPTREND
        st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Continuous Increase)")
        uptrend = df_main[
            (df_main['RATE_1MO'] > df_main['RATE_2MO']) & 
            (df_main['RATE_2MO'] > df_main['RATE_3MO']) & 
            (df_main['RATE_3MO'] > 0)
        ].copy()

        if not uptrend.empty:
            st.warning(f"Terdeteksi {len(uptrend)} komponen dengan tren kenaikan.")
        else:
            st.success(f"✅ Tidak ada uptrend removal rate pada periode {p_month} {p_year}.")

        st.dataframe(
            uptrend[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], 
            use_container_width=True, hide_index=True,
            column_config={"RATE_3MO": "RATE PREV. 3MO", "RATE_2MO": "RATE PREV. 2MO", "RATE_1MO": "RATE PREV. 1MO"}
        )

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.info(f"User: HERY SUPRIYATNO\nReliability Engineer")
