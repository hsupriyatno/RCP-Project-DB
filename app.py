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
        # A. AMBIL REFERENCE PERIOD DARI EXCEL (Contoh: February 2026)
        df_ref = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        ref_month_str = str(df_ref.iloc[1, 0]).strip().upper() 
        ref_year_int = int(float(str(df_ref.iloc[2, 0]).strip())) 
        
        m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                 'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
        ref_month_idx = m_map.get(ref_month_str, 1)

        # B. LOGIKA REVISI: MUNDUR 1 BULAN (N-1)
        # Jika Input Feb 2026, maka target_date menjadi Jan 2026
        input_date = datetime(ref_year_int, ref_month_idx, 1)
        target_date = input_date - relativedelta(months=1)
        
        calc_month_idx = target_date.month
        calc_year_int = target_date.year
        calc_month_name = target_date.strftime('%B').upper()

        # C. LOAD TABEL UTAMA (REMOVAL RATE)
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        h_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                h_idx = i
                break
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df_main.columns = [str(c).strip().upper() for c in df_main.columns]
        
        df_main['RATE_1MO'] = pd.to_numeric(df_main.iloc[:, 14], errors='coerce').fillna(0)
        df_main['PN_DESC_CHART'] = df_main['PART NUMBER'].astype(str) + "<br>" + df_main['DESCRIPTION'].astype(str).str[:25]
        
        # D. LOAD HISTORY (COMPONENT REPLACEMENT)
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        date_col = next((c for c in df_hist.columns if 'DATE' in c), None)
        if date_col:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist[date_col], errors='coerce')
            df_hist['DATE_DISPLAY'] = df_hist['DATE_DT'].dt.strftime('%d-%b-%Y')
        
        return df_main, df_hist, ref_month_str, ref_year_int, calc_month_name, calc_year_int, calc_month_idx
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", 2026, "N/A", 2026, 1

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

if st.sidebar.button("🔄 Sync with Excel"):
    st.cache_data.clear()

# Load data dengan variabel tambahan untuk periode kalkulasi (N-1)
df_main, df_history, p_month, p_year, c_month, c_year, c_m_idx = load_all_data(FILE_PATH)

try:
    if not df_main.empty:
        st.title("📊 Reliability Analysis Dashboard")
        
        # Header Informasi Period
        col_header1, col_header2 = st.columns(2)
        col_header1.info(f"📁 **Excel Period (A2/A3):** {p_month} {p_year}")
        col_header2.success(f"⚙️ **Data Processing For:** {c_month} {c_year} (N-1)")
        
        st.divider()

        # 4. CHART TOP 10 (Tetap mengambil Kolom O karena Excel Bapak biasanya sudah menghitung rate bulan lalu di kolom tersebut)
        st.subheader(f"📈 Top 10 Removal Rate ({c_month} {c_year})")
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

        # 6. DETAIL & HISTORY (Filter menggunakan c_m_idx dan c_year agar lari ke January)
        if event_top10.selection.rows:
            sel_row = top_10.iloc[event_top10.selection.rows[0]]
            pn_selected = str(sel_row['PART NUMBER']).strip()
            
            pn_col_h = next((c for c in df_history.columns if 'PART' in c), df_history.columns[1])
            
            # FILTER HISTORY: Mengunci ke bulan N-1 (Januari jika input Februari)
            hist_match = df_history[
                (df_history[pn_col_h].astype(str).str.strip() == pn_selected) & 
                (df_history['DATE_DT'].dt.month == c_m_idx) & 
                (df_history['DATE_DT'].dt.year == c_year)
            ].copy().sort_values(by='DATE_DT', ascending=False)
            
            qty_rem = len(hist_match)

            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            c1.metric("Description", sel_row.get('DESCRIPTION', 'N/A'))
            c2.metric("Rate Period", f"{sel_row['RATE_1MO']:.2f}")
            c3.metric("Qty Removed", f"{qty_rem} EA")
            
            st.write(f"**Actual Removal Records in {c_month} {c_year}:**")
            if not hist_match.empty:
                st.dataframe(hist_match[['DATE_DISPLAY', 'REASON OF REMOVAL', 'TSN', 'TSO']], use_container_width=True, hide_index=True)
            else:
                st.info(f"Tidak ada record pelepasan di sheet history untuk periode {c_month} {c_year}.")

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.info(f"User: HERY SUPRIYATNO\nReliability Engineer")
