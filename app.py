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
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .js-plotly-plot .plotly .nsewdrag { pointer-events: none !important; }
    .js-plotly-plot .plotly .hoverlayer { pointer-events: auto !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. LOGIKA PERIODE
def get_period_labels(bulan_str, tahun_str):
    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    inv_map = {v: k for k, v in m_map.items()}
    try:
        curr_m = m_map.get(bulan_str.upper(), 12)
        curr_y = int(float(tahun_str))
        period_label = f"{bulan_str.capitalize()} {curr_y}"
        dt_curr = datetime(curr_y, curr_m, 1)
        dt_prev = dt_curr - timedelta(days=1)
        analysis_label = f"{inv_map[dt_prev.month].capitalize()} {dt_prev.year}"
        return period_label, analysis_label, dt_prev.month, dt_prev.year
    except:
        return f"{bulan_str} {tahun_str}", "N/A", 1, 2026

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_all_data(file_name):
    try:
        df_info = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_info.iloc[1, 0]).strip()
        thn_raw = str(df_info.iloc[2, 0]).strip().replace('.0', '')
        
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        h_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                h_idx = i
                break
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df_main.columns = [str(c).strip().upper() for c in df_main.columns]
        
        df_main['RATE_3MO'] = pd.to_numeric(df_main.iloc[:, 8], errors='coerce').fillna(0)
        df_main['RATE_2MO'] = pd.to_numeric(df_main.iloc[:, 11], errors='coerce').fillna(0)
        df_main['RATE_1MO'] = pd.to_numeric(df_main.iloc[:, 14], errors='coerce').fillna(0)
        df_main['QTY_VAL'] = pd.to_numeric(df_main.iloc[:, 13], errors='coerce').fillna(0)
        df_main['PN_DESC_CHART'] = df_main['PART NUMBER'].astype(str) + "<br>" + df_main['DESCRIPTION'].astype(str).str[:25]
        
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        date_col = next((c for c in df_hist.columns if 'DATE' in c), None)
        if date_col:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist[date_col], errors='coerce')
            df_hist['DATE_DISPLAY'] = df_hist['DATE_DT'].dt.strftime('%d-%b-%Y')
        
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH)
    period_txt, analysis_txt, m_idx, y_idx = get_period_labels(bln_ref, thn_ref)

    if not df_main.empty:
        st.title("📊 Reliability Analysis Dashboard")
        st.markdown(f"**Period Month:** {period_txt} | **Analysis Month:** {analysis_txt}")
        st.divider()

        # 4. CHART TOP 10
        st.subheader("📈 Top 10 Removal Rate (Current Month)")
        top_10 = df_main.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
        fig = px.bar(top_10, x='PN_DESC_CHART', y='RATE_1MO', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        fig.update_layout(dragmode=False, xaxis_tickangle=-45, margin=dict(b=100), xaxis_title=None)
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        # --- FITUR UTAMA: SELECTION DARI TOP 10 TABLE ---
        with st.expander("📊 Click to View Top 10 Data Table Summary", expanded=False):
            event_top10 = st.dataframe(
                top_10[['PART NUMBER', 'DESCRIPTION', 'QTY_VAL', 'RATE_1MO']], 
                use_container_width=True, hide_index=True,
                on_select="rerun", selection_mode="single-row",
                column_config={"QTY_VAL": "QTY REM", "RATE_1MO": "RATE"}
            )

        st.divider()

        # 5. COMPONENT EXPLORER
        st.subheader("🔍 Component Explorer")
        search = st.text_input("Cari Part Number atau Deskripsi:")
        filtered = df_main.copy()
        if search:
            mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            filtered = df_main[mask]

        event_explorer = st.dataframe(
            filtered[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
            use_container_width=True, hide_index=True, 
            on_select="rerun", selection_mode="single-row"
        )

        # 6. LOGIKA PENAMPILAN DETAIL & HISTORY (DARI KEDUA TABEL)
        sel_row = None
        # Cek jika ada pilihan dari Top 10 Table
        if event_top10.selection.rows:
            sel_row = top_10.iloc[event_top10.selection.rows[0]]
        # Cek jika ada pilihan dari Explorer (ini akan menimpa pilihan Top 10 jika keduanya diklik)
        elif event_explorer.selection.rows:
            sel_row = filtered.iloc[event_explorer.selection.rows[0]]

        if sel_row is not None:
            pn_selected = str(sel_row['PART NUMBER']).strip()
            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            c1.metric("Description", sel_row.get('DESCRIPTION', 'N/A'))
            c2.metric("Current Rate", f"{sel_row['RATE_1MO']:.2f}")
            c3.metric("Total Qty Rem", f"{int(sel_row['QTY_VAL'])} EA")
            
            st.write(f"**Part Removal History ({analysis_txt}):**")
            
            if not df_history.empty:
                pn_col_h = next((c for c in df_history.columns if 'PART' in c), df_history.columns[1])
                hist_match = df_history[
                    (df_history[pn_col_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE_DT'].dt.month == m_idx) & 
                    (df_history['DATE_DT'].dt.year == y_idx)
                ].copy()
                
                if not hist_match.empty:
                    st.dataframe(
                        hist_match[['DATE_DISPLAY', 'REASON OF REMOVAL', 'TSN', 'TSO']], 
                        use_container_width=True, hide_index=True,
                        column_config={"DATE_DISPLAY": "DATE"}
                    )
                else:
                    st.info(f"Tidak ada record removal untuk {pn_selected} pada periode {analysis_txt}.")

        st.divider()

        # 7. UPTREND PART REMOVAL
        st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Continuous Increase)")
        uptrend = df_main[(df_main['RATE_3MO'] > 0) & (df_main['RATE_2MO'] > df_main['RATE_3MO']) & (df_main['RATE_1MO'] > df_main['RATE_2MO'])].copy()
        if not uptrend.empty:
            st.warning(f"Terdeteksi {len(uptrend)} komponen dengan tren kenaikan.")
            st.dataframe(uptrend[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], use_container_width=True, hide_index=True)
        else:
            st.success("Analisis Tren: Stabil.")

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.info(f"User: HERY SUPRIYATNO\nReliability Engineer")
