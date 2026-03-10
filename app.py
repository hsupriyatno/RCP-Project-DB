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
    /* Memungkinkan scroll menembus chart di HP */
    .js-plotly-plot .plotly .nsewdrag { pointer-events: none !important; }
    .js-plotly-plot .plotly .hoverlayer { pointer-events: auto !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNGSI PEMBERSIH DATA DINAMIS
def clean_dynamic_columns(df):
    for i in range(len(df)):
        row_values = [str(val).upper() for val in df.iloc[i].values]
        if 'PART NUMBER' in row_values:
            new_cols = df.iloc[i].values
            df = df.iloc[i+1:].copy()
            df.columns = [str(c).strip().upper() for c in new_cols]
            break
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).fillna(0)
    # Convert numeric columns
    cols_to_fix = ['RATE', 'QTY REM', 'RATE PREVIOUS 3 MO', 'RATE PREVIOUS 2 MO', 'RATE PREVIOUS 1 MO']
    for col in df.columns:
        if any(x in col for x in cols_to_fix):
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        # Load period & raw data
        df_raw = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        
        # Ambil Periode dari Sheet Utama (Baris awal)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Proses Data Utama
        df_main = clean_dynamic_columns(df_raw)
        
        # Load History Replacement
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
        curr = datetime(int(float(tahun)), m_map.get(bulan, 12), 1)
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
    full_period = f"{target_m_name} {target_y}"

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Reporting Month: {bln_ref} {thn_ref} | Analysis Period: {full_period}")

    # 5. CHART TOP 10
    if 'PART NUMBER' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        fig.update_layout(dragmode=False, hovermode="closest", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    st.divider()

    # 6. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    event = st.dataframe(filtered, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # 7. PART REMOVAL DETAIL
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
        c1, c2, c3 = st.columns([5, 1, 1])
        c1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        c2.metric("Current Rate", f"{row.get('RATE', 0):.2f}")
        c3.metric("Total Qty", f"{int(row.get('QTY REM', 0))} EA")

        if not df_history.empty and 'DATE_DT' in df_history.columns:
            hist_match = df_history[
                (df_history.iloc[:, 1].astype(str).str.strip() == pn_selected) & 
                (df_history['DATE_DT'].dt.month == target_m) & 
                (df_history['DATE_DT'].dt.year == target_y)
            ].copy()
            if not hist_match.empty:
                st.dataframe(hist_match[['DATE_STR', 'REASON OF REMOVAL', 'TSN', 'TSO']], use_container_width=True, hide_index=True)

    # 8. NEW: UPTREND PART REMOVAL (EARLY WARNING)
    st.write("---")
    st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Continuous Increase)")
    
    # Identifikasi kolom I, L, O (berdasarkan urutan di sheet REMOVAL RATE CALCULATION)
    # Dalam dataframe yang sudah dibersihkan, kita ambil kolom RATE historis
    try:
        # Kita buat kriteria: Rate Prev 3 < Rate Prev 2 < Rate Prev 1
        # Mengabaikan Rate 0
        uptrend_df = df_main[
            (df_main['RATE PREVIOUS 3 MO'] > 0) & 
            (df_main['RATE PREVIOUS 2 MO'] > df_main['RATE PREVIOUS 3 MO']) & 
            (df_main['RATE PREVIOUS 1 MO'] > df_main['RATE PREVIOUS 2 MO'])
        ].copy()

        if not uptrend_df.empty:
            st.warning(f"Ditemukan {len(uptrend_df)} komponen dengan tren kenaikan removal rate. Engineer segera evaluasi!")
            display_cols = ['PART NUMBER', 'DESCRIPTION', 'RATE PREVIOUS 3 MO', 'RATE PREVIOUS 2 MO', 'RATE PREVIOUS 1 MO']
            st.dataframe(
                uptrend_df[display_cols], 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "RATE PREVIOUS 3 MO": "Rate (I)",
                    "RATE PREVIOUS 2 MO": "Rate (L)",
                    "RATE PREVIOUS 1 MO": "Rate (O) 🚩"
                }
            )
        else:
            st.success("Tidak ada komponen yang mengalami kenaikan tren 3 bulan berturut-turut.")
    except Exception as e:
        st.info("Logika Uptrend memerlukan kolom data historis (I, L, O) di sheet Excel.")

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.markdown("---")
st.sidebar.info("Dashboard v1.6 | Hery Supriyatno")
