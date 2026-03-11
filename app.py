Siap, Pak Hery. Saya tangkap poinnya:

Batang Grafik: Dirampingkan agar tidak terlalu lebar.

Sumbu X: Menampilkan kombinasi Part Number dan Description agar informasinya lengkap dalam sekali lihat.

Fix Error: Memastikan tidak ada lagi error merah 'DATE' yang mengganggu.

Berikut adalah kode v2.1 yang sudah diperbarui:

Python
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
        return period_label, analysis_label
    except:
        return f"{bulan_str} {tahun_str}", "N/A"

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_reliability_data(file_name):
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
        
        df = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=h_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Mapping Kolom Berdasarkan Index (I=8, L=11, O=14, N=13 untuk QTY)
        df['RATE_3MO'] = pd.to_numeric(df.iloc[:, 8], errors='coerce').fillna(0)
        df['RATE_2MO'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0)
        df['RATE_1MO'] = pd.to_numeric(df.iloc[:, 14], errors='coerce').fillna(0)
        df['QTY_REM_VAL'] = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
        
        # Buat label gabungan PN + DESC untuk sumbu X
        df['PN_DESC'] = df['PART NUMBER'].astype(str) + "<br>" + df['DESCRIPTION'].astype(str).str[:25]
        
        return df, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal Load Data: {e}")
        return pd.DataFrame(), "N/A", "N/A"

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    df_main, bln_ref, thn_ref = load_reliability_data(FILE_PATH)
    period_txt, analysis_txt = get_period_labels(bln_ref, thn_ref)

    if not df_main.empty:
        st.title("📊 Reliability Analysis Dashboard")
        st.markdown(f"**Period Month:** {period_txt} | **Analysis Month:** {analysis_txt}")
        st.divider()

        # 4. CHART TOP 10 (RAMPING & LABEL LENGKAP)
        st.subheader("📈 Top 10 Removal Rate (Current Month)")
        top_10 = df_main.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
        
        fig = px.bar(top_10, x='PN_DESC', y='RATE_1MO', text_auto='.2f', 
                     labels={'RATE_1MO': 'Rate', 'PN_DESC': 'Part Number & Description'})
        
        # Mengatur lebar batang (width) dan warna
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        
        fig.update_layout(
            dragmode=False, 
            xaxis_tickangle=-45, 
            margin=dict(t=10, b=100), # Beri ruang bawah untuk label panjang
            xaxis_title=None
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        # --- TABLE SUMMARY TOP 10 ---
        with st.expander("📊 Click to View Top 10 Data Table"):
            st.dataframe(
                top_10[['PART NUMBER', 'DESCRIPTION', 'QTY_REM_VAL', 'RATE_1MO']], 
                use_container_width=True, hide_index=True,
                column_config={"QTY_REM_VAL": "QTY REM", "RATE_1MO": "RATE"}
            )

        st.divider()

        # 5. COMPONENT EXPLORER
        st.subheader("🔍 Component Explorer")
        search = st.text_input("Cari Part Number atau Deskripsi:")
        filtered = df_main.copy()
        if search:
            mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            filtered = df_main[mask]

        event = st.dataframe(
            filtered[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
            use_container_width=True, hide_index=True, 
            on_select="rerun", selection_mode="single-row"
        )

        # 6. PART REMOVAL DETAIL (SISTEM ERROR 'DATE' FIX)
        if event.selection.rows:
            sel = filtered.iloc[event.selection.rows[0]]
            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {sel['PART NUMBER']}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.metric("Description", sel.get('DESCRIPTION', 'N/A'))
            with c2: st.metric("Current Rate", f"{sel['RATE_1MO']:.2f}")
            with c3: st.metric("Total Qty Rem", f"{int(sel['QTY_REM_VAL'])} EA")
            
            st.write("**Recent Removal Trend (Rates):**")
            trend_data = pd.DataFrame({
                "Periode": ["3 Months Ago (I)", "2 Months Ago (L)", "Current Analysis (O)"],
                "Removal Rate": [sel['RATE_3MO'], sel['RATE_2MO'], sel['RATE_1MO']]
            })
            st.table(trend_data)

        st.divider()

        # 7. UPTREND PART REMOVAL
        st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Increase)")
        uptrend = df_main[
            (df_main['RATE_3MO'] > 0) & 
            (df_main['RATE_2MO'] > df_main['RATE_3MO']) & 
            (df_main['RATE_1MO'] > df_main['RATE_2MO'])
        ].copy()

        if not uptrend.empty:
            st.warning(f"Terdeteksi {len(uptrend)} komponen dengan tren kenaikan.")
            st.dataframe(
                uptrend[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], 
                use_container_width=True, hide_index=True,
                column_config={"RATE_3MO": "Rate (I)", "RATE_2MO": "Rate (L)", "RATE_1MO": "Rate (O) 🚩"}
            )
        else:
            st.success("Analisis Tren: Stabil (Tidak ada kenaikan beruntun).")

except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.info(f"User: HERY SUPRIYATNO\nAviation Reliability Engineer")
