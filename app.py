import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi & CSS (Huruf Metrik Kecil)
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    [data-testid="stMetricLabel"] { font-size: 14px !important; }
    [data-testid="stMetricValue"] { font-size: 22px !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

def clean_dynamic_columns(df):
    """Memperbaiki kolom 'Unnamed' dengan mencari baris header yang benar secara otomatis"""
    # Cari baris yang mengandung kata 'PART NUMBER'
    for i in range(len(df)):
        row_values = [str(val).upper() for val in df.iloc[i].values]
        if 'PART NUMBER' in row_values:
            # Jadikan baris ini sebagai header
            new_cols = df.iloc[i].values
            df = df.iloc[i+1:].copy()
            df.columns = [str(c).strip().upper() for c in new_cols]
            break
    
    # Hapus baris/kolom kosong
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    # Pastikan kolom RATE adalah angka
    if 'RATE' in df.columns:
        df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
    return df

# 2. Fungsi Load Data
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        # Load Periode (A2 & A3)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load Main Data (Tanpa header dulu agar bisa dibersihkan secara dinamis)
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        # Load History
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        # Pastikan data di-load meskipun header belum pas (Unnamed:0, Unnamed:1, dll)
        # Kami perbaiki nanti saat filter dinamis di Bagian 6
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Error Loading Data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# 3. Logika Bulan
def get_period_info(bulan, tahun):
    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    m_num = m_map.get(bulan, 12)
    curr = datetime(int(tahun) if str(tahun).isdigit() else 2026, m_num, 1)
    prev = curr - timedelta(days=1)
    p_name = [k for k, v in m_map.items() if v == prev.month][0]
    return prev.month, prev.year, p_name

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
try:
    xls = pd.ExcelFile(FILE_PATH)
    st.sidebar.title("Navigation")
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet:", xls.sheet_names, index=xls.sheet_names.index("REMOVAL RATE CALCULATION") if "REMOVAL RATE CALCULATION" in xls.sheet_names else 0)
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, target_m_name = get_period_info(bln_ref, thn_ref)
    full_period = f"{target_m_name} {target_y}"

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Excel Period: {bln_ref} {thn_ref} | Displaying: {full_period}")

# 4. CHART (Final: Ramping, Kuning Solid, 2 Desimal, & Label Ganda)
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        
        # Gabungkan P/N dan Deskripsi (menggunakan <br> untuk baris baru di chart)
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        
        # Pengaturan ketebalan batang dan warna solid Airfast
        fig.update_traces(
            marker_color='#F2B200', 
            width=0.4 # Batang lebih ramping
        ) 
        
        fig.update_layout(
            xaxis_title="PART NUMBER & DESCRIPTION",
            yaxis_title="RATE",
            xaxis_tickangle=-45,
            coloraxis_showscale=False,
            showlegend=False,
            bargap=0.5 # Jarak antar batang lebih lebar agar terlihat elegan
        )
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. DETAIL (Fix: Solusi untuk 'not in index' error)
    if event.selection.rows:
        row = filtered.iloc[event.selection.rows[0]]
        pn = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn}")
        
        # Metric Cards (Huruf Kecil via CSS)
        c1, c2, c3 = st.columns(3)
        c1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        c2.metric("Current Rate", f"{row.get('RATE', 0):.2f}")
        c3.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        if not df_history.empty:
            # Cari kolom P/N di history secara fleksibel
            col_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
            
            if col_h:
                match = df_history[(df_history[col_h].astype(str).str.strip() == pn) & 
                                   (df_history['DATE'].dt.month == target_m) & 
                                   (df_history['DATE'].dt.year == target_y)]
                
                # FIX: Hanya tampilkan kolom yang benar-benar ada di Excel untuk menghindari error 'not in index'
                potential_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                existing_cols = [c for c in potential_cols if c in df_history.columns]
                
                if not match.empty:
                    st.table(match[existing_cols])
                else:
                    st.info(f"No removal records for {pn} in {full_period}.")
