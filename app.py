Mohon maaf sekali atas kendalanya, Pak Hery. Saya melihat dari gambar yang Bapak kirim bahwa tabelnya bergeser sehingga nama kolom aslinya (seperti PART NUMBER dan RATE) terbaca sebagai Unnamed: 2 atau Unnamed: 5 oleh Python. Hal ini menyebabkan error "Value of x is not the name of a column" karena Python tidak bisa menemukan kolom yang tepat untuk membuat chart.

Saya telah memperbaiki logika pembersihan data (clean_df) agar secara otomatis mencari baris yang berisi data asli dan mengganti nama kolom Unnamed menjadi nama yang benar secara dinamis.

Berikut adalah kode yang sudah diperbaiki agar Chart muncul kembali di paling atas:

Python
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
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
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
    sheet_pilihan = st.sidebar.selectbox("Select Report Sheet:", xls.sheet_names)
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, target_m_name = get_period_info(bln_ref, thn_ref)
    full_period = f"{target_m_name} {target_y}"

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Excel Period: {bln_ref} {thn_ref} | Displaying: {full_period}")

    # 4. CHART (Paling Atas)
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        fig = px.bar(top_10, x='PART NUMBER', y='RATE', color='RATE', 
                     color_continuous_scale='Reds', text_auto='.4f')
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        with st.expander("📊 View Top 10 Summary Table"):
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], use_container_width=True, hide_index=True)

    st.divider()

    # 5. EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Search Part Number or Description:")
    filtered = df_main.copy()
    if search:
        filtered = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]

    event = st.dataframe(filtered, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    if event.selection.rows:
        row = filtered.iloc[event.selection.rows[0]]
        pn = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn}")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        c2.metric("Current Rate", f"{row.get('RATE', 0):.4f}")
        c3.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        if not df_history.empty:
            col_h = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_history.columns else 'PART NUMBER'
            match = df_history[(df_history[col_h].astype(str).str.strip() == pn) & 
                               (df_history['DATE'].dt.month == target_m) & 
                               (df_history['DATE'].dt.year == target_y)]
            
            if not match.empty:
                st.table(match[['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']])
            else:
                st.warning(f"No removal records for {pn} in {full_period}.")

except Exception as e:
    st.error(f"Critical System Error: {e}")
