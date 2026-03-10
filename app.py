Tenang Pak Hery, saya mengerti. Memang lebih mudah kalau langsung copy-paste kode utuh yang sudah diperbaiki semua bagian error-nya (termasuk menghapus kolom index di kiri).

Berikut adalah kode komplit (Full Code) yang sudah saya perbaiki berdasarkan screenshot error yang Bapak kirimkan tadi:

Menghapus Kolom Index: Ditambahkan hide_index=True di semua tabel.

Fix Error 'DATE': Menambahkan pengecekan kolom sebelum memproses data history agar tidak crash.

Fix Error 'range' & 'NaN': Membersihkan data dari nilai kosong (NaN) agar grafik dan perhitungan tidak error.

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
    .stMetric { 
        background-color: #ffffff; 
        padding: 10px; 
        border-radius: 10px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
    }
    [data-testid="stDataFrame"] div[data-testid="stTable"] th { text-align: left !important; }
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
    # Hapus kolom/baris yang benar-benar kosong
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    # Ganti NaN dengan 0 agar tidak error saat kalkulasi
    df = df.fillna(0) 
    if 'RATE' in df.columns:
        df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
    if 'QTY REM' in df.columns:
        df['QTY REM'] = pd.to_numeric(df['QTY REM'], errors='coerce').fillna(0)
    return df

# 3. FUNGSI LOAD DATA
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        # Load period dari sheet utama
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().split('.')[0] # Menangani '2026.0'
        
        # Load data report
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        # Load history removal
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        
        # Proteksi kolom DATE
        if 'DATE' in df_hist.columns:
            df_hist['DATE_DT'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            df_hist['DATE_STR'] = df_hist['DATE_DT'].dt.strftime('%d-%m-%Y')
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# 4. LOGIKA PERIODE (Bulan Analisis = Bulan Report - 1)
def get_period_info(bulan, tahun):
    m_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
             'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    m_num = m_map.get(bulan, 12)
    try:
        curr = datetime(int(tahun), m_num, 1)
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
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        fig.update_layout(xaxis_title="PN & DESC", yaxis_title="RATE", xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📊 View Top 10 Data (Tanpa Index)"):
            # Menghapus index di sini
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'QTY REM', 'RATE']], use_container_width=True, hide_index=True)

    st.divider()

    # 6. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:")
    
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    # Menghapus index di tabel utama
    event = st.dataframe(filtered, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # 7. PART REMOVAL DETAIL
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
        
        m1, m2, m3 = st.columns([5, 1, 1])
        with m1: st.metric("Description", row.get('DESCRIPTION', 'N/A'))
        with m2: st.metric("Current Rate", f"{row.get('RATE', 0):.2f}")
        with m3: st.metric("Total Qty Rem", f"{int(row.get('QTY REM', 0))} EA")

        if not df_history.empty:
            col_pn_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
            
            # Cek kolom DATE_DT untuk menghindari error 'DATE'
            if col_pn_h and 'DATE_DT' in df_history.columns:
                hist_match = df_history[
                    (df_history[col_pn_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE_DT'].dt.month == target_m) & 
                    (df_history['DATE_DT'].dt.year == target_y)
                ].copy()
                
                if not hist_match.empty:
                    hist_match['DATE'] = hist_match['DATE_STR']
                    cols_to_show = [c for c in ['DATE', 'REASON OF REMOVAL', 'TSN', 'TSO'] if c in hist_match.columns]
                    
                    # Menghapus index di tabel detail
                    st.dataframe(
                        hist_match[cols_to_show], 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "DATE": st.column_config.Column(width="small"),
                            "REASON OF REMOVAL": st.column_config.Column(width="large")
                        }
                    )
                else:
                    st.info(f"Tidak ada record removal untuk {pn_selected} pada {full_period}.")
            else:
                st.warning("Kolom 'DATE' tidak ditemukan di sheet History.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")

st.sidebar.markdown("---")
st.sidebar.info(f"Logged in as: {st.session_state.get('user', 'HERY SUPRIYATNO')}")
