import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. FUNGSI PEMBERSIH DATA (Auto-detect Header)
def clean_dynamic_columns(df):
    for i in range(len(df)):
        row_values = [str(val).upper() for val in df.iloc[i].values]
        if 'PART NUMBER' in row_values:
            new_cols = df.iloc[i].values
            df = df.iloc[i+1:].copy()
            df.columns = [str(c).strip().upper() for c in new_cols]
            break
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    if 'RATE' in df.columns:
        df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
    return df

# 3. FUNGSI LOAD DATA (Format Tanggal dd-mm-yyyy)
@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        # Load Periode
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load Main Data
        df_main = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        # Load History & Format Tanggal
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            # Simpan versi string dd-mm-yyyy untuk tampilan tabel
            df_hist['DATE_STR'] = df_hist['DATE'].dt.strftime('%d-%m-%Y')
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

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
        return 11, 2025, "NOVEMBER"

# --- MAIN APP START ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    st.sidebar.title("Navigation")
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names)
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, target_m_name = get_period_info(bln_ref, thn_ref)
    full_period = f"{target_m_name} {target_y}"

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")

    # 4. CHART DASHBOARD (Top 10 Rate)
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        st.subheader(f"📈 Top 10 Removal Rate Comparison ({full_period})")
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        fig.update_layout(xaxis_title="PART NUMBER & DESCRIPTION", yaxis_title="RATE", bargap=0.5)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Search Part Number or Description:")
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    # Menghilangkan kolom paling kiri (hide_index=True)
    event = st.dataframe(filtered, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

    # 6. DETAIL VIEW (Proporsi 1:5:1:1 & No Remark)
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        # Judul Center
        st.markdown(f"<h3 style='text-align: center;'>🛠️ PART REMOVAL DETAIL: {pn_selected}</h3>", unsafe_allow_html=True)
        
        # Proporsi Kolom 1:5:1:1
        _, c1, c2, c3 = st.columns([1, 5, 1, 1]) 
        c1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        c2.metric("Current Rate", f"{row.get('RATE', 0):.2f}")
        c3.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        if not df_history.empty:
            col_pn_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
            if col_pn_h:
                hist_match = df_history[
                    (df_history[col_pn_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE'].dt.month == target_m) & 
                    (df_history['DATE'].dt.year == target_y)
                ].copy()
                
                if not hist_match.empty:
                    # Tampilkan Tanggal Format dd-mm-yyyy
                    if 'DATE_STR' in hist_match.columns:
                        hist_match['DATE'] = hist_match['DATE_STR']
                    
                    # Remark Dihilangkan sesuai permintaan
                    potential_cols = ['DATE', 'REASON OF REMOVAL', 'TSN', 'TSO']
                    existing_cols = [c for c in potential_cols if c in hist_match.columns]
                    
                    st.dataframe(hist_match[existing_cols], use_container_width=True, hide_index=True)
                else:
                    st.info(f"Tidak ada record removal untuk {pn_selected} pada {full_period}.")
            else:
                st.warning("Kolom Part Number tidak ditemukan di sheet history.")

except Exception as e:
    st.error(f"Sistem Error: {e}")
