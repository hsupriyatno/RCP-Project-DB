import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# 2. INJEKSI CSS GLOBAL (MENGUNCI RATA KIRI)
st.markdown("""
    <style>
        [data-testid="stDataFrame"] div[data-testid="stTable"] th { text-align: left !important; }
        [data-testid="stDataFrame"] div[data-testid="stTable"] td { text-align: left !important; }
        .stMarkdown, .stSubheader { text-align: left !important; }
        .left-metric { text-align: left; margin-bottom: 20px; }
        .left-metric p { margin: 0; color: gray; font-size: 14px; }
        .left-metric h2 { margin: 0; font-size: 28px; }
    </style>
""", unsafe_allow_html=True)

# 3. FUNGSI PEMBERSIH & LOAD DATA (DIPERBAIKI)
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

@st.cache_data
def load_all_data(file_name, sheet_name):
    try:
        xls_file = pd.ExcelFile(file_name)
        # Ambil Periode dari Sheet Kalkulasi Utama
        df_crit = pd.read_excel(xls_file, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load Data Utama untuk Explorer
        df_main = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
        df_main = clean_dynamic_columns(df_main)
        
        # Load History Replacement
        df_hist = pd.read_excel(xls_file, sheet_name="COMPONENT REPLACEMENT")
        df_hist.columns = [str(c).strip().upper() for c in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            df_hist['DATE_STR'] = df_hist['DATE'].dt.strftime('%d-%m-%Y')
            
        return df_main, df_hist, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Error Loading Data: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# --- MAIN APP START ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names)
    
    # Load semua variabel agar tidak ada NameError
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)

    if not df_main.empty:
        st.title(f"📊 Reliability Analysis: {sheet_pilihan}")

        # 4. CHART: TOP 10 REMOVAL RATE (KEMBALI DIMUNCULKAN)
        if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
            st.subheader(f"📈 Top 10 Removal Rate Comparison")
            top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
            top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + " - " + top_10['DESCRIPTION'].astype(str)
            
            fig = px.bar(top_10, x='RATE', y='LABEL', orientation='h', text_auto='.2f')
            fig.update_layout(yaxis={'categoryorder':'total ascending'}, height=400)
            st.plotly_chart(fig, use_container_width=True)

        st.divider()

        # 5. COMPONENT EXPLORER
        st.subheader("🔍 Component Explorer")
        search = st.text_input("Search Part Number or Description:")
        filtered = df_main.copy()
        if search:
            mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            filtered = df_main[mask]

        event = st.dataframe(filtered, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")

        # 6. DETAIL (DIPASTIKAN RATA KIRI)
        if event.selection.rows:
            selected_idx = event.selection.rows[0]
            row = filtered.iloc[selected_idx]
            pn_selected = str(row['PART NUMBER']).strip()
            
            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1:
                st.markdown(f"<div class='left-metric'><p>Description</p><h2>{row.get('DESCRIPTION', 'N/A')}</h2></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='left-metric'><p>Current Rate</p><h2>{row.get('RATE', 0):.2f}</h2></div>", unsafe_allow_html=True)
            with c3:
                st.markdown(f"<div class='left-metric'><p>Total Qty Rem</p><h2>{row.get('QTY REM', 0)} EA</h2></div>", unsafe_allow_html=True)

            # 7. HISTORY TABLE
            if not df_history.empty:
                col_pn_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
                if col_pn_h:
                    hist_match = df_history[df_history[col_pn_h].astype(str).str.strip() == pn_selected].copy()
                    if not hist_match.empty:
                        st.dataframe(hist_match[['DATE_STR', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']], use_container_width=True, hide_index=True)

    else:
        st.warning("Data pada sheet ini kosong atau format tidak sesuai.")

except Exception as e:
    st.error(f"Sistem Error: {e}")
