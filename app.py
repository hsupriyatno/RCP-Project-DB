import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# 1. PAGE CONFIG & CSS STYLING
st.set_page_config(page_title="Reliability Dashboard - Airfast", layout="wide")

# CSS untuk memaksa semua tabel rata tengah (Center)
st.markdown("""
    <style>
    [data-testid="stDataFrame"] td, 
    [data-testid="stDataFrame"] th {
        text-align: center !important;
    }
    .main {
        background-color: #f5f7f9;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. HELPER FUNCTIONS
def get_period_info(bln, thn):
    months = {
        "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
        "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12
    }
    m_idx = months.get(str(bln).upper(), 1)
    y_val = int(thn)
    return m_idx, y_val, f"{str(bln).capitalize()} {thn}"

@st.cache_data
def load_all_data(file_path, sheet_name):
    df_main = pd.read_excel(file_path, sheet_name=sheet_name, range="A5:Z500")
    df_main.columns = [str(c).strip().upper() for c in df_main.columns]
    
    # Ambil referensi bulan/tahun dari cell A1 & B1
    df_ref = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1, header=None)
    bln_ref = df_ref.iloc[0, 0]
    thn_ref = df_ref.iloc[0, 1]
    
    # Load History (Sheet 3)
    df_h = pd.read_excel(file_path, sheet_name=2)
    df_h.columns = [str(c).strip().upper() for c in df_h.columns]
    if 'DATE' in df_h.columns:
        df_h['DATE'] = pd.to_datetime(df_h['DATE'], errors='coerce')
        df_h['DATE_STR'] = df_h['DATE'].dt.strftime('%d-%m-%Y')
        
    return df_main, df_h, bln_ref, thn_ref

# 3. MAIN LOGIC
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    st.sidebar.title("Navigation")
    
    # Gunakan key unik agar tidak Duplicate ID
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names, key="main_sheet_selector")
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, full_period = get_period_info(bln_ref, thn_ref)

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Analysis Period: {full_period}")

    # 4. CHART (Yellow Ramping)
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f', title=f"Top 10 Removal Rate ({full_period})")
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. COMPONENT EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number atau Deskripsi:", key="comp_search")
    
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    # Tabel Utama
    event = st.dataframe(
        filtered, 
        use_container_width=True, 
        hide_index=True, 
        on_select="rerun", 
        selection_mode="single-row", 
        key="main_data_table"
    )

    # 6. PART REMOVAL DETAIL (Rasio 1:5:1:1 & Center)
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ PART REMOVAL DETAIL: {pn_selected}")
        
        # Metrik Atas
        m1, m2, m3 = st.columns([5, 1, 1])
        with m1:
            st.metric("Description", row.get('DESCRIPTION', 'N/A'))
        with m2:
            st.metric("Current Rate", f"{row.get('RATE', 0):.2f}")
        with m3:
            st.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        if not df_history.empty:
            col_pn_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
            if col_pn_h:
                hist_match = df_history[
                    (df_history[col_pn_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE'].dt.month == target_m) & 
                    (df_history['DATE'].dt.year == target_y)
                ].copy()
                
                if not hist_match.empty:
                    if 'DATE_STR' in hist_match.columns:
                        hist_match['DATE'] = hist_match['DATE_STR']
                    
                    # Kolom Tanpa REMARK
                    display_cols = ['DATE', 'REASON OF REMOVAL', 'TSN', 'TSO']
                    existing_cols = [c for c in display_cols if c in hist_match.columns]
                    
                    # Konfigurasi Lebar 1:5:1:1
                    # Note: 'alignment' dihapus karena menyebabkan error di versi Bapak
                    st.dataframe(
                        hist_match[existing_cols], 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "DATE": st.column_config.Column(width="small"),
                            "REASON OF REMOVAL": st.column_config.Column(width="large"),
                            "TSN": st.column_config.Column(width="small"),
                            "TSO": st.column_config.Column(width="small")
                        }
                    )
                else:
                    st.info(f"Tidak ada record removal untuk {pn_selected} pada {full_period}.")

except Exception as e:
    st.error(f"Terjadi kesalahan sistem: {e}")

st.sidebar.markdown("---")
st.sidebar.info(f"Logged in as: HERY SUPRIYATNO") # Nama sesuai identitas user
