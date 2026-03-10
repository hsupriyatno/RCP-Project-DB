import streamlit as st
import pandas as pd
import plotly.express as px

# 1. KONFIGURASI HALAMAN & STYLE
st.set_page_config(page_title="Reliability Dashboard - Airfast", layout="wide")

# CSS untuk memaksa teks tabel rata tengah (Center)
st.markdown("""
    <style>
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th {
        text-align: center !important;
    }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNGSI LOAD DATA (DENGAN PENANGANAN DATA KOSONG/NaN)
@st.cache_data
def load_all_data(file_path, sheet_name):
    # Menggunakan skiprows untuk menghindari error 'range'
    df_m = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=4)
    df_m.columns = [str(c).strip().upper() for c in df_m.columns]
    
    # FIX: Tangani data kosong agar tidak error "float NaN to integer"
    df_m = df_m.fillna(0)
    
    # Ambil referensi periode
    df_ref = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1, header=None)
    bln_ref = str(df_ref.iloc[0, 0]).strip()
    thn_ref = str(df_ref.iloc[0, 1]).strip()
    
    # Load History (Sheet ke-3)
    df_h = pd.read_excel(file_path, sheet_name=2)
    df_h.columns = [str(c).strip().upper() for c in df_h.columns]
    df_h = df_h.fillna(0) # Tangani NaN di history juga
    
    if 'DATE' in df_h.columns:
        df_h['DATE'] = pd.to_datetime(df_h['DATE'], errors='coerce')
        df_h['DATE_STR'] = df_h['DATE'].dt.strftime('%d-%m-%Y')
        
    return df_m, df_h, bln_ref, thn_ref

def get_period_info(bln, thn):
    months = {
        "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
        "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12
    }
    try:
        m_idx = months.get(bln.upper(), 1)
        y_val = int(float(thn)) # Handle jika thn terbaca sebagai float
        return m_idx, y_val, f"{bln.capitalize()} {y_val}"
    except:
        return 1, 2026, "Unknown Period"

# 3. LOGIKA UTAMA
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    xls = pd.ExcelFile(FILE_PATH)
    st.sidebar.title("Navigation")
    
    # Key unik untuk mencegah Duplicate ID
    sheet_pilihan = st.sidebar.selectbox("Pilih Sheet Report:", xls.sheet_names, key="nav_select")
    
    df_main, df_history, bln_ref, thn_ref = load_all_data(FILE_PATH, sheet_pilihan)
    target_m, target_y, full_period = get_period_info(bln_ref, thn_ref)

    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Periode: {full_period}")

    # 4. CHART TOP 10
    if 'PART NUMBER' in df_main.columns and 'RATE' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10).copy()
        top_10['LABEL'] = top_10['PART NUMBER'].astype(str) + "<br>" + top_10['DESCRIPTION'].astype(str)
        
        fig = px.bar(top_10, x='LABEL', y='RATE', text_auto='.2f')
        fig.update_traces(marker_color='#F2B200', width=0.4) 
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. EXPLORER
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Search PN/Description:", key="search_comp")
    
    filtered = df_main.copy()
    if search:
        mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered = df_main[mask]

    event = st.dataframe(
        filtered, 
        use_container_width=True, 
        hide_index=True, 
        on_select="rerun", 
        selection_mode="single-row", 
        key="data_explorer"
    )

    # 6. DETAIL (RASIO 1:5:1:1)
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ DETAIL REMOVAL: {pn_selected}")
        
        # Metrik Ringkas
        m1, m2, m3 = st.columns([5, 1, 1])
        m1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        m2.metric("Rate", f"{row.get('RATE', 0):.2f}")
        m3.metric("Total Qty", f"{int(row.get('QTY REM', 0))} EA")

        if not df_history.empty:
            col_pn_h = next((c for c in df_history.columns if 'PART' in c.upper()), None)
            if col_pn_h:
                hist_match = df_history[
                    (df_history[col_pn_h].astype(str).str.strip() == pn_selected) & 
                    (df_history['DATE'].dt.month == target_m) & 
                    (df_history['DATE'].dt.year == target_y)
                ].copy()
                
                if not hist_match.empty:
                    # Tampilkan Tanpa REMARK
                    disp = ['DATE_STR', 'REASON OF REMOVAL', 'TSN', 'TSO']
                    existing = [c for c in disp if c in hist_match.columns]
                    
                    st.dataframe(
                        hist_match[existing], 
                        use_container_width=True, 
                        hide_index=True,
                        column_config={
                            "DATE_STR": st.column_config.Column("DATE", width="small"),
                            "REASON OF REMOVAL": st.column_config.Column(width="large"),
                            "TSN": st.column_config.Column(width="small"),
                            "TSO": st.column_config.Column(width="small")
                        }
                    )

# PENUTUP BLOK AGAR TIDAK SYNTAX ERROR
except Exception as e:
    st.error(f"Sistem Error: {e}")

st.sidebar.markdown("---")
st.sidebar.info("User: HERY SUPRIYATNO")
