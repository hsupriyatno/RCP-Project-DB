import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Standar Profesional
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

# Custom CSS untuk tampilan lebih rapi
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# 2. Fungsi Load Data (Optimized)
@st.cache_data
def load_reliability_data(file_name, sheet_name):
    try:
        # Load kriteria periode dari sheet utama (A2 & A3)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_raw = str(df_crit.iloc[1, 0]).strip().upper()
        thn_raw = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load data tabel utama
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() for col in df.columns]
        
        if 'RATE' in df.columns:
            df['RATE'] = pd.to_numeric(df['RATE'], errors='coerce').fillna(0)
            
        return df, bln_raw, thn_raw
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame(), "N/A", "N/A"

@st.cache_data
def load_replacement_history(file_name):
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        if 'DATE' in df_hist.columns:
            df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
        return df_hist
    except:
        return pd.DataFrame()

# 3. Logika Filter Periode (Mundur 1 Bulan)
def get_previous_month(bulan, tahun):
    months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
                  'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
    try:
        m_num = months_map.get(bulan, 12)
        current_dt = datetime(int(tahun), m_num, 1)
        prev_dt = current_dt - timedelta(days=1)
        return prev_dt.month, prev_dt.year, [k for k, v in months_map.items() if v == prev_dt.month][0]
    except:
        return 11, 2025, "NOVEMBER"

# --- MAIN APP ---
try:
    FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(FILE_PATH)
    
    # Sidebar Navigation
    st.sidebar.image("https://www.airfastindonesia.com/images/logo.png", width=150) # Opsional: Link logo Airfast
    st.sidebar.title("Navigation")
    sheet_pilihan = st.sidebar.selectbox("Select Report Sheet:", xls.sheet_names)
    
    # Load Data & Criteria
    df_main, bln_ref, thn_ref = load_reliability_data(FILE_PATH, sheet_pilihan)
    target_m_num, target_y, target_m_name = get_previous_month(bln_ref, thn_ref)
    df_history = load_replacement_history(FILE_PATH)

    # Header Dashboard
    st.title(f"📊 Reliability Analysis: {sheet_pilihan}")
    st.caption(f"Reporting Period: {bln_ref} {thn_ref} | Analysis Data: {target_m_name} {target_y}")

    # 4. SECTION: HIGHLIGHT TOP 10 (Hanya muncul jika ada kolom RATE)
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.subheader(f"🏆 Top 10 Highest Removal Rate")
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], use_container_width=True, hide_index=True)
            
        with col_b:
            fig = px.bar(top_10, x='RATE', y='PART NUMBER', orientation='h', 
                         color='RATE', color_continuous_scale='Reds',
                         title=f"Removal Rate Comparison ({target_m_name})")
            fig.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 5. SECTION: EXPLORER & DRILL-DOWN
    st.subheader("🔍 Component Explorer")
    search_query = st.text_input("Search by Part Number or Description", placeholder="e.g. 300-2000-01")
    
    filtered_df = df_main.copy()
    if search_query:
        filtered_df = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)]

    # Main Data Table
    event = st.dataframe(
        filtered_df, 
        use_container_width=True, 
        hide_index=True, 
        on_select="rerun", 
        selection_mode="single-row"
    )

    # Detail View saat diklik
    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        row = filtered_df.iloc[selected_idx]
        pn_selected = str(row['PART NUMBER']).strip()
        
        st.write("---")
        st.subheader(f"🛠️ Maintenance Detail: {pn_selected}")
        
        # Dashboard Mini untuk Part Terpilih
        m1, m2, m3 = st.columns(3)
        m1.metric("Description", row.get('DESCRIPTION', 'N/A'))
        m2.metric("Current Rate", f"{row.get('RATE', 0):.4f}")
        m3.metric("Total Qty Rem", f"{row.get('QTY REM', 0)} EA")

        # History Table (Filtered by Month-1)
        if not df_history.empty and 'PART NUMBER OFF' in df_history.columns:
            hist_match = df_history[
                (df_history['PART NUMBER OFF'].astype(str).str.strip() == pn_selected) &
                (df_history['DATE'].dt.month == target_m_num) &
                (df_history['DATE'].dt.year == target_y)
            ]
            
            if not hist_match.empty:
                st.table(hist_match[['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']])
            else:
                st.warning(f"No removal records found in {target_m_name} {target_y} for this part.")

except Exception as e:
    st.error(f"Critical System Error: {e}")
