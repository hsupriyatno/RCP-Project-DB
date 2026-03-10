import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Dasar
st.set_page_config(page_title="Reliability Airfast DHC6-300", layout="wide")

# Fungsi untuk membaca header yang dinamis (mengatasi 'Unnamed' columns)
def clean_df(df):
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = [str(col).strip() for col in df.columns]
    return df

# 2. Fungsi Load Data Utama
@st.cache_data
def load_all_data(file_name):
    try:
        # Load Kriteria Periode (A2 & A3)
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_ref = str(df_crit.iloc[1, 0]).strip().upper()
        thn_ref = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load Data Kalkulasi (Header di Baris 2 Excel -> header=1)
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=1)
        df_main = clean_df(df_main)
        
        # Load Data History (Header di Baris 1 Excel -> header=0)
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist = clean_df(df_hist)
        
        return df_main, df_hist, bln_ref, thn_ref
    except Exception as e:
        st.error(f"Error Loading File: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# --- PROSES DATA ---
FILE_NAME = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
df_main, df_hist, bulan_ref, tahun_ref = load_all_data(FILE_NAME)

# Logika Bulan (Mundur 1 Bulan untuk History)
months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
              'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}
m_num = months_map.get(bulan_ref, 12)
current_date = datetime(int(tahun_ref) if tahun_ref != "N/A" else 2026, m_num, 1)
prev_date = current_date - timedelta(days=1)
target_m_num = prev_date.month
target_y = prev_date.year
target_m_name = [k for k, v in months_map.items() if v == target_m_num][0]

# --- TAMPILAN DASHBOARD ---
st.title("✈️ Reliability Analysis Dashboard")
st.info(f"Analysis Focus: **{target_m_name} {target_y}** (Based on {bulan_ref} {tahun_ref} Report)")

if not df_main.empty:
    # A. Top 10 Section
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        df_main['RATE'] = pd.to_numeric(df_main['RATE'], errors='coerce').fillna(0)
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        col1, col2 = st.columns([1, 1.2])
        with col1:
            st.subheader("🏆 Top 10 Removal Rate")
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], hide_index=True)
        with col2:
            fig = px.bar(top_10, x='RATE', y='PART NUMBER', orientation='h', 
                         title="Chart: Highest Removal Rates", color='RATE', color_continuous_scale='Reds')
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # B. Explorer Section
    st.subheader("🔍 Component Explorer")
    search = st.text_input("Cari Part Number:")
    
    filtered_df = df_main
    if search:
        filtered_df = df_main[df_main['PART NUMBER'].astype(str).str.contains(search, case=False)]
    
    # Selection Table
    selection = st.dataframe(filtered_df, use_container_width=True, hide_index=True, 
                             on_select="rerun", selection_mode="single-row")

    # C. Detail History (Muncul Saat Baris Diklik)
    if selection.selection.rows:
        selected_row = filtered_df.iloc[selection.selection.rows[0]]
        pn = str(selected_row['PART NUMBER']).strip()
        
        st.markdown(f"### 🛠️ Maintenance Details: {pn}")
        
        # Cek ketersediaan kolom di History secara aman
        if not df_hist.empty:
            # Pastikan kolom DATE valid
            if 'DATE' in df_hist.columns:
                df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
            # Cari kolom P/N yang sesuai (mungkin 'PART NUMBER OFF' atau 'PART NUMBER')
            col_pn_hist = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_hist.columns else 'PART NUMBER'
            
            if col_pn_hist in df_hist.columns:
                history_match = df_hist[
                    (df_hist[col_pn_hist].astype(str).str.strip() == pn) &
                    (df_hist['DATE'].dt.month == target_m_num) &
                    (df_hist['DATE'].dt.year == target_y)
                ]
                
                # Menampilkan kolom yang ada saja (menghindari 'not in index' error)
                cols_to_show = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                available_cols = [c for c in cols_to_show if c in df_hist.columns]
                
                if not history_match.empty:
                    st.table(history_match[available_cols])
                else:
                    st.warning(f"Tidak ada catatan removal untuk P/N {pn} pada {target_m_name} {target_y}.")
            else:
                st.error("Kolom 'PART NUMBER OFF' tidak ditemukan di sheet Replacement.")
