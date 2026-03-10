import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability Dashboard Airfast", layout="wide")

def clean_df(df):
    """Membersihkan dataframe dari baris/kolom kosong dan spasi pada header"""
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = [str(col).strip() for col in df.columns]
    return df

# 2. Fungsi Load Data (Update nama sheet ke COMPONENT REPLACEMENT)
@st.cache_data
def load_all_data(file_name):
    try:
        # Load Kriteria Periode dari sheet utama
        df_crit = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None, nrows=3, usecols="A")
        bln_ref = str(df_crit.iloc[1, 0]).strip().upper()
        thn_ref = str(df_crit.iloc[2, 0]).strip().replace('.0', '')
        
        # Load Data Kalkulasi (Header di baris 2)
        df_main = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=1)
        df_main = clean_df(df_main)
        
        # FIX: Menggunakan nama sheet 'COMPONENT REPLACEMENT' sesuai koreksi Bapak
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        df_hist = clean_df(df_hist)
        
        return df_main, df_hist, bln_ref, thn_ref
    except Exception as e:
        st.error(f"Gagal memuat file: {e}")
        return pd.DataFrame(), pd.DataFrame(), "N/A", "N/A"

# --- EKSEKUSI DATA ---
FILE_NAME = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
df_main, df_hist, bulan_ref, tahun_ref = load_all_data(FILE_NAME)

# Mapping Bulan untuk Filter History
months_map = {'JANUARY':1,'FEBRUARY':2,'MARCH':3,'APRIL':4,'MAY':5,'JUNE':6,
              'JULY':7,'AUGUST':8,'SEPTEMBER':9,'OCTOBER':10,'NOVEMBER':11,'DECEMBER':12}

# --- INTERFACE ---
st.title("✈️ Aviation Reliability Dashboard")
st.markdown(f"**Reporting Period:** {bulan_ref} {tahun_ref}")

if not df_main.empty:
    # Visualisasi Top 10
    if 'RATE' in df_main.columns and 'PART NUMBER' in df_main.columns:
        df_main['RATE'] = pd.to_numeric(df_main['RATE'], errors='coerce').fillna(0)
        top_10 = df_main.sort_values(by='RATE', ascending=False).head(10)
        
        c1, c2 = st.columns([1, 1.2])
        with c1:
            st.subheader("Top 10 Removal Rates")
            st.dataframe(top_10[['PART NUMBER', 'DESCRIPTION', 'RATE']], hide_index=True)
        with c2:
            fig = px.bar(top_10, x='RATE', y='PART NUMBER', orientation='h', color='RATE',
                         color_continuous_scale='Reds', title="Highest Removal Rates Chart")
            st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # Komponen Explorer
    st.subheader("🔍 Component History Explorer")
    search = st.text_input("Ketik Part Number atau Nama Komponen:")
    
    filtered_df = df_main
    if search:
        filtered_df = df_main[df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
    
    # Tabel Utama (Klik baris untuk melihat detail)
    selection = st.dataframe(filtered_df, use_container_width=True, hide_index=True, 
                             on_select="rerun", selection_mode="single-row")

    # Tampilkan Detail jika baris dipilih
    if selection.selection.rows:
        selected_idx = selection.selection.rows[0]
        row_data = filtered_df.iloc[selected_idx]
        pn = str(row_data['PART NUMBER']).strip()
        
        st.info(f"### Maintenance Detail for P/N: {pn}")
        
        if not df_hist.empty:
            # Pastikan kolom DATE terbaca sebagai waktu
            if 'DATE' in df_hist.columns:
                df_hist['DATE'] = pd.to_datetime(df_hist['DATE'], errors='coerce')
            
            # Cari kolom P/N yang benar di sheet history
            col_pn_hist = 'PART NUMBER OFF' if 'PART NUMBER OFF' in df_hist.columns else 'PART NUMBER'
            
            if col_pn_hist in df_hist.columns:
                # Filter data history berdasarkan Part Number
                history_match = df_hist[df_hist[col_pn_hist].astype(str).str.strip() == pn]
                
                # Hanya tampilkan kolom yang ada di Excel Bapak
                target_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                available_cols = [c for c in target_cols if c in df_hist.columns]
                
                if not history_match.empty:
                    st.write(f"Ditemukan {len(history_match)} record di sheet COMPONENT REPLACEMENT:")
                    st.table(history_match[available_cols])
                else:
                    st.warning(f"Tidak ada data removal yang tercatat untuk P/N {pn}.")
            else:
                st.error("Kolom identifier (PART NUMBER OFF) tidak ditemukan di sheet history.")
