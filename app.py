import streamlit as st
import pandas as pd

# 1. Judul Aplikasi
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi untuk Membaca File Excel Bapak
@st.cache_data
def load_data():
    # Pastikan nama file ini sama persis dengan yang Bapak upload ke GitHub
    file_name = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    # Membaca data dari Sheet 'DATA' (atau ganti sesuai nama sheet Bapak)
    df = pd.read_excel(file_name, sheet_name='DATA')
    return df

# 3. Menampilkan Data ke Layar
try:
    df = load_data()
    st.write("Daftar Komponen:")
    st.dataframe(df)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
