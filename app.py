import streamlit as st
import pandas as pd

# Konfigurasi halaman agar tampilan lebar (bagus untuk tabel besar)
st.set_page_config(page_title="Reliability Dashboard DHC6", layout="wide")

st.title("✈️ Aircraft Component Reliability Dashboard")
st.subheader("DHC6-300 Maintenance Monitoring")

@st.cache_data
def load_data():
    # Pastikan nama file di bawah ini sama persis dengan yang Bapak upload ke GitHub
    file_name = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    
    # Membaca Sheet 'REPORT' (asumsi ini adalah ringkasan yang ingin dipantau)
    # Jika ingin sheet lain, ganti sheet_name-nya
    df = pd.read_excel(file_name, sheet_name='REPORT', skiprows=1)
    
    # Membersihkan data yang kosong
    df = df.dropna(how='all', axis=0)
    return df

try:
    data = load_data()
    
    # Fitur Pencarian berdasarkan Part Number atau Description
    search = st.text_input("🔍 Cari Part Number atau Description Component:")
    
    if search:
        # Mencari di semua kolom
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True)
    else:
        st.write("Menampilkan semua data komponen:")
        st.dataframe(data, use_container_width=True)

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.info("Pastikan file 'COMPONENT_RELIABILITY_DHC6-300.xlsm' sudah di-upload ke GitHub.")
