import streamlit as st
import pandas as pd

# Setting tampilan lebar agar tabel Excel enak dilihat di HP/Laptop
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")

st.title("✈️ Reliability Dashboard DHC6-300")

@st.cache_data
def load_data(sheet_name):
    # Nama file harus sama persis dengan yang di-upload ke GitHub
    file_name = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    # Membaca sheet tertentu, mulai dari baris ke-2 (skiprow=1) jika ada header di baris 1
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    return df

try:
    # Nama file asli Bapak
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    
    # Pilihan Sheet agar Bapak bisa pindah halaman di HP
    xls = pd.ExcelFile(file_target)
    sheet_pilihan = st.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    
    # Load data berdasarkan sheet yang dipilih
    data = load_data(sheet_pilihan)
    
    # Kolom Pencarian
    search = st.text_input(f"🔍 Cari data di {sheet_pilihan}:")
    
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True)
    else:
        st.dataframe(data, use_container_width=True)

except Exception as e:
    st.error(f"File belum terbaca: {e}")
    st.info("Pastikan file 'COMPONENT_RELIABILITY_DHC6-300.xlsm' sudah di-upload ke GitHub Bapak.")
