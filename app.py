import streamlit as st
import pandas as pd

# Konfigurasi tampilan
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")

st.title("✈️ Reliability Dashboard DHC6-300")


try:
    # Nama file harus sama persis dengan yang ada di GitHub
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # Pilih Sheet di Sidebar kiri
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    
    data = load_data(file_target, sheet_pilihan)
    
    st.subheader(f"📊 Data Sheet: {sheet_pilihan}")
    
    # Fitur Pencarian Part Number atau Description
    search = st.text_input("🔍 Cari Part Number / Description:")
    
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True, hide_index=True)
    else:
        st.dataframe(data, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.info("Pastikan file 'COMPONENT_RELIABILITY_DHC6-300.xlsm' sudah di-upload ke GitHub.")





