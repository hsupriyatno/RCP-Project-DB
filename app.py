import streamlit as st
import pandas as pd

st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

@st.cache_data
def load_data(file_name, sheet_name):
    # Mengambil header dari baris ke-2 (index 1) agar judul kolom benar
    df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
    # Hapus baris & kolom yang 100% kosong
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    # Bersihkan nama kolom Unnamed
    df.columns = ["" if "Unnamed" in str(col) else col for col in df.columns]
    return df

try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    
    st.markdown(f"### 📊 Laporan: {sheet_pilihan}")
    
    data = load_data(file_target, sheet_pilihan)
    search = st.text_input("🔍 Cari Part Number / Description:")
    
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True, hide_index=True)
    else:
        st.dataframe(data, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
