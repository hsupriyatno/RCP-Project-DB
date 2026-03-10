import streamlit as st
import pandas as pd

st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

@st.cache_data
def load_data(file_name, sheet_name):
    # Membaca data, kita asumsikan data mulai dari baris yang ada isinya
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    # Menghapus kolom/baris yang benar-benar kosong
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    return df

try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # Letakkan pilihan sheet di Sidebar (Samping) agar layar utama luas
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    
    data = load_data(file_target, sheet_pilihan)
    
    st.subheader(f"📊 Data Sheet: {sheet_pilihan}")
    
    # Kolom Pencarian
    search = st.text_input(f"🔍 Cari Part Number / Description:")
    
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True, hide_index=True)
    else:
        st.dataframe(data, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Pilih sheet yang valid atau cek isi file: {e}")
