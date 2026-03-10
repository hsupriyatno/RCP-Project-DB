import streamlit as st
import pandas as pd

# Konfigurasi tampilan
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")

st.title("✈️ Reliability Dashboard DHC6-300")

@st.cache_data
def load_data(file_name, sheet_name):
    # 1. Membaca Excel tanpa menghapus apapun dulu
    # Kita gunakan header=0 agar Python mencoba mengambil baris pertama sebagai judul
    df = pd.read_excel(file_name, sheet_name=sheet_name)
    
    # 2. Hapus baris yang 100% kosong (sering ada di paling bawah atau atas)
    df = df.dropna(how='all', axis=0)
    
    # 3. Hapus kolom yang 100% kosong (kolom hantu di sebelah kanan)
    df = df.dropna(how='all', axis=1)
    
    # 4. Merapikan nama kolom 'Unnamed' agar tidak merusak tampilan, 
    # tapi tetap mempertahankan datanya agar tidak hilang.
    new_columns = []
    for i, col in enumerate(df.columns):
        if "Unnamed" in str(col):
            new_columns.append(f"Col_{i}") # Beri nama sementara agar data tetap muncul
        else:
            new_columns.append(col)
    df.columns = new_columns
    
    return df
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



