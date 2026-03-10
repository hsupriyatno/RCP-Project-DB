import streamlit as st
import pandas as pd

st.set_page_config(page_title="Database Alamat", layout="wide")
st.title("📱 Database Alamat")

@st.cache_data
def load_data():
    # Membaca file
    df = pd.read_csv('data.csv', skiprows=1)
    
    # 1. Menghapus kolom pertama yang kosong (Unnamed: 0)
    df = df.iloc[:, 1:]
    
    # 2. Menghapus baris yang semuanya kosong
    df = df.dropna(how='all')
    
    # 3. Memastikan NIK dibaca sebagai teks (bukan angka ilmiah)
    if 'NIK' in df.columns:
        df['NIK'] = df['NIK'].astype(str).str.split('.').str[0]
        
    return df

try:
    data = load_data()

    # Kotak Pencarian
    search = st.text_input("🔍 Cari berdasarkan Nama atau NIK:")
    
    if search:
        # Mencari di semua kolom
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        filtered_df = data[mask]
        st.success(f"Ditemukan {len(filtered_df)} data")
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    else:
        st.info("Daftar Alamat Lengkap:")
        # Menggunakan dataframe agar bisa di-scroll jika datanya banyak
        st.dataframe(data, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")