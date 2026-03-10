import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Data
@st.cache_data
def load_data(file_name, sheet_name):
    # Mengambil header dari baris ke-2 (index 1) agar judul kolom benar
    df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
    # Hapus baris & kolom yang 100% kosong
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    # Bersihkan nama kolom Unnamed agar tampilan rapi
    df.columns = ["" if "Unnamed" in str(col) else col for col in df.columns]
    return df

# 3. Alur Utama Aplikasi
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # Sidebar untuk pilih Sheet
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    st.markdown(f"### 📊 Laporan: {sheet_pilihan}")
    
    # Memuat data
    data = load_data(file_target, sheet_pilihan)
    
    # Fitur Pencarian
    search = st.text_input("🔍 Cari Part Number / Description:")
    
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_data = data[mask]
    else:
        display_data = data

    # Menampilkan Tabel
    st.dataframe(display_data, use_container_width=True, hide_index=True)

    # --- BAGIAN GRAFIK ---
    # Grafik ini akan otomatis muncul di bawah tabel
    st.markdown("---")
    st.subheader("📈 Visualisasi Tren Removal (Top 10)")

    try:
        # Cek apakah kolom yang dibutuhkan ada di sheet ini
        # Sesuaikan 'PART NUMBER' dan 'RATE' dengan judul kolom di Excel Bapak
        if 'PART NUMBER' in display_data.columns and 'RATE' in display_data.columns:
            # Ambil 10 data teratas untuk grafik
            chart_data = display_data.head(10)
            
            fig = px.bar(
                chart_data, 
                x='PART NUMBER', 
                y='RATE', 
                title=f"Top 10 Components on {sheet_pilihan}",
                color='RATE',
                labels={'RATE': 'Removal Rate', 'PART NUMBER': 'P/N'},
                color_continuous_scale='Reds',
                template='plotly_white'
            )
            
            # Menampilkan grafik di Streamlit
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("💡 Grafik otomatis akan muncul jika sheet ini memiliki kolom 'PART NUMBER' dan 'RATE'.")
            
    except Exception as e:
        st.warning(f"Tidak dapat membuat grafik pada sheet ini: {e}")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.info("Pastikan file Excel dan library sudah sesuai.")
