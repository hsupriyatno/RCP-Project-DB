import streamlit as st
import pandas as pd
import plotly.express as px

# ... (kode load data sebelumnya tetap sama)

    # TAMPILKAN TABEL
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        st.dataframe(data[mask], use_container_width=True, hide_index=True)
        display_data = data[mask] # Data yang difilter untuk grafik
    else:
        st.dataframe(data, use_container_width=True, hide_index=True)
        display_data = data # Data penuh untuk grafik

    # --- BAGIAN GRAFIK ---
    st.markdown("---") # Garis pembatas
    st.subheader("📈 Visualisasi Tren Removal")

    try:
        # Contoh: Membuat grafik Bar untuk 10 data teratas
        # Ganti 'PART NUMBER' dan 'RATE' sesuai nama kolom di Excel Bapak
        fig = px.bar(
            display_data.head(10), 
            x='PART NUMBER', 
            y='RATE', 
            title="Top 10 Components by Removal Rate",
            color='RATE',
            color_continuous_scale='Reds'
        )
        st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.info("Pilih sheet yang memiliki data numerik untuk menampilkan grafik.")







