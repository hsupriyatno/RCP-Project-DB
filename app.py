import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Halaman & Judul
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

@st.cache_data
def load_data(file_name, sheet_name):
    df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = ["" if "Unnamed" in str(col) else col for col in df.columns]
    return df

try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # Sidebar
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    st.markdown(f"### 📊 REPORT: TOP 10 HIGHEST REMOVAL RATE ({sheet_pilihan})")
    
    # Load Data
    data = load_data(file_target, sheet_pilihan)
    
    # Pencarian
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        mask = data.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_data = data[mask]
    else:
        display_data = data

    # --- BAGIAN TABEL INTERAKTIF ---
    st.info("💡 Klik pada baris tabel untuk melihat detail history removal.")
    event = st.dataframe(
        display_data, 
        use_container_width=True, 
        hide_index=True,
        on_select="rerun",  
        selection_mode="single-row" # GANTI '_' MENJADI '-' DISINI
    )

    # --- LOGIKA DETAIL (POPUP) ---
    if len(event.selection.rows) > 0:
        index_terpilih = event.selection.rows[0]
        row_data = display_data.iloc[index_terpilih]
        pn_terpilih = row_data['PART NUMBER']
        
        st.markdown(f"### 🔍 Detailed History: {pn_terpilih}")
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.write(f"**Description:** {row_data['DESCRIPTION']}")
                st.metric("Total Removal", f"{row_data['QTY REM']} EA")
            with col2:
                st.markdown("**📅 History Record:**")
                history_cols = ['DATE', 'REASON OF REMOVAL', 'REMARK']
                available = [c for c in history_cols if c in display_data.columns]
                if available:
                    st.table(display_data[display_data['PART NUMBER'] == pn_terpilih][available])
                else:
                    st.write("Kolom detail history tidak ditemukan.")

    # --- BAGIAN GRAFIK ---
    st.markdown("---")
    st.subheader("📈 Visualisasi Tren Removal (Top 10)")
    
    if 'PART NUMBER' in display_data.columns and 'RATE' in display_data.columns:
        chart_data = display_data.head(10).copy()
        chart_data['Label'] = chart_data['PART NUMBER'].astype(str) + " - " + chart_data['DESCRIPTION'].astype(str)
        
        fig = px.bar(
            chart_data, x='Label', y='RATE', text='RATE',
            color='RATE', color_continuous_scale='Reds',
            height=600, template='plotly_white'
        )
        fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True, key=f"chart_{sheet_pilihan}")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.info("Pastikan file Excel sudah di-upload dan format kolom sesuai.")





