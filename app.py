import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Data Utama
@st.cache_data
def load_data(file_name, sheet_name):
    df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    df.columns = ["" if "Unnamed" in str(col) else col for col in df.columns]
    return df

# 3. Fungsi Load Khusus Sheet History (COMPONENT REPLACEMENT)
@st.cache_data
def load_history(file_name):
    # Membaca database history dari sheet sebelah
    try:
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=1)
        df_hist = df_hist.dropna(how='all', axis=0)
        return df_hist
    except:
        return pd.DataFrame() # Balikkan tabel kosong jika sheet tidak ketemu

# 4. Alur Utama
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    # Sidebar
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    st.markdown(f"### 📊 REPORT: TOP 10 HIGHEST REMOVAL RATE ({sheet_pilihan})")
    
    # Muat kedua data sekaligus
    data_utama = load_data(file_target, sheet_pilihan)
    data_history = load_history(file_target)
    
    # Fitur Pencarian
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        mask = data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_data = data_utama[mask]
    else:
        display_data = data_utama

    # --- BAGIAN TABEL INTERAKTIF ---
    st.info("💡 Klik pada baris tabel di bawah untuk melihat history dari sheet Replacement.")
    event = st.dataframe(
        display_data, 
        use_container_width=True, 
        hide_index=True,
        on_select="rerun",  
        selection_mode="single-row"
    )

    # --- LOGIKA DRILL-DOWN (Mencari ke sheet COMPONENT REPLACEMENT) ---
    if len(event.selection.rows) > 0:
        index_terpilih = event.selection.rows[0]
        row_data = display_data.iloc[index_terpilih]
        pn_terpilih = str(row_data['PART NUMBER'])
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn_terpilih}")
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.info(f"**Description:**\n\n{row_data['DESCRIPTION']}")
                st.metric("Total Qty Removal", f"{row_data['QTY REM']} EA")
            with col2:
                st.markdown("**📅 Records found in 'COMPONENT REPLACEMENT':**")
                # Filter data history berdasarkan P/N yang diklik
                detail_pn = data_history[data_history['PART NUMBER'].astype(str) == pn_terpilih]
                
                # Kolom yang ingin ditampilkan (Pastikan nama kolom sesuai dengan di Excel)
                cols_to_show = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                available = [c for c in cols_to_show if c in detail_pn.columns]
                
                if not detail_pn.empty:
                    st.table(detail_pn[available])
                else:
                    st.warning(f"Data P/N {pn_terpilih} tidak ditemukan di sheet COMPONENT REPLACEMENT.")

    # --- BAGIAN GRAFIK ---
    st.markdown("---")
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
    st.error(f"Terjadi kesalahan struktur: {e}")





