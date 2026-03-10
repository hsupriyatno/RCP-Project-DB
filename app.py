import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Konfigurasi Dasar
st.set_page_config(page_title="Reliability DHC6-300", layout="wide")
st.title("✈️ Reliability Dashboard DHC6-300")

# 2. Fungsi Load Data Utama (Sheet Laporan)
@st.cache_data
def load_data(file_name, sheet_name):
    try:
        # Sheet laporan tetap menggunakan header=1 jika judulnya di baris ke-2
        df = pd.read_excel(file_name, sheet_name=sheet_name, header=1)
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        df.columns = [str(col).strip() if "Unnamed" not in str(col) else "" for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal memuat sheet {sheet_name}: {e}")
        return pd.DataFrame()

# 3. Fungsi Load History (Sheet COMPONENT REPLACEMENT)
@st.cache_data
def load_history(file_name):
    try:
        # DISESUAIKAN: header=0 karena judul 'PART NUMBER OFF' ada di baris ke-1
        df_hist = pd.read_excel(file_name, sheet_name="COMPONENT REPLACEMENT", header=0)
        # Bersihkan spasi di nama kolom
        df_hist.columns = [str(col).strip() for col in df_hist.columns]
        return df_hist
    except Exception as e:
        return pd.DataFrame()

# 4. Alur Utama Dashboard
try:
    file_target = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'
    xls = pd.ExcelFile(file_target)
    
    sheet_pilihan = st.sidebar.selectbox("Pilih Halaman (Sheet):", xls.sheet_names)
    st.markdown(f"### 📊 REPORT: TOP 10 HIGHEST REMOVAL RATE ({sheet_pilihan})")
    
    data_utama = load_data(file_target, sheet_pilihan)
    data_history = load_history(file_target)
    
    search = st.text_input("🔍 Cari Part Number / Description:")
    if search:
        mask = data_utama.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_data = data_utama[mask]
    else:
        display_data = data_utama

    st.info("💡 Klik baris tabel untuk melihat detail history dari COMPONENT REPLACEMENT.")
    event = st.dataframe(
        display_data, 
        use_container_width=True, 
        hide_index=True,
        on_select="rerun",  
        selection_mode="single-row"
    )

    # --- LOGIKA DETAIL (DRILL-DOWN) ---
    if event.selection.rows:
        index_terpilih = event.selection.rows[0]
        row_data = display_data.iloc[index_terpilih]
        pn_terpilih = str(row_data['PART NUMBER']).strip()
        
        st.markdown(f"### 🔍 Detailed History for P/N: {pn_terpilih}")
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.info(f"**Description:**\n\n{row_data.get('DESCRIPTION', 'N/A')}")
                st.metric("Total Qty Removal", f"{row_data.get('QTY REM', 0)} EA")
            with col2:
                # Kolom target di sheet history (Kolom D)
                target_col = 'PART NUMBER OFF'
                
                if target_col in data_history.columns:
                    # Filter history berdasarkan P/N
                    detail_pn = data_history[data_history[target_col].astype(str).strip() == pn_terpilih]
                    
                    # Kolom yang ingin ditampilkan
                    cols_to_show = ['DATE', 'REASON OF REMOVAL', 'REMARK', 'TSN', 'TSO']
                    available = [c for c in cols_to_show if c in data_history.columns]
                    
                    if not detail_pn.empty:
                        st.table(detail_pn[available])
                    else:
                        st.warning(f"Tidak ditemukan history untuk P/N {pn_terpilih} di kolom {target_col}.")
                else:
                    # Diagnostik: Menunjukkan judul kolom yang terbaca oleh Python
                    st.error(f"Kolom '{target_col}' tidak ditemukan di baris pertama.")
                    st.write("Kolom yang terbaca:", list(data_history.columns))

    # --- GRAFIK ---
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
    st.error(f"Terjadi kesalahan: {e}")
