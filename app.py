Mohon maaf, Pak Hery. Saya mengerti, maksud Bapak adalah menggabungkan fitur Uptrend yang baru tanpa menghapus fitur Top 10 Summary dan Detail Removal yang sudah berjalan baik sebelumnya.

Berikut adalah kode v1.9 yang menggabungkan semuanya secara utuh:

Top 10 Removal Summary (Tabel di bawah Chart).

Part Removal Detail (Muncul saat baris di Explorer diklik).

Uptrend Part Removal (Tabel Early Warning di paling bawah).

Akses Langsung Kolom I, L, O dari sheet REMOVAL RATE CALCULATION.

Python
import streamlit as st
import pandas as pd
import plotly.express as px

# 1. KONFIGURASI HALAMAN & CSS
st.set_page_config(page_title="Reliability Dashboard | Airfast Indonesia", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    [data-testid="stMetricLabel"] { font-size: 14px !important; }
    [data-testid="stMetricValue"] { font-size: 22px !important; }
    .stMetric { background-color: #ffffff; padding: 10px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    /* Agar chart stabil saat di-scroll jari di HP */
    .js-plotly-plot .plotly .nsewdrag { pointer-events: none !important; }
    .js-plotly-plot .plotly .hoverlayer { pointer-events: auto !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. FUNGSI LOAD DATA (MENGAMBIL I, L, O)
@st.cache_data
def load_reliability_data(file_name):
    try:
        df_raw = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=None)
        
        header_idx = 0
        for i, row in df_raw.iterrows():
            if 'PART NUMBER' in [str(x).upper() for x in row.values]:
                header_idx = i
                break
        
        df = pd.read_excel(file_name, sheet_name="REMOVAL RATE CALCULATION", header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Mapping Kolom Berdasarkan Index Excel (I=8, L=11, O=14)
        df['RATE_3MO'] = pd.to_numeric(df.iloc[:, 8], errors='coerce').fillna(0)  # Kolom I
        df['RATE_2MO'] = pd.to_numeric(df.iloc[:, 11], errors='coerce').fillna(0) # Kolom L
        df['RATE_1MO'] = pd.to_numeric(df.iloc[:, 14], errors='coerce').fillna(0) # Kolom O
        
        # Ambil kolom QTY REM dan DESCRIPTION jika ada
        if 'QTY REM' not in df.columns:
            df['QTY REM'] = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0) # Asumsi dekat kolom O

        return df
    except Exception as e:
        st.error(f"Gagal memuat data: {e}")
        return pd.DataFrame()

# --- MAIN APP ---
FILE_PATH = 'COMPONENT_RELIABILITY_DHC6-300.xlsm'

try:
    df_main = load_reliability_data(FILE_PATH)

    if not df_main.empty:
        st.title("📊 Reliability Analysis Dashboard")
        
        # 3. CHART TOP 10 (STABIL)
        st.subheader("📈 Top 10 Removal Rate (Current Month)")
        top_10 = df_main.sort_values(by='RATE_1MO', ascending=False).head(10).copy()
        
        fig = px.bar(top_10, x='PART NUMBER', y='RATE_1MO', text_auto='.2f', 
                     hover_data=['DESCRIPTION'], labels={'RATE_1MO': 'Rate'})
        fig.update_traces(marker_color='#F2B200')
        fig.update_layout(dragmode=False, xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        # --- TAMBAHAN: TABEL SUMMARY TOP 10 (FITUR LAMA YANG KEMBALI) ---
        with st.expander("📊 Click to View Top 10 Data Table Summary"):
            st.dataframe(
                top_10[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], 
                use_container_width=True, 
                hide_index=True
            )

        st.divider()

        # 4. COMPONENT EXPLORER
        st.subheader("🔍 Component Explorer")
        search = st.text_input("Cari Part Number atau Deskripsi:")
        
        filtered = df_main.copy()
        if search:
            mask = df_main.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            filtered = df_main[mask]

        event = st.dataframe(
            filtered[['PART NUMBER', 'DESCRIPTION', 'RATE_1MO']], 
            use_container_width=True, 
            hide_index=True, 
            on_select="rerun", 
            selection_mode="single-row"
        )

        # --- TAMBAHAN: PART REMOVAL DETAIL (FITUR LAMA YANG KEMBALI) ---
        if event.selection.rows:
            selected_row = filtered.iloc[event.selection.rows[0]]
            st.write("---")
            st.subheader(f"🛠️ PART REMOVAL DETAIL: {selected_row['PART NUMBER']}")
            
            c1, c2, c3 = st.columns([4, 1, 1])
            with c1: st.metric("Description", selected_row.get('DESCRIPTION', 'N/A'))
            with c2: st.metric("Current Rate (O)", f"{selected_row['RATE_1MO']:.2f}")
            with c3: st.metric("Qty Rem", f"{int(selected_row.get('QTY REM', 0))} EA")
            
            # Menampilkan tren 3 bulan untuk part yang dipilih
            st.write("**3-Month Rate Trend:**")
            st.dataframe(
                pd.DataFrame({
                    "Period": ["3 Months Ago (I)", "2 Months Ago (L)", "Current (O)"],
                    "Rate": [selected_row['RATE_3MO'], selected_row['RATE_2MO'], selected_row['RATE_1MO']]
                }), use_container_width=True, hide_index=True
            )

        st.divider()

        # 5. UPTREND PART REMOVAL (EARLY WARNING)
        st.subheader("⚠️ UPTREND PART REMOVAL (3-Month Increase)")
        uptrend_df = df_main[
            (df_main['RATE_3MO'] > 0) & 
            (df_main['RATE_2MO'] > df_main['RATE_3MO']) & 
            (df_main['RATE_1MO'] > df_main['RATE_2MO'])
        ].copy()

        if not uptrend_df.empty:
            st.warning(f"Ditemukan {len(uptrend_df)} komponen dengan tren kenaikan removal rate.")
            st.dataframe(
                uptrend_df[['PART NUMBER', 'DESCRIPTION', 'RATE_3MO', 'RATE_2MO', 'RATE_1MO']], 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "RATE_3MO": "Rate (I)",
                    "RATE_2MO": "Rate (L)",
                    "RATE_1MO": "Rate (O) 🚩"
                }
            )
        else:
            st.success("Tidak ada komponen yang mengalami kenaikan berturut-turut (Uptrend).")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")

st.sidebar.info(f"Aviation Reliability Dashboard\nUser: Hery Supriyatno")
