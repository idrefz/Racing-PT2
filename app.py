import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import shutil

# Konfigurasi
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# Fungsi Helper
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(file_content):
    return hashlib.md5(file_content).hexdigest()

def compare_with_previous(current_df):
    delta_dict = {}
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if len(history_df) >= 2:
            previous_hash = history_df.iloc[-2]['file_hash']
            previous_file = os.path.join(DATA_FOLDER, f"previous_{previous_hash}.xlsx")
            if os.path.exists(previous_file):
                previous_df = pd.read_excel(previous_file)
                current_golive = current_df[current_df['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                previous_golive = previous_df[previous_df['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                for witel in current_golive.index:
                    delta_dict[witel] = current_golive[witel] - previous_golive.get(witel,0)
    return delta_dict

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_hash = get_file_hash(open(LATEST_FILE, "rb").read()) if os.path.exists(LATEST_FILE) else ""
    history_df = pd.DataFrame(columns=['timestamp', 'file_hash'])
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty:
            previous_hash = history_df.iloc[-1]['file_hash']
            previous_file = os.path.join(DATA_FOLDER, f"previous_{previous_hash}.xlsx")
            if os.path.exists(LATEST_FILE):
                shutil.copy(LATEST_FILE, previous_file)
    history_df = pd.concat([
        history_df,
        pd.DataFrame([[now, file_hash]], columns=['timestamp', 'file_hash'])
    ], ignore_index=True)
    history_df.to_csv(HISTORY_FILE, index=False)

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty:
            return history_df.iloc[-1]['timestamp'], history_df.iloc[-1]['file_hash']
    return None, None

def validate_data(df):
    required = ['Regional', 'Witel', 'Status Proyek', 'Total Port', 'Datel', 'Ticket ID', 'Nama Proyek']
    missing = [col for col in required if col not in df.columns]
    if missing:
        return False, f"Kolom hilang: {', '.join(missing)}"
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except ValueError:
        return False, "Kolom 'Total Port' harus numerik"
    return True, "Data valid"

def create_pivot_tables(df):
    delta_values = compare_with_previous(df)
    df['LoP'] = 1

    # Pivot Witel
    witel_pivot = pd.pivot_table(
        df, values=['LoP', 'Total Port'], index='Witel', columns='Status Proyek',
        aggfunc='sum', fill_value=0, margins=True, margins_name='Grand Total'
    )
    witel_pivot.columns = ['_'.join(col) for col in witel_pivot.columns]
    witel_pivot['% Go Live'] = (witel_pivot.get('Total Port_Go Live',0) / witel_pivot['Total Port_Grand Total']).fillna(0)*100
    witel_pivot['Penambahan Go Live'] = [delta_values.get(witel,0) for witel in witel_pivot.index]
    witel_pivot['RANK'] = witel_pivot['Total Port_Grand Total'].rank(ascending=False, method='dense').astype('Int64')
    witel_pivot.loc['Grand Total', 'RANK'] = pd.NA

    # Pivot Datel
    datel_pivot = pd.pivot_table(
        df, values=['LoP', 'Total Port'], index=['Witel', 'Datel'], columns='Status Proyek',
        aggfunc='sum', fill_value=0
    )
    datel_pivot.columns = ['_'.join(col) for col in datel_pivot.columns]
    datel_pivot['Total Port'] = datel_pivot.get('Total Port_On Going',0) + datel_pivot.get('Total Port_Go Live',0)
    datel_pivot['% Go Live'] = (datel_pivot.get('Total Port_Go Live',0) / datel_pivot['Total Port']).fillna(0)*100
    datel_pivot['RANK'] = datel_pivot.groupby('Witel')['Total Port'].rank(ascending=False, method='min').astype('Int64')

    return witel_pivot, datel_pivot

# UI Setup
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")
st.sidebar.title("Navigasi")
view_mode = st.sidebar.radio("Pilih Mode", ["Dashboard", "Upload Data"])

# Dashboard
if view_mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)
        regional_list = ['All'] + sorted(df['Regional'].dropna().unique())
        selected_region = st.selectbox("Pilih Regional", regional_list)
        if selected_region != 'All':
            df = df[df['Regional'] == selected_region]
        st.subheader("🚀 Detail Proyek Go Live")
        go_live_df = df[df['Status Proyek']=='Go Live'][['Witel','Datel','Nama Proyek','Total Port','Status Proyek']]
        st.dataframe(go_live_df,use_container_width=True)

        witel_pivot, datel_pivot = create_pivot_tables(df)

        st.subheader("📌 Rekap per Witel")
        st.dataframe(witel_pivot,use_container_width=True)

        st.subheader("📌 Rekap per Datel")
        for witel in datel_pivot.index.get_level_values(0).unique():
            st.write(f"**{witel}**")
            st.dataframe(datel_pivot.loc[witel],use_container_width=True)
            fig = px.bar(datel_pivot.loc[witel].reset_index(), x='Datel', y=['Total Port_On Going','Total Port_Go Live'], title=f'Status Port di {witel}')
            st.plotly_chart(fig,use_container_width=True)
    else:
        st.warning("Belum ada data. Silakan upload.")

# Upload Data
else:
    st.title("📤 Upload Data Harian")
    last_upload, last_hash = get_last_upload_info()
    if last_upload:
        st.info(f"Terakhir upload: {last_upload}")
    uploaded_file = st.file_uploader("Upload file Excel harian", type="xlsx")
    if uploaded_file:
        current_hash = get_file_hash(uploaded_file.getvalue())
        if last_hash and current_hash == last_hash:
            st.success("✅ Data sama dengan upload terakhir.")
        else:
            df = load_excel(uploaded_file)
            is_valid, msg = validate_data(df)
            if not is_valid:
                st.error(msg)
            else:
                save_file(LATEST_FILE, uploaded_file)
                record_upload_history()
                st.success("✅ Berhasil upload dan update!")
                st.balloons()
                st.dataframe(df.head(),use_container_width=True)
                st.metric("Total Port",df['Total Port'].sum())

if view_mode=="Upload Data" and os.path.exists(HISTORY_FILE):
    st.sidebar.subheader("History Upload")
    history_df = pd.read_csv(HISTORY_FILE)
    st.sidebar.dataframe(history_df.tail(5),hide_index=True)
