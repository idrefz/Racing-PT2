import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import shutil

# Configuration
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")

# Create folder if not exists
os.makedirs(DATA_FOLDER, exist_ok=True)

# Helper Functions
@st.cache_data
def load_excel(path):
    return pd.read_excel(path)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(file_content: bytes):
    return hashlib.md5(file_content).hexdigest()

def compare_with_previous(current_df):
    """
    Compare current data with previous upload to calculate delta in Go Live ports
    Returns a dict {Witel: delta}
    """
    delta_dict = {}
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if len(history_df) >= 2:
            previous_hash = history_df.iloc[-2]['file_hash']
            prev_file = os.path.join(DATA_FOLDER, f"previous_{previous_hash}.xlsx")
            if os.path.exists(prev_file):
                prev_df = pd.read_excel(prev_file)
                curr_go = current_df[current_df['Status Proyek'] == 'Go Live'].groupby('Witel')['Total Port'].sum()
                prev_go = prev_df[prev_df['Status Proyek'] == 'Go Live'].groupby('Witel')['Total Port'].sum()
                for w in curr_go.index:
                    delta_dict[w] = int(curr_go[w] - prev_go.get(w, 0))
                # keep previous file intact
    return delta_dict

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # read last hash
    last_hash = ""
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        if not hist.empty:
            last_hash = hist.iloc[-1]['file_hash']
    # compute new hash
    with open(LATEST_FILE, "rb") as f:
        new_hash = get_file_hash(f.read())
    # archive previous file copy
    if last_hash and last_hash != new_hash:
        prev_copy = os.path.join(DATA_FOLDER, f"previous_{last_hash}.xlsx")
        shutil.copy(LATEST_FILE, prev_copy)
    # append history
    new_entry = pd.DataFrame([[now, new_hash]], columns=['timestamp','file_hash'])
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        hist = pd.concat([hist, new_entry], ignore_index=True)
    else:
        hist = new_entry
    hist.to_csv(HISTORY_FILE, index=False)

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        df = pd.read_csv(HISTORY_FILE)
        if not df.empty:
            return df.iloc[-1]['timestamp'], df.iloc[-1]['file_hash']
    return None, None

def validate_data(df):
    required = ['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    missing = [c for c in required if c not in df.columns]
    if missing:
        return False, f"Kolom hilang: {', '.join(missing)}"
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except Exception:
        return False, "Kolom 'Total Port' harus numerik"
    return True, "Data valid"

def create_pivot_tables(df):
    try:
        delta = compare_with_previous(df)
        df['LoP'] = 1

        # Witel pivot
        w = pd.pivot_table(
            df,
            values=['LoP','Total Port'],
            index='Witel',
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        )
        w.columns = ['_'.join(col) for col in w.columns]
        if 'Total Port_Go Live' in w.columns:
            w['%'] = w['Total Port_Go Live'] / w['Total Port_Grand Total'] * 100
        else:
            w['%'] = 0
        w['Penambahan GOLIVE H-1 vs HI'] = w.index.map(lambda x: delta.get(x, 0))
        w['RANK'] = w['%'].rank(ascending=False, method='dense')
        w = w.round(0).astype({c:int for c in w.columns if c not in ['%']})
        w.loc['Grand Total','RANK'] = None

        # Datel pivot
        d = pd.pivot_table(
            df,
            values=['LoP','Total Port'],
            index=['Witel','Datel'],
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0
        )
        d.columns = ['_'.join(col) for col in d.columns]
        if 'Total Port_Go Live' in d.columns:
            d['%'] = d['Total Port_Go Live'] / (
                d.get('Total Port_On Going',0) + d['Total Port_Go Live']
            ) * 100
        else:
            d['%'] = 0
        d['RANK'] = d.groupby(level=0)['Total Port_Go Live'].rank(ascending=False, method='min')
        d = d.round(0).astype({c:int for c in d.columns if c not in ['%']})

        return w, d
    except Exception as e:
        st.error(f"Error membuat pivot: {e}")
        return None, None

# UI Setup
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")

# Sidebar
st.sidebar.title("Navigasi")
view_mode = st.sidebar.radio("Pilih Mode", ["Dashboard","Upload Data"])

# Dashboard
if view_mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)

        # Regional filter
        regions = ['All'] + sorted(df['Regional'].dropna().unique())
        sel_reg = st.selectbox("Pilih Regional", regions)
        if sel_reg != 'All':
            df = df[df['Regional'] == sel_reg]

        # ===== Tabel Go Live =====
        go_live_df = df[df['Status Proyek'] == 'Go Live']
        if not go_live_df.empty:
            st.subheader("📋 Detail Proyek Go Live")
            st.dataframe(
                go_live_df[
                    ['Witel','Datel','Nama Proyek','Total Port','Status Proyek']
                ],
                use_container_width=True
            )
        else:
            st.info("Tidak ada data Go Live untuk pilihan Regional saat ini.")
        # ===== end Go Live table =====

        # Pivot & Visualisasi
        witel_pivot, datel_pivot = create_pivot_tables(df)
        if witel_pivot is not None:
            # Witel summary
            wp = witel_pivot.copy()
            wp = wp.reset_index().rename(columns={
                'LoP_On Going':'On Going_Lop','Total Port_On Going':'On Going_Port',
                'LoP_Go Live':'Go Live_Lop','Total Port_Go Live':'Go Live_Port',
                'LoP_Grand Total':'Total Lop','Total Port_Grand Total':'Total Port'
            })
            wp['%'] = wp['%'].round(0)
            wp['RANK'] = wp['RANK'].fillna(0).astype(int)
            st.subheader("📊 Rekap Racing PT2 per WITEL")
            st.dataframe(wp, use_container_width=True)

            # Datel per Witel
            if datel_pivot is not None:
                dp = datel_pivot.copy()
                dp = dp.reset_index().rename(columns={
                    'LoP_On Going':'On Going_Lop','Total Port_On Going':'On Going_Port',
                    'LoP_Go Live':'Go Live_Lop','Total Port_Go Live':'Go Live_Port'
                })
                dp['Total Lop'] = dp['On Going_Lop'] + dp['Go Live_Lop']
                dp['Total Port'] = dp['On Going_Port'] + dp['Go Live_Port']
                dp['%'] = dp['%'].round(0)
                dp['RANK'] = dp['RANK'].astype(int)

                witels = dp['Witel'].unique()
                tabs = st.tabs([f"🏆 {w}" for w in witels])
                for i, w in enumerate(witels):
                    with tabs[i]:
                        df_w = dp[dp['Witel']==w].sort_values('RANK')
                        st.dataframe(df_w, use_container_width=True)
                        fig = px.bar(
                            df_w,
                            x='Datel',
                            y=['On Going_Port','Go Live_Port'],
                            title=f'Port Status per DATEL - {w}',
                            labels={'value':'Port Count','variable':'Status'}
                        )
                        st.plotly_chart(fig, use_container_width=True)

            # Overall charts
            st.subheader("📈 Visualisasi Data")
            plot_df = wp[wp['Witel']!='Grand Total'].sort_values('Total Port', ascending=False)
            col1, col2 = st.columns(2)
            with col1:
                fig1 = px.bar(plot_df, x='Witel', y='Total Port', text='Total Port')
                fig1.update_traces(textposition='outside')
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                fig2 = px.bar(plot_df, x='Witel', y='%', text='%')
                fig2.update_traces(textposition='outside')
                fig2.update_yaxes(range=[0,100])
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("📈 Perubahan Port Go Live (H-1 vs HI)")
            fig3 = px.bar(plot_df, x='Witel', y='Penambahan GOLIVE H-1 vs HI', text='Penambahan GOLIVE H-1 vs HI')
            fig3.update_traces(textposition='outside')
            st.plotly_chart(fig3, use_container_width=True)

    else:
        st.warning("Belum ada data yang diupload. Silakan ke halaman Upload Data.")

# Upload Data View
else:
    st.title("📤 Upload Data Harian")
    last_time, last_hash = get_last_upload_info()
    if last_time:
        st.info(f"Terakhir upload: {last_time}")

    uploaded = st.file_uploader("Upload file Excel harian", type="xlsx")
    if uploaded:
        try:
            curr_hash = get_file_hash(uploaded.getvalue())
            if last_hash and curr_hash == last_hash:
                st.success("✅ Data sama dengan upload terakhir. Tidak perlu upload ulang.")
            else:
                df_new = pd.read_excel(uploaded)
                valid, msg = validate_data(df_new)
                if not valid:
                    st.error(msg)
                else:
                    # save & history
                    save_file(LATEST_FILE, uploaded)
                    record_upload_history()

                    st.success("✅ File berhasil diupload dan dashboard diperbarui!")
                    st.balloons()

                    # ===== Tabel Go Live pasca-upload =====
                    go_live_df2 = df_new[df_new['Status Proyek']=='Go Live']
                    if not go_live_df2.empty:
                        st.subheader("📋 Detail Proyek Go Live (Upload Terbaru)")
                        st.dataframe(
                            go_live_df2[
                                ['Witel','Datel','Nama Proyek','Total Port','Status Proyek']
                            ],
                            use_container_width=True
                        )
                    # ===== end =====

                    # Quick metrics
                    st.subheader("📋 Ringkasan Data")
                    cols = st.columns(4)
                    cols[0].metric("Total Projek", len(df_new))
                    cols[1].metric("Total Port", int(df_new['Total Port'].sum()))
                    cols[2].metric("WITEL", df_new['Witel'].nunique())
                    cols[3].metric("DATEL", df_new['Datel'].nunique())

                    st.subheader("🖥️ Preview Data")
                    st.dataframe(df_new.head(), use_container_width=True)
        except Exception as e:
            st.error(f"Kesalahan saat memproses file: {e}")

# Sidebar: upload history
if view_mode == "Upload Data" and os.path.exists(HISTORY_FILE):
    hist_df = pd.read_csv(HISTORY_FILE)
    st.sidebar.subheader("History Upload")
    st.sidebar.dataframe(hist_df[['timestamp']].tail(5), hide_index=True)
