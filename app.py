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
os.makedirs(DATA_FOLDER, exist_ok=True)

# Helper Functions
@st.cache_data
def load_excel(path):
    return pd.read_excel(path)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(content: bytes):
    return hashlib.md5(content).hexdigest()

def compare_with_previous(df_curr):
    delta = {}
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        if len(hist) >= 2:
            prev_hash = hist.iloc[-2]['file_hash']
            prev_path = os.path.join(DATA_FOLDER, f"previous_{prev_hash}.xlsx")
            if os.path.exists(prev_path):
                df_prev = pd.read_excel(prev_path)
                go_curr = (df_curr[df_curr['Status Proyek']=='Go Live']
                           .groupby('Witel')['Total Port'].sum())
                go_prev = (df_prev[df_prev['Status Proyek']=='Go Live']
                           .groupby('Witel')['Total Port'].sum())
                for w in go_curr.index:
                    delta[w] = int(go_curr[w] - go_prev.get(w, 0))
    return delta

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # last hash
    last_hash = ""
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        if not h.empty:
            last_hash = h.iloc[-1]['file_hash']
    # new hash
    with open(LATEST_FILE,'rb') as f:
        new_hash = get_file_hash(f.read())
    # archive previous if changed
    if last_hash and last_hash != new_hash:
        shutil.copy(LATEST_FILE, os.path.join(DATA_FOLDER, f"previous_{last_hash}.xlsx"))
    # append history
    entry = pd.DataFrame([[now, new_hash]], columns=['timestamp','file_hash'])
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        h = pd.concat([h, entry], ignore_index=True)
    else:
        h = entry
    h.to_csv(HISTORY_FILE, index=False)

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        if not h.empty:
            return h.iloc[-1]['timestamp'], h.iloc[-1]['file_hash']
    return None, None

def validate_data(df):
    req = ['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    missing = [c for c in req if c not in df.columns]
    if missing:
        return False, f"Kolom hilang: {', '.join(missing)}"
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except:
        return False, "Kolom 'Total Port' harus numerik"
    return True, "Data valid"

def create_pivots(df):
    delta = compare_with_previous(df)
    df['LoP'] = 1

    # --- Witel pivot ---
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
    # flatten columns
    w.columns = ['_'.join(col) for col in w.columns]
    # percentage
    w['%'] = (w.get('Total Port_Go Live',0) / w['Total Port_Grand Total']).fillna(0) * 100
    # delta
    w['Penambahan GOLIVE H-1 vs HI'] = [delta.get(witel,0) for witel in w.index]
    # rank overall
    w['RANK'] = w['%'].rank(ascending=False, method='dense')
    # ensure integer types
    for col in w.columns:
        if col not in ['%']:
            w[col] = w[col].round(0).astype(int)
    w.loc['Grand Total','RANK'] = None

    # --- Datel pivot ---
    d = pd.pivot_table(
        df,
        values=['LoP','Total Port'],
        index=['Witel','Datel'],
        columns='Status Proyek',
        aggfunc='sum',
        fill_value=0
    )
    d.columns = ['_'.join(col) for col in d.columns]
    # percentage per Datel
    total_on = d.get('Total Port_On Going',0)
    total_gl = d.get('Total Port_Go Live',0)
    d['%'] = (total_gl / (total_on + total_gl)).fillna(0) * 100

    # ---- FIXED: correct ranking per Witel ----
    if 'Total Port_Go Live' in d.columns:
        # group the series, then rank
        d['RANK'] = d['Total Port_Go Live'].groupby(level=0).rank(ascending=False, method='min')
    else:
        d['RANK'] = 0
    # cast to int where appropriate
    for col in d.columns:
        if col not in ['%']:
            d[col] = d[col].round(0).astype(int)

    return w, d

# --- UI Setup ---
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")
st.sidebar.title("Navigasi")
mode = st.sidebar.radio("Pilih Mode", ["Dashboard","Upload Data"])

if mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)

        # Regional filter
        regions = ['All'] + sorted(df['Regional'].dropna().unique())
        sel = st.selectbox("Pilih Regional", regions)
        if sel != 'All':
            df = df[df['Regional']==sel]

        # Go Live detail
        go = df[df['Status Proyek']=='Go Live']
        if not go.empty:
            st.subheader("📋 Detail Proyek Go Live")
            st.dataframe(go[['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                         use_container_width=True)
        else:
            st.info("Tidak ada data Go Live untuk pilihan Regional saat ini.")

        # Pivots & Visualisasi
        witel_pivot, datel_pivot = create_pivots(df)

        # Witel summary
        wp = witel_pivot.reset_index().rename(columns={
            'LoP_On Going':'On Going_Lop','Total Port_On Going':'On Going_Port',
            'LoP_Go Live':'Go Live_Lop','Total Port_Go Live':'Go Live_Port',
            'LoP_Grand Total':'Total Lop','Total Port_Grand Total':'Total Port'
        })
        st.subheader("📊 Rekap Racing PT2 per WITEL")
        st.dataframe(wp, use_container_width=True)

        # Datel per Witel
        dp = datel_pivot.reset_index().rename(columns={
            'LoP_On Going':'On Going_Lop','Total Port_On Going':'On Going_Port',
            'LoP_Go Live':'Go Live_Lop','Total Port_Go Live':'Go Live_Port'
        })
        # pastikan kolom ada
        for c in ['On Going_Lop','Go Live_Lop','On Going_Port','Go Live_Port']:
            if c not in dp.columns:
                dp[c] = 0
        dp['Total Lop']  = dp['On Going_Lop'] + dp['Go Live_Lop']
        dp['Total Port'] = dp['On Going_Port'] + dp['Go Live_Port']

        witels = dp['Witel'].unique()
        tabs = st.tabs([f"🏆 {w}" for w in witels])
        for i, w in enumerate(witels):
            with tabs[i]:
                sub = dp[dp['Witel']==w].sort_values('RANK')
                st.dataframe(sub, use_container_width=True)
                fig = px.bar(
                    sub,
                    x='Datel',
                    y=['On Going_Port','Go Live_Port'],
                    title=f'Port Status per DATEL - {w}',
                    labels={'value':'Port Count','variable':'Status'}
                )
                st.plotly_chart(fig, use_container_width=True)

        # Overall charts
        st.subheader("📈 Visualisasi Data")
        plot = wp[wp['Witel']!='Grand Total'].sort_values('Total Port', ascending=False)
        c1, c2 = st.columns(2)
        with c1:
            f1 = px.bar(plot, x='Witel', y='Total Port', text='Total Port')
            f1.update_traces(textposition='outside')
            st.plotly_chart(f1, use_container_width=True)
        with c2:
            f2 = px.bar(plot, x='Witel', y='%', text='%')
            f2.update_traces(textposition='outside')
            f2.update_yaxes(range=[0,100])
            st.plotly_chart(f2, use_container_width=True)

        st.subheader("📈 Perubahan Port Go Live (H-1 vs Hari Ini)")
        f3 = px.bar(plot, x='Witel', y='Penambahan GOLIVE H-1 vs HI', text='Penambahan GOLIVE H-1 vs HI')
        f3.update_traces(textposition='outside')
        st.plotly_chart(f3, use_container_width=True)

    else:
        st.warning("Belum ada data yang diupload. Silakan ke halaman Upload Data.")

else:  # Upload Data View
    st.title("📤 Upload Data Harian")
    last_ts, last_h = get_last_upload_info()
    if last_ts:
        st.info(f"Terakhir upload: {last_ts}")

    upl = st.file_uploader("Upload file Excel harian", type="xlsx")
    if upl:
        try:
            curr_h = get_file_hash(upl.getvalue())
            if last_h and curr_h == last_h:
                st.success("✅ Data sama dengan upload terakhir. Tidak perlu upload ulang.")
            else:
                df_new = pd.read_excel(upl)
                ok, msg = validate_data(df_new)
                if not ok:
                    st.error(msg)
                else:
                    save_file(LATEST_FILE, upl)
                    record_upload_history()
                    st.success("✅ File berhasil diupload dan dashboard diperbarui!")
                    st.balloons()

                    # Tabel Go Live pasca-upload
                    go2 = df_new[df_new['Status Proyek']=='Go Live']
                    if not go2.empty:
                        st.subheader("📋 Detail Proyek Go Live (Upload Terbaru)")
                        st.dataframe(go2[['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                                     use_container_width=True)

                    # Ringkasan cepat
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
if mode == "Upload Data" and os.path.exists(HISTORY_FILE):
    hist = pd.read_csv(HISTORY_FILE)
    st.sidebar.subheader("History Upload")
    st.sidebar.dataframe(hist[['timestamp']].tail(5), hide_index=True)
