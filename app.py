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
                go_curr = df_curr[df_curr['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                go_prev = df_prev[df_prev['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                for w in go_curr.index:
                    delta[w] = int(go_curr[w] - go_prev.get(w, 0))
    return delta

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    last_hash = ""
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        if not h.empty:
            last_hash = h.iloc[-1]['file_hash']
    with open(LATEST_FILE,'rb') as f:
        new_hash = get_file_hash(f.read())
    if last_hash and last_hash != new_hash:
        shutil.copy(LATEST_FILE, os.path.join(DATA_FOLDER, f"previous_{last_hash}.xlsx"))
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
    w['% Go Live'] = (w.get('Total Port_Go Live',0) / w['Total Port_Grand Total']).fillna(0) * 100
    w['Δ Go Live'] = [delta.get(witel, 0) for witel in w.index]
    w['RANK'] = w['Total Port_Grand Total'].rank(ascending=False, method='dense')

    # Pastikan kolom default ada sebelum reset
    for c in ['LoP_On Going','Total Port_On Going','LoP_Go Live','Total Port_Go Live','LoP_Grand Total','Total Port_Grand Total']:
        if c not in w.columns:
            w[c] = 0

    # Reset index & rename
    wp = w.reset_index().rename(columns={
        'LoP_On Going':'OnGoing_LoP',
        'Total Port_On Going':'OnGoing_Port',
        'LoP_Go Live':'GoLive_LoP',
        'Total Port_Go Live':'GoLive_Port',
        'LoP_Grand Total':'Total_LoP',
        'Total Port_Grand Total':'Total_Port'
    })

    # Buang rank di Grand Total
    wp.loc[wp['Witel']=='Grand Total','RANK'] = None

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
    d['Total Port'] = d.get('Total Port_On Going',0) + d.get('Total Port_Go Live',0)
    d['% Go Live'] = (d.get('Total Port_Go Live',0) / d['Total Port']).fillna(0) * 100
    d['RANK'] = d['Total Port'].groupby(level=0).rank(ascending=False, method='min').astype(int)

    return wp, d.reset_index()

# UI Setup
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

        # Buat pivots
        wp, dp = create_pivots(df)

        # ---- Styling Witel ----
        st.subheader("📊 Rekap Racing PT2 per WITEL")

        def highlight_witel(row):
            if row['Witel']=='Grand Total':
                return ['background-color: #fffae6']*len(row)
            if row['RANK']==1:
                return ['background-color: #e6f3ff']*len(row)
            return ['']*len(row)

        styled_wp = (
            wp.style
              .format({
                  'Total_Port':'{:,.0f}',
                  '% Go Live':'{:.1f}%',
                  'OnGoing_LoP':'{:,.0f}',
                  'OnGoing_Port':'{:,.0f}',
                  'GoLive_LoP':'{:,.0f}',
                  'GoLive_Port':'{:,.0f}',
                  'Total_LoP':'{:,.0f}',
                  'Δ Go Live':'{:,.0f}',
              })
              # untuk kolom RANK, pakai fungsi agar None menjadi kosong
              .format({'RANK': lambda v: '' if pd.isna(v) else f"{int(v)}"})
              .apply(highlight_witel, axis=1)
        )
        st.dataframe(styled_wp, use_container_width=True, height=350)

        # ---- Datel per Witel ----
        st.subheader("🏆 Racing per DATEL")
        dp = dp.rename(columns={
            'LoP_On Going':'OnGoing_LoP',
            'Total Port_On Going':'OnGoing_Port',
            'LoP_Go Live':'GoLive_LoP',
            'Total Port_Go Live':'GoLive_Port'
        })
        dp['Total_Port'] = dp.get('OnGoing_Port',0) + dp.get('GoLive_Port',0)

        witels = dp['Witel'].unique()
        tabs = st.tabs([f"🏆 {w}" for w in witels])
        for i, w in enumerate(witels):
            with tabs[i]:
                sub = dp[dp['Witel']==w].sort_values('RANK')
                styled_dp = (
                    sub.style
                       .format({
                           'Total_Port':'{:,.0f}',
                           '% Go Live':'{:.1f}%',
                           'OnGoing_LoP':'{:,.0f}',
                           'OnGoing_Port':'{:,.0f}',
                           'GoLive_LoP':'{:,.0f}',
                           'GoLive_Port':'{:,.0f}',
                       })
                       .format({'RANK': lambda v: f"{int(v)}"})
                       .applymap(lambda v: 'background-color: #e6f3ff' if v==1 else '', subset=['RANK'])
                )
                st.dataframe(styled_dp, use_container_width=True, height=300)

                fig = px.bar(
                    sub,
                    x='Datel',
                    y=['OnGoing_Port','GoLive_Port'],
                    title=f'Port Status per DATEL - {w}',
                    labels={'value':'Port Count','variable':'Status'}
                )
                st.plotly_chart(fig, use_container_width=True)

        # ---- Charts Summary ----
        st.subheader("📈 Visualisasi Data")
        plot = wp[wp['Witel']!='Grand Total'].sort_values('Total_Port', ascending=False)
        c1, c2 = st.columns(2)
        with c1:
            f1 = px.bar(plot, x='Witel', y='Total_Port', text='Total_Port')
            f1.update_traces(textposition='outside')
            st.plotly_chart(f1, use_container_width=True)
        with c2:
            f2 = px.bar(plot, x='Witel', y='% Go Live', text='% Go Live')
            f2.update_traces(textposition='outside')
            f2.update_yaxes(range=[0,100])
            st.plotly_chart(f2, use_container_width=True)

        st.subheader("📈 Perubahan Port Go Live (H-1 vs Hari Ini)")
        f3 = px.bar(plot, x='Witel', y='Δ Go Live', text='Δ Go Live')
        f3.update_traces(textposition='outside')
        st.plotly_chart(f3, use_container_width=True)

    else:
        st.warning("Belum ada data yang diupload. Silakan ke halaman Upload Data.")

else:
    # Upload Data View (tidak berubah) ...
    st.title("📤 Upload Data Harian")
    last_ts, last_h = get_last_upload_info()
    if last_ts:
        st.info(f"Terakhir upload: {last_ts}")

    uploaded = st.file_uploader("Upload file Excel harian", type="xlsx")
    if uploaded:
        try:
            curr_h = get_file_hash(uploaded.getvalue())
            if last_h and curr_h == last_h:
                st.success("✅ Data sama dengan upload terakhir. Tidak perlu upload ulang.")
            else:
                df_new = pd.read_excel(uploaded)
                ok, msg = validate_data(df_new)
                if not ok:
                    st.error(msg)
                else:
                    save_file(LATEST_FILE, uploaded)
                    record_upload_history()
                    st.success("✅ File berhasil diupload dan dashboard diperbarui!")
                    st.balloons()

                    go2 = df_new[df_new['Status Proyek']=='Go Live']
                    if not go2.empty:
                        st.subheader("📋 Detail Proyek Go Live (Upload Terbaru)")
                        st.dataframe(go2[['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                                     use_container_width=True)

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

# Sidebar history
if mode=="Upload Data" and os.path.exists(HISTORY_FILE):
    hist = pd.read_csv(HISTORY_FILE)
    st.sidebar.subheader("History Upload")
    st.sidebar.dataframe(hist[['timestamp']].tail(5), hide_index=True)
