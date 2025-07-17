import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import shutil

# ── CONFIG ─────────────────────────────────────────────────────────────────────
DATA_FOLDER   = "data_daily_uploads"
LATEST_FILE   = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE  = os.path.join(DATA_FOLDER, "upload_history.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# ── HELPERS ────────────────────────────────────────────────────────────────────
@st.cache_data
def load_excel(path):
    return pd.read_excel(path)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(content: bytes) -> str:
    return hashlib.md5(content).hexdigest()

def compare_with_previous(df_curr: pd.DataFrame) -> dict[str,int]:
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

def get_last_upload_info() -> tuple[str,str] | tuple[None,None]:
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        if not h.empty:
            return h.iloc[-1]['timestamp'], h.iloc[-1]['file_hash']
    return None, None

def validate_data(df: pd.DataFrame) -> tuple[bool,str]:
    req = ['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    missing = [c for c in req if c not in df.columns]
    if missing:
        return False, f"Kolom hilang: {', '.join(missing)}"
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except:
        return False, "Kolom 'Total Port' harus numerik"
    return True, "Data valid"

def create_pivots(df: pd.DataFrame):
    delta = compare_with_previous(df)
    df['LoP'] = 1

    # ─ Witel pivot ─────────────────────────────────
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
    # flatten cols
    w.columns = ['_'.join(col) for col in w.columns]
    # derived metrics
    w['% Go Live']     = (w.get('Total Port_Go Live',0) / w['Total Port_Grand Total']).fillna(0)*100
    w['Δ Go Live']     = [delta.get(wit,0) for wit in w.index]
    w['RANK']          = w['Total Port_Grand Total'].rank(ascending=False, method='dense')

    # ensure expected cols exist
    for c in ['LoP_On Going','Total Port_On Going',
              'LoP_Go Live','Total Port_Go Live',
              'LoP_Grand Total','Total Port_Grand Total']:
        if c not in w.columns:
            w[c] = 0

    # reset & rename for display
    wp = w.reset_index().rename(columns={
        'LoP_On Going'            :'OnGoing_LoP',
        'Total Port_On Going'     :'OnGoing_Port',
        'LoP_Go Live'             :'GoLive_LoP',
        'Total Port_Go Live'      :'GoLive_Port',
        'LoP_Grand Total'         :'Total_LoP',
        'Total Port_Grand Total'  :'Total_Port'
    })
    wp.loc[wp['Witel']=='Grand Total','RANK'] = None

    # ─ Datel pivot ─────────────────────────────────
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
    d['% Go Live']  = (d.get('Total Port_Go Live',0) / d['Total Port']).fillna(0)*100
    d['RANK']       = d['Total Port'].groupby(level=0).rank(ascending=False, method='min')

    dp = d.reset_index()
    return wp, dp

# ── UI ─────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")
st.sidebar.title("Navigasi")
mode = st.sidebar.radio("Pilih Mode", ["Dashboard","Upload Data"])

if mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")

    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)

        # —— Regional filter —— 
        regions = ['All'] + sorted(df['Regional'].dropna().unique())
        sel = st.selectbox("Pilih Regional", regions)
        if sel != 'All':
            df = df[df['Regional']==sel]

        # —— Go Live detail —— 
        go = df[df['Status Proyek']=='Go Live']
        if not go.empty:
            st.subheader("📋 Detail Proyek Go Live")
            st.dataframe(
                go[['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                use_container_width=True
            )
        else:
            st.info("Tidak ada data Go Live untuk pilihan ini.")

        # —— Generate pivots —— 
        wp, dp = create_pivots(df)

        # —— Tabel Witel —— 
        st.subheader("📈 Rekap per WITEL")
        def style_witel(row):
            if row.Witel == 'Grand Total':
                return ['background-color:#fffae6;font-weight:bold']*len(row)
            elif row.RANK == 1:
                return ['background-color:#e6f3ff']*len(row)
            else:
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
                  'Δ Go Live':'{:,.0f}'
              })
              .format({'RANK': lambda v: '' if pd.isna(v) else f"{int(v)}"})
              .apply(style_witel, axis=1)
        )
        st.dataframe(styled_wp, use_container_width=True, height=350)

        # —— Tabel Datel —— 
        st.subheader("🏆 Rekap per DATEL")
        dp['RANK'] = dp['RANK'].fillna(0).astype(int)
        def style_datel(val):
            return 'background-color:#e6f3ff' if val == 1 else ''

        witels = dp['Witel'].unique()
        tabs = st.tabs([f"{w}" for w in witels])
        for i, w in enumerate(witels):
            with tabs[i]:
                sub = dp[dp['Witel']==w].sort_values('RANK')
                styled_dp = (
                    sub.style
                       .format({
                           'Total Port':'{:,.0f}',
                           '% Go Live':'{:.1f}%',
                           'LoP':'{:,.0f}'
                       })
                       .applymap(style_datel, subset=['RANK'])
                       .format({'RANK': lambda v: str(v)})
                )
                st.dataframe(styled_dp, use_container_width=True, height=300)

                # grafik per Datel
                fig = px.bar(
                    sub,
                    x='Datel',
                    y=['Total Port'],
                    color='Status Proyek',
                    title=f'Port per Status di {w}',
                    labels={'value':'Port','variable':'Status Proyek'}
                )
                st.plotly_chart(fig, use_container_width=True)

        # —— Ringkasan Charts —— 
        st.subheader("🚀 Visualisasi Keseluruhan")
        overall = wp[wp['Witel']!='Grand Total'].sort_values('Total_Port', ascending=False)
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.bar(overall, x='Witel', y='Total_Port', text='Total_Port', title="Total Port per Witel")
            fig1.update_traces(textposition='outside')
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = px.bar(overall, x='Witel', y='% Go Live', text='% Go Live', title="% Go Live per Witel")
            fig2.update_traces(textposition='outside')
            fig2.update_yaxes(range=[0,100])
            st.plotly_chart(fig2, use_container_width=True)

    else:
        st.warning("Belum ada data. Silakan upload di menu 'Upload Data'.")

else:
    # —— Upload Data View —— 
    st.title("📤 Upload Data Harian")
    last_ts, last_h = get_last_upload_info()
    if last_ts:
        st.info(f"Terakhir upload: {last_ts}")

    uploaded = st.file_uploader("Pilih file .xlsx", type="xlsx")
    if uploaded:
        try:
            curr_h = get_file_hash(uploaded.getvalue())
            if last_h and curr_h == last_h:
                st.success("✅ Data sama dengan upload terakhir.")
            else:
                df_new = pd.read_excel(uploaded)
                ok, msg = validate_data(df_new)
                if not ok:
                    st.error(msg)
                else:
                    save_file(LATEST_FILE, uploaded)
                    record_upload_history()
                    st.success("✅ Berhasil upload dan update dashboard!")
                    st.balloons()

                    # langsung tunjukkan Go Live pasca upload
                    go2 = df_new[df_new['Status Proyek']=='Go Live']
                    if not go2.empty:
                        st.subheader("📋 Go Live (Upload Terbaru)")
                        st.dataframe(go2[['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                                     use_container_width=True)

                    # ringkasan cepat
                    cols = st.columns(4)
                    cols[0].metric("Total Proyek", len(df_new))
                    cols[1].metric("Total Port", int(df_new['Total Port'].sum()))
                    cols[2].metric("Witel Unik", df_new['Witel'].nunique())
                    cols[3].metric("Datel Unik", df_new['Datel'].nunique())

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

# —— Sidebar History —— 
if mode=="Upload Data" and os.path.exists(HISTORY_FILE):
    hist = pd.read_csv(HISTORY_FILE)
    st.sidebar.subheader("History Upload")
    st.sidebar.dataframe(hist[['timestamp']].tail(5), hide_index=True)
