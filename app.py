import streamlit as st
import pandas as pd
import os, hashlib, shutil
from datetime import datetime
import plotly.express as px

# ───── CONFIG ───────────────────────────────────────────
DATA_FOLDER  = "data_daily_uploads"
LATEST_FILE  = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# ───── HELPERS ──────────────────────────────────────────
@st.cache_data
def load_excel(path): return pd.read_excel(path)

def save_file(path, upl): open(path,"wb").write(upl.getbuffer())
def get_hash(b): return hashlib.md5(b).hexdigest()

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        h=pd.read_csv(HISTORY_FILE)
        if not h.empty: return h.iloc[-1]["timestamp"], h.iloc[-1]["file_hash"]
    return None,None

def record_history():
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_hash=get_hash(open(LATEST_FILE,"rb").read()) if os.path.exists(LATEST_FILE) else ""
    if os.path.exists(HISTORY_FILE):
        hist=pd.read_csv(HISTORY_FILE)
        if not hist.empty:
            prev_hash=hist.iloc[-1]['file_hash']
            shutil.copy(LATEST_FILE, os.path.join(DATA_FOLDER,f"previous_{prev_hash}.xlsx"))
            hist=pd.concat([hist, pd.DataFrame([[now,new_hash]],columns=['timestamp','file_hash'])])
    else:
        hist=pd.DataFrame([[now,new_hash]],columns=['timestamp','file_hash'])
    hist.to_csv(HISTORY_FILE,index=False)

def compare_with_previous(df_now):
    delta={}
    if os.path.exists(HISTORY_FILE):
        hist=pd.read_csv(HISTORY_FILE)
        if len(hist)>=2:
            prev_hash=hist.iloc[-2]['file_hash']
            prev_path=os.path.join(DATA_FOLDER,f"previous_{prev_hash}.xlsx")
            if os.path.exists(prev_path):
                df_prev=pd.read_excel(prev_path)
                now_gl=df_now[df_now['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                prev_gl=df_prev[df_prev['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
                for w in now_gl.index: delta[w]=int(now_gl[w]-prev_gl.get(w,0))
    return delta

def validate(df):
    req=['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    miss=[c for c in req if c not in df.columns]
    if miss: return False,f"Kolom hilang: {', '.join(miss)}"
    try: df['Total Port']=pd.to_numeric(df['Total Port'],errors='raise')
    except: return False,"'Total Port' harus numerik"
    return True,"OK"

# ───── PIVOTS ───────────────────────────────────────────
def make_pivots(df):
    delta=compare_with_previous(df)
    df['LoP']=1
    w=df.pivot_table(values=['LoP','Total Port'], index='Witel', columns='Status Proyek',
                     aggfunc='sum', fill_value=0, margins=True, margins_name='Grand Total')
    w.columns=['_'.join(c) for c in w.columns]
    w['%']=(w.get('Total Port_Go Live',0)/w['Total Port_Grand Total']).fillna(0)*100
    w['Penambahan']= [delta.get(w,0) for w in w.index]
    w['RANK']=w['Total Port_Grand Total'].rank(ascending=False,method='dense').astype('Int64')
    w.loc['Grand Total','RANK']=pd.NA

    d=df.pivot_table(values=['LoP','Total Port'], index=['Witel','Datel'], columns='Status Proyek',
                     aggfunc='sum', fill_value=0)
    d.columns=['_'.join(c) for c in d.columns]
    d['Total Port']=d.get('Total Port_On Going',0)+d.get('Total Port_Go Live',0)
    d['%']=(d.get('Total Port_Go Live',0)/d['Total Port']).fillna(0)*100
    d['RANK']=d.groupby(level=0)['Total Port'].rank(ascending=False,method='min').astype('Int64')
    return w,d

# ───── UI ───────────────────────────────────────────────
st.set_page_config("Delta Ticket Harian",layout="wide")
page=st.sidebar.radio("Mode",["Dashboard","Upload Data"])

# ======= DASHBOARD =======
if page=="Dashboard":
    st.title("📊 Dashboard Deployment PT2 IHLD")
    if not os.path.exists(LATEST_FILE):
        st.info("Belum ada data. Silakan upload.")
        st.stop()

    df=load_excel(LATEST_FILE)
    regs=['All']+sorted(df['Regional'].dropna().unique())
    reg_sel=st.selectbox("Filter Regional",regs)
    if reg_sel!="All":
        df=df[df['Regional']==reg_sel]

    st.subheader("🚀 Proyek Go Live")
    st.dataframe(df[df['Status Proyek']=="Go Live"]
                 [['Witel','Datel','Nama Proyek','Total Port','Status Proyek']],
                 use_container_width=True, height=280)

    witel_pivot, datel_pivot = make_pivots(df)

    # tampilan kolom wajib
    needed=['LoP_On Going','Total Port_On Going','LoP_Go Live','Total Port_Go Live',
            'LoP_Grand Total','Total Port_Grand Total']
    for c in needed:
        if c not in witel_pivot: witel_pivot[c]=0

    wdf=pd.DataFrame({
        'Witel': witel_pivot.index,
        'On Going_Lop': witel_pivot['LoP_On Going'],
        'On Going_Port': witel_pivot['Total Port_On Going'],
        'Go Live_Lop':  witel_pivot['LoP_Go Live'],
        'Go Live_Port': witel_pivot['Total Port_Go Live'],
        'Total Lop':    witel_pivot['LoP_Grand Total'],
        'Total Port':   witel_pivot['Total Port_Grand Total'],
        '%':            witel_pivot['%'].round(1),
        'Penambahan GOLIVE H-1 vs HI': witel_pivot['Penambahan'],
        'RANK':         witel_pivot['RANK']
    })

    st.subheader("📌 Rekapitulasi per Witel")
    def style_w(r):
        if r['Witel']=='Grand Total': return ['background-color:#fff4b2;font-weight:bold']*len(r)
        if r['RANK']==1: return ['background-color:#d4f1f9']*len(r)
        return ['']*len(r)

    st.dataframe(
        wdf.style
           .format({'%':'{:.1f}%','RANK':lambda v:'' if pd.isna(v) else f'{int(v)}',
                    **{c:'{:,.0f}' for c in ['On Going_Lop','On Going_Port','Go Live_Lop',
                                             'Go Live_Port','Total Lop','Total Port',
                                             'Penambahan GOLIVE H-1 vs HI']}})
           .apply(style_w,axis=1),
        use_container_width=True, height=350)

    # ---- Datel tabs ----
    st.subheader("🏆 Rekap per Datel")
    w_nonGT=wdf[wdf['Witel']!='Grand Total']['Witel']
    tabs=st.tabs(w_nonGT.tolist())
    for i,w in enumerate(w_nonGT):
        with tabs[i]:
            sub=datel_pivot[datel_pivot.index.get_level_values(0)==w].reset_index()
            sub=sub.sort_values('RANK')

            # pastikan kolom tersedia
            for c in ['Total Port_On Going','Total Port_Go Live','Total Port','%','RANK']:
                if c not in sub: sub[c]=0

            show=sub[['Datel','Total Port_On Going','Total Port_Go Live',
                      'Total Port','%','RANK']]
            st.dataframe(show.style.format({
                'Total Port_On Going':'{:,.0f}',
                'Total Port_Go Live':'{:,.0f}',
                'Total Port':'{:,.0f}',
                '%':'{:.1f}%',
                'RANK':'{:,.0f}'
            }).applymap(
                lambda v:'background-color:#d4f1f9' if v==1 else '', subset=['RANK']),
                use_container_width=True, height=260)

            fig=px.bar(sub, x='Datel',
                       y=['Total Port_On Going','Total Port_Go Live'],
                       barmode='stack',
                       labels={'value':'Port','variable':'Status'},
                       title=f'Status Port di {w}',
                       color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig,use_container_width=True)

    # ---- Summary charts ----
    st.subheader("🎯 Ringkasan Grafik")
    base=wdf[wdf['Witel']!='Grand Total'].sort_values('Total Port',ascending=False)
    c1,c2=st.columns(2)
    with c1:
        fig_tot=px.bar(base,x='Witel',y='Total Port',
                       text='Total Port',title='Total Port per Witel',
                       color='Witel',color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_tot.update_traces(textposition='outside')
        st.plotly_chart(fig_tot,use_container_width=True)
    with c2:
        st.plotly_chart(px.pie(base,names='Witel',values='Total Port',
                               title='Distribusi Total Port',hole=.35),
                        use_container_width=True)

# ======= UPLOAD PAGE =======
else:
    st.title("📤 Upload Data Harian")
    last,last_hash=get_last_upload_info()
    if last: st.info(f"Terakhir upload: {last}")

    upl=st.file_uploader("Pilih file .xlsx",type="xlsx")
    if upl:
        if last_hash and get_hash(upl.getvalue())==last_hash:
            st.success("✅ Data sama dengan upload terakhir.")
        else:
            df=load_excel(upl)
            ok,msg=validate(df)
            if not ok: st.error(msg)
            else:
                save_file(LATEST_FILE,upl)
                record_history()
                st.success("✅ Upload berhasil & dashboard diperbarui!")
                st.balloons()
                st.dataframe(df.head(),use_container_width=True)

# sidebar: riwayat
if page=="Upload Data" and os.path.exists(HISTORY_FILE):
    st.sidebar.subheader("Riwayat Upload")
    st.sidebar.dataframe(pd.read_csv(HISTORY_FILE).tail(5),hide_index=True)
