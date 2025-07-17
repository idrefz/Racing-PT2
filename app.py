import streamlit as st
import pandas as pd
import os, hashlib, shutil
from datetime import datetime
import plotly.express as px

# ───────────────────────── CONFIG ─────────────────────────
DATA_FOLDER  = "data_daily_uploads"
LATEST_FILE  = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# ────────────────────── BASIC HELPERS ─────────────────────
@st.cache_data
def load_excel(path): return pd.read_excel(path)

def save_file(path, upl): open(path, "wb").write(upl.getbuffer())
def md5(b): return hashlib.md5(b).hexdigest()

def last_upload_info():
    if os.path.exists(HISTORY_FILE):
        h=pd.read_csv(HISTORY_FILE)
        if not h.empty: return h.iloc[-1]['timestamp'], h.iloc[-1]['file_hash']
    return None,None

def record_history():
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_hash=md5(open(LATEST_FILE,'rb').read())
    if os.path.exists(HISTORY_FILE):
        h=pd.read_csv(HISTORY_FILE)
        if not h.empty:
            prev_hash=h.iloc[-1]['file_hash']
            shutil.copy(LATEST_FILE, os.path.join(DATA_FOLDER,f"previous_{prev_hash}.xlsx"))
            h=pd.concat([h, pd.DataFrame([[now,new_hash]],columns=['timestamp','file_hash'])])
    else:
        h=pd.DataFrame([[now,new_hash]],columns=['timestamp','file_hash'])
    h.to_csv(HISTORY_FILE,index=False)

def load_prev_df():
    if os.path.exists(HISTORY_FILE):
        h=pd.read_csv(HISTORY_FILE)
        if len(h)>=2:
            ph=h.iloc[-2]['file_hash']
            pth=os.path.join(DATA_FOLDER,f"previous_{ph}.xlsx")
            if os.path.exists(pth):
                return pd.read_excel(pth)
    return pd.DataFrame()

def validate(df):
    req=['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    miss=[c for c in req if c not in df.columns]
    if miss: return False,f"Kolom hilang: {', '.join(miss)}"
    try: df['Total Port']=pd.to_numeric(df['Total Port'],errors='raise')
    except: return False,"'Total Port' harus numerik"
    return True,"OK"

# ────────────────────── PIVOT BUILDER ────────────────────
def build_pivots(df_now, df_prev):
    df_now['LoP']=1

    # Witel pivot
    wp=df_now.pivot_table(values=['LoP','Total Port'], index='Witel', columns='Status Proyek',
                          aggfunc='sum', fill_value=0, margins=True, margins_name='Grand Total')
    wp.columns=['_'.join(c) for c in wp.columns]
    wp['%']=(wp.get('Total Port_Go Live',0)/wp['Total Port_Grand Total']).fillna(0)*100

    # Rank start at 1 (exclude Grand Total)
    ranks=wp.loc[wp.index!='Grand Total','Total Port_Grand Total'].rank(ascending=False,
                                                                        method='dense').astype('Int64')
    wp['RANK']=pd.NA
    wp.loc[ranks.index,'RANK']=ranks

    # Δ Go Live
    delta={}
    if not df_prev.empty:
        now_gl=df_now[df_now['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
        prev_gl=df_prev[df_prev['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
        for w in now_gl.index: delta[w]=int(now_gl[w]-prev_gl.get(w,0))
    wp['Δ Go Live']=[delta.get(w,0) for w in wp.index]

    # Datel pivot
    dp=df_now.pivot_table(values=['LoP','Total Port'], index=['Witel','Datel'], columns='Status Proyek',
                          aggfunc='sum', fill_value=0)
    dp.columns=['_'.join(c) for c in dp.columns]
    dp['Total Port']=dp.get('Total Port_On Going',0)+dp.get('Total Port_Go Live',0)
    dp['%']=(dp.get('Total Port_Go Live',0)/dp['Total Port']).fillna(0)*100
    dp['RANK']=dp.groupby(level=0)['Total Port'].rank(ascending=False, method='min').astype('Int64')
    return wp, dp

# ───────────────────────── UI ────────────────────────────
st.set_page_config("Delta Ticket Harian",layout="wide")
page=st.sidebar.radio("Mode",["Dashboard","Upload Data"])

# =============== DASHBOARD ===============
if page=="Dashboard":
    st.title("📊 Dashboard Deployment PT2 IHLD")

    if not os.path.exists(LATEST_FILE):
        st.info("Belum ada data. Silakan upload.")
        st.stop()

    df_now=load_excel(LATEST_FILE)
    df_prev=load_prev_df()

    # Regional filter
    regs=['All']+sorted(df_now['Regional'].dropna().unique())
    sel=st.selectbox("Filter Regional",regs)
    if sel!="All":
        df_now=df_now[df_now['Regional']==sel]
        df_prev=df_prev[df_prev['Regional']==sel] if not df_prev.empty else df_prev

    # Go Live HI (baru)
    if not df_prev.empty:
        prev_ticket=set(df_prev[df_prev['Status Proyek']=='Go Live']['Ticket ID'])
        golive_hi=df_now[(df_now['Status Proyek']=='Go Live') & (~df_now['Ticket ID'].isin(prev_ticket))]
    else:
        golive_hi=df_now[df_now['Status Proyek']=='Go Live']
    st.subheader("🚀 Proyek Go Live – HI (baru)")
    if golive_hi.empty:
        st.success("Tidak ada Go Live baru.")
    else:
        st.dataframe(golive_hi[['Witel','Datel','Nama Proyek','Total Port','Ticket ID']],
                     use_container_width=True, height=250)

    wp, dp = build_pivots(df_now, df_prev)

    # Pastikan kolom exist
    for c in ['LoP_On Going','Total Port_On Going','LoP_Go Live','Total Port_Go Live',
              'LoP_Grand Total','Total Port_Grand Total']:
        if c not in wp: wp[c]=0

    wdf=pd.DataFrame({
        'Witel': wp.index,
        'On Going_Lop': wp['LoP_On Going'],
        'On Going_Port': wp['Total Port_On Going'],
        'Go Live_Lop':  wp['LoP_Go Live'],
        'Go Live_Port': wp['Total Port_Go Live'],
        'Total Lop':    wp['LoP_Grand Total'],
        'Total Port':   wp['Total Port_Grand Total'],
        '%':            wp['%'].round(1),
        'Penambahan GOLIVE H-1 vs HI': wp['Δ Go Live'],
        'RANK':         wp['RANK']
    })

    # ===== Header warna =====
    orange=['On Going_Lop','On Going_Port']
    green =['Go Live_Lop','Go Live_Port']
    purple=['Total Lop','Total Port']
    black =['%','Penambahan GOLIVE H-1 vs HI']
    yellow=['RANK']
    def header_css(bg, fontcolor='#fff'):
        return [('background-color',bg),('color',fontcolor),('font-weight','bold')]

    header_styles={}
    header_styles.update({c:header_css('#FFA500') for c in orange})       # orange
    header_styles.update({c:header_css('#2ECC71') for c in green})        # green
    header_styles.update({c:header_css('#9B59B6') for c in purple})       # purple
    header_styles.update({c:header_css('#000000','#ffffff') for c in black}) # black
    header_styles.update({c:header_css('#F7DC6F','#000000') for c in yellow}) # yellow

    def baris_style(r):
        if r['Witel']=='Grand Total':
            return ['background-color:#fff4b2;font-weight:bold']*len(r)
        if r['RANK']==1:
            return ['background-color:#d4f1f9']*len(r)
        return ['']*len(r)

    sty = (wdf.style
              .format({'%':'{:.1f}%','RANK':lambda v:'' if pd.isna(v) else f'{int(v)}',
                       **{c:'{:,.0f}' for c in orange+green+purple+['Penambahan GOLIVE H-1 vs HI']}})
              .apply(baris_style,axis=1)
              .set_table_styles(header_styles, axis=1)  # apply header colors
          )

    st.subheader("📌 Rekapitulasi per Witel")
    st.dataframe(sty, use_container_width=True, height=360)

    # ===== Datel Tabs =====
    st.subheader("🏆 Rekap per Datel")
    non_gt=wdf[wdf['Witel']!='Grand Total']['Witel']
    tabs=st.tabs(non_gt.tolist())
    for idx,w in enumerate(non_gt):
        with tabs[idx]:
            sub=dp[dp.index.get_level_values(0)==w].reset_index().sort_values('RANK')
            for c in ['Total Port_On Going','Total Port_Go Live','Total Port','%','RANK']:
                if c not in sub: sub[c]=0
            show=sub[['Datel','Total Port_On Going','Total Port_Go Live','Total Port','%','RANK']]
            st.dataframe(show.style.format({
                'Total Port_On Going':'{:,.0f}','Total Port_Go Live':'{:,.0f}',
                'Total Port':'{:,.0f}','%':'{:.1f}%','RANK':'{:,.0f}'
            }).applymap(lambda v:'background-color:#d4f1f9' if v==1 else '',subset=['RANK']),
            use_container_width=True,height=260)

            fig=px.bar(sub, x='Datel',
                       y=['Total Port_On Going','Total Port_Go Live'],
                       barmode='stack',
                       labels={'value':'Port','variable':'Status'},
                       title=f'Status Port di {w}',
                       color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig,use_container_width=True)

    # ===== Summary Charts =====
    st.subheader("🎯 Ringkasan Grafik")
    base=wdf[wdf['Witel']!='Grand Total'].sort_values('Total Port',ascending=False)
    c1,c2=st.columns(2)
    with c1:
        bar=px.bar(base,x='Witel',y='Total Port',text='Total Port',
                   color='Witel',title='Total Port per Witel',
                   color_discrete_sequence=px.colors.qualitative.Pastel)
        bar.update_traces(textposition='outside')
        st.plotly_chart(bar,use_container_width=True)
    with c2:
        st.plotly_chart(px.pie(base,names='Witel',values='Total Port',
                               title='Distribusi Total Port',hole=.35),
                        use_container_width=True)

# =============== UPLOAD ===============
else:
    st.title("📤 Upload Data Harian")
    last,last_hash=last_upload_info()
    if last: st.info(f"Terakhir upload: {last}")

    upl=st.file_uploader("Pilih file .xlsx",type="xlsx")
    if upl:
        if last_hash and md5(upl.getvalue())==last_hash:
            st.success("✅ Data sama dengan upload terakhir.")
        else:
            df=load_excel(upl)
            ok,msg=validate(df)
            if not ok:
                st.error(msg)
            else:
                save_file(LATEST_FILE,upl)
                record_history()
                st.success("✅ Upload berhasil & dashboard diperbarui!")
                st.balloons()
                st.dataframe(df.head(),use_container_width=True)

# Sidebar history
if page=="Upload Data" and os.path.exists(HISTORY_FILE):
    st.sidebar.subheader("Riwayat Upload")
    st.sidebar.dataframe(pd.read_csv(HISTORY_FILE).tail(5),hide_index=True)
