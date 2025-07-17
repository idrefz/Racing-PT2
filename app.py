import streamlit as st, pandas as pd, os, hashlib, shutil
from datetime import datetime
import plotly.express as px

# ── CONFIG ───────────────────────────
FOLDER="data_daily_uploads"
LATEST=os.path.join(FOLDER,"latest.xlsx")
HIST  =os.path.join(FOLDER,"upload_history.csv")
os.makedirs(FOLDER,exist_ok=True)

# ── BASIC FUNCS ──────────────────────
@st.cache_data
def load_xlsx(p): return pd.read_excel(p)
def md5(b): return hashlib.md5(b).hexdigest()
def save_file(p,u): open(p,"wb").write(u.getbuffer())

def last_info():
    if os.path.exists(HIST):
        h=pd.read_csv(HIST); 
        if not h.empty: return h.iloc[-1]['timestamp'],h.iloc[-1]['file_hash']
    return None,None

def record_hist():
    ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new=md5(open(LATEST,'rb').read())
    if os.path.exists(HIST):
        h=pd.read_csv(HIST)
        if not h.empty:
            prev_hash=h.iloc[-1]['file_hash']
            shutil.copy(LATEST,os.path.join(FOLDER,f"previous_{prev_hash}.xlsx"))
            h=pd.concat([h,pd.DataFrame([[ts,new]],columns=['timestamp','file_hash'])])
    else:
        h=pd.DataFrame([[ts,new]],columns=['timestamp','file_hash'])
    h.to_csv(HIST,index=False)

def load_prev():
    if os.path.exists(HIST):
        h=pd.read_csv(HIST)
        if len(h)>=2:
            ph=h.iloc[-2]['file_hash']
            fp=os.path.join(FOLDER,f"previous_{ph}.xlsx")
            if os.path.exists(fp): return pd.read_excel(fp)
    return pd.DataFrame()

# ── VALIDATE ────────────────────────
def validate(df):
    need=['Regional','Witel','Status Proyek','Total Port','Datel','Ticket ID','Nama Proyek']
    miss=[c for c in need if c not in df.columns]
    if miss: return False,f"Kolom hilang: {', '.join(miss)}"
    try: df['Total Port']=pd.to_numeric(df['Total Port'],errors='raise')
    except: return False,"'Total Port' harus numerik"
    return True,"OK"

# ── PIVOTS & RANK ───────────────────
def pivots(df_now, df_prev):
    df_now['LoP']=1
    wp=df_now.pivot_table(values=['LoP','Total Port'],index='Witel',columns='Status Proyek',
                          aggfunc='sum',fill_value=0,margins=True,margins_name='Grand Total')
    wp.columns=['_'.join(c) for c in wp.columns]
    wp['%']=(wp.get('Total Port_Go Live',0)/wp['Total Port_Grand Total']).fillna(0)*100
    ranks=wp.loc[wp.index!='Grand Total','Total Port_Grand Total'].rank(ascending=False,
                                                                        method='dense').astype('Int64')
    wp['RANK']=pd.NA; wp.loc[ranks.index,'RANK']=ranks
    delta={}
    if not df_prev.empty:
        n=df_now[df_now['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
        p=df_prev[df_prev['Status Proyek']=='Go Live'].groupby('Witel')['Total Port'].sum()
        for w in n.index: delta[w]=int(n[w]-p.get(w,0))
    wp['Δ Go Live']=[delta.get(w,0) for w in wp.index]

    dp=df_now.pivot_table(values=['LoP','Total Port'],index=['Witel','Datel'],columns='Status Proyek',
                          aggfunc='sum',fill_value=0)
    dp.columns=['_'.join(c) for c in dp.columns]
    dp['Total Port']=dp.get('Total Port_On Going',0)+dp.get('Total Port_Go Live',0)
    dp['%']=(dp.get('Total Port_Go Live',0)/dp['Total Port']).fillna(0)*100
    dp['RANK']=dp.groupby(level=0)['Total Port'].rank(ascending=False,method='min').astype('Int64')
    return wp,dp

# ── FORMATTER SAFE ──────────────────
def int_fmt(v): 
    try: return f"{int(v):,}"
    except: return ""
def pct_fmt(v):
    try: return f"{float(v):.1f}%"
    except: return ""

# ── UI ──────────────────────────────
st.set_page_config("Delta Ticket Harian",layout="wide")
page=st.sidebar.radio("Mode",["Dashboard","Upload Data"])

# ========== DASHBOARD ==========
if page=="Dashboard":
    st.title("📊 Dashboard Deployment PT2 IHLD")
    if not os.path.exists(LATEST): st.stop()

    df_now=load_xlsx(LATEST); df_prev=load_prev()

    reg=st.selectbox("Filter Regional",['All']+sorted(df_now['Regional'].dropna().unique()))
    if reg!="All":
        df_now=df_now[df_now['Regional']==reg]; 
        df_prev=df_prev[df_prev['Regional']==reg] if not df_prev.empty else df_prev

    # Go-Live HI
    if not df_prev.empty:
        old=set(df_prev[df_prev['Status Proyek']=='Go Live']['Ticket ID'])
        hi=df_now[(df_now['Status Proyek']=='Go Live') & (~df_now['Ticket ID'].isin(old))]
    else:
        hi=df_now[df_now['Status Proyek']=='Go Live']
    st.subheader("🚀 Proyek Go Live – HI (baru)")
    st.dataframe(hi[['Witel','Datel','Nama Proyek','Total Port','Ticket ID']],
                 use_container_width=True,height=220)

    wp,dp=pivots(df_now,df_prev)
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
        '%':            wp['%'],
        'Penambahan GOLIVE H-1 vs HI': wp['Δ Go Live'],
        'RANK':         wp['RANK']
    })

    st.subheader("📌 Rekapitulasi per Witel")

    # Header style list
    palette={'orange':'#FFA500','green':'#2ECC71','purple':'#9B59B6',
             'black':'#000000','yellow':'#F7DC6F'}
    groups={'orange':['On Going_Lop','On Going_Port'],
            'green' :['Go Live_Lop','Go Live_Port'],
            'purple':['Total Lop','Total Port'],
            'black' :['%','Penambahan GOLIVE H-1 vs HI'],
            'yellow':['RANK']}
    tbl_styles=[]
    for color,cols in groups.items():
        for c in cols:
            cls=c.replace(" ","_")
            tbl_styles.append({'selector':f"th.col_heading.level0.{cls}",
                               'props':[('background-color',palette[color]),
                                        ('color','#fff' if color!='yellow' else '#000'),
                                        ('font-weight','bold')]})

    def row_hl(r):
        if r['Witel']=='Grand Total': return ['background-color:#fff4b2;font-weight:bold']*len(r)
        if r['RANK']==1:              return ['background-color:#d4f1f9']*len(r)
        return ['']*len(r)

    sty=(wdf.style
            .format({c:int_fmt for c in wdf.columns if c not in ['%','RANK','Witel']})
            .format({'%':pct_fmt,'RANK':lambda v:'' if pd.isna(v) else int(v)})
            .apply(row_hl,axis=1)
            .set_table_styles(tbl_styles))
    st.dataframe(sty,use_container_width=True,height=360)

    # Datel tabs
    st.subheader("🏆 Rekap per Datel")
    for w in wdf[wdf['Witel']!='Grand Total']['Witel']:
        with st.expander(f"📌 {w}",expanded=False):
            sub=dp[dp.index.get_level_values(0)==w].reset_index().sort_values('RANK')
            for c in ['Total Port_On Going','Total Port_Go Live','Total Port','%','RANK']:
                if c not in sub: sub[c]=0
            st.dataframe(sub[['Datel','Total Port_On Going','Total Port_Go Live',
                              'Total Port','%','RANK']].style.format({
                                  'Total Port_On Going':int_fmt,'Total Port_Go Live':int_fmt,
                                  'Total Port':int_fmt,'%':pct_fmt,'RANK':lambda v:int(v)}),
                         use_container_width=True,height=240)
            fig=px.bar(sub,x='Datel',
                       y=['Total Port_On Going','Total Port_Go Live'],barmode='stack',
                       labels={'value':'Port','variable':'Status'},
                       title=f'Status Port di {w}',
                       color_discrete_sequence=px.colors.qualitative.Set2)
            st.plotly_chart(fig,use_container_width=True)

# ========== UPLOAD ==========
else:
    st.title("📤 Upload Data Harian")
    last,lh=last_info()
    if last: st.info(f"Terakhir upload: {last}")

    upl=st.file_uploader("Pilih file .xlsx",type="xlsx")
    if upl:
        if lh and md5(upl.getvalue())==lh:
            st.success("✅ Data sama dengan upload terakhir.")
        else:
            df=load_xlsx(upl); ok,msg=validate(df)
            if not ok: st.error(msg)
            else:
                save_file(LATEST,upl); record_hist()
                st.success("✅ Upload berhasil & dashboard diperbarui!"); st.balloons()
                st.dataframe(df.head(),use_container_width=True)

# Sidebar history
if page=="Upload Data" and os.path.exists(HIST):
    st.sidebar.subheader("Riwayat Upload")
    st.sidebar.dataframe(pd.read_csv(HIST).tail(5),hide_index=True)
