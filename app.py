import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import shutil

# ───────────────────────── CONFIG ──────────────────────────
DATA_FOLDER  = "data_daily_uploads"
LATEST_FILE  = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")
os.makedirs(DATA_FOLDER, exist_ok=True)

# ──────────────────────── HELPERS ──────────────────────────
@st.cache_data
def load_excel(path):              # baca Excel sekali, cache
    return pd.read_excel(path)

def save_file(path, uploaded):
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())

def get_file_hash(binary: bytes):
    return hashlib.md5(binary).hexdigest()

def compare_with_previous(df_curr: pd.DataFrame) -> dict[str, int]:
    """hitung selisih port Go Live dengan upload sebelumnya"""
    delta = {}
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        if len(hist) >= 2:
            prev_hash = hist.iloc[-2]["file_hash"]
            prev_path = os.path.join(DATA_FOLDER, f"previous_{prev_hash}.xlsx")
            if os.path.exists(prev_path):
                df_prev = pd.read_excel(prev_path)
                go_curr = df_curr[df_curr["Status Proyek"]=="Go Live"].groupby("Witel")["Total Port"].sum()
                go_prev = df_prev[df_prev["Status Proyek"]=="Go Live"].groupby("Witel")["Total Port"].sum()
                for w in go_curr.index:
                    delta[w] = int(go_curr[w] - go_prev.get(w, 0))
    return delta

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_hash = get_file_hash(open(LATEST_FILE,"rb").read()) if os.path.exists(LATEST_FILE) else ""
    if os.path.exists(HISTORY_FILE):
        hist = pd.read_csv(HISTORY_FILE)
        if not hist.empty:
            last_hash = hist.iloc[-1]["file_hash"]
            shutil.copy(LATEST_FILE, os.path.join(DATA_FOLDER, f"previous_{last_hash}.xlsx"))
            hist = pd.concat([hist, pd.DataFrame([[now,new_hash]], columns=["timestamp","file_hash"])])
    else:
        hist = pd.DataFrame([[now,new_hash]], columns=["timestamp","file_hash"])
    hist.to_csv(HISTORY_FILE, index=False)

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        h = pd.read_csv(HISTORY_FILE)
        if not h.empty:
            return h.iloc[-1]["timestamp"], h.iloc[-1]["file_hash"]
    return None, None

def validate_data(df):
    req = ["Regional","Witel","Status Proyek","Total Port","Datel","Ticket ID","Nama Proyek"]
    miss = [c for c in req if c not in df.columns]
    if miss:
        return False, f"Kolom hilang: {', '.join(miss)}"
    try:
        df["Total Port"] = pd.to_numeric(df["Total Port"], errors="raise")
    except ValueError:
        return False, "Kolom 'Total Port' harus numerik"
    return True,"Data valid"

# ────────────── PIVOT & METRIC CALCULATION ────────────────
def create_pivots(df: pd.DataFrame):
    delta = compare_with_previous(df)
    df["LoP"] = 1

    # Witel pivot
    wp = pd.pivot_table(
        df, values=["LoP","Total Port"], index="Witel", columns="Status Proyek",
        aggfunc="sum", fill_value=0, margins=True, margins_name="Grand Total"
    )
    wp.columns = ['_'.join(col) for col in wp.columns]
    wp["% Go Live"]     = (wp.get("Total Port_Go Live",0) / wp["Total Port_Grand Total"]).fillna(0)*100
    wp["Δ Go Live"]     = [delta.get(w,0) for w in wp.index]
    wp["RANK"]          = wp["Total Port_Grand Total"].rank(ascending=False, method="dense").astype("Int64")
    wp.loc["Grand Total","RANK"] = pd.NA

    # Datel pivot
    dp = pd.pivot_table(
        df, values=["LoP","Total Port"], index=["Witel","Datel"], columns="Status Proyek",
        aggfunc="sum", fill_value=0
    )
    dp.columns = ['_'.join(col) for col in dp.columns]
    dp["Total Port"] = dp.get("Total Port_On Going",0)+dp.get("Total Port_Go Live",0)
    dp["% Go Live"]  = (dp.get("Total Port_Go Live",0)/dp["Total Port"]).fillna(0)*100
    dp["RANK"]       = dp.groupby(level=0)["Total Port"].rank(ascending=False, method="min").astype("Int64")

    return wp.reset_index(), dp.reset_index()

# ───────────────────────── UI SETUP ────────────────────────
st.set_page_config("Delta Ticket Harian", layout="wide")
st.sidebar.title("Navigasi")
page = st.sidebar.radio("Pilih Mode", ["Dashboard","Upload Data"])

# ============ DASHBOARD ============ #
if page=="Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")
    if not os.path.exists(LATEST_FILE):
        st.info("Belum ada data. Silakan upload terlebih dulu.")
        st.stop()

    df = load_excel(LATEST_FILE)

    # Regional filter
    regions = ["All"]+sorted(df["Regional"].dropna().unique())
    sel_reg = st.selectbox("Pilih Regional", regions)
    if sel_reg!="All":
        df=df[df["Regional"]==sel_reg]

    # ---- Detail Go Live ----
    st.subheader("🚀 Detail Proyek Go Live")
    st.dataframe(
        df[df["Status Proyek"]=="Go Live"][["Witel","Datel","Nama Proyek","Total Port","Status Proyek"]],
        use_container_width=True, height=300
    )

    # ---- Pivots ----
    witel_df, datel_df = create_pivots(df)

    # ---- Tabel Witel ----
    st.subheader("📌 Rekap per Witel")
    def style_witel(r):
        if r["Witel"]=="Grand Total":
            return ['background-color:#ffeaa7;font-weight:bold']*len(r)
        if r["RANK"]==1:
            return ['background-color:#c7ecee']*len(r)
        return ['']*len(r)

    st.dataframe(
        witel_df.style
            .format({
                "Total Port_On Going":"{:,.0f}",
                "Total Port_Go Live":"{:,.0f}",
                "Total Port_Grand Total":"{:,.0f}",
                "% Go Live":"{:.1f}%",
                "Δ Go Live":"{:,.0f}",
                "RANK":lambda v: "" if pd.isna(v) else f"{int(v)}"
            })
            .apply(style_witel, axis=1),
        use_container_width=True, height=350
    )

    # ---- Tabel & Grafik Datel per Witel ----
    st.subheader("🏆 Rekap per Datel (tabs)")
    for w in witel_df.loc[witel_df["Witel"]!="Grand Total","Witel"]:
        with st.expander(f"🔸 {w}", expanded=False):
            sub = datel_df[datel_df["Witel"]==w].sort_values("RANK")
            st.dataframe(
                sub.style.format({
                    "Total Port_On Going":"{:,.0f}",
                    "Total Port_Go Live":"{:,.0f}",
                    "Total Port":"{:,.0f}",
                    "% Go Live":"{:.1f}%",
                    "RANK":"{:,.0f}"
                }).applymap(
                    lambda v: "background-color:#c7ecee" if v==1 else "",
                    subset=["RANK"]
                ),
                use_container_width=True, height=260
            )

            # pastikan kolom ada
            for col in ["Total Port_On Going","Total Port_Go Live"]:
                if col not in sub.columns:
                    sub[col]=0

            fig = px.bar(
                sub,
                x="Datel",
                y=["Total Port_On Going","Total Port_Go Live"],
                barmode="stack",
                title=f"Status Port di {w}",
                labels={"value":"Jumlah Port","variable":"Status"}
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---- Ringkasan Grafik ----
    st.subheader("🎯 Ringkasan Witel")
    top_plot = witel_df[witel_df["Witel"]!="Grand Total"].sort_values("Total Port_Grand Total", ascending=False)
    col1,col2 = st.columns(2)
    with col1:
        bar_tot = px.bar(top_plot, x="Witel", y="Total Port_Grand Total",
                         text="Total Port_Grand Total", title="Total Port per Witel",
                         color="Witel", color_discrete_sequence=px.colors.qualitative.Pastel)
        bar_tot.update_traces(textposition="outside")
        st.plotly_chart(bar_tot,use_container_width=True)
    with col2:
        pie = px.pie(top_plot, names="Witel", values="Total Port_Grand Total",
                     title="Distribusi Total Port", hole=.4)
        st.plotly_chart(pie,use_container_width=True)

# ============ UPLOAD ============ #
else:
    st.title("📤 Upload Data Harian")
    last_ts,last_hash = get_last_upload_info()
    if last_ts: st.info(f"Terakhir upload: {last_ts}")

    upl = st.file_uploader("Pilih file Excel (.xlsx)", type="xlsx")
    if upl:
        new_hash = get_file_hash(upl.getvalue())
        if last_hash and new_hash==last_hash:
            st.success("✅ Data sama dengan upload terakhir.")
            st.stop()

        df_new = load_excel(upl)
        ok,msg = validate_data(df_new)
        if not ok:
            st.error(msg)
        else:
            save_file(LATEST_FILE, upl)
            record_upload_history()
            st.success("✅ Upload berhasil & dashboard diperbarui!")
            st.balloons()

            st.subheader("Preview 5 baris")
            st.dataframe(df_new.head(),use_container_width=True)

            st.metric("Total Port", int(df_new["Total Port"].sum()))
            st.metric("Total Proyek", len(df_new))

# ───────────── Sidebar riwayat ─────────────
if page=="Upload Data" and os.path.exists(HISTORY_FILE):
    st.sidebar.subheader("Riwayat Upload")
    st.sidebar.dataframe(pd.read_csv(HISTORY_FILE).tail(5), hide_index=True)
