import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Config
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
YESTERDAY_FILE = os.path.join(DATA_FOLDER, "yesterday.xlsx")

# Buat folder jika belum ada
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# Helper to load Excel
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

# Helper: save as file
def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

# Helper: compare two dataframes
def compare_data(df_old, df_new):
    col_ticket = "Ticket ID"
    col_status = "Status Alokasi Alpro"

    df_old = df_old[[col_ticket, col_status]].dropna()
    df_new = df_new[[col_ticket, col_status]].dropna()

    old_set = set(df_old[col_ticket])
    new_set = set(df_new[col_ticket])

    new_tickets = new_set - old_set
    removed_tickets = old_set - new_set
    common = new_set & old_set

    df_old_common = df_old[df_old[col_ticket].isin(common)].set_index(col_ticket)
    df_new_common = df_new[df_new[col_ticket].isin(common)].set_index(col_ticket)

    status_diff = df_old_common.join(df_new_common, lsuffix="_H", rsuffix="_Hplus1")
    changed_status = status_diff[status_diff[f"{col_status}_H"] != status_diff[f"{col_status}_Hplus1"]]

    return {
        "total_old": len(df_old),
        "total_new": len(df_new),
        "new_count": len(new_tickets),
        "removed_count": len(removed_tickets),
        "changed_count": len(status_diff),
        "changed_df": changed_status.reset_index()
    }

# Style helpers
def highlight_columns(val):
    color = ""
    if val.name.startswith("On Going"):
        color = "background-color: #f28e2c"
    elif val.name.startswith("Go Live"):
        color = "background-color: #4caf50; color: white"
    elif val.name.startswith("Total"):
        color = "background-color: #1976d2; color: white"
    elif val.name.startswith("%"):
        color = "background-color: black; color: white"
    elif val.name.startswith("Penambahan"):
        color = "background-color: white"
    elif val.name.startswith("RANK"):
        color = "background-color: purple; color: white"
    return [color] * len(val)

# UI Starts
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")
st.title("📊 Delta Ticket Harian (H vs H+1)")

uploaded = st.file_uploader("Upload file hari ini (H+1)", type="xlsx")

if uploaded:
    df_new = load_excel(uploaded)
    df_new["LoP"] = 1
    df_new["Total Port"] = pd.to_numeric(df_new["Total Port"], errors="coerce")

    # Regional filter
    regional_list = df_new['Regional'].dropna().unique().tolist()
    selected_regional = st.selectbox("Pilih Regional", ["All"] + sorted(regional_list))

    if selected_regional != "All":
        df_new = df_new[df_new["Regional"] == selected_regional]

    if os.path.exists(LATEST_FILE):
        df_old = load_excel(LATEST_FILE)
        if selected_regional != "All":
            df_old = df_old[df_old["Regional"] == selected_regional]
        result = compare_data(df_old, df_new)

        st.subheader(":bar_chart: Ringkasan")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total H", result['total_old'])
        col2.metric("Total H+1", result['total_new'])
        col3.metric("Ticket Baru", result['new_count'])
        col4.metric("Ticket Hilang", result['removed_count'])
        col5.metric("Status Berubah", result['changed_count'])

        st.subheader(":pencil: Detail Perubahan Status")
        st.dataframe(result['changed_df'], use_container_width=True)
    else:
        st.warning("Tidak ditemukan data sebelumnya. Ini akan jadi referensi awal (H).")

    # Save backup kemarin
    if os.path.exists(LATEST_FILE):
        os.replace(LATEST_FILE, YESTERDAY_FILE)
    save_file(LATEST_FILE, uploaded)
    st.success("File berhasil disimpan sebagai referensi terbaru (H)")

    st.subheader("\U0001F4CA Tabel Komparasi Go Live Harian per Witel")

    if os.path.exists(YESTERDAY_FILE):
        df_yest = load_excel(YESTERDAY_FILE)
        df_yest["LoP"] = 1
        df_yest["Total Port"] = pd.to_numeric(df_yest["Total Port"], errors="coerce")

        if selected_regional != "All":
            df_yest = df_yest[df_yest["Regional"] == selected_regional]

        # Filter hanya Go Live
        h_today = df_new[df_new["Status Proyek"] == "Go Live"].groupby("Witel").agg({"LoP": "sum", "Total Port": "sum"}).rename(columns={"LoP": "GoLive_LoP_H", "Total Port": "GoLive_Port_H"})
        h_yest = df_yest[df_yest["Status Proyek"] == "Go Live"].groupby("Witel").agg({"LoP": "sum"}).rename(columns={"LoP": "GoLive_LoP_H-1"})

        on_going = df_new[df_new["Status Proyek"] == "On Going"].groupby("Witel").agg({"LoP": "sum", "Total Port": "sum"}).rename(columns={"LoP": "OnGoing_LoP", "Total Port": "OnGoing_Port"})
        total = df_new.groupby("Witel").agg({"LoP": "sum", "Total Port": "sum"}).rename(columns={"LoP": "Total_LoP", "Total Port": "Total_Port"})

        summary = pd.concat([on_going, h_today, h_yest, total], axis=1).fillna(0)
        summary["%"] = (summary["GoLive_Port_H"] / summary["Total_Port"] * 100).round(1).astype(str) + "%"
        summary["Penambahan GoLive"] = summary["GoLive_LoP_H"] - summary["GoLive_LoP_H-1"]
        summary["RANK"] = summary["Penambahan GoLive"].rank(ascending=False, method="min").astype(int)

        styled = summary.style.apply(highlight_columns, axis=1)
        st.dataframe(styled, use_container_width=True)
    else:
        st.info("Belum ada data H-1 untuk membandingkan progress harian.")

else:
    st.info("Silakan upload file Excel untuk diproses.")
