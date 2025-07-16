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

    # Pivot-style Table for Project Status
    st.subheader("\U0001F4CA Rekapitulasi Deployment per Witel")

    df_filtered = df_new.copy()
    df_filtered["LoP"] = 1
    df_filtered["Total Port"] = pd.to_numeric(df_filtered["Total Port"], errors="coerce")

    pivot_table = pd.pivot_table(
        df_filtered,
        values=["LoP", "Total Port"],
        index="Witel",
        columns="Status Proyek",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Total"
    )

    st.dataframe(pivot_table, use_container_width=True)

else:
    st.info("Silakan upload file Excel untuk diproses.")
