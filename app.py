import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib

# Config
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
YESTERDAY_FILE = os.path.join(DATA_FOLDER, "yesterday.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")

# Create folder if not exists
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# Initialize session state
if 'view_mode' not in st.session_state:
    st.session_state.view_mode = 'dashboard'
if 'upload_complete' not in st.session_state:
    st.session_state.upload_complete = False

# Helper functions
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(file_path):
    with open(file_path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_hash = get_file_hash(LATEST_FILE) if os.path.exists(LATEST_FILE) else ""
    
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
    else:
        history_df = pd.DataFrame(columns=['timestamp', 'file_hash'])
    
    history_df = pd.concat([
        history_df,
        pd.DataFrame([[now, file_hash]], columns=['timestamp', 'file_hash'])
    ], ignore_index=True)
    
    history_df.to_csv(HISTORY_FILE, index=False)

def get_last_upload_info():
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty:
            return history_df.iloc[-1]['timestamp'], history_df.iloc[-1]['file_hash']
    return None, None

def compare_data(df_old, df_new):
    col_ticket = "Ticket ID"
    col_status = "Status Proyek"
    col_port = "Total Port"
    col_witel = "Witel"
    col_datel = "Datel"
    col_project = "Nama Proyek"

    # Ensure we have the needed columns
    required_cols = [col_ticket, col_status, col_port, col_witel, col_datel, col_project]
    df_old = df_old[required_cols].dropna(subset=[col_ticket, col_status, col_port])
    df_new = df_new[required_cols].dropna(subset=[col_ticket, col_status, col_port])

    old_set = set(df_old[col_ticket])
    new_set = set(df_new[col_ticket])

    new_tickets = new_set - old_set
    removed_tickets = old_set - new_set
    common = new_set & old_set

    df_old_common = df_old[df_old[col_ticket].isin(common)].set_index(col_ticket)
    df_new_common = df_new[df_new[col_ticket].isin(common)].set_index(col_ticket)

    status_diff = df_old_common.join(df_new_common, lsuffix="_H", rsuffix="_Hplus1")
    
    # Find tickets that changed to Go Live
    changed_to_golive = status_diff[
        (status_diff[f"{col_status}_H"] != "Go Live") & 
        (status_diff[f"{col_status}_Hplus1"] == "Go Live")
    ]
    
    # Calculate total port added per Witel and Datel
    golive_port_by_witel = changed_to_golive.groupby(f"{col_witel}_Hplus1")[f"{col_port}_Hplus1"].sum()
    golive_port_by_datel = changed_to_golive.groupby(f"{col_datel}_Hplus1")[f"{col_port}_Hplus1"].sum()
    total_golive_port_added = changed_to_golive[f"{col_port}_Hplus1"].sum()

    # Prepare detailed changes
    changed_status = status_diff[status_diff[f"{col_status}_H"] != status_diff[f"{col_status}_Hplus1"]]
    detailed_changes = changed_status.reset_index()[[
        f"{col_ticket}", 
        f"{col_witel}_Hplus1",
        f"{col_datel}_Hplus1",
        f"{col_project}_Hplus1",
        f"{col_status}_H", 
        f"{col_port}_H",
        f"{col_status}_Hplus1"
    ]].rename(columns={
        f"{col_witel}_Hplus1": "Witel",
        f"{col_datel}_Hplus1": "Datel",
        f"{col_project}_Hplus1": "Nama Proyek",
        f"{col_status}_H": "Status Proyek H",
        f"{col_port}_H": "Total Port H",
        f"{col_status}_Hplus1": "Status Proyek H+1"
    })

    return {
        "total_old": len(df_old),
        "total_new": len(df_new),
        "new_count": len(new_tickets),
        "removed_count": len(removed_tickets),
        "changed_count": len(changed_status),
        "changed_df": detailed_changes,
        "golive_port_by_witel": golive_port_by_witel,
        "golive_port_by_datel": golive_port_by_datel,
        "total_golive_port_added": total_golive_port_added
    }

def show_dashboard(df_new):
    st.title("📊 Dashboard Monitoring Harian")
    
    # Check for required columns
    required_columns = ['Regional', 'Witel', 'Status Proyek', 'LoP', 'Total Port']
    missing_columns = [col for col in required_columns if col not in df_new.columns]
    
    if missing_columns:
        st.error(f"Data tidak valid. Kolom yang dibutuhkan tidak ditemukan: {', '.join(missing_columns)}")
        return
    
    # Regional filter
    regional_list = df_new['Regional'].dropna().unique().tolist()
    selected_regional = st.selectbox("Pilih Regional", ["All"] + sorted(regional_list))

    if selected_regional != "All":
        df_new = df_new[df_new["Regional"] == selected_regional]

    if os.path.exists(YESTERDAY_FILE):
        df_old = load_excel(YESTERDAY_FILE)
        if selected_regional != "All":
            df_old = df_old[df_old["Regional"] == selected_regional]
        
        result = compare_data(df_old, df_new)

        st.subheader(":bar_chart: Ringkasan Harian")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total H", result['total_old'])
        col2.metric("Total H+1", result['total_new'])
        col3.metric("Ticket Baru", result['new_count'])
        col4.metric("Ticket Hilang", result['removed_count'])
        col5.metric("Status Berubah", result['changed_count'])

    # Pivot-style Table for Project Status
    st.subheader("\U0001F4CA Rekapitulasi Deployment per Witel")
    
    try:
        # Create pivot table with proper column checks
        pivot_columns = ['LoP', 'Total Port']
        pivot_index = 'Witel'
        pivot_cols = 'Status Proyek'
        
        # Verify all pivot columns exist
        for col in pivot_columns + [pivot_index, pivot_cols]:
            if col not in df_new.columns:
                raise KeyError(f"Kolom '{col}' tidak ditemukan dalam data")
        
        pivot_table = pd.pivot_table(
            df_new,
            values=pivot_columns,
            index=pivot_index,
            columns=pivot_cols,
            aggfunc="sum",
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )
        
        # Rest of your pivot table processing...
        
    except KeyError as e:
        st.error(f"Error dalam memproses data: {str(e)}")
        st.warning("Pastikan file yang diupload memiliki format yang benar dengan kolom yang diperlukan")
        return
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
        return

    # Continue with the rest of your dashboard code...
