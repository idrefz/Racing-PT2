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
    # Create pivot tabletry:
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
    
    # Calculate additional metrics
    pivot_table['%'] = (pivot_table[('Total Port', 'Go Live')] / 
                        pivot_table[('Total Port', 'Grand Total')] * 100).round(0)
    
    # Calculate ranking EXCLUDING Grand Total
    witel_only = pivot_table[pivot_table.index != "Grand Total"]
    ranks = witel_only[('Total Port', 'Grand Total')].rank(ascending=False, method='min')
    pivot_table['RANK'] = pivot_table.index.map(ranks)
    
    # Add GOLIVE H-1 vs HI per Witel
    if os.path.exists(YESTERDAY_FILE):
        df_old = load_excel(YESTERDAY_FILE)
        result = compare_data(df_old, df_new)
        golive_port_added = result.get('golive_port_by_witel', pd.Series())
        pivot_table['Penambahan GOLIVE'] = pivot_table.index.map(
            lambda x: golive_port_added.get(x, 0)
        )
    else:
        pivot_table['Penambahan GOLIVE'] = 0

    # Create display table
    display_table = pd.DataFrame({
        'Witel': pivot_table.index,
        'On Going_Lop': pivot_table[('LoP', 'On Going')],
        'On Going_Port': pivot_table[('Total Port', 'On Going')],
        'Go Live_Lop': pivot_table[('LoP', 'Go Live')],
        'Go Live_Port': pivot_table[('Total Port', 'Go Live')],
        'Total Lop': pivot_table[('LoP', 'Grand Total')],
        'Total Port': pivot_table[('Total Port', 'Grand Total')],
        '%': pivot_table['%'],
        'Penambahan GOLIVE H-1 vs HI': pivot_table['Penambahan GOLIVE'],
        'RANK': pivot_table['RANK']
    })

    # Explicitly set Grand Total rank to empty
    display_table.loc[display_table['Witel'] == 'Grand Total', 'RANK'] = None

    # Format the table display
    st.dataframe(
        display_table.style.format({
            '%': '{:.0f}%',
            'On Going_Lop': '{:.0f}',
            'On Going_Port': '{:.0f}',
            'Go Live_Lop': '{:.0f}',
            'Go Live_Port': '{:.0f}',
            'Total Lop': '{:.0f}',
            'Total Port': '{:.0f}',
            'Penambahan GOLIVE H-1 vs HI': '{:.0f}',
            'RANK': '{:.0f}'
        }),
        use_container_width=True
    )

    # Breakdown by Datel with Witel filter
    st.subheader("📊 Breakdown per Datel")
    
    # Witel filter for Datel breakdown
    witel_list = df_new['Witel'].dropna().unique().tolist()
    selected_witel = st.selectbox("Pilih Witel untuk Breakdown Datel", ["All"] + sorted(witel_list))

    # Filter by selected Witel if not "All"
    if selected_witel != "All":
        df_filtered = df_new[df_new['Witel'] == selected_witel]
    else:
        df_filtered = df_new

    # Group by Datel and calculate metrics
    datel_table = df_filtered.groupby('Datel').agg(
        On_Going_LOP=('LoP', lambda x: x[df_filtered['Status Proyek'] == 'On Going'].sum()),
        On_Going_Port=('Total Port', lambda x: x[df_filtered['Status Proyek'] == 'On Going'].sum()),
        Go_Live_LOP=('LoP', lambda x: x[df_filtered['Status Proyek'] == 'Go Live'].sum()),
        Go_Live_Port=('Total Port', lambda x: x[df_filtered['Status Proyek'] == 'Go Live'].sum())
    ).reset_index()

    # Calculate totals and percentages
    datel_table['Total LOP'] = datel_table['On_Going_LOP'] + datel_table['Go_Live_LOP']
    datel_table['Total Port'] = datel_table['On_Going_Port'] + datel_table['Go_Live_Port']
    datel_table['%'] = (datel_table['Go_Live_Port'] / datel_table['Total Port'] * 100).round(0)

    # Add GOLIVE additions per Datel
    if os.path.exists(YESTERDAY_FILE):
        df_old = load_excel(YESTERDAY_FILE)
        result = compare_data(df_old, df_new)
        golive_port_added = result.get('golive_port_by_datel', pd.Series())
        datel_table['Penambahan'] = datel_table['Datel'].map(
            lambda x: golive_port_added.get(x, 0)
        )
    else:
        datel_table['Penambahan'] = 0

    # Calculate ranking
    datel_table['RANK'] = datel_table['Total Port'].rank(ascending=False, method='min')

    # Add Grand Total row
    grand_total = pd.DataFrame({
        'Datel': ['GRAND TOTAL'],
        'On_Going_LOP': [datel_table['On_Going_LOP'].sum()],
        'On_Going_Port': [datel_table['On_Going_Port'].sum()],
        'Go_Live_LOP': [datel_table['Go_Live_LOP'].sum()],
        'Go_Live_Port': [datel_table['Go_Live_Port'].sum()],
        'Total LOP': [datel_table['Total LOP'].sum()],
        'Total Port': [datel_table['Total Port'].sum()],
        '%': [(datel_table['Go_Live_Port'].sum() / datel_table['Total Port'].sum() * 100).round(0)],
        'Penambahan': [datel_table['Penambahan'].sum()],
        'RANK': [None]
    })

    datel_table = pd.concat([datel_table, grand_total], ignore_index=True)

    # Display the table
    st.dataframe(
        datel_table.style.format({
            'On_Going_LOP': '{:.0f}',
            'On_Going_Port': '{:.0f}',
            'Go_Live_LOP': '{:.0f}',
            'Go_Live_Port': '{:.0f}',
            'Total LOP': '{:.0f}',
            'Total Port': '{:.0f}',
            '%': '{:.0f}%',
            'Penambahan': '{:.0f}',
            'RANK': '{:.0f}'
        }),
        use_container_width=True,
        hide_index=True
    )

    # Visualizations
    st.subheader("\U0001F4C8 Visualisasi Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Bar chart for Total Port by Witel (excluding Grand Total)
        plot_df = display_table[display_table['Witel'] != 'Grand Total'].sort_values('Total Port', ascending=False)
        fig1 = px.bar(
            plot_df,
            x='Witel',
            y='Total Port',
            color='Witel',
            title='Total Port per Witel',
            text='Total Port'
        )
        fig1.update_traces(texttemplate='%{text:,}', textposition='outside')
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # Pie chart for Go Live vs On Going
        if not display_table[display_table['Witel'] == 'Grand Total'].empty:
            grand_total = display_table[display_table['Witel'] == 'Grand Total'].iloc[0]
            fig2 = px.pie(
                values=[grand_total['On Going_Port'], grand_total['Go Live_Port']],
                names=['On Going', 'Go Live'],
                title='Distribusi Port (Grand Total)',
                color=['On Going', 'Go Live'],
                color_discrete_map={'On Going':'red', 'Go Live':'green'}
            )
            st.plotly_chart(fig2, use_container_width=True)

# Navigation
st.sidebar.title("Menu Navigasi")
view_mode = st.sidebar.radio("Pilih Mode:", ["📊 Dashboard", "⬆️ Upload Data"])

# Main App Logic
if view_mode == "📊 Dashboard":
    if os.path.exists(LATEST_FILE):
        df_new = load_excel(LATEST_FILE)
        show_dashboard(df_new)
    else:
        st.warning("Belum ada data yang diupload. Silakan ke menu Upload Data.")

elif view_mode == "⬆️ Upload Data":
    st.title("⬆️ Upload Data Harian")
    
    # Show last upload info
    last_upload_time, last_hash = get_last_upload_info()
    if last_upload_time:
        st.info(f"Terakhir upload: {last_upload_time}")
    
    uploaded = st.file_uploader("Upload file hari ini (H+1)", type="xlsx")
    
    if uploaded:
        # Calculate hash of uploaded file
        uploaded_hash = hashlib.md5(uploaded.getvalue()).hexdigest()
        
        if last_hash and uploaded_hash == last_hash:
            st.warning("Data yang diupload sama dengan data terakhir. Tidak ada perubahan.")
            st.session_state.upload_complete = False
        else:
            df_new = load_excel(uploaded)
            df_new["LoP"] = 1
            df_new["Total Port"] = pd.to_numeric(df_new["Total Port"], errors="coerce")
            
            if os.path.exists(LATEST_FILE):
                df_old = load_excel(LATEST_FILE)
                result = compare_data(df_old, df_new)
                
                st.subheader("Perubahan dari Data Sebelumnya")
                col1, col2, col3 = st.columns(3)
                col1.metric("Ticket Baru", result['new_count'])
                col2.metric("Ticket Hilang", result['removed_count'])
                col3.metric("Status Berubah", result['changed_count'])
                
                if not result['changed_df'].empty:
                    st.dataframe(result['changed_df'], use_container_width=True)
                else:
                    st.info("Tidak ada perubahan status ticket")
            
            # Save files
            if os.path.exists(LATEST_FILE):
                os.replace(LATEST_FILE, YESTERDAY_FILE)
            save_file(LATEST_FILE, uploaded)
            record_upload_history()
            
            st.success("Data berhasil diupload dan disimpan!")
            st.session_state.upload_complete = True
            
            # Show preview
            st.subheader("Preview Data Terupload")
            st.dataframe(df_new.head())
    else:
        st.info("Silakan upload file Excel untuk memperbarui data.")
        st.session_state.upload_complete = False

    if st.session_state.upload_complete:
        if st.button("Lihat Dashboard"):
            st.session_state.view_mode = 'dashboard'
            st.experimental_rerun()
