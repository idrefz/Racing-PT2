import streamlit as st
import pandas as pd
import os
import plotly.express as px

# Config
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
YESTERDAY_FILE = os.path.join(DATA_FOLDER, "yesterday.xlsx")

# Create folder if not exists
if not os.path.exists(DATA_FOLDER):
    os.makedirs(DATA_FOLDER)

# Helper to load Excel
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

# Helper: save file
def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

# Helper: compare two dataframes
def compare_data(df_old, df_new):
    col_ticket = "Datel"
    col_status = "Status Proyek"
    col_port = "Total Port"

    # Ensure we have the needed columns
    df_old = df_old[[col_ticket, col_status, col_port]].dropna()
    df_new = df_new[[col_ticket, col_status, col_port]].dropna()

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
    
    # Calculate total port added from status changes
    golive_port_added = changed_to_golive[f"{col_port}_Hplus1"].sum()

    changed_status = status_diff[status_diff[f"{col_status}_H"] != status_diff[f"{col_status}_Hplus1"]]

    return {
        "total_old": len(df_old),
        "total_new": len(df_new),
        "new_count": len(new_tickets),
        "removed_count": len(removed_tickets),
        "changed_count": len(changed_status),
        "changed_df": changed_status.reset_index(),
        "golive_port_added": golive_port_added
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
    
    # Create pivot table
    pivot_table = pd.pivot_table(
        df_new,
        values=["LoP", "Total Port"],
        index="Witel",
        columns="Status Proyek",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    )
    
    # Calculate additional metrics
    pivot_table['%'] = (pivot_table[('Total Port', 'Go Live')] / 
                        pivot_table[('Total Port', 'Grand Total')] * 100).round(0)
    
    # Calculate ranking EXCLUDING Grand Total
    witel_only = pivot_table[pivot_table.index != "Grand Total"]
    ranks = witel_only[('Total Port', 'Grand Total')].rank(ascending=False, method='min')
    pivot_table['RANK'] = pivot_table.index.map(ranks)
    
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
        'Penambahan GOLIVE H-1 vs HI': result['golive_port_added'] if os.path.exists(LATEST_FILE) else 0,
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
            'RANK': '{:.0f}' if pd.notna(display_table['RANK']).any() else ''
        }),
        use_container_width=True
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

else:
    st.info("Silakan upload file Excel untuk diproses.")
