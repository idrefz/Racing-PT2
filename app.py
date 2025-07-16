import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

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

    # Pivot-style Table for Project Status - Updated to match your image
    st.subheader("\U0001F4CA Rekapitulasi Deployment per Witel (Format Visual)")
    
    # Create the pivot table similar to your image
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
    
    # Calculate additional columns like in your image
    pivot_table['%'] = (pivot_table[('Total Port', 'Go Live')] / pivot_table[('Total Port', 'Grand Total')] * 100).round(0)
    pivot_table['Penambahan GOLIVE H-1 vs HI'] = 0  # You would calculate this from your comparison logic
    pivot_table['RANK'] = pivot_table[('Total Port', 'Grand Total')].rank(ascending=False, method='min')
    
    # Format the table for display
    display_table = pd.DataFrame({
        'Witel': pivot_table.index,
        'On Going_Lop': pivot_table[('LoP', 'On Going')],
        'On Going_Port': pivot_table[('Total Port', 'On Going')],
        'Go Live_Lop': pivot_table[('LoP', 'Go Live')],
        'Go Live_Port': pivot_table[('Total Port', 'Go Live')],
        'Total Lop': pivot_table[('LoP', 'Grand Total')],
        'Total Port': pivot_table[('Total Port', 'Grand Total')],
        '%': pivot_table['%'],
        'Penambahan GOLIVE H-1 vs HI': pivot_table['Penambahan GOLIVE H-1 vs HI'],
        'RANK': pivot_table['RANK']
    })
    
    # Style the table
    def color_percent(val):
        color = 'green' if val >= 95 else 'orange' if val >= 85 else 'red'
        return f'color: {color}; font-weight: bold'
    
    styled_table = display_table.style \
        .applymap(color_percent, subset=['%']) \
        .background_gradient(subset=['Penambahan GOLIVE H-1 vs HI'], cmap='Blues') \
        .format({
            '%': '{:.0f}%',
            'On Going_Lop': '{:.0f}',
            'On Going_Port': '{:.0f}',
            'Go Live_Lop': '{:.0f}',
            'Go Live_Port': '{:.0f}',
            'Total Lop': '{:.0f}',
            'Total Port': '{:.0f}',
            'Penambahan GOLIVE H-1 vs HI': '{:.0f}',
            'RANK': '{:.0f}'
        })
    
    st.dataframe(styled_table, use_container_width=True, height=(len(display_table) + 1) * 35 + 3)
    
    # Add visualizations
    st.subheader("\U0001F4C8 Visualisasi Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Bar chart for Total Port by Witel
        fig1 = px.bar(
            display_table[display_table['Witel'] != 'Grand Total'].sort_values('Total Port', ascending=False),
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
