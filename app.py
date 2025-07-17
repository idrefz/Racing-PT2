import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime
import hashlib

# Configuration
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
YESTERDAY_FILE = os.path.join(DATA_FOLDER, "yesterday.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")

# Create folder if not exists
os.makedirs(DATA_FOLDER, exist_ok=True)

# Helper Functions
@st.cache_data
def load_excel(file):
    return pd.read_excel(file)

def save_file(path, uploaded_file):
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(file_content):
    return hashlib.md5(file_content).hexdigest()

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_hash = get_file_hash(open(LATEST_FILE, "rb").read()) if os.path.exists(LATEST_FILE) else ""
    
    history_df = pd.DataFrame(columns=['timestamp', 'file_hash'])
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
    
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

def validate_data(df):
    required_columns = ['Regional', 'Witel', 'Status Proyek', 'Total Port', 'Datel', 'Ticket ID']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        return False, f"Kolom yang dibutuhkan tidak ditemukan: {', '.join(missing_columns)}"
    
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except ValueError:
        return False, "Kolom 'Total Port' harus berisi nilai numerik"
    
    return True, "Data valid"

def compare_data(df_old, df_new):
    try:
        comparison_cols = ['Ticket ID', 'Status Proyek', 'Total Port', 'Witel', 'Datel']
        df_old = df_old[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])
        df_new = df_new[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])

        old_tickets = set(df_old['Ticket ID'])
        new_tickets = set(df_new['Ticket ID'])

        new_projects = new_tickets - old_tickets
        removed_projects = old_tickets - new_tickets
        common_tickets = new_tickets & old_tickets

        df_old_common = df_old[df_old['Ticket ID'].isin(common_tickets)].set_index('Ticket ID')
        df_new_common = df_new[df_new['Ticket ID'].isin(common_tickets)].set_index('Ticket ID')
        
        status_diff = df_old_common.join(df_new_common, lsuffix='_H', rsuffix='_H+1')
        
        changed_to_golive = status_diff[
            (status_diff['Status Proyek_H'] != 'Go Live') & 
            (status_diff['Status Proyek_H+1'] == 'Go Live')
        ]
        
        golive_additions = changed_to_golive.groupby('Witel_H+1')['Total Port_H+1'].sum()
        total_golive_added = golive_additions.sum()

        changed_status = status_diff[status_diff['Status Proyek_H'] != status_diff['Status Proyek_H+1']]
        
        return {
            'new_count': len(new_projects),
            'removed_count': len(removed_projects),
            'changed_count': len(changed_status),
            'golive_additions': golive_additions,
            'total_golive_added': total_golive_added,
            'changes': status_diff.reset_index()
        }
    except Exception as e:
        st.error(f"Error comparing data: {str(e)}")
        return None

def create_pivot_table(df):
    try:
        df['LoP'] = 1  # Each row is one project
        
        pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index='Witel',
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        )
        
        if ('Total Port', 'Go Live') in pivot.columns:
            pivot['%'] = (pivot[('Total Port', 'Go Live')] / 
                         pivot[('Total Port', 'Grand Total')]).fillna(0) * 100
        else:
            pivot['%'] = 0
            
        pivot['RANK'] = pivot[('Total Port', 'Grand Total')].rank(ascending=False, method='min')
        pivot.loc['Grand Total', 'RANK'] = None
        
        return pivot
    except Exception as e:
        st.error(f"Error creating pivot table: {str(e)}")
        return None

# UI Setup
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")
st.title("📊 Delta Ticket Harian (H vs H+1)")

# File Upload Section
uploaded_file = st.file_uploader("Upload file hari ini (H+1)", type="xlsx")

if uploaded_file:
    # Validate and process uploaded file
    try:
        df_new = load_excel(uploaded_file)
        is_valid, msg = validate_data(df_new)
        
        if not is_valid:
            st.error(msg)
        else:
            # Process numeric columns
            df_new['Total Port'] = pd.to_numeric(df_new['Total Port'], errors='coerce')
            
            # Regional filter
            regional_list = ['All'] + sorted(df_new['Regional'].dropna().unique().tolist())
            selected_region = st.selectbox("Pilih Regional", regional_list)
            
            if selected_region != 'All':
                df_new = df_new[df_new['Regional'] == selected_region]
            
            # Compare with previous data if available
            if os.path.exists(LATEST_FILE):
                df_old = load_excel(LATEST_FILE)
                if selected_region != 'All':
                    df_old = df_old[df_old['Regional'] == selected_region]
                
                comparison = compare_data(df_old, df_new)
                
                if comparison:
                    st.subheader(":bar_chart: Ringkasan Perubahan")
                    cols = st.columns(5)
                    cols[0].metric("Total H", len(df_old))
                    cols[1].metric("Total H+1", len(df_new))
                    cols[2].metric("Ticket Baru", comparison['new_count'])
                    cols[3].metric("Ticket Hilang", comparison['removed_count'])
                    cols[4].metric("Status Berubah", comparison['changed_count'])
                    
                    st.subheader(":pencil: Detail Perubahan Status")
                    st.dataframe(
                        comparison['changes'][[
                            'Witel_H+1', 'Datel_H+1', 'Status Proyek_H', 
                            'Total Port_H', 'Status Proyek_H+1'
                        ]].rename(columns={
                            'Witel_H+1': 'Witel',
                            'Datel_H+1': 'Datel',
                            'Status Proyek_H': 'Status H',
                            'Total Port_H': 'Total Port H',
                            'Status Proyek_H+1': 'Status H+1'
                        }),
                        use_container_width=True
                    )
            else:
                st.warning("Tidak ditemukan data sebelumnya. Ini akan jadi referensi awal (H).")
            
            # Save files
            if os.path.exists(LATEST_FILE):
                os.replace(LATEST_FILE, YESTERDAY_FILE)
            save_file(LATEST_FILE, uploaded_file)
            record_upload_history()
            st.success("File berhasil disimpan sebagai referensi terbaru (H)")
            
            # Create and display pivot table
            st.subheader("\U0001F4CA Rekapitulasi Deployment per Witel")
            pivot_table = create_pivot_table(df_new)
            
            if pivot_table is not None:
                # Prepare display table
                display_data = {
                    'Witel': pivot_table.index,
                    'On Going_Lop': pivot_table[('LoP', 'On Going')],
                    'On Going_Port': pivot_table[('Total Port', 'On Going')],
                    'Go Live_Lop': pivot_table[('LoP', 'Go Live')],
                    'Go Live_Port': pivot_table[('Total Port', 'Go Live')],
                    'Total Lop': pivot_table[('LoP', 'Grand Total')],
                    'Total Port': pivot_table[('Total Port', 'Grand Total')],
                    '%': pivot_table['%'],
                    'RANK': pivot_table['RANK']
                }
                
                # Add GOLIVE additions if available
                if 'comparison' in locals() and comparison:
                    display_data['Penambahan GOLIVE H-1 vs HI'] = [
                        comparison['golive_additions'].get(witel, 0) 
                        if witel != 'Grand Total' 
                        else comparison['total_golive_added']
                        for witel in pivot_table.index
                    ]
                else:
                    display_data['Penambahan GOLIVE H-1 vs HI'] = 0
                
                display_df = pd.DataFrame(display_data)
                
                # Format display
                st.dataframe(
                    display_df.style.format({
                        '%': '{:.0f}%',
                        'On Going_Lop': '{:.0f}',
                        'On Going_Port': '{:.0f}',
                        'Go Live_Lop': '{:.0f}',
                        'Go Live_Port': '{:.0f}',
                        'Total Lop': '{:.0f}',
                        'Total Port': '{:.0f}',
                        'Penambahan GOLIVE H-1 vs HI': '{:.0f}',
                        'RANK': '{:.0f}'
                    }).applymap(
                        lambda x: 'font-weight: bold' if isinstance(x, (int, float)) and x == display_df['Penambahan GOLIVE H-1 vs HI'].max() and x != 0 else '',
                        subset=['Penambahan GOLIVE H-1 vs HI']
                    ),
                    use_container_width=True
                )
                
                # Visualizations
                st.subheader("\U0001F4C8 Visualisasi Data")
                col1, col2 = st.columns(2)
                
                with col1:
                    # Total Port by Witel
                    plot_df = display_df[display_df['Witel'] != 'Grand Total'].sort_values('Total Port', ascending=False)
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
                    # GOLIVE additions
                    golive_plot = display_df[
                        (display_df['Witel'] != 'Grand Total') & 
                        (display_df['Penambahan GOLIVE H-1 vs HI'] > 0)
                    ]
                    if not golive_plot.empty:
                        fig2 = px.bar(
                            golive_plot,
                            x='Witel',
                            y='Penambahan GOLIVE H-1 vs HI',
                            color='Witel',
                            title='Penambahan GOLIVE H-1 vs HI',
                            text='Penambahan GOLIVE H-1 vs HI'
                        )
                        fig2.update_traces(texttemplate='%{text:,}', textposition='outside')
                        st.plotly_chart(fig2, use_container_width=True)
                    else:
                        st.info("Tidak ada penambahan GOLIVE pada periode ini")
    
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")
else:
    st.info("Silakan upload file Excel untuk memulai analisis")

# Show last upload info
last_upload, _ = get_last_upload_info()
if last_upload:
    st.sidebar.info(f"Terakhir upload: {last_upload}")
