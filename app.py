import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import shutil

# Configuration
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
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

def compare_with_previous(current_df):
    """
    Compare current data with previous upload to calculate delta in Go Live ports
    Returns a dictionary with Witel as key and delta as value
    """
    delta_dict = {}
    
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if len(history_df) >= 2:
            previous_hash = history_df.iloc[-2]['file_hash']
            previous_file = os.path.join(DATA_FOLDER, f"previous_{previous_hash}.xlsx")
            
            if os.path.exists(previous_file):
                previous_df = pd.read_excel(previous_file)
                
                current_golive = current_df[current_df['Status Proyek'] == 'Go Live'].groupby('Witel')['Total Port'].sum()
                previous_golive = previous_df[previous_df['Status Proyek'] == 'Go Live'].groupby('Witel')['Total Port'].sum()
                
                for witel in current_golive.index:
                    current_val = current_golive[witel]
                    previous_val = previous_golive.get(witel, 0)
                    delta_dict[witel] = current_val - previous_val
                
                shutil.copy(LATEST_FILE, previous_file)
    
    return delta_dict

def record_upload_history():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_hash = get_file_hash(open(LATEST_FILE, "rb").read()) if os.path.exists(LATEST_FILE) else ""
    
    history_df = pd.DataFrame(columns=['timestamp', 'file_hash'])
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty:
            previous_hash = history_df.iloc[-1]['file_hash']
            previous_file = os.path.join(DATA_FOLDER, f"previous_{previous_hash}.xlsx")
            if os.path.exists(LATEST_FILE):
                shutil.copy(LATEST_FILE, previous_file)
    
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

def create_pivot_tables(df):
    try:
        # Calculate delta from previous upload
        delta_values = compare_with_previous(df)
        
        # WITEL level summary
        df['LoP'] = 1
        
        # WITEL pivot
        witel_pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index='Witel',
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        )
        
        # Flatten multi-index columns
        witel_pivot.columns = ['_'.join(col) if isinstance(col, tuple) else col for col in witel_pivot.columns]
        
        # Calculate percentages
        if 'Total Port_Go Live' in witel_pivot.columns:
            witel_pivot['%'] = (witel_pivot['Total Port_Go Live'] / 
                               witel_pivot['Total Port_Grand Total']).fillna(0) * 100
        else:
            witel_pivot['%'] = 0
            
        # Add delta column
        witel_pivot['Penambahan GOLIVE H-1 vs HI'] = 0
        for witel in delta_values:
            if witel in witel_pivot.index:
                witel_pivot.at[witel, 'Penambahan GOLIVE H-1 vs HI'] = delta_values[witel]
        
        # Convert all numbers to integers except percentage
        witel_pivot = witel_pivot.round(0).astype(int, errors='ignore')
        witel_pivot['%'] = witel_pivot['%']  # Keep percentage as float for formatting
        
        witel_pivot['RANK'] = witel_pivot['%'].rank(ascending=False, method='dense')
        witel_pivot.loc['Grand Total', 'RANK'] = None
        
        # DATEL pivot
        datel_pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index=['Witel', 'Datel'],
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0
        )
        
        # Flatten multi-index columns
        datel_pivot.columns = ['_'.join(col) if isinstance(col, tuple) else col for col in datel_pivot.columns]
        
        # Calculate percentages
        if 'Total Port_Go Live' in datel_pivot.columns:
            datel_pivot['%'] = (datel_pivot['Total Port_Go Live'] / 
                               (datel_pivot['Total Port_On Going'] + 
                                datel_pivot['Total Port_Go Live'])).fillna(0) * 100
        else:
            datel_pivot['%'] = 0
            
        # Convert all numbers to integers except percentage
        datel_pivot = datel_pivot.round(0).astype(int, errors='ignore')
        datel_pivot['%'] = datel_pivot['%']  # Keep percentage as float for formatting
        
        datel_pivot['RANK'] = datel_pivot.groupby('Witel')['Total Port_Go Live'].rank(ascending=False, method='min')
        
        return witel_pivot, datel_pivot
    except Exception as e:
        st.error(f"Error creating pivot tables: {str(e)}")
        return None, None

# UI Setup
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")

# Sidebar Navigation
st.sidebar.title("Navigasi")
view_mode = st.sidebar.radio("Pilih Mode", ["Dashboard", "Upload Data"])

# Dashboard View
if view_mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment PT2 IHLD")
    
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)
        
        # Regional filter
        regional_list = ['All'] + sorted(df['Regional'].dropna().unique().tolist())
        selected_region = st.selectbox("Pilih Regional", regional_list)
        
        if selected_region != 'All':
            df = df[df['Regional'] == selected_region]
        
        # Create pivot tables
        witel_pivot, datel_pivot = create_pivot_tables(df)
        
        if witel_pivot is not None:
            # Prepare WITEL display table
            witel_display_data = {
                'Witel': witel_pivot.index,
                'On Going_Lop': witel_pivot['LoP_On Going'],
                'On Going_Port': witel_pivot['Total Port_On Going'],
                'Go Live_Lop': witel_pivot['LoP_Go Live'],
                'Go Live_Port': witel_pivot['Total Port_Go Live'],
                'Total Lop': witel_pivot['LoP_Grand Total'],
                'Total Port': witel_pivot['Total Port_Grand Total'],
                '%': witel_pivot['%'].round(0),
                'Penambahan GOLIVE H-1 vs HI': witel_pivot['Penambahan GOLIVE H-1 vs HI'],
                'RANK': witel_pivot['RANK']
            }
            
            witel_display_df = pd.DataFrame(witel_display_data)
            
            # Display WITEL summary with whole numbers
            st.subheader("📊 Rekapitulasi per WITEL")
            st.dataframe(
                witel_display_df.style.format({
                    '%': '{:.0f}%',
                    'On Going_Lop': '{:.0f}',
                    'On Going_Port': '{:.0f}',
                    'Go Live_Lop': '{:.0f}',
                    'Go Live_Port': '{:.0f}',
                    'Total Lop': '{:.0f}',
                    'Total Port': '{:.0f}',
                    'Penambahan GOLIVE H-1 vs HI': '{:.0f}',
                    'RANK': '{:.0f}'
                }).apply(
                    lambda x: ['font-weight: bold' if x.name == 'Grand Total' else '' for _ in x],
                    axis=1
                ),
                use_container_width=True,
                height=(len(witel_display_df) * 35 + 3)
            )
            
            # Display DATEL summary
            st.subheader("🏆 Racing per DATEL")
            
            if datel_pivot is not None:
                # Prepare DATEL display table
                datel_display_data = {
                    'Witel': datel_pivot.index.get_level_values(0),
                    'Datel': datel_pivot.index.get_level_values(1),
                    'On Going_Lop': datel_pivot['LoP_On Going'],
                    'On Going_Port': datel_pivot['Total Port_On Going'],
                    'Go Live_Lop': datel_pivot['LoP_Go Live'],
                    'Go Live_Port': datel_pivot['Total Port_Go Live'],
                    'Total Lop': datel_pivot['LoP_On Going'] + datel_pivot['LoP_Go Live'],
                    'Total Port': datel_pivot['Total Port_On Going'] + datel_pivot['Total Port_Go Live'],
                    '%': datel_pivot['%'].round(0),
                    'RANK': datel_pivot['RANK']
                }
                
                datel_display_df = pd.DataFrame(datel_display_data)
                
                # Group by WITEL for tabs
                witels = datel_display_df['Witel'].unique()
                tabs = st.tabs([f"🏆 {witel}" for witel in witels])
                
                for i, witel in enumerate(witels):
                    with tabs[i]:
                        witel_data = datel_display_df[datel_display_df['Witel'] == witel].sort_values('RANK')
                        
                        st.dataframe(
                            witel_data.style.format({
                                '%': '{:.0f}%',
                                'On Going_Lop': '{:.0f}',
                                'On Going_Port': '{:.0f}',
                                'Go Live_Lop': '{:.0f}',
                                'Go Live_Port': '{:.0f}',
                                'Total Lop': '{:.0f}',
                                'Total Port': '{:.0f}',
                                'RANK': '{:.0f}'
                            }).apply(
                                lambda x: ['background-color: #e6f3ff' if x.RANK == 1 else '' for _ in x],
                                axis=1
                            ),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        # Visualization for each WITEL
                        fig = px.bar(
                            witel_data,
                            x='Datel',
                            y=['On Going_Port', 'Go Live_Port'],
                            title=f'Port Status per DATEL - {witel}',
                            labels={'value': 'Port Count', 'variable': 'Status'},
                            color_discrete_map={
                                'On Going_Port': '#FFA15A',
                                'Go Live_Port': '#00CC96'
                            }
                        )
                        st.plotly_chart(fig, use_container_width=True)
            
            # Overall visualizations
            st.subheader("📈 Visualisasi Data")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Total Port by Witel
                plot_df = witel_display_df[witel_display_df['Witel'] != 'Grand Total'].sort_values('Total Port', ascending=False)
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
                # Completion percentage
                fig2 = px.bar(
                    plot_df,
                    x='Witel',
                    y='%',
                    color='Witel',
                    title='Persentase Go Live per Witel',
                    text='%'
                )
                fig2.update_traces(texttemplate='%{text:.0f}%', textposition='outside')
                fig2.update_yaxes(range=[0, 100])
                st.plotly_chart(fig2, use_container_width=True)
            
            # Delta visualization
            st.subheader("📈 Perubahan Port Go Live (H-1 vs Hari Ini)")
            fig3 = px.bar(
                plot_df,
                x='Witel',
                y='Penambahan GOLIVE H-1 vs HI',
                color='Witel',
                title='Penambahan Port Go Live vs Hari Sebelumnya',
                text='Penambahan GOLIVE H-1 vs HI',
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            fig3.update_traces(texttemplate='%{text:,}', textposition='outside')
            st.plotly_chart(fig3, use_container_width=True)
    else:
        st.warning("Belum ada data yang diupload. Silakan ke halaman Upload Data.")

# Upload View
else:
    st.title("📤 Upload Data Harian")
    
    # Show last upload info
    last_upload, last_hash = get_last_upload_info()
    if last_upload:
        st.info(f"Terakhir upload: {last_upload}")
    
    uploaded_file = st.file_uploader("Upload file Excel harian", type="xlsx")
    
    if uploaded_file:
        # Validate and process uploaded file
        try:
            current_hash = get_file_hash(uploaded_file.getvalue())
            
            if last_hash and current_hash == last_hash:
                st.success("✅ Data masih sama dengan upload terakhir. Tidak perlu upload ulang.")
            else:
                df = load_excel(uploaded_file)
                is_valid, msg = validate_data(df)
                
                if not is_valid:
                    st.error(msg)
                else:
                    # Process numeric columns
                    df['Total Port'] = pd.to_numeric(df['Total Port'], errors='coerce')
                    
                    # Save file and record history
                    save_file(LATEST_FILE, uploaded_file)
                    record_upload_history()
                    
                    st.success("✅ File berhasil diupload dan data dashboard telah diperbarui!")
                    st.balloons()
                    
                    # Show quick summary
                    st.subheader("📋 Ringkasan Data")
                    cols = st.columns(4)
                    cols[0].metric("Total Projek", len(df))
                    cols[1].metric("Total Port", int(df['Total Port'].sum()))
                    cols[2].metric("WITEL", df['Witel'].nunique())
                    cols[3].metric("DATEL", df['Datel'].nunique())
                    
                    # Show sample data
                    st.subheader("🖥️ Preview Data")
                    st.dataframe(df.head(), use_container_width=True)
        
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {str(e)}")

# Show upload history
if view_mode == "Upload Data" and os.path.exists(HISTORY_FILE):
    st.sidebar.subheader("History Upload")
    history_df = pd.read_csv(HISTORY_FILE)
    st.sidebar.dataframe(
        history_df[['timestamp']].tail(5),
        hide_index=True
    )
