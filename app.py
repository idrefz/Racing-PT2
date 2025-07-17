import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib

# Configuration
DATA_FOLDER = "data_daily_uploads"
os.makedirs(DATA_FOLDER, exist_ok=True)
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
YESTERDAY_FILE = os.path.join(DATA_FOLDER, "yesterday.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")

# Initialize session state
if 'view_mode' not in st.session_state:
    st.session_state.view_mode = 'dashboard'
if 'upload_complete' not in st.session_state:
    st.session_state.upload_complete = False

# Helper Functions
@st.cache_data
def load_excel(file):
    """Load Excel file with caching"""
    return pd.read_excel(file)

def save_file(path, uploaded_file):
    """Save uploaded file"""
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def get_file_hash(file_content):
    """Generate MD5 hash of file content"""
    return hashlib.md5(file_content).hexdigest()

def record_upload_history():
    """Record upload timestamp and file hash"""
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
    """Get last upload timestamp and file hash"""
    if os.path.exists(HISTORY_FILE):
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty:
            return history_df.iloc[-1]['timestamp'], history_df.iloc[-1]['file_hash']
    return None, None

def validate_data(df):
    """Validate uploaded data structure"""
    required_columns = {
        'Regional': 'text',
        'Witel': 'text', 
        'Status Proyek': 'text',
        'Total Port': 'numeric',
        'Datel': 'text'
    }
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return False, f"Kolom yang dibutuhkan tidak ditemukan: {', '.join(missing_columns)}"
    
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except ValueError:
        return False, "Kolom 'Total Port' harus berisi nilai numerik"
    
    return True, "Data valid"

def calculate_lop(df):
    """Calculate LoP (count of projects)"""
    df['LoP'] = 1  # Each row represents one project
    return df

def compare_data(df_old, df_new):
    """Compare two datasets and identify changes"""
    try:
        comparison_cols = ['Ticket ID', 'Status Proyek', 'Total Port', 'Witel', 'Datel']
        df_old = df_old[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])
        df_new = df_new[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])

        old_tickets = set(df_old['Ticket ID'])
        new_tickets = set(df_new['Ticket ID'])

        # Identify changes
        new_projects = new_tickets - old_tickets
        removed_projects = old_tickets - new_tickets
        common_tickets = new_tickets & old_tickets

        # Status changes
        df_old_common = df_old[df_old['Ticket ID'].isin(common_tickets)].set_index('Ticket ID')
        df_new_common = df_new[df_new['Ticket ID'].isin(common_tickets)].set_index('Ticket ID')
        
        status_diff = df_old_common.join(df_new_common, lsuffix='_H', rsuffix='_H+1')
        changed_to_golive = status_diff[
            (status_diff['Status Proyek_H'] != 'Go Live') & 
            (status_diff['Status Proyek_H+1'] == 'Go Live')
        ]
        
        # Calculate metrics
        golive_additions = {
            'by_witel': changed_to_golive.groupby('Witel_H+1')['Total Port_H+1'].sum(),
            'by_datel': changed_to_golive.groupby('Datel_H+1')['Total Port_H+1'].sum(),
            'total': changed_to_golive['Total Port_H+1'].sum()
        }

        return {
            'new_count': len(new_projects),
            'removed_count': len(removed_projects),
            'changed_count': len(status_diff[status_diff['Status Proyek_H'] != status_diff['Status Proyek_H+1']]),
            'golive_additions': golive_additions,
            'changes': status_diff.reset_index()
        }
    except Exception as e:
        st.error(f"Error comparing data: {str(e)}")
        return None

def create_pivot_table(df):
    """Create summary pivot table"""
    try:
        df = calculate_lop(df)
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
        
        # Calculate completion percentage
        if ('Total Port', 'Go Live') in pivot.columns:
            pivot['Completion %'] = (pivot[('Total Port', 'Go Live')] / 
                                   pivot[('Total Port', 'Grand Total')]).fillna(0) * 100
        else:
            pivot['Completion %'] = 0
            
        # Calculate rankings
        pivot['Rank'] = pivot[('Total Port', 'Grand Total')].rank(ascending=False, method='min')
        pivot.loc['Grand Total', 'Rank'] = None
        
        return pivot
    except Exception as e:
        st.error(f"Error creating pivot table: {str(e)}")
        return None

def show_dashboard(df):
    """Display the monitoring dashboard"""
    try:
        # Data validation
        is_valid, msg = validate_data(df)
        if not is_valid:
            st.error(msg)
            return
        
        # Regional filter
        regions = ['All'] + sorted(df['Regional'].dropna().unique().tolist())
        selected_region = st.selectbox("Filter Regional", regions)
        
        if selected_region != 'All':
            df = df[df['Regional'] == selected_region]

        # Create summary tables
        pivot_table = create_pivot_table(df)
        if pivot_table is None:
            return

        # Display KPIs
        st.subheader("📊 Key Performance Indicators")
        cols = st.columns(4)
        cols[0].metric("Total Projek", len(df))
        cols[1].metric("Total Port", df['Total Port'].sum())
        
        if os.path.exists(YESTERDAY_FILE):
            comparison = compare_data(load_excel(YESTERDAY_FILE), df)
            if comparison:
                cols[2].metric("Projek Baru", comparison['new_count'])
                cols[3].metric("Perubahan Status", comparison['changed_count'])

        # Witel Summary Table
        st.subheader("📋 Rekapitulasi per Witel")
        summary_df = pd.DataFrame({
            'Witel': pivot_table.index,
            'On Going (Projek)': pivot_table[('LoP', 'On Going')],
            'On Going (Port)': pivot_table[('Total Port', 'On Going')],
            'Go Live (Projek)': pivot_table[('LoP', 'Go Live')],
            'Go Live (Port)': pivot_table[('Total Port', 'Go Live')],
            'Total Projek': pivot_table[('LoP', 'Grand Total')],
            'Total Port': pivot_table[('Total Port', 'Grand Total')],
            'Progress (%)': pivot_table['Completion %'].round(1),
            'Ranking': pivot_table['Rank']
        })
        
        # Add GOLIVE additions if available
        if os.path.exists(YESTERDAY_FILE) and 'comparison' in locals():
            summary_df['Penambahan GOLIVE'] = summary_df['Witel'].map(
                comparison['golive_additions']['by_witel']
            ).fillna(0)
            summary_df.loc[summary_df['Witel'] == 'Grand Total', 'Penambahan GOLIVE'] = \
                comparison['golive_additions']['total']
        
        st.dataframe(
            summary_df.style.format({
                'Progress (%)': '{:.1f}%',
                'Penambahan GOLIVE': '{:.0f}',
                'Ranking': '{:.0f}'
            }),
            use_container_width=True
        )

        # Datel Breakdown
        st.subheader("📊 Breakdown per Datel")
        
        # Witel filter for Datel view
        witel_list = ['All'] + sorted(df['Witel'].unique().tolist())
        selected_witel = st.selectbox("Filter Witel", witel_list)
        
        filtered_df = df if selected_witel == 'All' else df[df['Witel'] == selected_witel]
        
        datel_summary = filtered_df.groupby('Datel').agg(
            On_Going_Projects=('LoP', lambda x: x[filtered_df['Status Proyek'] == 'On Going'].sum()),
            On_Going_Ports=('Total Port', lambda x: x[filtered_df['Status Proyek'] == 'On Going'].sum()),
            Go_Live_Projects=('LoP', lambda x: x[filtered_df['Status Proyek'] == 'Go Live'].sum()),
            Go_Live_Ports=('Total Port', lambda x: x[filtered_df['Status Proyek'] == 'Go Live'].sum())
        ).reset_index()
        
        datel_summary['Total Projects'] = datel_summary['On_Going_Projects'] + datel_summary['Go_Live_Projects']
        datel_summary['Total Ports'] = datel_summary['On_Going_Ports'] + datel_summary['Go_Live_Ports']
        datel_summary['Progress (%)'] = (datel_summary['Go_Live_Ports'] / datel_summary['Total Ports'] * 100).round(1)
        
        # Add grand total
        grand_total = pd.DataFrame({
            'Datel': ['GRAND TOTAL'],
            'On_Going_Projects': [datel_summary['On_Going_Projects'].sum()],
            'On_Going_Ports': [datel_summary['On_Going_Ports'].sum()],
            'Go_Live_Projects': [datel_summary['Go_Live_Projects'].sum()],
            'Go_Live_Ports': [datel_summary['Go_Live_Ports'].sum()],
            'Total Projects': [datel_summary['Total Projects'].sum()],
            'Total Ports': [datel_summary['Total Ports'].sum()],
            'Progress (%)': [(datel_summary['Go_Live_Ports'].sum() / datel_summary['Total Ports'].sum() * 100).round(1)]
        })
        
        st.dataframe(
            pd.concat([datel_summary, grand_total], ignore_index=True).style.format({
                'Progress (%)': '{:.1f}%'
            }),
            use_container_width=True
        )

        # Visualizations
        st.subheader("📈 Data Visualization")
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig1 = px.bar(
                summary_df[summary_df['Witel'] != 'Grand Total'].sort_values('Total Port', ascending=False),
                x='Witel',
                y='Total Port',
                title='Total Port per Witel',
                text='Total Port'
            )
            fig1.update_traces(texttemplate='%{text:,}')
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            fig2 = px.pie(
                names=['On Going', 'Go Live'],
                values=[summary_df['On Going (Port)'].sum(), summary_df['Go Live (Port)'].sum()],
                title='Port Distribution',
                hole=0.4
            )
            st.plotly_chart(fig2, use_container_width=True)

    except Exception as e:
        st.error(f"Error displaying dashboard: {str(e)}")

# Main App Interface
st.sidebar.title("Navigation")
view_mode = st.sidebar.radio("Select View", ["📊 Dashboard", "⬆️ Upload Data"])

if view_mode == "📊 Dashboard":
    st.title("Daily Monitoring Dashboard")
    
    if os.path.exists(LATEST_FILE):
        try:
            df = load_excel(LATEST_FILE)
            show_dashboard(df)
        except Exception as e:
            st.error(f"Failed to load data: {str(e)}")
    else:
        st.warning("No data available. Please upload data first.")

elif view_mode == "⬆️ Upload Data":
    st.title("Upload Daily Data")
    
    # Show upload history
    last_upload, last_hash = get_last_upload_info()
    if last_upload:
        st.info(f"Last upload: {last_upload}")
    
    uploaded_file = st.file_uploader("Upload today's data (Excel format)", type="xlsx")
    
    if uploaded_file:
        try:
            # Check for duplicate upload
            file_hash = get_file_hash(uploaded_file.getvalue())
            if last_hash and file_hash == last_hash:
                st.warning("Uploaded file is identical to last upload. No changes detected.")
                st.session_state.upload_complete = False
            else:
                # Load and validate data
                df = load_excel(uploaded_file)
                is_valid, msg = validate_data(df)
                
                if not is_valid:
                    st.error(msg)
                else:
                    # Show comparison with previous data
                    if os.path.exists(LATEST_FILE):
                        st.subheader("Changes from Previous Data")
                        comparison = compare_data(load_excel(LATEST_FILE), df)
                        
                        if comparison:
                            cols = st.columns(3)
                            cols[0].metric("New Projects", comparison['new_count'])
                            cols[1].metric("Removed Projects", comparison['removed_count'])
                            cols[2].metric("Status Changes", comparison['changed_count'])
                    
                    # Save new data
                    if os.path.exists(LATEST_FILE):
                        os.replace(LATEST_FILE, YESTERDAY_FILE)
                    
                    save_file(LATEST_FILE, uploaded_file)
                    record_upload_history()
                    
                    st.success("Data successfully uploaded!")
                    st.session_state.upload_complete = True
                    
                    # Preview uploaded data
                    st.subheader("Data Preview")
                    st.dataframe(df.head())
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    
    if st.session_state.get('upload_complete', False):
        if st.button("View Dashboard"):
            st.session_state.view_mode = 'dashboard'
            st.experimental_rerun()
