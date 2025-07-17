import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import hashlib
import shutil

# Configuration
DATA_FOLDER = "data_daily_uploads"
LATEST_FILE = os.path.join(DATA_FOLDER, "latest.xlsx")
HISTORY_FILE = os.path.join(DATA_FOLDER, "upload_history.csv")

# Create folder if not exists
os.makedirs(DATA_FOLDER, exist_ok=True)

# Initialize session state
if 'view_mode' not in st.session_state:
    st.session_state.view_mode = "Dashboard"

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
        return False, f"Missing required columns: {', '.join(missing_columns)}"
    
    try:
        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='raise')
    except ValueError:
        return False, "'Total Port' must contain numeric values"
    
    return True, "Data valid"

def create_pivot_tables(df):
    try:
        delta_values = compare_with_previous(df)
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
        
        witel_pivot.columns = ['_'.join(col).strip() for col in witel_pivot.columns.values]
        
        witel_pivot['Total_Port_All'] = witel_pivot['Total Port_Go Live'] + witel_pivot['Total Port_On Going']
        witel_pivot['Completion_Percentage'] = (witel_pivot['Total Port_Go Live'] / witel_pivot['Total_Port_All']).fillna(0) * 100
        
        witel_pivot['Delta_GoLive'] = 0
        for witel in delta_values:
            if witel in witel_pivot.index:
                witel_pivot.at[witel, 'Delta_GoLive'] = delta_values[witel]
        
        witel_pivot['Rank'] = witel_pivot['Total Port_Go Live'].rank(ascending=False, method='min').astype(int)
        witel_pivot.loc['Grand Total', 'Rank'] = None
        
        # DATEL pivot
        datel_pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index=['Witel', 'Datel'],
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0
        )
        datel_pivot.columns = ['_'.join(col).strip() for col in datel_pivot.columns.values]
        
        datel_pivot['Total_Port_All'] = datel_pivot['Total Port_Go Live'] + datel_pivot['Total Port_On Going']
        datel_pivot['Completion_Percentage'] = (datel_pivot['Total Port_Go Live'] / datel_pivot['Total_Port_All']).fillna(0) * 100
        datel_pivot['Rank'] = datel_pivot.groupby('Witel')['Total Port_Go Live'].rank(ascending=False, method='min').astype(int)
        
        return witel_pivot, datel_pivot
    except Exception as e:
        st.error(f"Error creating pivot tables: {str(e)}")
        return None, None

def create_witel_summary_plot(witel_data):
    plot_df = witel_data[witel_data['Witel'] != 'Grand Total'].sort_values('Total Port_Go Live', ascending=False)
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(
        go.Bar(
            x=plot_df['Witel'],
            y=plot_df['Total Port_Go Live'],
            name='Go Live',
            marker_color='#2ca02c',
            text=plot_df['Total Port_Go Live'],
            texttemplate='%{text:,}',
            textposition='outside'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Bar(
            x=plot_df['Witel'],
            y=plot_df['Total Port_On Going'],
            name='On Going',
            marker_color='#ff7f0e',
            text=plot_df['Total Port_On Going'],
            texttemplate='%{text:,}',
            textposition='outside'
        ),
        secondary_y=False
    )
    
    fig.add_trace(
        go.Scatter(
            x=plot_df['Witel'],
            y=plot_df['Completion_Percentage'],
            name='% Completion',
            line=dict(color='#1f77b4', width=3),
            text=plot_df['Completion_Percentage'].round(1).astype(str) + '%',
            mode='lines+markers+text',
            textposition='top center'
        ),
        secondary_y=True
    )
    
    for idx, row in plot_df.iterrows():
        if row['Delta_GoLive'] != 0:
            fig.add_annotation(
                x=row['Witel'],
                y=row['Total Port_Go Live'],
                text=f"Δ{row['Delta_GoLive']:+}",
                showarrow=True,
                arrowhead=1,
                ax=0,
                ay=-40,
                font=dict(color='green' if row['Delta_GoLive'] > 0 else 'red')
            )
    
    fig.update_layout(
        title='<b>Port Deployment Summary by WITEL</b>',
        xaxis_title='WITEL',
        yaxis_title='Port Count',
        yaxis2_title='Completion %',
        yaxis2=dict(range=[0, 105]),
        hovermode='x unified',
        barmode='stack',
        height=600,
        template='plotly_white'
    )
    
    return fig

def display_witel_summary(witel_pivot):
    st.subheader("🏆 WITEL Performance Ranking")
    
    display_cols = {
        'Witel': witel_pivot.index,
        'Rank': witel_pivot['Rank'],
        'Go Live (Port)': witel_pivot['Total Port_Go Live'],
        'On Going (Port)': witel_pivot['Total Port_On Going'],
        'Total Port': witel_pivot['Total_Port_All'],
        '% Completion': witel_pivot['Completion_Percentage'],
        'Daily Δ': witel_pivot['Delta_GoLive']
    }
    
    display_df = pd.DataFrame(display_cols).set_index('Witel')
    
    def color_negative_red(val):
        color = 'red' if val < 0 else 'green' if val > 0 else 'gray'
        return f'color: {color}'
    
    styled_df = display_df.style.format({
        'Go Live (Port)': '{:,}',
        'On Going (Port)': '{:,}',
        'Total Port': '{:,}',
        '% Completion': '{:.1f}%',
        'Daily Δ': '{:+}'
    }).applymap(color_negative_red, subset=['Daily Δ'])
    
    def highlight_top3(s):
        styles = []
        for i in range(len(s)):
            if s.name != 'Grand Total' and s.iloc[i] <= 3:
                styles.append('background-color: #e6f3ff; font-weight: bold')
            else:
                styles.append('')
        return styles
    
    styled_df = styled_df.apply(highlight_top3, subset=['Rank'])
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=(len(display_df) * 35 + 3)
    )
    
    st.plotly_chart(create_witel_summary_plot(display_df.reset_index()), use_container_width=True)

def display_datel_details(datel_pivot):
    st.subheader("📊 DATEL-Level Performance")
    witels = datel_pivot.index.get_level_values(0).unique()
    tabs = st.tabs([f"🏅 {witel}" for witel in witels])
    
    for i, witel in enumerate(witels):
        with tabs[i]:
            witel_data = datel_pivot.loc[witel].sort_values('Total Port_Go Live', ascending=False)
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total DATELs", len(witel_data))
            col2.metric("Total Go Live", f"{witel_data['Total Port_Go Live'].sum():,}")
            col3.metric("Total On Going", f"{witel_data['Total Port_On Going'].sum():,}")
            col4.metric("Overall Completion", 
                       f"{(witel_data['Total Port_Go Live'].sum() / 
                         (witel_data['Total Port_Go Live'].sum() + 
                          witel_data['Total Port_On Going'].sum()) * 100):.1f}%")
            
            st.dataframe(
                witel_data.style.format({
                    'Total Port_Go Live': '{:,}',
                    'Total Port_On Going': '{:,}',
                    'Total_Port_All': '{:,}',
                    'Completion_Percentage': '{:.1f}%'
                }).apply(
                    lambda x: ['background-color: #e6f3ff' if x.name == witel_data.index[0] else '' 
                             for _ in x],
                    axis=1
                ),
                use_container_width=True
            )
            
            fig = px.bar(
                witel_data.reset_index(),
                x='Datel',
                y=['Total Port_Go Live', 'Total Port_On Going'],
                title=f'Port Deployment Status - {witel}',
                labels={'value': 'Port Count', 'variable': 'Status'},
                color_discrete_map={
                    'Total Port_Go Live': '#2ca02c',
                    'Total Port_On Going': '#ff7f0e'
                },
                text_auto=True
            )
            fig.update_layout(barmode='stack')
            st.plotly_chart(fig, use_container_width=True)

# UI Setup
st.set_page_config(page_title="PT2 IHLD Deployment Dashboard", layout="wide")

# Sidebar Navigation
with st.sidebar:
    st.title("Navigasi")
    st.session_state.view_mode = st.radio(
        "Pilih Mode",
        ["Dashboard", "Upload Data"],
        key='nav_radio'
    )
    
    if os.path.exists(HISTORY_FILE):
        st.subheader("History Upload")
        history_df = pd.read_csv(HISTORY_FILE)
        st.dataframe(
            history_df[['timestamp']].tail(5),
            hide_index=True,
            use_container_width=True
        )

# Dashboard View
if st.session_state.view_mode == "Dashboard":
    st.title("📈 PT2 IHLD Deployment Dashboard")
    
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)
        
        with st.sidebar.expander("🔍 Filter Options", expanded=True):
            regional_list = ['All'] + sorted(df['Regional'].dropna().unique().tolist())
            selected_region = st.selectbox("Select Regional", regional_list)
            
            if selected_region != 'All':
                df = df[df['Regional'] == selected_region]
            
            status_filter = st.multiselect(
                "Filter by Status",
                options=df['Status Proyek'].unique(),
                default=df['Status Proyek'].unique()
            )
            df = df[df['Status Proyek'].isin(status_filter)]
        
        witel_pivot, datel_pivot = create_pivot_tables(df)
        
        if witel_pivot is not None:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total WITEL", len(witel_pivot)-1)
            col2.metric("Total Go Live Ports", f"{witel_pivot.loc['Grand Total', 'Total Port_Go Live']:,}")
            col3.metric("Total On Going Ports", f"{witel_pivot.loc['Grand Total', 'Total Port_On Going']:,}")
            col4.metric("Overall Completion", 
                       f"{(witel_pivot.loc['Grand Total', 'Total Port_Go Live'] / 
                         witel_pivot.loc['Grand Total', 'Total_Port_All'] * 100):.1f}%")
            
            display_witel_summary(witel_pivot)
            display_datel_details(datel_pivot)
    else:
        st.warning("No data available. Please upload data in the Upload section.")

# Upload View
else:
    st.title("📤 Upload Daily Data")
    
    last_upload, last_hash = get_last_upload_info()
    if last_upload:
        with st.expander("ℹ️ Last Upload Info"):
            st.write(f"**Last Upload:** {last_upload}")
            if os.path.exists(LATEST_FILE):
                df_info = pd.read_excel(LATEST_FILE, nrows=1)
                st.write(f"**Detected Columns:** {', '.join(df_info.columns)}")
    
    uploaded_file = st.file_uploader("Upload daily Excel file", type="xlsx")
    
    if uploaded_file:
        with st.spinner("Processing file..."):
            try:
                current_hash = get_file_hash(uploaded_file.getvalue())
                
                if last_hash and current_hash == last_hash:
                    st.success("✅ Data identical to last upload. No need to re-upload.")
                else:
                    df = load_excel(uploaded_file)
                    is_valid, msg = validate_data(df)
                    
                    if not is_valid:
                        st.error(f"Validation failed: {msg}")
                    else:
                        df = df.dropna(subset=['Witel', 'Datel'])
                        df['Total Port'] = pd.to_numeric(df['Total Port'], errors='coerce').fillna(0)
                        
                        if st.checkbox("Confirm new data upload"):
                            save_file(LATEST_FILE, uploaded_file)
                            record_upload_history()
                            
                            st.success("✅ File uploaded successfully!")
                            st.balloons()
                            
                            with st.expander("📋 Data Summary", expanded=True):
                                cols = st.columns(4)
                                cols[0].metric("Total Projects", len(df))
                                cols[1].metric("Total Ports", f"{int(df['Total Port'].sum()):,}")
                                cols[2].metric("WITELs", df['Witel'].nunique())
                                cols[3].metric("DATELs", df['Datel'].nunique())
                                
                                st.write("**Status Distribution:**")
                                status_counts = df['Status Proyek'].value_counts()
                                fig = px.pie(status_counts, 
                                           values=status_counts.values,
                                           names=status_counts.index)
                                st.plotly_chart(fig, use_container_width=True)
            
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
