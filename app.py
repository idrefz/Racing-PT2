import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
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
    required_columns = ['Regional', 'Witel', 'Status Proyek', 'Total Port', 'Datel', 'Ticket ID', 'Nama Proyek']
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
        comparison_cols = ['Ticket ID', 'Status Proyek', 'Total Port', 'Witel', 'Datel', 'Nama Proyek']
        df_old = df_old[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])
        df_new = df_new[comparison_cols].dropna(subset=['Ticket ID', 'Status Proyek', 'Total Port'])

        merged = pd.merge(df_old, df_new, on='Ticket ID', suffixes=('_H-1', '_HI'))
        
        changed_to_golive = merged[
            (merged['Status Proyek_H-1'] != 'Go Live') & 
            (merged['Status Proyek_HI'] == 'Go Live')
        ]
        
        golive_additions = changed_to_golive.groupby('Witel_HI')['Total Port_HI'].sum()
        total_golive_added = golive_additions.sum()
        
        return {
            'changed_to_golive': changed_to_golive,
            'golive_additions': golive_additions,
            'total_golive_added': total_golive_added,
            'changes': merged
        }
    except Exception as e:
        st.error(f"Error comparing data: {str(e)}")
        return None

def create_pivot_tables(df):
    try:
        df['LoP'] = 1
        
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
        
        witel_pivot.columns = ['_'.join(col) if isinstance(col, tuple) else col for col in witel_pivot.columns]
        
        if 'Total Port_Go Live' in witel_pivot.columns:
            witel_pivot['%'] = (witel_pivot['Total Port_Go Live'] / 
                               witel_pivot['Total Port_Grand Total']).fillna(0) * 100
        else:
            witel_pivot['%'] = 0
            
        witel_pivot['RANK'] = witel_pivot['%'].rank(ascending=False, method='dense')
        witel_pivot.loc['Grand Total', 'RANK'] = None
        
        datel_pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index=['Witel', 'Datel'],
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0
        )
        
        datel_pivot.columns = ['_'.join(col) if isinstance(col, tuple) else col for col in datel_pivot.columns]
        
        if 'Total Port_Go Live' in datel_pivot.columns:
            datel_pivot['%'] = (datel_pivot['Total Port_Go Live'] / 
                               (datel_pivot['Total Port_On Going'] + 
                                datel_pivot['Total Port_Go Live'])).fillna(0) * 100
        else:
            datel_pivot['%'] = 0
            
        datel_pivot['RANK'] = datel_pivot.groupby('Witel')['Total Port_Go Live'].rank(ascending=False, method='min')
        
        return witel_pivot, datel_pivot
    except Exception as e:
        st.error(f"Error creating pivot tables: {str(e)}")
        return None, None

def display_witel_table(witel_display_df, comparison=None):
    style = witel_display_df.style.format({
        '%': '{:.0f}%',
        'On Going_Lop': '{:.0f}',
        'On Going_Port': '{:.0f}',
        'Go Live_Lop': '{:.0f}',
        'Go Live_Port': '{:.0f}',
        'Total Lop': '{:.0f}',
        'Total Port': '{:.0f}',
        'RANK': '{:.0f}'
    }).apply(
        lambda x: ['font-weight: bold' if x.name == 'Grand Total' else '' for _ in x],
        axis=1
    )
    
    if comparison and 'GOLIVE H-1 vs HI' in witel_display_df.columns:
        style = style.apply(
            lambda x: ['background-color: #e6ffe6' if (x.name == 'GOLIVE H-1 vs HI' and val > 0) 
                      else '' for val in x],
            axis=0
        ).format({
            'GOLIVE H-1 vs HI': '{:.0f}'
        })
    
    st.dataframe(
        style,
        use_container_width=True,
        height=(len(witel_display_df) * 35 + 3)
    )

def display_golive_changes(comparison):
    if not comparison['changed_to_golive'].empty:
        # 1. Bar Chart
        st.subheader("📈 Penambahan GOLIVE H-1 vs HI per Witel")
        
        additions_df = pd.DataFrame({
            'Witel': list(comparison['golive_additions'].keys()),
            'Penambahan Port': list(comparison['golive_additions'].values())
        }).sort_values('Penambahan Port', ascending=False)
        
        grand_total = pd.DataFrame({
            'Witel': ['GRAND TOTAL'],
            'Penambahan Port': [comparison['total_golive_added']]
        })
        additions_df = pd.concat([additions_df, grand_total])
        
        fig = px.bar(
            additions_df,
            x='Witel',
            y='Penambahan Port',
            color='Penambahan Port',
            text='Penambahan Port',
            title='Penambahan Port Go Live (H-1 vs HI)',
            color_continuous_scale='greens'
        )
        fig.update_traces(texttemplate='%{y}', textposition='outside')
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        
        # 2. Detailed Changes Table
        st.subheader("📋 Detail Perubahan Status ke Go Live")
        changes_df = comparison['changed_to_golive'][[
            'Witel_HI', 'Datel_HI', 'Nama Proyek_HI',
            'Status Proyek_H-1', 'Total Port_H-1',
            'Status Proyek_HI', 'Total Port_HI'
        ]].rename(columns={
            'Witel_HI': 'WITEL',
            'Datel_HI': 'DATEL',
            'Nama Proyek_HI': 'NAMA PROYEK',
            'Status Proyek_H-1': 'STATUS H-1',
            'Total Port_H-1': 'PORT H-1',
            'Status Proyek_HI': 'STATUS HI',
            'Total Port_HI': 'PORT HI'
        })
        
        st.dataframe(
            changes_df.style.apply(
                lambda x: ['background-color: #e6ffe6' if x['STATUS HI'] == 'Go Live' else '' 
                          for i, x in changes_df.iterrows()],
                axis=1
            ),
            use_container_width=True
        )
    else:
        st.info("Tidak ada perubahan status ke Go Live pada periode ini")

# UI Setup
st.set_page_config(page_title="Delta Ticket Harian", layout="wide")

# Sidebar Navigation
st.sidebar.title("Navigasi")
view_mode = st.sidebar.radio("Pilih Mode", ["Dashboard", "Upload Data"])

# Dashboard View
if view_mode == "Dashboard":
    st.title("📊 Dashboard Monitoring Deployment")
    
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)
        
        regional_list = ['All'] + sorted(df['Regional'].dropna().unique().tolist())
        selected_region = st.selectbox("Pilih Regional", regional_list)
        
        if selected_region != 'All':
            df = df[df['Regional'] == selected_region]
        
        witel_pivot, datel_pivot = create_pivot_tables(df)
        
        if witel_pivot is not None:
            witel_display_data = {
                'Witel': witel_pivot.index,
                'On Going_Lop': witel_pivot['LoP_On Going'],
                'On Going_Port': witel_pivot['Total Port_On Going'],
                'Go Live_Lop': witel_pivot['LoP_Go Live'],
                'Go Live_Port': witel_pivot['Total Port_Go Live'],
                'Total Lop': witel_pivot['LoP_Grand Total'],
                'Total Port': witel_pivot['Total Port_Grand Total'],
                '%': witel_pivot['%'].round(0),
                'RANK': witel_pivot['RANK']
            }
            
            # Add GOLIVE H-1 vs HI column if comparison data exists
            comparison = None
            if os.path.exists(YESTERDAY_FILE):
                df_old = load_excel(YESTERDAY_FILE)
                if selected_region != 'All':
                    df_old = df_old[df_old['Regional'] == selected_region]
                
                comparison = compare_data(df_old, df)
                if comparison:
                    witel_display_data['GOLIVE H-1 vs HI'] = [
                        comparison['golive_additions'].get(witel, 0) 
                        if witel != 'Grand Total' 
                        else comparison['total_golive_added']
                        for witel in witel_pivot.index
                    ]
            
            witel_display_df = pd.DataFrame(witel_display_data)
            
            # Display WITEL summary
            st.subheader("📊 Rekapitulasi per WITEL")
            display_witel_table(witel_display_df, comparison)
            
            # Show GOLIVE H-1 vs HI visualization if comparison data exists
            if comparison:
                display_golive_changes(comparison)
            
            # Display DATEL summary
            st.subheader("🏆 Racing per DATEL")
            
            if datel_pivot is not None:
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
            st.subheader("📊 Visualisasi Data")
            col1, col2 = st.columns(2)
            
            with col1:
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
    else:
        st.warning("Belum ada data yang diupload. Silakan ke halaman Upload Data.")

# Upload View
else:
    st.title("📤 Upload Data Harian")
    
    last_upload, last_hash = get_last_upload_info()
    if last_upload:
        st.info(f"Terakhir upload: {last_upload}")
    
    uploaded_file = st.file_uploader("Upload file Excel harian", type="xlsx")
    
    if uploaded_file:
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
                    df['Total Port'] = pd.to_numeric(df['Total Port'], errors='coerce')
                    
                    if os.path.exists(LATEST_FILE):
                        os.replace(LATEST_FILE, YESTERDAY_FILE)
                    save_file(LATEST_FILE, uploaded_file)
                    record_upload_history()
                    
                    st.success("✅ File berhasil diupload dan data dashboard telah diperbarui!")
                    st.balloons()
                    
                    st.subheader("📋 Ringkasan Data")
                    cols = st.columns(4)
                    cols[0].metric("Total Projek", len(df))
                    cols[1].metric("Total Port", df['Total Port'].sum())
                    cols[2].metric("WITEL", df['Witel'].nunique())
                    cols[3].metric("DATEL", df['Datel'].nunique())
                    
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
