import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Improved Helper Functions
def create_pivot_tables(df):
    """Enhanced pivot table creation with proper ranking"""
    try:
        # Calculate delta from previous upload
        delta_values = compare_with_previous(df)
        
        # Create LoP (Line of Project) column
        df['LoP'] = 1
        
        # WITEL level summary - improved calculation
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
        witel_pivot.columns = ['_'.join(col).strip() for col in witel_pivot.columns.values]
        
        # Calculate percentages and totals
        witel_pivot['Total_Port_All'] = witel_pivot['Total Port_Go Live'] + witel_pivot['Total Port_On Going']
        witel_pivot['Completion_Percentage'] = (witel_pivot['Total Port_Go Live'] / 
                                              witel_pivot['Total_Port_All']).fillna(0) * 100
        
        # Add delta column
        witel_pivot['Delta_GoLive'] = 0
        for witel in delta_values:
            if witel in witel_pivot.index:
                witel_pivot.at[witel, 'Delta_GoLive'] = delta_values[witel]
        
        # Calculate ranking based on Go Live Ports (descending)
        witel_pivot['Rank'] = witel_pivot['Total Port_Go Live'].rank(ascending=False, method='min').astype(int)
        witel_pivot.loc['Grand Total', 'Rank'] = None
        
        # Format numbers (integers except percentages)
        for col in ['LoP_Go Live', 'LoP_On Going', 'LoP_Grand Total',
                  'Total Port_Go Live', 'Total Port_On Going', 'Total_Port_All',
                  'Delta_GoLive']:
            witel_pivot[col] = witel_pivot[col].round(0).astype(int)
        
        # DATEL level summary
        datel_pivot = pd.pivot_table(
            df,
            values=['LoP', 'Total Port'],
            index=['Witel', 'Datel'],
            columns='Status Proyek',
            aggfunc='sum',
            fill_value=0
        )
        datel_pivot.columns = ['_'.join(col).strip() for col in datel_pivot.columns.values]
        
        # Calculate datel-level metrics
        datel_pivot['Total_Port_All'] = datel_pivot['Total Port_Go Live'] + datel_pivot['Total Port_On Going']
        datel_pivot['Completion_Percentage'] = (datel_pivot['Total Port_Go Live'] / 
                                             datel_pivot['Total_Port_All']).fillna(0) * 100
        
        # Calculate ranking within each Witel
        datel_pivot['Rank'] = datel_pivot.groupby('Witel')['Total Port_Go Live'].rank(
            ascending=False, method='min').astype(int)
        
        return witel_pivot, datel_pivot
    
    except Exception as e:
        st.error(f"Error in pivot table creation: {str(e)}")
        return None, None

# Enhanced Visualization Functions
def create_witel_summary_plot(witel_data):
    """Create professional summary plot for WITEL level"""
    # Prepare data
    plot_df = witel_data[witel_data['Witel'] != 'Grand Total'].sort_values('Total Port_Go Live', ascending=False)
    
    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Add bars
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
    
    # Add line for completion percentage
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
    
    # Add delta annotations
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
    
    # Update layout
    fig.update_layout(
        title='<b>Port Deployment Summary by WITEL</b>',
        xaxis_title='WITEL',
        yaxis_title='Port Count',
        yaxis2_title='Completion %',
        yaxis2=dict(range=[0, 105]),
        hovermode='x unified',
        barmode='stack',
        height=600,
        template='plotly_white',
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        )
    )
    
    return fig

def create_datel_radar_chart(datel_data):
    """Create radar chart for DATEL comparison"""
    # Prepare data - top 5 DATELs by Go Live ports
    top_datels = datel_data.sort_values('Total Port_Go Live', ascending=False).head(5)
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatterpolar(
        r=top_datels['Total Port_Go Live'],
        theta=top_datels.index.get_level_values(1),
        fill='toself',
        name='Go Live Ports',
        line_color='#2ca02c'
    ))
    
    fig.add_trace(go.Scatterpolar(
        r=top_datels['Total Port_On Going'],
        theta=top_datels.index.get_level_values(1),
        fill='toself',
        name='On Going Ports',
        line_color='#ff7f0e'
    ))
    
    fig.update_layout(
        title='<b>Top 5 DATELs Performance Comparison</b>',
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, top_datels['Total Port_Go Live'].max() * 1.2]
            )),
        showlegend=True,
        template='plotly_white',
        height=500
    )
    
    return fig

# Enhanced UI Components
def display_witel_summary(witel_pivot):
    """Improved WITEL summary display with styled dataframe"""
    st.subheader("🏆 WITEL Performance Ranking")
    
    # Prepare display dataframe
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
    
    # Apply styling
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
    
    # Highlight top performers
    def highlight_top3(s):
        styles = []
        for i in range(len(s)):
            if s.name != 'Grand Total' and s.iloc[i] <= 3:
                styles.append('background-color: #e6f3ff; font-weight: bold')
            else:
                styles.append('')
        return styles
    
    styled_df = styled_df.apply(highlight_top3, subset=['Rank'])
    
    # Display the styled dataframe
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=(len(display_df) * 35 + 3)
    )
    
    # Add visualization
    st.plotly_chart(create_witel_summary_plot(display_df.reset_index()), use_container_width=True)

def display_datel_details(datel_pivot):
    """Enhanced DATEL details display"""
    st.subheader("📊 DATEL-Level Performance")
    
    # Group by WITEL for tabs
    witels = datel_pivot.index.get_level_values(0).unique()
    tabs = st.tabs([f"🏅 {witel}" for witel in witels])
    
    for i, witel in enumerate(witels):
        with tabs[i]:
            witel_data = datel_pivot.loc[witel].sort_values('Total Port_Go Live', ascending=False)
            
            # Create metrics columns
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total DATELs", len(witel_data))
            col2.metric("Total Go Live", f"{witel_data['Total Port_Go Live'].sum():,}")
            col3.metric("Total On Going", f"{witel_data['Total Port_On Going'].sum():,}")
            col4.metric("Overall Completion", 
                       f"{(witel_data['Total Port_Go Live'].sum() / 
                         (witel_data['Total Port_Go Live'].sum() + 
                          witel_data['Total Port_On Going'].sum()) * 100):.1f}%")
            
            # Display DATEL table
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
                use_container_width=True,
                column_order=[
                    'Total Port_Go Live', 'Total Port_On Going', 
                    'Total_Port_All', 'Completion_Percentage', 'Rank'
                ]
            )
            
            # Add visualization
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

# In your Dashboard view section, replace the display code with:
if view_mode == "Dashboard":
    st.title("📈 PT2 IHLD Deployment Dashboard")
    
    if os.path.exists(LATEST_FILE):
        df = load_excel(LATEST_FILE)
        
        # Regional filter with improved UI
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
        
        # Create and display pivot tables
        witel_pivot, datel_pivot = create_pivot_tables(df)
        
        if witel_pivot is not None:
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total WITEL", len(witel_pivot)-1)
            col2.metric("Total Go Live Ports", f"{witel_pivot.loc['Grand Total', 'Total Port_Go Live']:,}")
            col3.metric("Total On Going Ports", f"{witel_pivot.loc['Grand Total', 'Total Port_On Going']:,}")
            col4.metric("Overall Completion", 
                       f"{(witel_pivot.loc['Grand Total', 'Total Port_Go Live'] / 
                         witel_pivot.loc['Grand Total', 'Total_Port_All'] * 100):.1f}%")
            
            # Display enhanced views
            display_witel_summary(witel_pivot)
            display_datel_details(datel_pivot)
            
            # Additional visualizations
            st.subheader("📌 Performance Insights")
            st.plotly_chart(create_datel_radar_chart(datel_pivot), use_container_width=True)
    else:
        st.warning("No data available. Please upload data in the Upload section.")
