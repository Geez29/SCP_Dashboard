# app.py - Enhanced SCP Savings Dashboard with McKinsey-style visuals

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import numpy as np

# Configure Streamlit page
st.set_page_config(
    page_title="SCP Savings Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# McKinsey-style CSS
st.markdown("""
<style>
    .main-header {
        font-size: 42px;
        color: #003366;
        text-align: center;
        font-weight: 700;
        margin-bottom: 30px;
        font-family: 'Helvetica Neue', sans-serif;
    }
    .insight-container {
        display: flex;
        justify-content: space-between;
        margin-bottom: 30px;
        gap: 20px;
    }
    .insight-box {
        background: linear-gradient(135deg, #003366 0%, #004080 100%);
        color: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 8px 24px rgba(0,51,102,0.15);
        flex: 1;
        text-align: center;
        border-left: 5px solid #00ccff;
    }
    .insight-positive {
        background: linear-gradient(135deg, #006633 0%, #00804d 100%);
        border-left: 5px solid #00ff80;
    }
    .insight-negative {
        background: linear-gradient(135deg, #cc0000 0%, #e60000 100%);
        border-left: 5px solid #ff6666;
    }
    .insight-title {
        font-size: 16px;
        font-weight: 500;
        margin-bottom: 10px;
        opacity: 0.9;
    }
    .insight-value {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .insight-subtitle {
        font-size: 12px;
        opacity: 0.8;
    }
    .filter-container {
        background-color: #f8fafc;
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #003366;
        margin-bottom: 20px;
    }
    .chart-container {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Load data function
@st.cache_data
def load_data():
    try:
        file_path = "SCP_Savings_FY26_dummy_v3.xlsx"
        df = pd.read_excel(file_path, sheet_name="Savings_WIP_Data")
        return df
    except FileNotFoundError:
        st.error("❌ Excel file 'SCP_Savings_FY26_dummy_v3.xlsx' not found.")
        return None
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return None

# File upload option
with st.sidebar:
    st.header("📁 Data Upload")
    uploaded_file = st.file_uploader(
        "Upload Excel File", 
        type=['xlsx', 'xls'],
        help="Upload your SCP Savings file"
    )

# Load data
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Savings_WIP_Data")
    st.sidebar.success("✅ File uploaded successfully!")
else:
    df = load_data()

if df is not None:
    # Data preprocessing
    df.rename(columns={
        "Difference (PA)-Finance": "Savings_Finance",
        "Difference (PA) -SCP": "Savings_SCP",
    }, inplace=True)
    
    # Ensure numeric columns
    df["Savings_Finance"] = pd.to_numeric(df["Savings_Finance"], errors="coerce").fillna(0)
    df["Savings_SCP"] = pd.to_numeric(df["Savings_SCP"], errors="coerce").fillna(0)
    
    # Convert date columns
    date_columns = ["Contract Start", "Contract End"]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Dashboard Header
    st.markdown('<h1 class="main-header">SCP Savings Dashboard</h1>', unsafe_allow_html=True)

    # FILTERS SECTION
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)
    st.subheader("🔍 Dashboard Filters")
    
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    
    with filter_col1:
        # Contract Start Date Filter
        if "Contract Start" in df.columns:
            min_start_date = df["Contract Start"].min()
            max_start_date = df["Contract Start"].max()
            if pd.notna(min_start_date) and pd.notna(max_start_date):
                start_date_filter = st.date_input(
                    "Contract Start Date (From)",
                    value=min_start_date.date(),
                    min_value=min_start_date.date(),
                    max_value=max_start_date.date()
                )
            else:
                start_date_filter = None
        else:
            start_date_filter = None

    with filter_col2:
        # Contract End Date Filter
        if "Contract End" in df.columns:
            min_end_date = df["Contract End"].min()
            max_end_date = df["Contract End"].max()
            if pd.notna(min_end_date) and pd.notna(max_end_date):
                end_date_filter = st.date_input(
                    "Contract End Date (To)",
                    value=max_end_date.date(),
                    min_value=min_end_date.date(),
                    max_value=max_end_date.date()
                )
            else:
                end_date_filter = None
        else:
            end_date_filter = None

    with filter_col3:
        # FY of Savings-Finance Filter
        if "FY of Savings-Finance" in df.columns:
            finance_fy_options = ["All"] + sorted(df["FY of Savings-Finance"].dropna().unique().tolist())
            finance_fy_filter = st.selectbox(
                "FY of Savings-Finance",
                options=finance_fy_options,
                index=0
            )
        else:
            finance_fy_filter = "All"

    with filter_col4:
        # FY of Savings-SCP Filter
        if "FY of Savings-SCP" in df.columns:
            scp_fy_options = ["All"] + sorted(df["FY of Savings-SCP"].dropna().unique().tolist())
            scp_fy_filter = st.selectbox(
                "FY of Savings-SCP",
                options=scp_fy_options,
                index=0
            )
        else:
            scp_fy_filter = "All"

    # Domain filter (full width)
    if "Domain" in df.columns:
        domain_options = ["All"] + sorted(df["Domain"].dropna().unique().tolist())
        domain_filter = st.selectbox(
            "🏢 Domain Filter",
            options=domain_options,
            index=0
        )
    else:
        domain_filter = "All"
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Apply filters
    filtered_df = df.copy()
    
    # Date filters
    if start_date_filter and "Contract Start" in df.columns:
        filtered_df = filtered_df[filtered_df["Contract Start"] >= pd.Timestamp(start_date_filter)]
    
    if end_date_filter and "Contract End" in df.columns:
        filtered_df = filtered_df[filtered_df["Contract End"] <= pd.Timestamp(end_date_filter)]
    
    # FY filters
    if finance_fy_filter != "All":
        filtered_df = filtered_df[filtered_df["FY of Savings-Finance"] == finance_fy_filter]
    
    if scp_fy_filter != "All":
        filtered_df = filtered_df[filtered_df["FY of Savings-SCP"] == scp_fy_filter]
    
    # Domain filter
    if domain_filter != "All":
        filtered_df = filtered_df[filtered_df["Domain"] == domain_filter]

    # Calculate insights from filtered data
    total_finance_savings = filtered_df["Savings_Finance"].sum()
    positive_finance = filtered_df.loc[filtered_df["Savings_Finance"] > 0, "Savings_Finance"].sum()
    negative_finance = filtered_df.loc[filtered_df["Savings_Finance"] < 0, "Savings_Finance"].sum()
    
    total_scp_savings = filtered_df["Savings_SCP"].sum()
    positive_scp = filtered_df.loc[filtered_df["Savings_SCP"] > 0, "Savings_SCP"].sum()
    negative_scp = filtered_df.loc[filtered_df["Savings_SCP"] < 0, "Savings_SCP"].sum()

    # INSIGHTS PANEL - McKinsey Style
    st.markdown("### 📊 Executive Summary")
    
    insights_col1, insights_col2, insights_col3, insights_col4, insights_col5, insights_col6 = st.columns(6)
    
    with insights_col1:
        st.markdown(f"""
        <div class="insight-box">
            <div class="insight-title">Total Finance Savings</div>
            <div class="insight-value">${total_finance_savings:,.0f}</div>
            <div class="insight-subtitle">Filtered Results</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col2:
        st.markdown(f"""
        <div class="insight-box insight-positive">
            <div class="insight-title">Positive Finance</div>
            <div class="insight-value">${positive_finance:,.0f}</div>
            <div class="insight-subtitle">Gains</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col3:
        st.markdown(f"""
        <div class="insight-box insight-negative">
            <div class="insight-title">Negative Finance</div>
            <div class="insight-value">${negative_finance:,.0f}</div>
            <div class="insight-subtitle">Losses</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col4:
        st.markdown(f"""
        <div class="insight-box">
            <div class="insight-title">Total SCP Savings</div>
            <div class="insight-value">${total_scp_savings:,.0f}</div>
            <div class="insight-subtitle">Filtered Results</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col5:
        st.markdown(f"""
        <div class="insight-box insight-positive">
            <div class="insight-title">Positive SCP</div>
            <div class="insight-value">${positive_scp:,.0f}</div>
            <div class="insight-subtitle">Gains</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col6:
        st.markdown(f"""
        <div class="insight-box insight-negative">
            <div class="insight-title">Negative SCP</div>
            <div class="insight-value">${negative_scp:,.0f}</div>
            <div class="insight-subtitle">Losses</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # CHARTS SECTION
    st.markdown("### 📈 Financial Analysis")

    # McKinsey blue gradient colors
    mckinsey_blues = ['#001f3f', '#003366', '#004080', '#0066cc', '#3399ff', '#66b3ff', '#99ccff']
    
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        # Gradient Vertical Bar Chart for Finance Savings by FY
        if "FY of Savings-Finance" in filtered_df.columns:
            finance_fy_data = filtered_df.groupby("FY of Savings-Finance")["Savings_Finance"].sum().reset_index()
            finance_fy_data = finance_fy_data.sort_values("FY of Savings-Finance")
            
            # Create gradient colors based on data points
            n_bars = len(finance_fy_data)
            colors = [mckinsey_blues[i % len(mckinsey_blues)] for i in range(n_bars)]
            
            fig_finance = go.Figure()
            
            for i, row in finance_fy_data.iterrows():
                fig_finance.add_trace(go.Bar(
                    x=[row["FY of Savings-Finance"]],
                    y=[row["Savings_Finance"]],
                    name=row["FY of Savings-Finance"],
                    marker_color=colors[i],
                    text=f"${row['Savings_Finance']:,.0f}",
                    textposition="outside",
                    showlegend=False,
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-Finance']}<br><b>Savings:</b> ${row['Savings_Finance']:,.0f}<extra></extra>"
                ))
            
            fig_finance.update_layout(
                title={
                    'text': "Finance Savings by Fiscal Year",
                    'x': 0.5,
                    'font': {'size': 20, 'color': '#003366', 'family': 'Helvetica Neue'}
                },
                xaxis_title="Fiscal Year",
                yaxis_title="Savings ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=12, color="#003366"),
                height=500,
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor='lightgray', gridwidth=0.5)
            )
            
            st.plotly_chart(fig_finance, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with chart_col2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        # Gradient Vertical Bar Chart for SCP Savings by FY
        if "FY of Savings-SCP" in filtered_df.columns:
            scp_fy_data = filtered_df.groupby("FY of Savings-SCP")["Savings_SCP"].sum().reset_index()
            scp_fy_data = scp_fy_data.sort_values("FY of Savings-SCP")
            
            # Create gradient colors
            n_bars = len(scp_fy_data)
            colors = [mckinsey_blues[i % len(mckinsey_blues)] for i in range(n_bars)]
            
            fig_scp = go.Figure()
            
            for i, row in scp_fy_data.iterrows():
                fig_scp.add_trace(go.Bar(
                    x=[row["FY of Savings-SCP"]],
                    y=[row["Savings_SCP"]],
                    name=row["FY of Savings-SCP"],
                    marker_color=colors[i],
                    text=f"${row['Savings_SCP']:,.0f}",
                    textposition="outside",
                    showlegend=False,
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-SCP']}<br><b>Savings:</b> ${row['Savings_SCP']:,.0f}<extra></extra>"
                ))
            
            fig_scp.update_layout(
                title={
                    'text': "SCP Savings by Fiscal Year",
                    'x': 0.5,
                    'font': {'size': 20, 'color': '#003366', 'family': 'Helvetica Neue'}
                },
                xaxis_title="Fiscal Year",
                yaxis_title="Savings ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=12, color="#003366"),
                height=500,
                xaxis=dict(showgrid=False),
                yaxis=dict(showgrid=True, gridcolor='lightgray', gridwidth=0.5)
            )
            
            st.plotly_chart(fig_scp, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # HORIZONTAL BAR CHART FOR DOMAINS
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown("### 🏢 Savings by Domain")
    
    if "Domain" in filtered_df.columns:
        domain_finance = filtered_df.groupby("Domain")["Savings_Finance"].sum().reset_index()
        domain_finance = domain_finance.sort_values("Savings_Finance", ascending=True)
        
        # Create gradient colors for domains
        n_domains = len(domain_finance)
        domain_colors = [mckinsey_blues[i % len(mckinsey_blues)] for i in range(n_domains)]
        
        fig_domain = go.Figure()
        
        fig_domain.add_trace(go.Bar(
            x=domain_finance["Savings_Finance"],
            y=domain_finance["Domain"],
            orientation='h',
            marker=dict(
                color=domain_colors,
                line=dict(color='white', width=1)
            ),
            text=[f"${val:,.0f}" for val in domain_finance["Savings_Finance"]],
            textposition="outside",
            hovertemplate="<b>Domain:</b> %{y}<br><b>Savings:</b> $%{x:,.0f}<extra></extra>"
        ))
        
        fig_domain.update_layout(
            title={
                'text': "Finance Savings by Domain (Filtered)",
                'x': 0.5,
                'font': {'size': 20, 'color': '#003366', 'family': 'Helvetica Neue'}
            },
            xaxis_title="Savings ($)",
            yaxis_title="Domain",
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(size=12, color="#003366"),
            height=max(400, n_domains * 40),
            showlegend=False,
            xaxis=dict(showgrid=True, gridcolor='lightgray', gridwidth=0.5),
            yaxis=dict(showgrid=False)
        )
        
        st.plotly_chart(fig_domain, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

    # DATA TABLE SECTION
    st.markdown("### 📋 Filtered Data Table")
    
    # Show filter summary
    filter_summary = f"Showing {len(filtered_df)} records"
    if len(filtered_df) != len(df):
        filter_summary += f" (filtered from {len(df)} total)"
    st.info(filter_summary)
    
    # Display columns selection
    display_columns = st.multiselect(
        "Select columns to display:",
        options=list(filtered_df.columns),
        default=["Domain", "Forecast ID", "Brand Vendor Term Description", "Contract Start", "Contract End", 
                "FY of Savings-Finance", "FY of Savings-SCP", "Savings_Finance", "Savings_SCP"],
        help="Choose which columns to show in the table below"
    )
    
    if display_columns:
        # Format display data
        display_df = filtered_df[display_columns].copy()
        
        # Format savings columns
        savings_cols = ["Savings_Finance", "Savings_SCP"]
        for col in savings_cols:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(lambda x: f"${x:,.0f}" if pd.notnull(x) else "")
        
        # Format date columns
        date_cols = ["Contract Start", "Contract End"]
        for col in date_cols:
            if col in display_df.columns:
                display_df[col] = display_df[col].dt.strftime('%Y-%m-%d')
        
        st.dataframe(display_df, use_container_width=True, height=400)
        
        # Download buttons
        col1, col2 = st.columns(2)
        
        with col1:
            csv_data = filtered_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Filtered Data (CSV)",
                data=csv_data,
                file_name=f"scp_savings_filtered_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
        
        with col2:
            csv_display = display_df.to_csv(index=False)
            st.download_button(
                label="📥 Download Display Data (CSV)",
                data=csv_display,
                file_name=f"scp_savings_display_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )

else:
    st.error("⚠️ No data available!")
    st.info("Please upload an Excel file or ensure the data file is in the correct location.")

# Footer
st.markdown("---")
st.markdown("**SCP Savings Dashboard** | Built with Streamlit | McKinsey-style Analytics")
