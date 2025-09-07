# app.py - Executive SCP Savings Dashboard with OneDrive Integration

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import numpy as np
import requests
from io import BytesIO

# Configure Streamlit page
st.set_page_config(
    page_title="SCP Savings Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Executive-style CSS with font sizes reduced by 1.5 points
st.markdown("""
<style>
    .main-header {
        font-size: 42px;
        color: #003366;
        text-align: center;
        font-weight: 700;
        margin-bottom: 40px;
        font-family: 'Helvetica Neue', sans-serif;
        letter-spacing: -0.5px;
    }
    .insight-container {
        display: flex;
        justify-content: space-between;
        margin-bottom: 40px;
        gap: 15px;
    }
    .insight-box {
        background: linear-gradient(135deg, #003366 0%, #004080 100%);
        color: white;
        padding: 30px 20px;
        border-radius: 15px;
        box-shadow: 0 8px 32px rgba(0,51,102,0.2);
        flex: 1;
        text-align: center;
        border: 1px solid rgba(255,255,255,0.1);
        transition: transform 0.3s ease;
    }
    .insight-box:hover {
        transform: translateY(-5px);
    }
    .insight-gains {
        background: linear-gradient(135deg, #1e7e34 0%, #28a745 100%);
    }
    .insight-risks {
        background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
    }
    .insight-title {
        font-size: 12.5px;
        font-weight: 500;
        margin-bottom: 8px;
        opacity: 0.95;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .insight-value {
        font-size: 26.5px;
        font-weight: 700;
        margin-bottom: 8px;
        line-height: 1;
    }
    .insight-subtitle {
        font-size: 9.5px;
        opacity: 0.85;
        font-style: italic;
    }
    .filter-section {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        padding: 25px;
        border-radius: 15px;
        border: 1px solid #cbd5e0;
        margin-bottom: 30px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    }
    .chart-container {
        background-color: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        margin-bottom: 25px;
        border: 1px solid #e2e8f0;
    }
    .section-header {
        font-size: 24px;
        color: #003366;
        font-weight: 600;
        margin-bottom: 20px;
        font-family: 'Helvetica Neue', sans-serif;
    }
    .data-summary {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
        font-weight: 500;
    }
    .onedrive-config {
        background: linear-gradient(135deg, #e6f3ff 0%, #ccebff 100%);
        padding: 20px;
        border-radius: 10px;
        border-left: 4px solid #003366;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# OneDrive file configuration
def convert_onedrive_url(share_url):
    """Convert OneDrive sharing URL to direct download URL"""
    try:
        # For OneDrive URLs, attempt to convert to direct download
        if "onedrive.live.com" in share_url:
            # Extract file ID and convert to direct download
            if "resid=" in share_url:
                # Try to construct direct download URL
                base_url = share_url.split("?")[0]
                direct_url = base_url + "?download=1"
                return direct_url
        elif "1drv.ms" in share_url:
            # Handle shortened OneDrive URLs
            return share_url + "&download=1" if "?" in share_url else share_url + "?download=1"
        
        return share_url
    except:
        return share_url

# Load data from OneDrive
@st.cache_data
def load_data_from_onedrive(onedrive_url):
    try:
        # Convert sharing URL to direct download
        direct_url = convert_onedrive_url(onedrive_url)
        
        # Try loading directly first
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            response = requests.get(direct_url, headers=headers, timeout=30, allow_redirects=True)
            
            if response.status_code == 200:
                df = pd.read_excel(BytesIO(response.content), sheet_name="Savings_WIP_Data")
                return df, "OneDrive file loaded successfully"
        except Exception as e:
            # Try alternative approach
            try:
                response = requests.get(onedrive_url, headers=headers, timeout=30)
                if response.status_code == 200:
                    df = pd.read_excel(BytesIO(response.content), sheet_name="Savings_WIP_Data")
                    return df, "OneDrive file loaded successfully (alternative method)"
            except:
                pass
        
        # Fallback to local file for demonstration
        try:
            df = pd.read_excel("SCP_Savings_FY26_dummy_v3.xlsx", sheet_name="Savings_WIP_Data")
            return df, "Using local file (OneDrive connection issue - please check URL)"
        except:
            return None, "Unable to load data from OneDrive or local file"
            
    except Exception as e:
        return None, f"Error loading data: {str(e)}"

# Dashboard Header
st.markdown('<h1 class="main-header">Executive SCP Savings Dashboard</h1>', unsafe_allow_html=True)

# OneDrive Configuration Section
with st.sidebar:
    st.markdown('<div class="onedrive-config">', unsafe_allow_html=True)
    st.header("🔗 OneDrive Configuration")
    
    # *** THIS IS THE LINE TO UPDATE YOUR ONEDRIVE URL ***
    default_url = "https://onedrive.live.com/:x:/g/personal/9E1C07238F947303/EbI62L-aBvdDgxmyFIMOdugB5BoH7r7ATZcU1ywNSR1Psw?resid=9E1C07238F947303!sbfd83ab2069a43f78319b214830e76e8&ithint=file%2Cxlsx&e=bSpS5T&migratedtospo=true&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy85ZTFjMDcyMzhmOTQ3MzAzL0ViSTYyTC1hQnZkRGd4bXlGSU1PZHVnQjVCb0g3cjdBVFpjVTF5d05TUjFQc3c_ZT1iU3BTNVQ"
    
    onedrive_url = st.text_input(
        "OneDrive File URL:",
        value=default_url,
        help="Paste your OneDrive Excel file sharing URL here"
    )
    
    if st.button("🔄 Reload Data"):
        st.cache_data.clear()
        st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Data source info
    st.info("📊 Data Source: OneDrive Integration")
    st.markdown("---")
    st.markdown("**Dashboard Features:**")
    st.markdown("• Executive Summary")
    st.markdown("• Strategic Analytics")
    st.markdown("• Domain Analysis")
    st.markdown("• Portfolio Overview")

# Load data from OneDrive
df, load_message = load_data_from_onedrive(onedrive_url)

if df is not None:
    st.success(f"✅ {load_message}")
    
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

    # FILTERS SECTION
    st.markdown('<div class="filter-section">', unsafe_allow_html=True)
    st.markdown('<h3 class="section-header">📊 Business Intelligence Filters</h3>', unsafe_allow_html=True)
    
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    
    with filter_col1:
        # Contract Start Date Filter
        if "Contract Start" in df.columns:
            min_start_date = df["Contract Start"].min()
            max_start_date = df["Contract Start"].max()
            if pd.notna(min_start_date) and pd.notna(max_start_date):
                start_date_filter = st.date_input(
                    "Contract Start Date",
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
                    "Contract End Date",
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
                "Finance FY",
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
                "SCP FY",
                options=scp_fy_options,
                index=0
            )
        else:
            scp_fy_filter = "All"

    # Domain filter (full width)
    if "Domain" in df.columns:
        domain_options = ["All Domains"] + sorted(df["Domain"].dropna().unique().tolist())
        domain_filter = st.selectbox(
            "🏢 Business Domain",
            options=domain_options,
            index=0
        )
    else:
        domain_filter = "All Domains"
    
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
    if domain_filter != "All Domains":
        filtered_df = filtered_df[filtered_df["Domain"] == domain_filter]

    # Calculate insights from filtered data
    total_finance_savings = filtered_df["Savings_Finance"].sum()
    gains_finance = filtered_df.loc[filtered_df["Savings_Finance"] > 0, "Savings_Finance"].sum()
    risks_finance = filtered_df.loc[filtered_df["Savings_Finance"] < 0, "Savings_Finance"].sum()
    
    total_scp_savings = filtered_df["Savings_SCP"].sum()
    gains_scp = filtered_df.loc[filtered_df["Savings_SCP"] > 0, "Savings_SCP"].sum()
    risks_scp = filtered_df.loc[filtered_df["Savings_SCP"] < 0, "Savings_SCP"].sum()

    # EXECUTIVE SUMMARY PANEL
    st.markdown('<h2 class="section-header">📈 Executive Summary</h2>', unsafe_allow_html=True)
    
    insights_col1, insights_col2, insights_col3, insights_col4, insights_col5, insights_col6 = st.columns(6)
    
    with insights_col1:
        st.markdown(f"""
        <div class="insight-box">
            <div class="insight-title">Net Finance Impact</div>
            <div class="insight-value">${total_finance_savings:,.0f}</div>
            <div class="insight-subtitle">Total Portfolio</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col2:
        st.markdown(f"""
        <div class="insight-box insight-gains">
            <div class="insight-title">Finance Upside</div>
            <div class="insight-value">${gains_finance:,.0f}</div>
            <div class="insight-subtitle">Value Creation</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col3:
        st.markdown(f"""
        <div class="insight-box insight-risks">
            <div class="insight-title">Finance Exposure</div>
            <div class="insight-value">${abs(risks_finance):,.0f}</div>
            <div class="insight-subtitle">Risk Management</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col4:
        st.markdown(f"""
        <div class="insight-box">
            <div class="insight-title">Net SCP Impact</div>
            <div class="insight-value">${total_scp_savings:,.0f}</div>
            <div class="insight-subtitle">Total Portfolio</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col5:
        st.markdown(f"""
        <div class="insight-box insight-gains">
            <div class="insight-title">SCP Upside</div>
            <div class="insight-value">${gains_scp:,.0f}</div>
            <div class="insight-subtitle">Value Creation</div>
        </div>
        """, unsafe_allow_html=True)

    with insights_col6:
        st.markdown(f"""
        <div class="insight-box insight-risks">
            <div class="insight-title">SCP Exposure</div>
            <div class="insight-value">${abs(risks_scp):,.0f}</div>
            <div class="insight-subtitle">Risk Management</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # STRATEGIC ANALYTICS SECTION
    st.markdown('<h2 class="section-header">📊 Strategic Analytics</h2>', unsafe_allow_html=True)

    # McKinsey blue gradient colors
    mckinsey_blues = ['#001f3f', '#003366', '#004080', '#0066cc', '#3399ff', '#66b3ff', '#99ccff']
    
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        # Finance Savings by FY
        if "FY of Savings-Finance" in filtered_df.columns:
            finance_fy_data = filtered_df.groupby("FY of Savings-Finance")["Savings_Finance"].sum().reset_index()
            finance_fy_data = finance_fy_data.sort_values("FY of Savings-Finance")
            
            # Create gradient colors
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
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-Finance']}<br><b>Impact:</b> ${row['Savings_Finance']:,.0f}<extra></extra>"
                ))
            
            fig_finance.update_layout(
                title={
                    'text': "Finance Impact by Fiscal Year",
                    'x': 0.5,
                    'font': {'size': 18, 'color': '#003366', 'family': 'Helvetica Neue'}
                },
                xaxis_title="Fiscal Year",
                yaxis_title="Financial Impact ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=11, color="#003366"),
                height=450,
                xaxis=dict(showgrid=False, tickangle=0),
                yaxis=dict(showgrid=True, gridcolor='#f0f0f0', gridwidth=1)
            )
            
            st.plotly_chart(fig_finance, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with chart_col2:
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        
        # SCP Savings by FY
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
                    hovertemplate=f"<b>FY:</b> {row['FY of Savings-SCP']}<br><b>Impact:</b> ${row['Savings_SCP']:,.0f}<extra></extra>"
                ))
            
            fig_scp.update_layout(
                title={
                    'text': "SCP Impact by Fiscal Year",
                    'x': 0.5,
                    'font': {'size': 18, 'color': '#003366', 'family': 'Helvetica Neue'}
                },
                xaxis_title="Fiscal Year",
                yaxis_title="SCP Impact ($)",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font=dict(size=11, color="#003366"),
                height=450,
                xaxis=dict(showgrid=False, tickangle=0),
                yaxis=dict(showgrid=True, gridcolor='#f0f0f0', gridwidth=1)
            )
            
            st.plotly_chart(fig_scp, use_container_width=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # DOMAIN ANALYSIS
    st.markdown('<div class="chart-container">', unsafe_allow_html=True)
    st.markdown('<h3 class="section-header">🏢 Business Domain Analysis</h3>', unsafe_allow_html=True)
    
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
            hovertemplate="<b>Domain:</b> %{y}<br><b>Financial Impact:</b> $%{x:,.0f}<extra></extra>"
        ))
        
        fig_domain.update_layout(
            title={
                'text': "Financial Impact by Business Domain",
                'x': 0.5,
                'font': {'size': 18, 'color': '#003366', 'family': 'Helvetica Neue'}
            },
            xaxis_title="Financial Impact ($)",
            yaxis_title="Business Domain",
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(size=11, color="#003366"),
            height=max(400, n_domains * 45),
            showlegend=False,
            xaxis=dict(showgrid=True, gridcolor='#f0f0f0', gridwidth=1),
            yaxis=dict(showgrid=False)
        )
        
        st.plotly_chart(fig_domain, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

    # PORTFOLIO OVERVIEW
    st.markdown('<h2 class="section-header">📋 Portfolio Overview</h2>', unsafe_allow_html=True)
    
    # Portfolio summary metrics
    portfolio_col1, portfolio_col2, portfolio_col3, portfolio_col4 = st.columns(4)
    
    with portfolio_col1:
        st.metric("Active Contracts", len(filtered_df))
    
    with portfolio_col2:
        avg_finance = filtered_df["Savings_Finance"].mean()
        st.metric("Avg Finance Impact", f"${avg_finance:,.0f}")
    
    with portfolio_col3:
        avg_scp = filtered_df["Savings_SCP"].mean()
        st.metric("Avg SCP Impact", f"${avg_scp:,.0f}")
    
    with portfolio_col4:
        if "Domain" in filtered_df.columns:
            unique_domains = filtered_df["Domain"].nunique()
            st.metric("Business Domains", unique_domains)

    # Data export section
    st.markdown("### 💾 Data Export")
    
    # Summary for executives
    total_records = len(filtered_df)
    if total_records != len(df):
        st.markdown(f'<div class="data-summary">Portfolio Analysis: {total_records:,} contracts selected from {len(df):,} total records</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="data-summary">Complete Portfolio Analysis: {total_records:,} active contracts</div>', unsafe_allow_html=True)

    # Download options
    export_col1, export_col2 = st.columns(2)
    
    with export_col1:
        # Executive summary export
        summary_data = {
            'Metric': ['Net Finance Impact', 'Finance Upside', 'Finance Exposure', 'Net SCP Impact', 'SCP Upside', 'SCP Exposure'],
            'Value': [total_finance_savings, gains_finance, abs(risks_finance), total_scp_savings, gains_scp, abs(risks_scp)]
        }
        summary_df = pd.DataFrame(summary_data)
        csv_summary = summary_df.to_csv(index=False)
        st.download_button(
            label="📊 Download Executive Summary",
            data=csv_summary,
            file_name=f"executive_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )
    
    with export_col2:
        # Full portfolio export
        csv_data = filtered_df.to_csv(index=False)
        st.download_button(
            label="📁 Download Portfolio Data",
            data=csv_data,
            file_name=f"portfolio_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

else:
    # Error state
    st.error(f"⚠️ {load_message}")
    st.info("Please check your OneDrive URL and ensure the file is accessible.")
    
    # Show example URL format
    st.markdown("### 📝 OneDrive URL Format")
    st.code("https://onedrive.live.com/:x:/g/personal/[user-id]/[file-params]")
    
    st.markdown("### 🔧 Troubleshooting")
    st.markdown("• Ensure the OneDrive file has appropriate sharing permissions")
    st.markdown("• Verify the Excel file contains 'Savings_WIP_Data' sheet")
    st.markdown("• Check if the file URL is accessible from external applications")
    st.markdown("• Contact IT support if OneDrive API integration is required")

# Footer
st.markdown("---")
st.markdown("**Executive SCP Savings Dashboard** | OneDrive Integration | Strategic Portfolio Analytics | Confidential")
