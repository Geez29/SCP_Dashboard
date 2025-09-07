# app.py - STREAMLIT VERSION (Complete conversion from your Dash code)

import streamlit as st
import pandas as pd
import plotly.express as px

# Configure Streamlit page
st.set_page_config(
    page_title="SCP Savings Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for styling (similar to your blue theme)
st.markdown("""
<style>
    .main-header {
        font-size: 36px;
        color: #003366;
        text-align: center;
        font-weight: bold;
        margin-bottom: 30px;
    }
    .insight-box {
        background-color: #e6f0ff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    .metric-positive {
        color: green;
        font-size: 18px;
        font-weight: bold;
    }
    .metric-negative {
        color: red;
        font-size: 18px;
        font-weight: bold;
    }
    .metric-total {
        color: #003366;
        font-size: 18px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Load Excel data function
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
uploaded_file = st.sidebar.file_uploader(
    "📁 Upload Excel File (Optional)", 
    type=['xlsx', 'xls'],
    help="Upload your SCP Savings file if the default file is not found"
)

# Load data
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name="Savings_WIP_Data")
    st.sidebar.success("✅ File uploaded successfully!")
else:
    df = load_data()

# Main dashboard
if df is not None:
    # Rename columns for easier handling
    df.rename(columns={
        "Difference (PA)-Finance": "Savings_Finance",
        "Difference (PA) -SCP": "Savings_SCP",
    }, inplace=True)

    # Ensure numeric
    df["Savings_Finance"] = pd.to_numeric(df["Savings_Finance"], errors="coerce").round(0)
    df["Savings_SCP"] = pd.to_numeric(df["Savings_SCP"], errors="coerce").round(0)

    # Dashboard Header
    st.markdown('<h1 class="main-header">SCP Savings Dashboard</h1>', unsafe_allow_html=True)

    # Calculate insights
    total_finance_savings = int(df["Savings_Finance"].sum())
    positive_finance = int(df.loc[df["Savings_Finance"] > 0, "Savings_Finance"].sum())
    negative_finance = int(df.loc[df["Savings_Finance"] < 0, "Savings_Finance"].sum())

    # Insights section
    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f'<p class="metric-total">💡 Total Finance Savings: ${total_finance_savings:,}</p>', 
                   unsafe_allow_html=True)
    
    with col2:
        st.markdown(f'<p class="metric-positive">✅ Positive Finance Savings: ${positive_finance:,}</p>', 
                   unsafe_allow_html=True)
    
    with col3:
        st.markdown(f'<p class="metric-negative">⚠️ Negative Finance Savings: ${negative_finance:,}</p>', 
                   unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

    # Data Table Section
    st.subheader("📊 Savings Data Table")
    
    # Create display dataframe with formatting
    df_display = df.copy()
    df_display["Savings_Finance"] = df_display["Savings_Finance"].apply(
        lambda x: f"${int(x):,}" if pd.notnull(x) else ""
    )
    df_display["Savings_SCP"] = df_display["Savings_SCP"].apply(
        lambda x: f"${int(x):,}" if pd.notnull(x) else ""
    )
    
    # Display table with conditional formatting
    def highlight_negative(val):
        """Highlight negative values in red, positive in green"""
        if isinstance(val, str) and val.startswith('$-'):
            return 'color: red; font-weight: bold'
        elif isinstance(val, str) and val.startswith('$') and not val.startswith('$-'):
            return 'color: green; font-weight: bold'
        return ''
    
    styled_df = df_display.style.applymap(
        highlight_negative, 
        subset=['Savings_Finance', 'Savings_SCP']
    ).set_table_styles([
        {'selector': 'th', 'props': [
            ('background-color', '#003366'),
            ('color', 'white'),
            ('font-weight', 'bold'),
            ('text-align', 'center'),
            ('padding', '8px')
        ]},
        {'selector': 'td', 'props': [
            ('text-align', 'center'),
            ('padding', '8px')
        ]}
    ])
    
    st.dataframe(styled_df, use_container_width=True, height=400)

    # Charts Section
    st.subheader("📈 Visualizations")
    
    # Create Finance Chart
    fig_finance = px.bar(
        df,
        x="FY of Savings-Finance",
        y="Savings_Finance",
        text="Savings_Finance",
        title="Yearly Finance Savings",
        color_discrete_sequence=["#003366"]
    )
    
    fig_finance.update_traces(
        texttemplate="$%{text:,.0f}", 
        textposition="outside"
    )
    
    fig_finance.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font=dict(size=20, color="#003366", family="Arial"),
        font=dict(size=14, color="black"),
        height=500
    )
    
    fig_finance.update_yaxes(title="Savings ($)")
    fig_finance.update_xaxes(title="Fiscal Year")
    
    # Create SCP Chart
    fig_scp = px.bar(
        df,
        x="FY of Savings-SCP",
        y="Savings_SCP",
        text="Savings_SCP",
        title="Yearly SCP Savings",
        color_discrete_sequence=["#336699"]
    )
    
    fig_scp.update_traces(
        texttemplate="$%{text:,.0f}", 
        textposition="outside"
    )
    
    fig_scp.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font=dict(size=20, color="#003366", family="Arial"),
        font=dict(size=14, color="black"),
        height=500
    )
    
    fig_scp.update_yaxes(title="Savings ($)")
    fig_scp.update_xaxes(title="Fiscal Year")
    
    # Display charts in columns
    col1, col2 = st.columns(2)
    
    with col1:
        st.plotly_chart(fig_finance, use_container_width=True)
    
    with col2:
        st.plotly_chart(fig_scp, use_container_width=True)

    # Additional Analytics Section
    st.subheader("📊 Additional Analytics")
    
    analytics_col1, analytics_col2, analytics_col3 = st.columns(3)
    
    with analytics_col1:
        # Summary metrics
        st.metric(
            "Average Finance Savings", 
            f"${df['Savings_Finance'].mean():,.0f}",
            delta=f"{((positive_finance + negative_finance) / len(df)):,.0f}"
        )
    
    with analytics_col2:
        st.metric(
            "Total SCP Savings", 
            f"${df['Savings_SCP'].sum():,.0f}",
            delta=f"{df['Savings_SCP'].mean():,.0f}"
        )
    
    with analytics_col3:
        positive_count = len(df[df['Savings_Finance'] > 0])
        negative_count = len(df[df['Savings_Finance'] < 0])
        st.metric(
            "Positive vs Negative",
            f"{positive_count}/{negative_count}",
            delta=f"{positive_count - negative_count} difference"
        )

    # Download section
    st.subheader("💾 Download Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Download raw data as CSV
        csv_data = df.to_csv(index=False)
        st.download_button(
            label="📥 Download Raw Data (CSV)",
            data=csv_data,
            file_name="scp_savings_raw_data.csv",
            mime="text/csv"
        )
    
    with col2:
        # Download formatted data as CSV
        csv_display = df_display.to_csv(index=False)
        st.download_button(
            label="📥 Download Formatted Data (CSV)",
            data=csv_display,
            file_name="scp_savings_formatted_data.csv",
            mime="text/csv"
        )

else:
    # Error state - no data available
    st.error("⚠️ No data available!")
    st.info("Please ensure 'SCP_Savings_FY26_dummy_v3.xlsx' is in the same directory as this app, or upload a file using the sidebar.")
    
    # Show sample data structure
    st.subheader("Expected Data Structure")
    sample_data = {
        'FY of Savings-Finance': ['FY2024', 'FY2025', 'FY2026'],
        'FY of Savings-SCP': ['FY2024', 'FY2025', 'FY2026'],
        'Difference (PA)-Finance': [10000, -5000, 15000],
        'Difference (PA) -SCP': [8000, -3000, 12000]
    }
    sample_df = pd.DataFrame(sample_data)
    st.dataframe(sample_df, use_container_width=True)

# Footer
st.markdown("---")
st.markdown(
    "Built with ❤️ using Streamlit | SCP Savings Dashboard", 
    help="This dashboard analyzes SCP savings data across fiscal years"
)
