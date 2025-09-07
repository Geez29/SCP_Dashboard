import dash
from dash import dcc, html, dash_table
import pandas as pd
import plotly.express as px

# Load Excel data
file_path = "SCP_Savings_FY26_dummy_v3.xlsx"
df = pd.read_excel(file_path, sheet_name="Savings_WIP_Data")

# Rename for easier handling
df.rename(columns={
    "Difference (PA)-Finance": "Savings_Finance",
    "Difference (PA) -SCP": "Savings_SCP",
}, inplace=True)

# Ensure numeric
df["Savings_Finance"] = pd.to_numeric(df["Savings_Finance"], errors="coerce").round(0)
df["Savings_SCP"] = pd.to_numeric(df["Savings_SCP"], errors="coerce").round(0)

# Copy for display formatting
df_display = df.copy()
df_display["Savings_Finance"] = df_display["Savings_Finance"].apply(lambda x: f"${int(x):,}" if pd.notnull(x) else "")
df_display["Savings_SCP"] = df_display["Savings_SCP"].apply(lambda x: f"${int(x):,}" if pd.notnull(x) else "")

# Insights
total_finance_savings = int(df["Savings_Finance"].sum())
positive_finance = int(df.loc[df["Savings_Finance"] > 0, "Savings_Finance"].sum())
negative_finance = int(df.loc[df["Savings_Finance"] < 0, "Savings_Finance"].sum())

# Finance Chart (Blue Theme)
fig_finance = px.bar(
    df,
    x="FY of Savings-Finance",
    y="Savings_Finance",
    text="Savings_Finance",
    title="Yearly Finance Savings",
    color_discrete_sequence=["#003366"]
)
fig_finance.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
fig_finance.update_layout(
    plot_bgcolor="white",
    paper_bgcolor="white",
    title_font=dict(size=20, color="#003366", family="Arial"),
    font=dict(size=14, color="black")
)
fig_finance.update_yaxes(title="Savings ($)")
fig_finance.update_xaxes(title="Fiscal Year")

# SCP Chart (Blue Theme)
fig_scp = px.bar(
    df,
    x="FY of Savings-SCP",
    y="Savings_SCP",
    text="Savings_SCP",
    title="Yearly SCP Savings",
    color_discrete_sequence=["#336699"]
)
fig_scp.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
fig_scp.update_layout(
    plot_bgcolor="white",
    paper_bgcolor="white",
    title_font=dict(size=20, color="#003366", family="Arial"),
    font=dict(size=14, color="black")
)
fig_scp.update_yaxes(title="Savings ($)")
fig_scp.update_xaxes(title="Fiscal Year")

# Dash app
app = dash.Dash(__name__)
app.layout = html.Div(style={"fontFamily": "Arial, sans-serif", "backgroundColor": "#f8f9fa", "padding": "20px"}, children=[
    html.H2("SCP Savings Dashboard", style={"textAlign": "center", "color": "#003366"}),

    # Insights
    html.Div([
        html.P(f"💡 Total Finance Savings: ${total_finance_savings:,}",
               style={"color": "#003366", "fontSize": "18px", "fontWeight": "bold"}),
        html.P(f"✅ Positive Finance Savings: ${positive_finance:,}",
               style={"color": "green", "fontSize": "18px", "fontWeight": "bold"}),
        html.P(f"⚠️ Negative Finance Savings: ${negative_finance:,}",
               style={"color": "red", "fontSize": "18px", "fontWeight": "bold"}),
    ], style={"backgroundColor": "#e6f0ff", "padding": "20px", "borderRadius": "10px",
              "boxShadow": "0 4px 8px rgba(0,0,0,0.1)", "marginBottom": "20px"}),

    # Data Table
    dash_table.DataTable(
        id="savings-table",
        columns=[{"name": col, "id": col} for col in df_display.columns],
        data=df_display.to_dict("records"),
        style_cell={"textAlign": "center", "padding": "8px"},
        style_header={"backgroundColor": "#003366", "color": "white", "fontWeight": "bold"},
        style_data_conditional=[
            # Finance column red if negative
            {"if": {"filter_query": "{Savings_Finance} contains \"-\""}, "color": "red", "fontWeight": "bold"},
            # SCP column red if negative
            {"if": {"filter_query": "{Savings_SCP} contains \"-\""}, "color": "red", "fontWeight": "bold"},
            # Default Finance column green
            {"if": {"column_id": "Savings_Finance"}, "color": "green"},
            # Default SCP column green
            {"if": {"column_id": "Savings_SCP"}, "color": "green"},
        ],
        page_size=10,
        style_table={"overflowX": "auto"}
    ),

    html.Br(),
    dcc.Graph(figure=fig_finance),
    html.Br(),
    dcc.Graph(figure=fig_scp),
])

if __name__ == "__main__":
    app.run(debug=True)
