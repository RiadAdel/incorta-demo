import streamlit as st

from utils import LOGO_SVG, build_pdf, build_pptx

st.set_page_config(
    layout="wide",
    page_title="Incorta Sales Dashboard",
    page_icon="incorta-logo.svg",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer {visibility: hidden;}
.stAppDeployButton,
[data-testid="stAppDeployButton"],
[data-testid="stDeployButton"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebarCollapsedControl"] { visibility: visible !important; z-index: 999999 !important; }
.stApp { background: linear-gradient(150deg, #0d0a2e 0%, #1a1350 60%, #241c55 100%); }
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; }

/* ── Header box ── */
[data-testid="stHorizontalBlock"]:first-of-type {
    background: linear-gradient(135deg, #241c55 0%, #3d2f8a 55%, #5b4fcf 100%);
    border-radius: 20px; padding: 16px 28px !important;
    border: 1px solid rgba(139,92,246,0.35);
    box-shadow: 0 8px 40px rgba(36,28,85,0.7), inset 0 1px 0 rgba(255,255,255,0.12);
    align-items: center !important; margin-bottom: 28px; position: relative; overflow: hidden;
}
[data-testid="stHorizontalBlock"]:first-of-type::before {
    content: ''; position: absolute; top: -60px; right: -40px; width: 320px; height: 320px;
    background: radial-gradient(circle, rgba(245,158,11,0.18) 0%, transparent 65%); pointer-events: none;
}
[data-testid="stHorizontalBlock"]:first-of-type::after {
    content: ''; position: absolute; bottom: -60px; left: 35%; width: 250px; height: 250px;
    background: radial-gradient(circle, rgba(192,132,252,0.18) 0%, transparent 65%); pointer-events: none;
}

/* ── Header buttons ── */
[data-testid="stHorizontalBlock"]:first-of-type button {
    background: linear-gradient(135deg, rgba(245,158,11,0.22), rgba(124,58,237,0.32)) !important;
    border: 1px solid rgba(245,158,11,0.65) !important; color: white !important;
    border-radius: 12px !important; font-weight: 700 !important; font-size: 0.9rem !important;
    transition: all 0.2s !important;
}
[data-testid="stHorizontalBlock"]:first-of-type button:hover {
    background: linear-gradient(135deg, rgba(245,158,11,0.42), rgba(124,58,237,0.52)) !important;
    border-color: rgba(245,158,11,0.95) !important; box-shadow: 0 0 16px rgba(245,158,11,0.3) !important;
}

/* ── Popover panel ── */
[data-testid="stPopoverBody"] {
    background: linear-gradient(135deg, #1a1350, #241c55) !important;
    border: 1px solid rgba(139,92,246,0.45) !important; border-radius: 14px !important; padding: 8px !important;
}
[data-testid="stPopoverBody"] button {
    background: linear-gradient(135deg, rgba(36,28,85,0.9), rgba(91,79,207,0.4)) !important;
    border: 1px solid rgba(139,92,246,0.5) !important; color: white !important;
    border-radius: 10px !important; font-weight: 600 !important;
}
[data-testid="stPopoverBody"] button:hover {
    border-color: rgba(245,158,11,0.75) !important;
    background: linear-gradient(135deg, rgba(36,28,85,1), rgba(124,58,237,0.55)) !important;
}

/* ── Header text ── */
.dash-title {
    font-size: 2rem; font-weight: 800; margin: 0;
    background: linear-gradient(90deg, #ffffff 0%, #c4b5fd 60%, #f59e0b 100%);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
}
.dash-subtitle { color: rgba(255,255,255,0.55); font-size: 0.88rem; margin: 6px 0 0 0; letter-spacing: 0.4px; }

/* ── KPI cards ── */
[data-testid="metric-container"] {
    background: linear-gradient(135deg, rgba(36,28,85,0.85) 0%, rgba(91,79,207,0.25) 100%);
    border: 1px solid rgba(139,92,246,0.35); border-radius: 18px; padding: 22px 24px !important;
    box-shadow: 0 4px 28px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.1);
    backdrop-filter: blur(12px); transition: border-color 0.25s, box-shadow 0.25s;
}
[data-testid="metric-container"]:hover {
    border-color: rgba(245,158,11,0.55);
    box-shadow: 0 6px 36px rgba(245,158,11,0.18), inset 0 1px 0 rgba(255,255,255,0.12);
}
[data-testid="stMetricLabel"] p {
    color: rgba(255,255,255,0.65) !important; font-size: 0.8rem !important;
    text-transform: uppercase; letter-spacing: 0.8px; font-weight: 600 !important;
}
[data-testid="stMetricValue"] { color: #ffffff !important; font-size: 1.9rem !important; font-weight: 800 !important; }

/* ── Headings ── */
h2, h3 {
    color: #c4b5fd !important;
    -webkit-text-fill-color: #c4b5fd !important;
    background: none !important;
    font-weight: 700 !important;
}

/* ── Chart + Table cards ── */
[data-testid="stPlotlyChart"] > div,
[data-testid="stDataFrame"] {
    background: rgba(20,15,60,0.55); border: 1px solid rgba(139,92,246,0.2);
    border-radius: 18px; padding: 8px; box-shadow: 0 4px 24px rgba(0,0,0,0.25);
}

/* ── Customer HTML table ── */
.customer-table-wrap {
    max-height: 640px; overflow: auto;
    border-radius: 18px; border: 1px solid rgba(139,92,246,0.2);
    background: rgba(20,15,60,0.55); box-shadow: 0 4px 24px rgba(0,0,0,0.25);
    padding: 0;
}
.customer-table {
    width: 100%; border-collapse: separate; border-spacing: 0;
    color: #e8e6f0; font-size: 0.88rem;
}
.customer-table thead th {
    position: sticky; top: 0; z-index: 2;
    background: #241c55; color: #c4b5fd;
    font-weight: 700; text-transform: uppercase; letter-spacing: 0.6px;
    font-size: 0.74rem;
    padding: 14px 14px;
    border-bottom: 1px solid rgba(139,92,246,0.35);
    text-align: left;
}
.customer-table thead th:first-child { text-align: center; width: 48px; }
.customer-table tbody th {
    background: rgba(36,28,85,0.4); color: rgba(196,181,253,0.7);
    font-weight: 600; text-align: center; width: 48px;
    padding: 10px 8px; border-bottom: 1px solid rgba(139,92,246,0.1);
}
.customer-table tbody td {
    padding: 10px 14px;
    border-bottom: 1px solid rgba(139,92,246,0.1);
}
.customer-table tbody tr:nth-child(odd) td,
.customer-table tbody tr:nth-child(odd) th { background-color: rgba(36,28,85,0.18); }
.customer-table tbody tr:hover td,
.customer-table tbody tr:hover th { background-color: rgba(124,58,237,0.18); }
.customer-table td.right, .customer-table th.right { text-align: right; }
.customer-table td.amber { color: #f59e0b; font-weight: 600; }

/* ── Divider ── */
hr { border: none !important; height: 1px !important; margin: 28px 0 !important;
     background: linear-gradient(90deg, transparent, rgba(139,92,246,0.5), transparent) !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a1350 0%, #0d0a2e 100%) !important;
    border-right: 1px solid rgba(139,92,246,0.2) !important;
}
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p {
    color: rgba(255,255,255,0.75) !important;
}


</style>
""", unsafe_allow_html=True)

# ── Data ─────────────────────────────────────────────────────────────────────
conn = st.connection("postgresql", type="sql")
df = conn.query(
    """
    SELECT
        "AccountNumber", "CustomerID", "Sales_Order_Count",
        "Average_Sales_Order", "Total_Sales", "Rank",
        "StateProvinceName", "City"
    FROM "default"."Online_Store"."Customers"
    """,
    ttl="10m",
)

# ── Sidebar: global filters ──────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Filters")
    all_states = sorted(df["StateProvinceName"].dropna().unique().tolist())
    selected_states = st.multiselect(
        "State / Province",
        options=all_states,
        placeholder="All states",
    )
    st.caption("Incorta Sales Dashboard · v1.0")

filtered_df = df[df["StateProvinceName"].isin(selected_states)] if selected_states else df

kpi_customers  = filtered_df["CustomerID"].nunique()
kpi_total      = filtered_df["Total_Sales"].sum()
kpi_avg_sales  = filtered_df["Total_Sales"].mean()
kpi_avg_orders = filtered_df["Sales_Order_Count"].mean()

# Expose filtered data to pages
st.session_state.filtered_df = filtered_df

# ── Header ───────────────────────────────────────────────────────────────────
logo_col, btn_col = st.columns([5, 1], vertical_alignment="center")

with logo_col:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:24px;padding:8px 0;">
        <div style="flex-shrink:0;">{LOGO_SVG}</div>
        <div>
            <p class="dash-title">Customer sales dashboard</p>
            <p class="dash-subtitle">Online store · Performance analytics</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

with btn_col:
    with st.popover("Download", icon=":material/download:", use_container_width=True):
        st.download_button(
            "PDF report",
            icon=":material/picture_as_pdf:",
            data=build_pdf(filtered_df, kpi_customers, kpi_total, kpi_avg_sales, kpi_avg_orders),
            file_name="incorta_dashboard.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.download_button(
            "PowerPoint",
            icon=":material/slideshow:",
            data=build_pptx(filtered_df, kpi_customers, kpi_total, kpi_avg_sales, kpi_avg_orders),
            file_name="incorta_dashboard.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

# ── KPIs ─────────────────────────────────────────────────────────────────────
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total customers",       f"{kpi_customers:,}")
col2.metric("Total sales",           f"${kpi_total:,.0f}")
col3.metric("Avg sales / customer",  f"${kpi_avg_sales:,.0f}")
col4.metric("Avg orders / customer", f"{kpi_avg_orders:,.1f}")

# ── Navigation ───────────────────────────────────────────────────────────────
page = st.navigation([
    st.Page("app_pages/geography.py",   title="Geography",   icon=":material/public:"),
    st.Page("app_pages/performance.py", title="Performance", icon=":material/analytics:"),
    st.Page("app_pages/customers.py",   title="Customers",   icon=":material/groups:"),
], position="sidebar")

page.run()
