"""
Nashik iCenter Orderbook & Daily Material Incoming Dashboard
============================================================
Upload your daily Orderbook (.xlsm) and Daily Material Incoming (.xlsx) files
to get an instant project-wise summary, component shortage analysis, and
GRN reconciliation view.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="iCenter Orderbook Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px; border-radius: 12px; color: white;
        text-align: center; margin: 5px 0;
    }
    .metric-card h2 { margin: 0; font-size: 2rem; }
    .metric-card p  { margin: 5px 0 0 0; font-size: 0.9rem; opacity: 0.9; }
    .shortage  { background: linear-gradient(135deg, #f5576c 0%, #ff6b6b 100%); }
    .available { background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); color: #1a1a2e; }
    .incoming  { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color: #1a1a2e; }
    .project   { background: linear-gradient(135deg, #fa709a 0%, #fee140 100%); color: #1a1a2e; }
    div[data-testid="stMetric"] { background: #f8f9fa; padding: 12px; border-radius: 8px; border: 1px solid #e9ecef; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_data
def load_orderbook(file):
    """Load OpenOrdersBOM sheet from the Orderbook .xlsm file."""
    xl = pd.ExcelFile(file, engine="openpyxl")
    sheets = xl.sheet_names

    if "OpenOrdersBOM" not in sheets:
        st.error("❌ Sheet 'OpenOrdersBOM' not found! Available sheets: " + ", ".join(sheets))
        return None, sheets

    df = pd.read_excel(xl, sheet_name="OpenOrdersBOM")
    # Clean key columns
    df["Component Code"] = df["Component Code"].astype(str).str.strip()
    df["Required Quantity"] = pd.to_numeric(df["Required Quantity"], errors="coerce").fillna(0)
    df["Quantity Issued"] = pd.to_numeric(df["Quantity Issued"], errors="coerce").fillna(0)
    df["Open Qty"] = df["Required Quantity"] - df["Quantity Issued"]
    df["Project Num"] = df["Project Num"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["Project Name"] = df["Project Name"].astype(str).str.strip()

    # Ensure numeric columns
    for col in ["On Hand Quantity", "Incoming PO Qty", "Total Available", "Variance",
                 "Item Cost", "Total Demand", "Net Extended Available Qty",
                 "Ordered Quantity", "Manufacturing Lead Time"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df, sheets


@st.cache_data
def load_grn(file):
    """Load Daily GRN sheet from the Daily Material Incoming .xlsx file."""
    xl = pd.ExcelFile(file, engine="openpyxl")
    sheets = xl.sheet_names

    if "Daily GRN" not in sheets:
        st.error("❌ Sheet 'Daily GRN' not found! Available sheets: " + ", ".join(sheets))
        return None, sheets

    df = pd.read_excel(xl, sheet_name="Daily GRN")
    # Clean key columns
    df["Item"] = df["Item"].astype(str).str.strip()
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Status column (Unnamed: 1)
    status_col = "Unnamed: 1" if "Unnamed: 1" in df.columns else None
    if status_col:
        df["GRN_Status"] = df[status_col].astype(str).str.strip()
    else:
        df["GRN_Status"] = "Deliver"

    # Remove header-like rows
    df = df[df["Item"] != "Item"].copy()
    df = df[df["Item"] != "nan"].copy()
    df = df.dropna(subset=["Date"])

    return df, sheets


def build_component_reconciliation(df_oob, df_grn_today):
    """
    Build a reconciliation table:
    - Component Code from OpenOrdersBOM with Required, Issued, Open Qty
    - Matched against today's GRN received qty
    """
    # Aggregate OOB by Component Code
    oob_agg = df_oob.groupby("Component Code", as_index=False).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "On Hand Quantity": "first",
        "Total Available": "first",
        "Component Desc": "first",
    })

    # Aggregate GRN today by Item (only "Deliver" status)
    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_agg = grn_delivered.groupby("Item", as_index=False).agg({
        "Qty": "sum",
        "Supplier": "first",
        "Order Number": "first",
    }).rename(columns={"Item": "Component Code", "Qty": "Received Today"})

    # Merge
    merged = oob_agg.merge(grn_agg, on="Component Code", how="left")
    merged["Received Today"] = merged["Received Today"].fillna(0)
    merged["Still Pending"] = merged["Open Qty"] - merged["Received Today"]
    merged["Status"] = np.where(
        merged["Still Pending"] <= 0, "✅ Fulfilled",
        np.where(merged["Received Today"] > 0, "🔶 Partially Received", "❌ Not Received")
    )

    return merged


def build_project_summary(df_oob, df_grn_today):
    """
    Build a project-wise summary:
    - Total unique components required per project
    - Total required qty, issued qty, open qty per project
    - Components received today per project
    """
    # Aggregate by project
    proj_agg = df_oob.groupby(["Project Num", "Project Name"], as_index=False).agg({
        "Component Code": "nunique",
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "Order Number": "nunique",
        "ITEM": "nunique",
        "Availability": lambda x: (x == "Available").sum(),
    }).rename(columns={
        "Component Code": "Unique Components",
        "Order Number": "Unique Orders",
        "ITEM": "Unique Items (FG)",
        "Availability": "Components Available",
    })

    # Get components per project
    proj_components = df_oob.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
    proj_components.columns = ["Project Num", "Project Name", "Component Set"]

    # GRN delivered today items
    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_items_today = set(grn_delivered["Item"].unique())
    grn_qty_by_item = grn_delivered.groupby("Item")["Qty"].sum().to_dict()

    # Match
    proj_components["Components Received Today"] = proj_components["Component Set"].apply(
        lambda s: len(s & grn_items_today)
    )
    proj_components["Qty Received Today"] = proj_components["Component Set"].apply(
        lambda s: sum(grn_qty_by_item.get(item, 0) for item in s & grn_items_today)
    )

    proj_summary = proj_agg.merge(
        proj_components[["Project Num", "Project Name", "Components Received Today", "Qty Received Today"]],
        on=["Project Num", "Project Name"], how="left"
    )
    proj_summary["Fulfillment %"] = np.where(
        proj_summary["Required Quantity"] > 0,
        ((proj_summary["Quantity Issued"] / proj_summary["Required Quantity"]) * 100).round(1),
        100
    )
    proj_summary = proj_summary.sort_values("Open Qty", ascending=False)

    return proj_summary


# ═══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR — FILE UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════

st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/ef/Emerson_Electric_Company.svg/320px-Emerson_Electric_Company.svg.png", width=200)
st.sidebar.markdown("---")
st.sidebar.header("📂 Upload Daily Files")

uploaded_ob = st.sidebar.file_uploader(
    "Orderbook (.xlsm)",
    type=["xlsm", "xlsx"],
    help="Upload the Nashik iCenter Orderbook file"
)
uploaded_grn = st.sidebar.file_uploader(
    "Daily Material Incoming (.xlsx)",
    type=["xlsx", "xlsm"],
    help="Upload the Daily Material Incoming file"
)

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN CONTENT
# ═══════════════════════════════════════════════════════════════════════════════

st.title("📦 Nashik iCenter — Orderbook & Material Dashboard")
st.caption(f"Report Date: {datetime.now().strftime('%d %B %Y, %A')}")

if not uploaded_ob or not uploaded_grn:
    st.info("👈 Upload both files from the sidebar to begin.")
    st.markdown("""
    ### How to use this dashboard:
    1. **Upload** your daily Orderbook `.xlsm` file
    2. **Upload** your Daily Material Incoming `.xlsx` file
    3. The dashboard will automatically:
       - Match **Component Code** (Orderbook) with **Item** (GRN)
       - Calculate **Required − Issued = Open Qty** and compare against **GRN Received Qty**
       - Show a **project-wise summary** of material requirements and what was received today
       - Highlight shortages and fulfilled items
    """)
    st.stop()

# ── Load data ────────────────────────────────────────────────────────────────
with st.spinner("Loading Orderbook..."):
    df_oob, ob_sheets = load_orderbook(uploaded_ob)
with st.spinner("Loading Daily GRN..."):
    df_grn, grn_sheets = load_grn(uploaded_grn)

if df_oob is None or df_grn is None:
    st.stop()

# ── Date filter for GRN ─────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.header("📅 Filter by Date")

available_dates = sorted(df_grn["Date"].dt.date.dropna().unique(), reverse=True)
if len(available_dates) > 0:
    selected_date = st.sidebar.date_input(
        "Select GRN Date",
        value=available_dates[0],
        min_value=min(available_dates),
        max_value=max(available_dates),
        help="Filter Daily GRN to this date"
    )
else:
    selected_date = date.today()

df_grn_today = df_grn[df_grn["Date"].dt.date == selected_date].copy()

# ── Project filter ───────────────────────────────────────────────────────────
st.sidebar.markdown("---")
st.sidebar.header("🏗️ Filter by Project")

all_projects = sorted(df_oob["Project Name"].dropna().unique())
all_projects = [p for p in all_projects if p != "nan"]
selected_projects = st.sidebar.multiselect(
    "Select Projects (leave empty for all)",
    options=all_projects,
    default=[],
    help="Filter to specific projects"
)

if selected_projects:
    df_oob_filtered = df_oob[df_oob["Project Name"].isin(selected_projects)].copy()
else:
    df_oob_filtered = df_oob.copy()

# ═══════════════════════════════════════════════════════════════════════════════
#  KPI CARDS
# ═══════════════════════════════════════════════════════════════════════════════

grn_delivered_today = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
grn_rejected_today = df_grn_today[df_grn_today["GRN_Status"].str.contains("Reject", case=False, na=False)]

total_open_qty = df_oob_filtered["Open Qty"].sum()
total_received_today = grn_delivered_today["Qty"].sum()
total_unique_components_needed = df_oob_filtered["Component Code"].nunique()
components_received_today = len(set(grn_delivered_today["Item"]) & set(df_oob_filtered["Component Code"]))
total_projects = df_oob_filtered[["Project Num", "Project Name"]].drop_duplicates().shape[0]
total_rejected_today = grn_rejected_today["Qty"].sum()

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    st.markdown(f'<div class="metric-card shortage"><h2>{total_open_qty:,.0f}</h2><p>Total Open Qty</p></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card incoming"><h2>{total_received_today:,.0f}</h2><p>Received Today</p></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card available"><h2>{components_received_today}</h2><p>Components Received (Matching)</p></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="metric-card project"><h2>{total_projects}</h2><p>Active Projects</p></div>', unsafe_allow_html=True)
with c5:
    st.markdown(f'<div class="metric-card"><h2>{total_unique_components_needed}</h2><p>Unique Components Needed</p></div>', unsafe_allow_html=True)
with c6:
    st.markdown(f'<div class="metric-card shortage"><h2>{total_rejected_today:,.0f}</h2><p>Rejected Today</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB LAYOUT
# ═══════════════════════════════════════════════════════════════════════════════

tab1, tab2, tab3, tab4 = st.tabs([
    "🏗️ Project-wise Summary",
    "🔍 Component Reconciliation",
    "📥 Today's GRN Detail",
    "📊 Analytics"
])

# ── TAB 1: Project-wise Summary ─────────────────────────────────────────────
with tab1:
    st.subheader("Project-wise Material Requirement & Receipt Summary")
    st.caption(f"GRN Date: {selected_date.strftime('%d-%b-%Y')}")

    proj_summary = build_project_summary(df_oob_filtered, df_grn_today)

    # Style the dataframe
    display_cols = [
        "Project Num", "Project Name", "Unique Orders", "Unique Items (FG)",
        "Unique Components", "Components Available",
        "Required Quantity", "Quantity Issued", "Open Qty", "Fulfillment %",
        "Components Received Today", "Qty Received Today"
    ]
    proj_display = proj_summary[display_cols].copy()

    # Color code
    def color_fulfillment(val):
        if val >= 90:
            return "background-color: #d4edda"
        elif val >= 50:
            return "background-color: #fff3cd"
        else:
            return "background-color: #f8d7da"

    style_method = getattr(proj_display.style, "map", None) or proj_display.style.applymap
    styled = style_method(
        color_fulfillment, subset=["Fulfillment %"]
    ).format({
        "Required Quantity": "{:,.0f}",
        "Quantity Issued": "{:,.0f}",
        "Open Qty": "{:,.0f}",
        "Fulfillment %": "{:.1f}%",
        "Qty Received Today": "{:,.0f}",
    })

    st.dataframe(styled, use_container_width=True, height=500)

    # Download button
    csv_proj = proj_display.to_csv(index=False)
    st.download_button("📥 Download Project Summary", csv_proj, "project_summary.csv", "text/csv")

    # ── Expandable: Project Detail ──
    st.markdown("---")
    st.subheader("🔎 Drill-down: Project Component Detail")
    selected_proj_detail = st.selectbox("Select a project to view details:", all_projects)

    if selected_proj_detail:
        proj_detail = df_oob_filtered[df_oob_filtered["Project Name"] == selected_proj_detail]

        # Get GRN matches for this project
        proj_components_set = set(proj_detail["Component Code"].unique())
        grn_for_proj = grn_delivered_today[grn_delivered_today["Item"].isin(proj_components_set)]

        grn_qty_map = grn_for_proj.groupby("Item")["Qty"].sum().to_dict()

        proj_comp_detail = proj_detail.groupby(["Component Code", "Component Desc"], as_index=False).agg({
            "Required Quantity": "sum",
            "Quantity Issued": "sum",
            "Open Qty": "sum",
            "On Hand Quantity": "first",
            "Incoming PO Qty": "sum",
            "Total Available": "first",
            "Availability": "first",
            "Supplier": "first",
        })
        proj_comp_detail["Received Today"] = proj_comp_detail["Component Code"].map(grn_qty_map).fillna(0)
        proj_comp_detail["Net Pending"] = proj_comp_detail["Open Qty"] - proj_comp_detail["Received Today"]
        proj_comp_detail["Status"] = np.where(
            proj_comp_detail["Net Pending"] <= 0, "✅ Fulfilled",
            np.where(proj_comp_detail["Received Today"] > 0, "🔶 Partial", "❌ Pending")
        )
        proj_comp_detail = proj_comp_detail.sort_values("Net Pending", ascending=False)

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Total Components", len(proj_comp_detail))
        col_b.metric("Received Today", int(proj_comp_detail["Received Today"].gt(0).sum()))
        col_c.metric("Still Pending", int(proj_comp_detail["Net Pending"].gt(0).sum()))

        st.dataframe(proj_comp_detail, use_container_width=True, height=400)


# ── TAB 2: Component Reconciliation ─────────────────────────────────────────
with tab2:
    st.subheader("Component-Level Reconciliation: Open Qty vs Today's GRN")
    st.caption("Required Quantity − Quantity Issued = Open Qty → compared with GRN Received Today")

    reconciliation = build_component_reconciliation(df_oob_filtered, df_grn_today)

    # Filter options
    status_filter = st.multiselect(
        "Filter by Status:",
        options=["✅ Fulfilled", "🔶 Partially Received", "❌ Not Received"],
        default=["🔶 Partially Received", "❌ Not Received"]
    )
    if status_filter:
        recon_display = reconciliation[reconciliation["Status"].isin(status_filter)]
    else:
        recon_display = reconciliation

    st.dataframe(
        recon_display[[
            "Component Code", "Component Desc", "Required Quantity", "Quantity Issued",
            "Open Qty", "Received Today", "Still Pending", "On Hand Quantity",
            "Total Available", "Supplier", "Status"
        ]].sort_values("Still Pending", ascending=False),
        use_container_width=True,
        height=500
    )

    csv_recon = recon_display.to_csv(index=False)
    st.download_button("📥 Download Reconciliation", csv_recon, "component_reconciliation.csv", "text/csv")

    # Summary stats
    st.markdown("---")
    rc1, rc2, rc3 = st.columns(3)
    fulfilled_count = (reconciliation["Status"] == "✅ Fulfilled").sum()
    partial_count = (reconciliation["Status"] == "🔶 Partially Received").sum()
    pending_count = (reconciliation["Status"] == "❌ Not Received").sum()
    rc1.metric("✅ Fulfilled", fulfilled_count)
    rc2.metric("🔶 Partially Received", partial_count)
    rc3.metric("❌ Not Received", pending_count)


# ── TAB 3: Today's GRN Detail ───────────────────────────────────────────────
with tab3:
    st.subheader(f"Today's GRN Receipts — {selected_date.strftime('%d-%b-%Y')}")

    if len(df_grn_today) == 0:
        st.warning(f"No GRN entries found for {selected_date.strftime('%d-%b-%Y')}. Try selecting a different date.")
    else:
        # Summary by status
        grn_status_summary = df_grn_today.groupby("GRN_Status", as_index=False).agg(
            Items=("Item", "nunique"),
            Total_Qty=("Qty", "sum"),
            Receipts=("Item", "count"),
        )
        st.dataframe(grn_status_summary, use_container_width=True)

        # Supplier-wise summary
        st.markdown("#### Supplier-wise Receipts")
        grn_supplier = grn_delivered_today.groupby("Supplier", as_index=False).agg(
            Items_Delivered=("Item", "nunique"),
            Total_Qty=("Qty", "sum"),
        ).sort_values("Total_Qty", ascending=False)
        st.dataframe(grn_supplier, use_container_width=True, height=300)

        # Full detail
        st.markdown("#### Full GRN Detail")
        grn_detail_cols = ["Item", "Item Description", "Qty", "Unit", "Date", "Order Number", "Supplier", "GRN_Status"]
        available_cols = [c for c in grn_detail_cols if c in df_grn_today.columns]
        st.dataframe(df_grn_today[available_cols], use_container_width=True, height=400)


# ── TAB 4: Analytics ─────────────────────────────────────────────────────────
with tab4:
    st.subheader("📊 Visual Analytics")

    col1, col2 = st.columns(2)

    with col1:
        # Project-wise Open Qty bar chart
        proj_summary_chart = build_project_summary(df_oob_filtered, df_grn_today)
        proj_top = proj_summary_chart.nlargest(15, "Open Qty")
        fig1 = px.bar(
            proj_top, x="Open Qty", y="Project Name", orientation="h",
            title="Top 15 Projects by Open Qty",
            color="Fulfillment %", color_continuous_scale="RdYlGn",
            labels={"Open Qty": "Open Quantity", "Project Name": "Project"},
        )
        fig1.update_layout(yaxis=dict(autorange="reversed"), height=500)
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        # Fulfillment % by project
        fig2 = px.bar(
            proj_top, x="Fulfillment %", y="Project Name", orientation="h",
            title="Fulfillment % by Project (Top 15 by Open Qty)",
            color="Fulfillment %", color_continuous_scale="RdYlGn",
            labels={"Fulfillment %": "Fulfillment %", "Project Name": "Project"},
        )
        fig2.update_layout(yaxis=dict(autorange="reversed"), height=500)
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    col3, col4 = st.columns(2)

    with col3:
        # Component status pie
        recon_full = build_component_reconciliation(df_oob_filtered, df_grn_today)
        status_counts = recon_full["Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        fig3 = px.pie(
            status_counts, values="Count", names="Status",
            title="Component Reconciliation Status",
            color_discrete_map={
                "✅ Fulfilled": "#2ecc71",
                "🔶 Partially Received": "#f39c12",
                "❌ Not Received": "#e74c3c",
            }
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Today's GRN by supplier (top 10)
        if len(grn_delivered_today) > 0:
            top_suppliers = grn_delivered_today.groupby("Supplier")["Qty"].sum().nlargest(10).reset_index()
            fig4 = px.bar(
                top_suppliers, x="Qty", y="Supplier", orientation="h",
                title=f"Top 10 Suppliers by Qty Delivered ({selected_date.strftime('%d-%b')})",
                color="Qty", color_continuous_scale="Blues",
            )
            fig4.update_layout(yaxis=dict(autorange="reversed"), height=400)
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("No deliveries for the selected date.")

    # Availability breakdown
    st.markdown("---")
    st.subheader("Component Availability Breakdown")
    avail_counts = df_oob_filtered["Availability"].value_counts().reset_index()
    avail_counts.columns = ["Availability", "Count"]
    fig5 = px.pie(avail_counts, values="Count", names="Availability",
                  title="Availability Status of All Components in Orderbook",
                  color_discrete_map={"Available": "#2ecc71", "Shortage": "#e74c3c"})
    st.plotly_chart(fig5, use_container_width=True)


# ── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Built for Nashik iCenter Supply Chain | Data refreshes on file upload")
