"""
Nashik iCenter Orderbook & Daily Material Incoming Dashboard
============================================================
Upload your daily Orderbook (.xlsm), Daily Material Incoming (.xlsx),
and Stock (.xlsx) files to get an instant project-wise summary,
component shortage analysis, and GRN reconciliation view.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
from pathlib import Path

# -- Page config ---------------------------------------------------------------
st.set_page_config(
    page_title="iCenter Orderbook Dashboard",
    page_icon="\U0001F4E6",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -- Custom CSS ----------------------------------------------------------------
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


# ==============================================================================
#  DEFAULT PROJECT PRIORITY SEQUENCE
# ==============================================================================

DEFAULT_PROJECT_SEQUENCE = pd.DataFrame({
    "Sr.": list(range(1, 46)),
    "Project Name": [
        "RLPP", "EPC4", "RLPP Ethylene Plant", "QG NFS",
        "RLPP PolyEthylene plant", "Meram RTU", "Meram F&G",
        "QatarGas North Field South", "KWIDF KOC", "QatarGas NFS 500JB",
        "DeltaVUpgradeAmin_Change", "Delta V_DCS_Upgrade_EGL", "Ceyhan",
        "NFS Balance of Plant", "Takara Bio CGCP3", "Marsa LNG.",
        "ENI Coral North LNG", "QatarGas NFS Safe area",
        "COMMON COOLING WATER SYSTEM.", "QOTL Booster Pump Station",
        "Aksa Senegal", "QURAYYAH", "LukOil PWT Ph-2", "AL-ALLISA Train 4",
        "BP PWRI2", "Domestic Gas Pipelines Kormor", "ESTIDAMA PKG6 F&G",
        "Cedar LNG Prototype", "EPC  4 New", "Maaden VMS in smelter",
        "Kyoto Fusioneering Unity2", "Gbaran Ubie EPU Phase 2",
        "Sales Gas Network Enhancemen", "Kaminho FPSO", "BP DS05 TMS",
        "TAZIZ Logistic", "PDO Rabiha", "KOC JPF-3", "OQ_FAHUD",
        "Air Liquid Allisa train 03", "Air Liquid Allisa train 2",
        "Air Liquide-ALLISA Train 7", "Hajr", "Alrar Boosting Phase", "WEP",
    ],
    "No. Of Cabinet": [
        11, 2, 1, 19, 1, 7, 10, 3, 31, 49, 2, 1, 10, 2, 1, 61, 62, 117,
        2, 1, 16, 6, 8, 10, 7, 6, 3, 1, 4, 8, 26, 2, 3, 82, 1, 5, 1, 31,
        1, 10, 10, 13, 9, 5, 13,
    ],
    "Prod. Open Dt.": [
        "18-Jul-23", "8-Apr-24", "11-Jun-24", "3-Jul-24", "27-Aug-24",
        "25-Dec-24", "25-Dec-24", "12-Feb-25", "12-Mar-25", "12-Mar-25",
        "7-May-25", "10-Jun-25", "4-Aug-25", "5-Aug-25", "7-Aug-25",
        "21-Aug-25", "22-Aug-25", "4-Sep-25", "4-Sep-25", "6-Oct-25",
        "8-Oct-25", "9-Oct-25", "16-Oct-25", "28-Oct-25", "24-Nov-25",
        "3-Dec-25", "8-Dec-25", "10-Dec-25", "18-Dec-25", "30-Dec-25",
        "7-Jan-26", "7-Jan-26", "9-Jan-26", "16-Jan-26", "29-Jan-26",
        "5-Feb-26", "5-Feb-26", "5-Feb-26", "5-Feb-26", "6-Feb-26",
        "6-Feb-26", "6-Feb-26", "19-Feb-26", "19-Feb-26", "19-Feb-26",
    ],
    "SO No.": [
        2000434, 2000441, 2000572, 2000576, 2000563, 2000685, 2000686,
        2000682, 2000709, 2000710, 2000712, 2000720, 2000771, 2000787,
        2000747, 2000816, 2000490, 2000746, 2000843, 2000864, 2000827,
        2000875, 2000844, 2000879, 2000601, 2000895, 2000839, 2000735,
        2000893, 2000891, 2000838, 2000900, 2000779, 2000847, 2000926,
        2000906, 2000911, 2000907, 2000922, 2000908, 2000909, 2000931,
        2000878, 2000903, 2000939,
    ],
})

SEQUENCE_COLUMNS = ["Sr.", "Project Name", "No. Of Cabinet", "Prod. Open Dt.", "SO No.", "Completed"]
SEQUENCE_STATE_FILE = Path(__file__).with_name("project_sequence_state.json")


# ==============================================================================
#  HELPER FUNCTIONS
# ==============================================================================

@st.cache_data
def load_orderbook(file):
    """Load Orderbook data by detecting the sheet via required columns."""
    xl = pd.ExcelFile(file, engine="openpyxl")
    sheets = xl.sheet_names

    required_cols = {
        "Component Code", "Required Quantity", "Quantity Issued", "Project Num", "Project Name"
    }
    matched_sheet = None
    matched_cols = set()
    best_sheet = None
    best_cols = set()

    for s in sheets:
        try:
            preview = pd.read_excel(xl, sheet_name=s, nrows=0)
        except Exception:
            continue
        cols = set(preview.columns.astype(str).str.strip())
        overlap = len(required_cols & cols)
        if overlap > len(best_cols):
            best_sheet = s
            best_cols = cols
        if required_cols.issubset(cols):
            matched_sheet = s
            matched_cols = cols
            break

    if matched_sheet is None:
        available = ", ".join(str(s) for s in sheets)
        if best_sheet is not None:
            missing = sorted(required_cols - best_cols)
            st.error(
                f"\u274C No sheet matches Orderbook format. Closest sheet: '{best_sheet}'. "
                f"Missing columns: {', '.join(missing)}. Available sheets: {available}"
            )
        else:
            st.error(f"\u274C Could not read any sheet from Orderbook file. Available sheets: {available}")
        return None, sheets

    df = pd.read_excel(xl, sheet_name=matched_sheet)
    df.columns = df.columns.astype(str).str.strip()
    df["Component Code"] = df["Component Code"].astype(str).str.strip()
    df["Required Quantity"] = pd.to_numeric(df["Required Quantity"], errors="coerce").fillna(0)
    df["Quantity Issued"] = pd.to_numeric(df["Quantity Issued"], errors="coerce").fillna(0)
    df["Open Qty"] = df["Required Quantity"] - df["Quantity Issued"]
    df["Project Num"] = df["Project Num"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["Project Name"] = df["Project Name"].astype(str).str.strip()

    # Clean numeric ID columns so they don't appear as floats (e.g. 12345.0 -> 12345)
    if "Order Number" in df.columns:
        df["Order Number"] = df["Order Number"].astype(str).str.replace(r"\.0$", "", regex=True)
    if "Work Order Number" in df.columns:
        df["Work Order Number"] = df["Work Order Number"].astype(str).str.replace(r"\.0$", "", regex=True)

    for col in ["On Hand Quantity", "Incoming PO Qty", "Total Available", "Variance",
                 "Item Cost", "Total Demand", "Net Extended Available Qty",
                 "Ordered Quantity", "Manufacturing Lead Time"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df, sheets


@st.cache_data
def load_grn(file):
    """Load GRN data by detecting the sheet via required columns."""
    xl = pd.ExcelFile(file, engine="openpyxl")
    sheets = xl.sheet_names

    required_cols = {"Item", "Qty", "Date"}
    matched_sheet = None
    best_sheet = None
    best_cols = set()

    for s in sheets:
        try:
            preview = pd.read_excel(xl, sheet_name=s, nrows=0)
        except Exception:
            continue
        cols = set(preview.columns.astype(str).str.strip())
        overlap = len(required_cols & cols)
        if overlap > len(best_cols):
            best_sheet = s
            best_cols = cols
        if required_cols.issubset(cols):
            matched_sheet = s
            break

    if matched_sheet is None:
        available = ", ".join(str(s) for s in sheets)
        if best_sheet is not None:
            missing = sorted(required_cols - best_cols)
            st.error(
                f"\u274C No sheet matches Daily GRN format. Closest sheet: '{best_sheet}'. "
                f"Missing columns: {', '.join(missing)}. Available sheets: {available}"
            )
        else:
            st.error(f"\u274C Could not read any sheet from GRN file. Available sheets: {available}")
        return None, sheets

    df = pd.read_excel(xl, sheet_name=matched_sheet)
    df.columns = df.columns.astype(str).str.strip()
    df["Item"] = df["Item"].astype(str).str.strip()
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    if "Order Number" in df.columns:
        df["Order Number"] = df["Order Number"].astype(str).str.replace(r"\.0$", "", regex=True)
    if "Supplier" in df.columns:
        df["Supplier"] = df["Supplier"].astype(str).str.strip()

    status_col = "Unnamed: 1" if "Unnamed: 1" in df.columns else None
    if status_col:
        df["GRN_Status"] = df[status_col].astype(str).str.strip()
    else:
        df["GRN_Status"] = "Deliver"

    df = df[df["Item"] != "Item"].copy()
    df = df[df["Item"] != "nan"].copy()
    df = df.dropna(subset=["Date"])

    return df, sheets


@st.cache_data
def load_stock(file):
    """Load Stock data by detecting the sheet/header via required columns."""
    xl = pd.ExcelFile(file, engine="openpyxl")
    sheets = xl.sheet_names

    required_cols = {"Item Number", "On Hand Quantity"}
    matched_sheet = None
    matched_header_row = None
    best_sheet = None
    best_cols = set()

    for s in sheets:
        try:
            df_raw = pd.read_excel(xl, sheet_name=s, header=None)
        except Exception:
            continue

        max_rows = min(12, len(df_raw))
        for i in range(max_rows):
            row_vals = df_raw.iloc[i].astype(str).str.strip().tolist()
            if "Item Number" not in row_vals:
                continue

            candidate = pd.read_excel(xl, sheet_name=s, header=i)
            candidate.columns = candidate.columns.astype(str).str.strip()
            cols = set(candidate.columns)
            overlap = len(required_cols & cols)
            if overlap > len(best_cols):
                best_sheet = s
                best_cols = cols

            if required_cols.issubset(cols):
                matched_sheet = s
                matched_header_row = i
                break
        if matched_sheet is not None:
            break

    if matched_sheet is None:
        available = ", ".join(str(s) for s in sheets)
        if best_sheet is not None:
            missing = sorted(required_cols - best_cols)
            st.error(
                f"\u274C No sheet matches Stock format. Closest sheet: '{best_sheet}'. "
                f"Missing columns: {', '.join(missing)}. Available sheets: {available}"
            )
        else:
            st.error(
                f"\u274C No valid Stock header found. Expected columns: Item Number, On Hand Quantity. "
                f"Available sheets: {available}"
            )
        return None

    df = pd.read_excel(xl, sheet_name=matched_sheet, header=matched_header_row)
    df.columns = df.columns.astype(str).str.strip()

    df["Item Number"] = df["Item Number"].astype(str).str.strip()
    df["On Hand Quantity"] = pd.to_numeric(df["On Hand Quantity"], errors="coerce").fillna(0)

    df = df[df["Item Number"] != "Item Number"].copy()
    df = df[df["Item Number"] != "nan"].copy()
    df = df[df["Item Number"] != ""].copy()

    return df


def get_stock_map(df_stock):
    """Aggregate stock by Item Number -> total On Hand Quantity."""
    if df_stock is None or len(df_stock) == 0:
        return {}
    stock_agg = df_stock.groupby("Item Number", as_index=False)["On Hand Quantity"].sum()
    return dict(zip(stock_agg["Item Number"], stock_agg["On Hand Quantity"]))


def ensure_arrow_compatible(df):
    """Ensure all object columns are consistently typed for Arrow serialization."""
    if df is None or df.empty:
        return df
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == "object":
            # Fill NaN/None values with empty string first
            df[col] = df[col].fillna("").astype(str)
            # Strip any leading/trailing whitespace
            try:
                df[col] = df[col].str.strip()
            except (AttributeError, TypeError):
                pass
    return df


def display_dataframe_arrow_safe(df, **kwargs):
    """Display dataframe with automatic Arrow compatibility handling."""
    # If it's a Styler object (from df.style), display directly without arrow processing
    if hasattr(df, 'data'):  # pandas Styler has a .data attribute
        st.dataframe(df, **kwargs)
    else:
        df_clean = ensure_arrow_compatible(df)
        st.dataframe(df_clean, **kwargs)


def _to_bool_series(series):
    """Convert mixed values to bool with common truthy/falsey string handling."""
    truthy = {"true", "1", "yes", "y", "completed", "done"}
    falsey = {"false", "0", "no", "n", "pending", ""}

    def _conv(v):
        if isinstance(v, bool):
            return v
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        if s in truthy:
            return True
        if s in falsey:
            return False
        return False

    return series.apply(_conv)


def normalize_project_key(value):
    """Normalize project name for case/space-insensitive matching."""
    if pd.isna(value):
        return ""
    return "".join(str(value).strip().lower().split())


def build_project_priority_map(project_sequence):
    """Build priority map using normalized project names."""
    seq = project_sequence.copy()
    seq["_project_key"] = seq["Project Name"].apply(normalize_project_key)
    seq = seq[seq["_project_key"] != ""].copy()
    seq = seq.sort_values("Sr.", kind="stable")
    seq = seq.drop_duplicates(subset=["_project_key"], keep="first")
    return dict(zip(seq["_project_key"], seq["Sr."]))


def normalize_sequence_df(df_in):
    """Normalize uploaded project sequence data into dashboard sequence schema."""
    if df_in is None or len(df_in) == 0:
        return None, "Uploaded sheet is empty."

    df = df_in.copy()
    df.columns = df.columns.astype(str).str.strip()

    if "Project Name" not in df.columns:
        return None, "Required column 'Project Name' not found in uploaded sheet."

    for col in SEQUENCE_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[SEQUENCE_COLUMNS].copy()
    df["Project Name"] = df["Project Name"].astype(str).str.strip()
    df = df[~df["Project Name"].isin(["", "nan", "None"])].copy()
    df["_project_key"] = df["Project Name"].apply(normalize_project_key)
    df = df[df["_project_key"] != ""].copy()
    df = df.drop_duplicates(subset=["_project_key"], keep="last")

    if len(df) == 0:
        return None, "No valid 'Project Name' values found in uploaded sheet."

    df["Sr."] = pd.to_numeric(df["Sr."], errors="coerce")
    missing_sr = df["Sr."].isna()
    if missing_sr.any():
        start = 1
        df.loc[missing_sr, "Sr."] = range(start, start + int(missing_sr.sum()))
    df["Sr."] = df["Sr."].astype(int)

    df["No. Of Cabinet"] = pd.to_numeric(df["No. Of Cabinet"], errors="coerce").fillna(0).astype(int)
    df["Prod. Open Dt."] = df["Prod. Open Dt."].fillna("").astype(str)
    df["SO No."] = pd.to_numeric(df["SO No."], errors="coerce")
    df["Completed"] = _to_bool_series(df["Completed"])

    df = df.sort_values("Sr.", kind="stable").reset_index(drop=True)
    df["Sr."] = range(1, len(df) + 1)
    df = df[SEQUENCE_COLUMNS].copy()
    return df, None


def load_saved_project_sequence():
    """Load project sequence from disk if available, else default."""
    if not SEQUENCE_STATE_FILE.exists():
        return DEFAULT_PROJECT_SEQUENCE.copy()

    try:
        saved_df = pd.read_json(SEQUENCE_STATE_FILE)
    except Exception:
        return DEFAULT_PROJECT_SEQUENCE.copy()

    normalized_df, err = normalize_sequence_df(saved_df)
    if err or normalized_df is None:
        return DEFAULT_PROJECT_SEQUENCE.copy()
    return normalized_df


def set_project_sequence(df):
    """Set session sequence and persist it to disk."""
    normalized_df, err = normalize_sequence_df(df)
    if err or normalized_df is None:
        return False

    st.session_state["project_sequence"] = normalized_df[SEQUENCE_COLUMNS].copy()
    try:
        st.session_state["project_sequence"].to_json(SEQUENCE_STATE_FILE, orient="records", indent=2)
    except Exception as ex:
        st.warning(f"Project sequence updated in session, but autosave failed: {ex}")
    return True


def load_sequence_from_excel(uploaded_file):
    """Load first sheet/header in Excel or CSV that contains 'Project Name' column."""
    # Check if it's a CSV file
    file_name = uploaded_file.name if hasattr(uploaded_file, 'name') else str(uploaded_file)
    if file_name.lower().endswith('.csv'):
        try:
            # Try reading CSV with standard header
            df = pd.read_csv(uploaded_file)
            df.columns = df.columns.astype(str).str.strip()
            if "Project Name" in df.columns:
                return df, "CSV"
            
            # Try different header rows (up to 12)
            uploaded_file.seek(0)
            for skip_rows in range(12):
                uploaded_file.seek(0)
                try:
                    df = pd.read_csv(uploaded_file, skiprows=skip_rows)
                    df.columns = df.columns.astype(str).str.strip()
                    if "Project Name" in df.columns:
                        return df, "CSV"
                except Exception:
                    continue
        except Exception:
            pass
        return None, None
    
    # Handle Excel files
    xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheets = xl.sheet_names

    for sheet in sheets:
        try:
            raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        except Exception:
            continue

        max_rows = min(12, len(raw))
        for header_row in range(max_rows):
            row_vals = raw.iloc[header_row].astype(str).str.strip().tolist()
            if "Project Name" in row_vals:
                parsed = pd.read_excel(xl, sheet_name=sheet, header=header_row)
                parsed.columns = parsed.columns.astype(str).str.strip()
                return parsed, sheet

    # Fallback: try normal header on each sheet
    for sheet in sheets:
        try:
            parsed = pd.read_excel(xl, sheet_name=sheet)
            parsed.columns = parsed.columns.astype(str).str.strip()
            if "Project Name" in parsed.columns:
                return parsed, sheet
        except Exception:
            continue

    return None, None


def render_sequence_upload_controls(section_key):
    """Render upload UI for replacing or merging project sequence."""
    st.markdown("---")
    st.markdown("#### 📥 Upload Project Sequence")
    seq_file = st.file_uploader(
        "Upload sequence file (.xlsx/.xlsm/.csv)",
        type=["xlsx", "xlsm", "csv"],
        key=f"{section_key}_seq_file",
        help="File must contain a 'Project Name' column. Other columns are optional: Sr., No. Of Cabinet, Prod. Open Dt., SO No., Completed."
    )
    apply_mode = st.radio(
        "Apply mode",
        options=["Replace current sequence", "Append / Merge into current sequence"],
        horizontal=True,
        key=f"{section_key}_seq_mode",
    )

    if st.button("Apply Uploaded Sequence", key=f"{section_key}_seq_apply", type="primary"):
        if not seq_file:
            st.warning("Please upload a sequence file first.")
            return

        loaded_df, loaded_sheet = load_sequence_from_excel(seq_file)
        if loaded_df is None:
            st.error("Could not find 'Project Name' column in the uploaded file.")
            return

        normalized_df, err = normalize_sequence_df(loaded_df)
        if err:
            st.error(f"Invalid sequence file: {err}")
            return
        if normalized_df is None:
            st.error("Could not normalize uploaded sequence.")
            return

        current_df = st.session_state.get("project_sequence", DEFAULT_PROJECT_SEQUENCE.copy()).copy()
        if "Completed" not in current_df.columns:
            current_df["Completed"] = False

        if apply_mode == "Replace current sequence":
            updated_df = normalized_df.copy()
        else:
            current_df["Project Name"] = current_df["Project Name"].astype(str).str.strip()
            current_df["_project_key"] = current_df["Project Name"].apply(normalize_project_key)
            incoming_projects = set(normalized_df["Project Name"].apply(normalize_project_key))
            current_kept = current_df[~current_df["_project_key"].isin(incoming_projects)].copy()
            current_kept = current_kept.drop(columns=["_project_key"], errors="ignore")
            updated_df = pd.concat([current_kept, normalized_df], ignore_index=True)
            updated_df = updated_df.sort_values("Sr.", kind="stable").reset_index(drop=True)
            updated_df["Sr."] = range(1, len(updated_df) + 1)

        set_project_sequence(updated_df)
        source_name = loaded_sheet if loaded_sheet else "uploaded file"
        st.success(
            f"Sequence updated from {source_name}. Total projects: {len(updated_df):,}."
        )
        st.rerun()


def build_component_reconciliation(df_oob, df_grn_today, stock_map=None):
    """Build a reconciliation table with stock data."""
    oob_agg = df_oob.groupby("Component Code", as_index=False).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "On Hand Quantity": "first",
        "Total Available": "first",
        "Component Desc": "first",
    })

    if stock_map:
        oob_agg["Stock Qty"] = oob_agg["Component Code"].map(stock_map).fillna(0)
    else:
        oob_agg["Stock Qty"] = oob_agg["On Hand Quantity"]

    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_agg = grn_delivered.groupby("Item", as_index=False).agg({
        "Qty": "sum",
        "Supplier": "first",
        "Order Number": "first",
    }).rename(columns={"Item": "Component Code", "Qty": "Received Today"})

    merged = oob_agg.merge(grn_agg, on="Component Code", how="left")
    merged["Received Today"] = merged["Received Today"].fillna(0)
    merged["Still Pending"] = merged["Open Qty"] - merged["Received Today"]
    merged["Status"] = np.where(
        merged["Still Pending"] <= 0, "\u2705 Fulfilled",
        np.where(merged["Received Today"] > 0, "\U0001F536 Partially Received", "\u274C Not Received")
    )

    return ensure_arrow_compatible(merged)


def build_project_summary(df_oob, df_grn_today, stock_map=None):
    """Build a project-wise summary with stock data."""
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

    proj_components = df_oob.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
    proj_components.columns = ["Project Num", "Project Name", "Component Set"]

    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_items_today = set(grn_delivered["Item"].unique())
    grn_qty_by_item = grn_delivered.groupby("Item")["Qty"].sum().to_dict()

    proj_components["Components Received Today"] = proj_components["Component Set"].apply(
        lambda s: len(s & grn_items_today)
    )
    proj_components["Qty Received Today"] = proj_components["Component Set"].apply(
        lambda s: sum(grn_qty_by_item.get(item, 0) for item in s & grn_items_today)
    )

    if stock_map:
        proj_components["Stock Qty"] = proj_components["Component Set"].apply(
            lambda s: sum(stock_map.get(item, 0) for item in s)
        )
    else:
        proj_components["Stock Qty"] = 0

    proj_summary = proj_agg.merge(
        proj_components[["Project Num", "Project Name", "Components Received Today", "Qty Received Today", "Stock Qty"]],
        on=["Project Num", "Project Name"], how="left"
    )
    proj_summary["Fulfillment %"] = np.where(
        proj_summary["Required Quantity"] > 0,
        ((proj_summary["Quantity Issued"] / proj_summary["Required Quantity"]) * 100).round(1),
        100
    )
    proj_summary = proj_summary.sort_values("Open Qty", ascending=False)

    return ensure_arrow_compatible(proj_summary)


def build_grn_by_project_sequence(df_oob, df_grn_today, project_sequence, stock_map=None):
    """For each GRN item, find which projects need it, ordered by priority."""
    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    if len(grn_delivered) == 0:
        return pd.DataFrame()

    priority_map = build_project_priority_map(project_sequence)

    grn_items = set(grn_delivered["Item"].unique())
    grn_qty_by_item = grn_delivered.groupby("Item").agg({
        "Qty": "sum",
        "Supplier": "first",
        "Order Number": "first",
    }).to_dict("index")

    oob_matching = df_oob[df_oob["Component Code"].isin(grn_items)].copy()
    if len(oob_matching) == 0:
        return pd.DataFrame()

    proj_comp = oob_matching.groupby(
        ["Project Name", "Project Num", "Component Code", "Component Desc"], as_index=False
    ).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "Order Number": lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v) and str(v).strip() not in ("", "nan")))),
    })

    proj_comp["GRN Qty Today"] = proj_comp["Component Code"].map(
        lambda c: grn_qty_by_item.get(c, {}).get("Qty", 0)
    )
    proj_comp["Supplier"] = proj_comp["Component Code"].map(
        lambda c: str(grn_qty_by_item.get(c, {}).get("Supplier", "")).strip()
    )
    proj_comp["GRN Order Number"] = proj_comp["Component Code"].map(
        lambda c: grn_qty_by_item.get(c, {}).get("Order Number", "")
    )

    if stock_map:
        proj_comp["Stock Qty"] = proj_comp["Component Code"].map(stock_map).fillna(0)
    else:
        proj_comp["Stock Qty"] = 0

    proj_comp["_project_key"] = proj_comp["Project Name"].apply(normalize_project_key)
    proj_comp["Priority"] = proj_comp["_project_key"].map(priority_map).fillna(9999).astype(int)
    proj_comp = proj_comp.sort_values(["Priority", "Component Code"]).reset_index(drop=True)

    # Filter out items with zero or negative Open Qty
    proj_comp = proj_comp[proj_comp["Open Qty"] > 0].copy()

    # Include projects while GRN+Stock availability remains (allows partial fulfillment and reuse)
    supply_map = {k: float(v["Qty"]) for k, v in grn_qty_by_item.items()}
    if stock_map:
        for k, v in stock_map.items():
            supply_map[k] = supply_map.get(k, 0) + float(v)

    keep_idx = []
    remaining_supply = dict(supply_map)
    for idx, row in proj_comp.sort_values(["Priority", "Component Code"]).iterrows():
        comp_code = row["Component Code"]
        avail = remaining_supply.get(comp_code, 0)
        if avail > 0:
            keep_idx.append(idx)
            req = float(row["Open Qty"])
            remaining_supply[comp_code] = max(0, avail - req)

    proj_comp = proj_comp.loc[keep_idx].copy() if keep_idx else proj_comp.iloc[0:0].copy()
    proj_comp = proj_comp.drop(columns=["_project_key"], errors="ignore")

    return ensure_arrow_compatible(proj_comp)


def build_stock_by_project_sequence(df_oob, stock_map, project_sequence):
    """Show stock vs open qty per project+component, ordered by project priority."""
    if not stock_map:
        return pd.DataFrame()

    priority_map = build_project_priority_map(project_sequence)

    stock_items = {k for k, v in stock_map.items() if v > 0}
    oob_matching = df_oob[df_oob["Component Code"].isin(stock_items)].copy()
    if len(oob_matching) == 0:
        return pd.DataFrame()

    proj_comp = oob_matching.groupby(
        ["Project Name", "Project Num", "Component Code", "Component Desc"], as_index=False
    ).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "Order Number": lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v) and str(v).strip() not in ("", "nan")))),
    })

    proj_comp["_project_key"] = proj_comp["Project Name"].apply(normalize_project_key)
    proj_comp["Priority"] = proj_comp["_project_key"].map(priority_map).fillna(9999).astype(int)
    proj_comp["Stock Qty"] = proj_comp["Component Code"].map(stock_map).fillna(0)

    # Filter out items with zero or negative Open Qty
    proj_comp = proj_comp[proj_comp["Open Qty"] > 0].copy()

    # Sort by priority first, then component code
    proj_comp = proj_comp.sort_values(["Priority", "Component Code"]).reset_index(drop=True)

    # Include projects while stock availability remains (allows partial fulfillment and reuse)
    keep_idx = []
    remaining_stock = dict(stock_map)
    for idx, row in proj_comp.sort_values(["Priority", "Component Code"]).iterrows():
        comp_code = row["Component Code"]
        avail = remaining_stock.get(comp_code, 0)
        if avail > 0:
            keep_idx.append(idx)
            req = float(row["Open Qty"])
            remaining_stock[comp_code] = max(0, avail - req)

    proj_comp = proj_comp.loc[keep_idx].copy() if keep_idx else proj_comp.iloc[0:0].copy()
    proj_comp = proj_comp.drop(columns=["_project_key"], errors="ignore")

    return ensure_arrow_compatible(proj_comp)


def build_project_summary_stock_only(df_oob, stock_map):
    """Build a project-wise summary using only Stock data (no GRN)."""
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

    proj_components = df_oob.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
    proj_components.columns = ["Project Num", "Project Name", "Component Set"]

    if stock_map:
        proj_components["Stock Qty"] = proj_components["Component Set"].apply(
            lambda s: sum(stock_map.get(item, 0) for item in s)
        )
        proj_components["Components In Stock"] = proj_components["Component Set"].apply(
            lambda s: sum(1 for item in s if stock_map.get(item, 0) > 0)
        )
    else:
        proj_components["Stock Qty"] = 0
        proj_components["Components In Stock"] = 0

    proj_summary = proj_agg.merge(
        proj_components[["Project Num", "Project Name", "Stock Qty", "Components In Stock"]],
        on=["Project Num", "Project Name"], how="left"
    )
    proj_summary["Fulfillment %"] = np.where(
        proj_summary["Required Quantity"] > 0,
        ((proj_summary["Quantity Issued"] / proj_summary["Required Quantity"]) * 100).round(1),
        100
    )
    proj_summary = proj_summary.sort_values("Open Qty", ascending=False)
    return proj_summary


def build_component_stock_analysis(df_oob, stock_map):
    """Build component-level stock analysis: Stock Qty vs Open Qty."""
    oob_agg = df_oob.groupby("Component Code", as_index=False).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "On Hand Quantity": "first",
        "Total Available": "first",
        "Component Desc": "first",
    })

    if stock_map:
        oob_agg["Stock Qty"] = oob_agg["Component Code"].map(stock_map).fillna(0)
    else:
        oob_agg["Stock Qty"] = oob_agg["On Hand Quantity"]

    oob_agg["Surplus / Deficit"] = oob_agg["Stock Qty"] - oob_agg["Open Qty"]
    oob_agg["Availability Status"] = np.where(
        oob_agg["Stock Qty"] >= oob_agg["Open Qty"],
        "\u2705 Surplus / Available",
        "\u274C Shortage"
    )
    return oob_agg


def build_project_summary_grn_only(df_oob, df_grn_today):
    """Build a project-wise summary using only GRN data (no Stock file)."""
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

    proj_components = df_oob.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
    proj_components.columns = ["Project Num", "Project Name", "Component Set"]

    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_items_today = set(grn_delivered["Item"].unique())
    grn_qty_by_item = grn_delivered.groupby("Item")["Qty"].sum().to_dict()

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


def build_component_reconciliation_grn_only(df_oob, df_grn_today):
    """Build reconciliation using only GRN data (no stock file)."""
    oob_agg = df_oob.groupby("Component Code", as_index=False).agg({
        "Required Quantity": "sum",
        "Quantity Issued": "sum",
        "Open Qty": "sum",
        "On Hand Quantity": "first",
        "Total Available": "first",
        "Component Desc": "first",
    })
    oob_agg["Stock Qty"] = oob_agg["On Hand Quantity"]

    grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_agg = grn_delivered.groupby("Item", as_index=False).agg({
        "Qty": "sum",
        "Supplier": "first",
        "Order Number": "first",
    }).rename(columns={"Item": "Component Code", "Qty": "Received Today"})

    merged = oob_agg.merge(grn_agg, on="Component Code", how="left")
    merged["Received Today"] = merged["Received Today"].fillna(0)
    merged["Still Pending"] = merged["Open Qty"] - merged["Received Today"]
    merged["Status"] = np.where(
        merged["Still Pending"] <= 0, "\u2705 Fulfilled",
        np.where(merged["Received Today"] > 0, "\U0001F536 Partially Received", "\u274C Not Received")
    )
    return ensure_arrow_compatible(merged)


def get_fulfillable_wo_map(df_oob, project_sequence, stock_map=None, grn_qty_map=None):
    """Return fulfillable Work Order Numbers per (Project, Component) AND per Project.

    Per-component logic:
    1. Build a supply pool per component (stock + GRN).
    2. For each component, collect all (project, Work Order Number, open_qty) rows
       and sort by project priority (asc) then Job Start Date (asc) then WO number.
    3. Walk through projects in priority order, including WOs while supply > 0.
       - Deduct allocated quantity from remaining supply after each WO.
       - Supports partial fulfillment (includes WO even if supply < requirement).
       - Stops when supply reaches 0.
     4. Build three mappings:
       - comp_map:  {(project_name, component_code): "WO1, WO2, …"}
       - proj_map:  {project_name: "WO1, WO2, …"}  (union of all fulfillable WOs)
         - proj_order_map: {project_name: "SO1, SO2, …"} (order numbers for those WOs)

    Uses 'Work Order Number' and 'Job Start Date' columns from the Orderbook.

    Returns:
        tuple  –  (comp_map, proj_map, proj_order_map)
    """
    wo_col = "Work Order Number"
    date_col = "Job Start Date"
    order_col = "Order Number"
    if wo_col not in df_oob.columns:
        return {}, {}, {}

    open_rows = df_oob[df_oob["Open Qty"] > 0].copy()
    if len(open_rows) == 0:
        return {}, {}, {}

    # Build available supply per component
    supply = {}
    if stock_map:
        supply = {k: v for k, v in stock_map.items()}
    if grn_qty_map:
        for k, v in grn_qty_map.items():
            supply[k] = supply.get(k, 0) + v

    # Priority map
    priority_map = build_project_priority_map(project_sequence)

    # Per (Project, WO, Component) demand, keeping Job Start Date
    agg_dict = {"Open Qty": "sum"}
    if date_col in open_rows.columns:
        agg_dict[date_col] = "first"
    if order_col in open_rows.columns:
        agg_dict[order_col] = lambda x: ", ".join(
            sorted(set(str(v) for v in x if pd.notna(v) and str(v).strip() not in ("", "nan")))
        )
    wo_comp = open_rows.groupby(
        ["Project Name", wo_col, "Component Code"], as_index=False
    ).agg(agg_dict)

    # Add sorting keys
    wo_comp["_project_key"] = wo_comp["Project Name"].apply(normalize_project_key)
    wo_comp["Priority"] = wo_comp["_project_key"].map(priority_map).fillna(9999).astype(int)
    if date_col in wo_comp.columns:
        wo_comp["_sort_date"] = pd.to_datetime(wo_comp[date_col], format="mixed", dayfirst=True, errors="coerce")
    else:
        wo_comp["_sort_date"] = pd.NaT

    wo_comp = wo_comp.sort_values(
        ["Priority", "_sort_date", wo_col]
    ).reset_index(drop=True)

    # Walk through each component's WOs in priority / date order
    # Include WOs while supply availability remains (allows partial fulfillment and reuse)
    remaining_supply = dict(supply)
    comp_result = {}   # {(project, component): [wo_list]}
    proj_result = {}   # {project: set(wo_list)}
    proj_order_result = {}  # {project: set(order_numbers)}

    for cc, grp in wo_comp.groupby("Component Code", sort=False):
        cc_rows = grp.sort_values(["Priority", "_sort_date", wo_col])
        for _, row in cc_rows.iterrows():
            avail = remaining_supply.get(cc, 0)
            if avail <= 0:
                break
            wo_str = str(row[wo_col])
            proj = row["Project Name"]
            req = float(row["Open Qty"])
            comp_result.setdefault((proj, cc), []).append(wo_str)
            proj_result.setdefault(proj, set()).add(wo_str)
            if order_col in row and pd.notna(row[order_col]):
                order_vals = [s.strip() for s in str(row[order_col]).split(",") if s.strip()]
                if order_vals:
                    proj_order_result.setdefault(proj, set()).update(order_vals)
            remaining_supply[cc] = max(0, avail - req)

    comp_map = {k: ", ".join(v) for k, v in comp_result.items()}

    # For project map, sort WOs by Job Start Date within each project
    proj_map = {}
    for proj, wo_set in proj_result.items():
        proj_wos = wo_comp[
            (wo_comp["Project Name"] == proj) & (wo_comp[wo_col].astype(str).isin(wo_set))
        ].drop_duplicates(subset=[wo_col]).sort_values(["_sort_date", wo_col])
        proj_map[proj] = ", ".join(proj_wos[wo_col].astype(str).tolist())

    proj_order_map = {
        proj: ", ".join(sorted(vals))
        for proj, vals in proj_order_result.items()
        if vals
    }

    return comp_map, proj_map, proj_order_map


# ==============================================================================
#  SIDEBAR -- FILE UPLOAD & MODE SELECTION
# ==============================================================================

st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/ef/Emerson_Electric_Company.svg/320px-Emerson_Electric_Company.svg.png", width=200)
st.sidebar.markdown("---")

st.sidebar.header("\U0001F4C2 Upload Files")
uploaded_ob = st.sidebar.file_uploader(
    "Orderbook (.xlsm) \u2014 Required",
    type=["xlsm", "xlsx"],
    help="Upload the Nashik iCenter Orderbook file (always required)"
)

st.sidebar.markdown("---")
st.sidebar.header("\U0001F527 Analysis Mode")
analysis_mode = st.sidebar.radio(
    "Select what to analyze:",
    options=[
        "\U0001F4E6 Stock Analysis",
        "\U0001F4E5 GRN Analysis",
        "\U0001F4CA Combined View (Stock + GRN)",
        "\U0001F4CB Supply Eligible Analysis",
    ],
    index=2,
    help="Choose which files to use for the dashboard"
)

# Show relevant file uploaders based on mode
uploaded_stock = None
uploaded_grn = None

if analysis_mode == "\U0001F4E6 Stock Analysis":
    st.sidebar.markdown("---")
    uploaded_stock = st.sidebar.file_uploader(
        "Stock (.xlsx) \u2014 Required",
        type=["xlsx", "xlsm"],
        help="Upload the Stock file for on-hand inventory data"
    )
elif analysis_mode == "\U0001F4E5 GRN Analysis":
    st.sidebar.markdown("---")
    uploaded_grn = st.sidebar.file_uploader(
        "Daily Material Incoming (.xlsx) \u2014 Required",
        type=["xlsx", "xlsm"],
        help="Upload the Daily Material Incoming file"
    )
elif analysis_mode == "\U0001F4CB Supply Eligible Analysis":
    st.sidebar.markdown("---")
    uploaded_stock = st.sidebar.file_uploader(
        "Stock (.xlsx) \u2014 Required",
        type=["xlsx", "xlsm"],
        help="Upload the Stock file for on-hand inventory data"
    )
else:  # Combined
    st.sidebar.markdown("---")
    uploaded_grn = st.sidebar.file_uploader(
        "Daily Material Incoming (.xlsx) \u2014 Required",
        type=["xlsx", "xlsm"],
        help="Upload the Daily Material Incoming file"
    )
    uploaded_stock = st.sidebar.file_uploader(
        "Stock (.xlsx) \u2014 Required",
        type=["xlsx", "xlsm"],
        help="Upload the Stock file for on-hand inventory data"
    )

# ==============================================================================
#  MAIN CONTENT
# ==============================================================================

st.title("\U0001F4E6 Nashik iCenter \u2014 Orderbook & Material Dashboard")
st.caption(f"Report Date: {datetime.now().strftime('%d %B %Y, %A')}")

# -- Validate required files per mode ------------------------------------------
if not uploaded_ob:
    st.info("\U0001F448 Upload the Orderbook file from the sidebar to begin.")
    st.stop()

if analysis_mode == "\U0001F4E6 Stock Analysis" and not uploaded_stock:
    st.info("\U0001F448 Upload the Stock file from the sidebar to use Stock Analysis mode.")
    st.stop()
elif analysis_mode == "\U0001F4E5 GRN Analysis" and not uploaded_grn:
    st.info("\U0001F448 Upload the Daily Material Incoming file from the sidebar to use GRN Analysis mode.")
    st.stop()
elif analysis_mode == "\U0001F4CB Supply Eligible Analysis" and not uploaded_stock:
    st.info("\U0001F448 Upload the Stock file from the sidebar to use Supply Eligible Analysis mode.")
    st.stop()
elif analysis_mode == "\U0001F4CA Combined View (Stock + GRN)" and (not uploaded_grn or not uploaded_stock):
    missing = []
    if not uploaded_grn:
        missing.append("Daily Material Incoming (.xlsx)")
    if not uploaded_stock:
        missing.append("Stock (.xlsx)")
    st.info(f"\U0001F448 Upload the following to use Combined View: **{', '.join(missing)}**")
    st.stop()

# -- Load Orderbook (always) --------------------------------------------------
with st.spinner("Loading Orderbook..."):
    df_oob, ob_sheets = load_orderbook(uploaded_ob)

if df_oob is None:
    st.stop()

# -- Filter to only "Production Open" Sales Status ----------------------------
df_oob_supply_open = pd.DataFrame()  # Keep Supply Eligible rows for separate analysis
if "Sales Status" in df_oob.columns:
    _before = len(df_oob)
    df_oob_supply_open = df_oob[df_oob["Sales Status"].astype(str).str.strip() == "Supply Eligible"].copy()
    df_oob = df_oob[df_oob["Sales Status"].astype(str).str.strip() == "Production Open"].copy()
    st.sidebar.info(f"\U0001F4CB Filtered to 'Production Open': {len(df_oob):,} / {_before:,} rows")
    if len(df_oob_supply_open) > 0:
        st.sidebar.info(f"\U0001F4E6 Supply Eligible rows available: {len(df_oob_supply_open):,}")

# -- Load GRN (if needed) -----------------------------------------------------
df_grn = None
df_grn_today = pd.DataFrame()
selected_date = date.today()

if uploaded_grn:
    with st.spinner("Loading Daily GRN..."):
        df_grn, grn_sheets = load_grn(uploaded_grn)
    if df_grn is None:
        st.stop()

# -- Load Stock (if needed) ----------------------------------------------------
df_stock = None
stock_map = {}
if uploaded_stock:
    with st.spinner("Loading Stock..."):
        df_stock = load_stock(uploaded_stock)
    if df_stock is not None:
        stock_map = get_stock_map(df_stock)
        st.sidebar.success(f"\u2705 Stock loaded: {len(df_stock):,} rows, {len(stock_map):,} unique items")

# -- Date filter for GRN (only if GRN loaded) ---------------------------------
if df_grn is not None:
    st.sidebar.markdown("---")
    st.sidebar.header("\U0001F4C5 Filter by Date")

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

# -- Project filter (always) --------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("\U0001F3D7\uFE0F Filter by Project")

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
    df_oob_supply_open_filtered = df_oob_supply_open[df_oob_supply_open["Project Name"].isin(selected_projects)].copy() if len(df_oob_supply_open) > 0 else pd.DataFrame()
else:
    df_oob_filtered = df_oob.copy()
    df_oob_supply_open_filtered = df_oob_supply_open.copy()

# -- Mode badge ----------------------------------------------------------------
mode_labels = {
    "\U0001F4E6 Stock Analysis": ("Stock Analysis", "available"),
    "\U0001F4E5 GRN Analysis": ("GRN Analysis", "incoming"),
    "\U0001F4CA Combined View (Stock + GRN)": ("Combined View", "project"),
    "\U0001F4CB Supply Eligible Analysis": ("Supply Eligible Analysis", "incoming"),
}
mode_name, mode_css = mode_labels[analysis_mode]
st.markdown(f'<div class="metric-card {mode_css}" style="padding:10px;margin-bottom:15px;"><p style="margin:0;font-size:1.1rem;font-weight:bold;">Mode: {mode_name}</p></div>', unsafe_allow_html=True)

# -- Load / initialize project sequence in session state -----------------------
if "project_sequence" not in st.session_state:
    st.session_state["project_sequence"] = load_saved_project_sequence()

project_sequence = st.session_state["project_sequence"]


# ##############################################################################
#  MODE: STOCK ANALYSIS
# ##############################################################################
if analysis_mode == "\U0001F4E6 Stock Analysis":

    # -- KPI Cards (Stock mode) ------------------------------------------------
    total_open_qty = df_oob_filtered["Open Qty"].sum()
    total_unique_components = df_oob_filtered["Component Code"].nunique()
    total_projects = df_oob_filtered[["Project Num", "Project Name"]].drop_duplicates().shape[0]
    components_in_stock = sum(1 for c in df_oob_filtered["Component Code"].unique() if stock_map.get(c, 0) > 0)
    total_stock_qty = sum(stock_map.get(c, 0) for c in df_oob_filtered["Component Code"].unique())
    shortage_components = sum(1 for c in df_oob_filtered.groupby("Component Code")["Open Qty"].sum().items()
                              if stock_map.get(c[0], 0) < c[1])

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.markdown(f'<div class="metric-card shortage"><h2>{total_open_qty:,.0f}</h2><p>Total Open Qty</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card available"><h2>{total_stock_qty:,.0f}</h2><p>Total Stock Qty (Matching)</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card incoming"><h2>{components_in_stock}</h2><p>Components In Stock</p></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card project"><h2>{total_projects}</h2><p>Active Projects</p></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="metric-card"><h2>{total_unique_components}</h2><p>Unique Components Needed</p></div>', unsafe_allow_html=True)
    with c6:
        st.markdown(f'<div class="metric-card shortage"><h2>{shortage_components}</h2><p>Components in Shortage</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # -- Tabs (Stock mode) -----------------------------------------------------
    tab_s1, tab_s2, tab_s3, tab_s4, tab_s5 = st.tabs([
        "\U0001F3D7\uFE0F Project-wise Summary",
        "\U0001F4E6 Stock by Project Priority",
        "\U0001F50D Component Stock Analysis",
        "\U0001F4CA Analytics",
        "\u2699\uFE0F Project Priority Sequence",
    ])

    # TAB S1: Project Summary (Stock)
    with tab_s1:
        st.subheader("Project-wise Material Requirement & Stock Summary")
        st.caption("Stock-based view \u2014 showing on-hand inventory vs open requirements")

        proj_summary = build_project_summary_stock_only(df_oob_filtered, stock_map)

        display_cols = [
            "Project Num", "Project Name", "Unique Orders", "Unique Items (FG)",
            "Unique Components", "Components Available", "Components In Stock",
            "Required Quantity", "Quantity Issued", "Open Qty", "Fulfillment %",
            "Stock Qty",
        ]
        available_cols = [c for c in display_cols if c in proj_summary.columns]
        proj_display = proj_summary[available_cols].copy()

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
            "Stock Qty": "{:,.0f}",
        })
        display_dataframe_arrow_safe(styled, width='stretch', height=500)

        csv_proj = proj_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Project Summary", csv_proj, "project_summary_stock.csv", "text/csv")

        # Drill-down
        st.markdown("---")
        st.subheader("\U0001F50E Drill-down: Project Component Detail")
        selected_proj_detail = st.selectbox("Select a project to view details:", all_projects, key="stock_drill")

        if selected_proj_detail:
            proj_detail = df_oob_filtered[df_oob_filtered["Project Name"] == selected_proj_detail]
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
            proj_comp_detail["Stock Qty"] = proj_comp_detail["Component Code"].map(stock_map).fillna(0)
            proj_comp_detail["Surplus / Deficit"] = proj_comp_detail["Stock Qty"] - proj_comp_detail["Open Qty"]
            proj_comp_detail["Status"] = np.where(
                proj_comp_detail["Surplus / Deficit"] >= 0, "\u2705 Available", "\u274C Shortage"
            )
            proj_comp_detail = proj_comp_detail.sort_values("Surplus / Deficit", ascending=True)

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Total Components", len(proj_comp_detail))
            col_b.metric("In Stock", int((proj_comp_detail["Stock Qty"] > 0).sum()))
            col_c.metric("Shortage", int((proj_comp_detail["Surplus / Deficit"] < 0).sum()))
            st.dataframe(proj_comp_detail, width='stretch', height=400)

    # TAB S2: Stock by Project Priority
    with tab_s2:
        st.subheader("Stock \u2014 Ordered by Project Priority")
        st.caption("Components cascade to projects by priority order. Remaining stock after each allocation can fulfill lower-priority projects.")

        project_sequence = st.session_state["project_sequence"]
        stock_by_proj = build_stock_by_project_sequence(df_oob_filtered, stock_map, project_sequence)

        if len(stock_by_proj) == 0:
            st.warning("No stock items matched with Orderbook components.")
        else:
            m1, m2, m3 = st.columns(3)
            m1.metric("Projects Impacted", stock_by_proj["Project Name"].nunique())
            m2.metric("Unique Components in Stock", stock_by_proj["Component Code"].nunique())
            m3.metric("Total Open Qty", f"{stock_by_proj['Open Qty'].sum():,.0f}")

            st.markdown("---")

            stock_proj_display = stock_by_proj.copy()

            # Add comma-separated fulfillable work orders per (project, component)
            wo_map_stock, wo_proj_map_stock, wo_proj_order_map_stock = get_fulfillable_wo_map(df_oob_filtered, project_sequence, stock_map=stock_map)
            stock_proj_display["Fulfillable Work Orders"] = stock_proj_display.apply(
                lambda r: wo_map_stock.get((r["Project Name"], r["Component Code"]), ""), axis=1
            )

            # Add Serial No. starting from 1
            stock_proj_display = stock_proj_display.reset_index(drop=True)
            stock_proj_display.insert(0, "Serial No.", range(1, len(stock_proj_display) + 1))

            stock_proj_display_cols = [
                "Serial No.", "Priority", "Project Name", "Order Number", "Component Code",
                "Component Desc", "Required Quantity", "Quantity Issued",
                "Open Qty", "Stock Qty", "Fulfillable Work Orders",
            ]
            available_display_cols = [c for c in stock_proj_display_cols if c in stock_proj_display.columns]

            def color_priority_stock(val):
                if val <= 5: return "background-color: #f8d7da"
                elif val <= 15: return "background-color: #fff3cd"
                elif val <= 30: return "background-color: #d4edda"
                else: return ""

            stock_proj_styled = stock_proj_display[available_display_cols].copy()
            style_method_sp = getattr(stock_proj_styled.style, "map", None) or stock_proj_styled.style.applymap
            stock_proj_styled_df = style_method_sp(
                color_priority_stock, subset=["Priority"]
            ).format({
                "Required Quantity": "{:,.0f}", "Quantity Issued": "{:,.0f}",
                "Open Qty": "{:,.0f}", "Stock Qty": "{:,.0f}",
            })
            st.dataframe(stock_proj_styled_df, width='stretch', height=600)

            csv_stock_proj = stock_proj_display[available_display_cols].to_csv(index=False)
            st.download_button("\U0001F4E5 Download Stock by Project Priority", csv_stock_proj, "stock_by_project_priority.csv", "text/csv")

            # Project-wise Work Order Summary
            if wo_proj_map_stock:
                st.markdown("---")
                st.subheader("\U0001F4CB Fulfillable Work Orders by Project")
                st.caption("Work Orders cascade by priority while stock remains. Partial fulfillment included. Sorted by Job Start Date.")
                wo_proj_norm = {normalize_project_key(k): v for k, v in wo_proj_map_stock.items()}
                wo_proj_order_norm = {normalize_project_key(k): v for k, v in wo_proj_order_map_stock.items()}
                proj_wo_rows = []
                for proj_name in project_sequence["Project Name"].str.strip():
                    proj_key = normalize_project_key(proj_name)
                    if proj_key in wo_proj_norm:
                        proj_wo_rows.append({
                            "Project Name": proj_name,
                            "Fulfillable Work Orders": wo_proj_norm[proj_key],
                            "Order Number": wo_proj_order_norm.get(proj_key, ""),
                        })
                if proj_wo_rows:
                    proj_wo_df = pd.DataFrame(proj_wo_rows)
                    proj_wo_df.insert(0, "Sr.", range(1, len(proj_wo_df) + 1))
                    st.dataframe(proj_wo_df, width='stretch', height=min(400, 35 * len(proj_wo_df) + 38))

    # TAB S3: Component Stock Analysis
    with tab_s3:
        st.subheader("Component-Level Stock Analysis: Stock Qty vs Open Qty")
        st.caption("Compares current stock on hand against open requirements")

        stock_analysis = build_component_stock_analysis(df_oob_filtered, stock_map)

        filter_col1, filter_col2 = st.columns(2)
        with filter_col1:
            avail_filter = st.multiselect(
                "Filter by Availability:",
                options=["\u2705 Surplus / Available", "\u274C Shortage"],
                default=[],
                key="stock_avail_filter"
            )

        stock_display = stock_analysis.copy()
        if avail_filter:
            stock_display = stock_display[stock_display["Availability Status"].isin(avail_filter)]

        stock_cols = [
            "Component Code", "Component Desc", "Required Quantity", "Quantity Issued",
            "Open Qty", "Stock Qty", "Surplus / Deficit", "Total Available", "Availability Status"
        ]
        available_stock_cols = [c for c in stock_cols if c in stock_display.columns]

        st.dataframe(
            stock_display[available_stock_cols].sort_values("Surplus / Deficit", ascending=True),
            width='stretch', height=500
        )

        csv_stock = stock_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Stock Analysis", csv_stock, "component_stock_analysis.csv", "text/csv")

        st.markdown("---")
        sc1, sc2, sc3, sc4 = st.columns(4)
        surplus_count = (stock_analysis["Availability Status"] == "\u2705 Surplus / Available").sum()
        shortage_count = (stock_analysis["Availability Status"] == "\u274C Shortage").sum()
        sc1.metric("\u2705 Surplus / Available", surplus_count)
        sc2.metric("\u274C Shortage", shortage_count)
        sc3.metric("Unique Items in Stock", len(stock_map))
        sc4.metric("Total Stock Qty", f"{sum(stock_map.values()):,.0f}")

    # TAB S4: Analytics (Stock)
    with tab_s4:
        st.subheader("\U0001F4CA Visual Analytics (Stock Mode)")

        # Build full project summary
        proj_summary_chart = build_project_summary_stock_only(df_oob_filtered, stock_map)
        proj_open_components = df_oob_filtered[df_oob_filtered["Open Qty"] > 0].groupby(
            ["Project Num", "Project Name"], as_index=False
        )["Component Code"].nunique().rename(columns={"Component Code": "Open Components"})
        proj_summary_chart = proj_summary_chart.merge(
            proj_open_components, on=["Project Num", "Project Name"], how="left"
        )
        proj_summary_chart["Open Components"] = proj_summary_chart["Open Components"].fillna(0).astype(int)
        proj_summary_chart["Component Fulfillment"] = np.where(
            proj_summary_chart["Unique Components"] > 0,
            (((proj_summary_chart["Unique Components"] - proj_summary_chart["Open Components"]) / proj_summary_chart["Unique Components"]) * 100).round(1),
            100
        )
        proj_summary_chart["Open/Total Label"] = (
            proj_summary_chart["Open Components"].astype(str) + "/" + proj_summary_chart["Unique Components"].astype(str)
        )
        # Add Open Qty per project for filtering/sorting
        proj_open_qty = df_oob_filtered.groupby("Project Name", as_index=False)["Open Qty"].sum().rename(columns={"Open Qty": "Total Open Qty"})
        proj_summary_chart = proj_summary_chart.merge(proj_open_qty, on="Project Name", how="left")
        proj_summary_chart["Total Open Qty"] = proj_summary_chart["Total Open Qty"].fillna(0)

        # Filters
        st.markdown("##### \U0001F50D Filters & Sorting")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            sort_option_s = st.selectbox("Sort by:", [
                "Open Components (desc)", "Open Components (asc)",
                "Fulfillment % (asc)", "Fulfillment % (desc)",
                "Total Open Qty (desc)", "Total Open Qty (asc)",
                "Project Name (A-Z)",
            ], key="stock_analytics_sort")
        with fc2:
            all_projects_s = proj_summary_chart["Project Name"].unique().tolist()
            selected_projects_s = st.multiselect(
                "Filter by Project:", options=all_projects_s, default=[], key="stock_analytics_proj_filter",
                placeholder="All projects shown"
            )
        with fc3:
            fulfillment_range_s = st.slider(
                "Fulfillment % range:", 0, 100, (0, 100), key="stock_analytics_fulfill_range"
            )

        # Apply filters
        proj_filtered = proj_summary_chart.copy()
        if selected_projects_s:
            proj_filtered = proj_filtered[proj_filtered["Project Name"].isin(selected_projects_s)]
        proj_filtered = proj_filtered[
            (proj_filtered["Component Fulfillment"] >= fulfillment_range_s[0]) &
            (proj_filtered["Component Fulfillment"] <= fulfillment_range_s[1])
        ]

        # Apply sorting
        sort_map = {
            "Open Components (desc)": ("Open Components", False),
            "Open Components (asc)": ("Open Components", True),
            "Fulfillment % (asc)": ("Component Fulfillment", True),
            "Fulfillment % (desc)": ("Component Fulfillment", False),
            "Total Open Qty (desc)": ("Total Open Qty", False),
            "Total Open Qty (asc)": ("Total Open Qty", True),
            "Project Name (A-Z)": ("Project Name", True),
        }
        sort_col, sort_asc = sort_map.get(sort_option_s, ("Open Components", False))
        proj_filtered = proj_filtered.sort_values(sort_col, ascending=sort_asc)

        st.caption(f"Showing {len(proj_filtered)} of {len(proj_summary_chart)} projects")

        chart_height = max(400, len(proj_filtered) * 28 + 100)

        col1, col2 = st.columns(2)
        with col1:
            fig1 = px.bar(
                proj_filtered, x="Open Components", y="Project Name", orientation="h",
                title=f"All Projects by Open Components ({len(proj_filtered)} projects)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Open Components": "Unique Components with Open Qty", "Project Name": "Project"},
                text="Open/Total Label",
            )
            fig1.update_traces(textposition="outside")
            fig1.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig1, width='stretch')

        with col2:
            proj_filtered_text = proj_filtered.copy()
            proj_filtered_text["Label"] = proj_filtered_text["Component Fulfillment"].apply(lambda x: f"{x:.1f}%")
            fig2 = px.bar(
                proj_filtered_text, x="Component Fulfillment", y="Project Name", orientation="h",
                title="Component Fulfillment % (Fulfilled / Total Unique Components)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Component Fulfillment": "Fulfillment %", "Project Name": "Project"},
                text="Label",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig2, width='stretch')

        # Downloadable table
        st.markdown("---")
        st.subheader("Project Analytics Table")
        analytics_table_s = proj_filtered[["Project Name", "Unique Components", "Open Components",
                                           "Component Fulfillment", "Total Open Qty"]].reset_index(drop=True)
        analytics_table_s.index = analytics_table_s.index + 1
        analytics_table_s.index.name = "Sr."
        st.dataframe(analytics_table_s, width='stretch', height=min(500, 40 + len(analytics_table_s) * 35))
        st.download_button("\U0001F4E5 Download Analytics Table", analytics_table_s.to_csv(), "stock_analytics.csv", "text/csv", key="dl_stock_analytics")

        st.markdown("---")

        col3, col4 = st.columns(2)
        with col3:
            stock_full = build_component_stock_analysis(df_oob_filtered, stock_map)
            avail_counts = stock_full["Availability Status"].value_counts().reset_index()
            avail_counts.columns = ["Status", "Count"]
            fig3 = px.pie(
                avail_counts, values="Count", names="Status",
                title="Component Stock Availability Status",
                color_discrete_map={
                    "\u2705 Surplus / Available": "#2ecc71",
                    "\u274C Shortage": "#e74c3c",
                }
            )
            st.plotly_chart(fig3, width='stretch')

        with col4:
            avail_counts2 = df_oob_filtered["Availability"].value_counts().reset_index()
            avail_counts2.columns = ["Availability", "Count"]
            fig4 = px.pie(avail_counts2, values="Count", names="Availability",
                          title="Orderbook Availability Status",
                          color_discrete_map={"Available": "#2ecc71", "Shortage": "#e74c3c"})
            st.plotly_chart(fig4, width='stretch')

        st.markdown("---")
        st.subheader("\U0001F4E6 Stock Summary")
        sc1, sc2 = st.columns(2)
        with sc1:
            st.metric("Unique Items in Stock", len(stock_map))
            st.metric("Total Stock Qty", f"{sum(stock_map.values()):,.0f}")
        with sc2:
            oob_components = set(df_oob_filtered["Component Code"].unique())
            stock_matching = {k: v for k, v in stock_map.items() if k in oob_components}
            st.metric("Stock Items Matching Orderbook", len(stock_matching))
            st.metric("Matching Stock Qty", f"{sum(stock_matching.values()):,.0f}")

    # TAB S5: Project Priority Sequence
    with tab_s5:
        st.subheader("\u2699\uFE0F Project Priority Sequence")
        st.caption("Edit the project priority order. Lower Sr. = higher priority.")
        render_sequence_upload_controls("stock")

        if "Completed" not in st.session_state["project_sequence"].columns:
            st.session_state["project_sequence"]["Completed"] = False

        show_completed = st.checkbox("Show completed projects", value=False, key="stock_show_completed")
        if show_completed:
            display_sequence = st.session_state["project_sequence"].copy()
        else:
            display_sequence = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == False].copy()

        col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4, col_ctrl5 = st.columns([3, 1, 1, 1, 1])
        with col_ctrl1:
            project_names = display_sequence["Project Name"].tolist()
            selected_project = st.selectbox("Select project to move:", project_names, key="stock_proj_select")
        with col_ctrl2:
            if st.button("\u2b06\ufe0f Move Up", key="stock_up"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx > 0:
                        sr_current = seq.at[idx, "Sr."]; sr_above = seq.at[idx-1, "Sr."]
                        seq.at[idx, "Sr."] = sr_above; seq.at[idx-1, "Sr."] = sr_current
                        set_project_sequence(seq.sort_values("Sr.").reset_index(drop=True))
                        st.rerun()
        with col_ctrl3:
            if st.button("\u2b07\ufe0f Move Down", key="stock_down"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx < len(seq) - 1:
                        sr_current = seq.at[idx, "Sr."]; sr_below = seq.at[idx+1, "Sr."]
                        seq.at[idx, "Sr."] = sr_below; seq.at[idx+1, "Sr."] = sr_current
                        set_project_sequence(seq.sort_values("Sr.").reset_index(drop=True))
                        st.rerun()
        with col_ctrl4:
            if st.button("\u2705 Mark Complete", key="stock_complete"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = True
                    set_project_sequence(seq); st.rerun()
        with col_ctrl5:
            if st.button("\u21a9\ufe0f Unmark", key="stock_unmark"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = False
                    set_project_sequence(seq); st.rerun()

        st.markdown("---")
        edited_sequence = st.data_editor(
            display_sequence, num_rows="dynamic", width='stretch', height=500,
            column_config={
                "Sr.": st.column_config.NumberColumn("Sr.", min_value=1, step=1, width="small"),
                "Project Name": st.column_config.TextColumn("Project Name", width="large"),
                "No. Of Cabinet": st.column_config.NumberColumn("No. Of Cabinet", min_value=0, step=1),
                "Prod. Open Dt.": st.column_config.TextColumn("Prod. Open Dt."),
                "SO No.": st.column_config.NumberColumn("SO No.", format="%d"),
                "Completed": st.column_config.CheckboxColumn("Completed", width="small"),
            },
            key="stock_seq_editor",
        )

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("\u2705 Apply Changes", type="primary", key="stock_apply"):
                if show_completed:
                    updated = edited_sequence.sort_values("Sr.").reset_index(drop=True)
                else:
                    completed_items = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == True]
                    updated = pd.concat([edited_sequence, completed_items], ignore_index=True).sort_values("Sr.").reset_index(drop=True)
                set_project_sequence(updated); st.rerun()
        with col_btn2:
            if st.button("\U0001F504 Reset to Default", key="stock_reset"):
                set_project_sequence(DEFAULT_PROJECT_SEQUENCE.copy()); st.rerun()
        with col_btn3:
            if st.button("\U0001F5D1\uFE0F Remove Completed", key="stock_remove"):
                seq = st.session_state["project_sequence"]
                seq = seq[seq["Completed"] == False].copy()
                seq["Sr."] = range(1, len(seq) + 1)
                set_project_sequence(seq); st.rerun()


# ##############################################################################
#  MODE: GRN ANALYSIS
# ##############################################################################
elif analysis_mode == "\U0001F4E5 GRN Analysis":

    # -- KPI Cards (GRN mode) --------------------------------------------------
    grn_delivered_today = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
    grn_rejected_today = df_grn_today[df_grn_today["GRN_Status"].str.contains("Reject", case=False, na=False)]

    total_open_qty = df_oob_filtered["Open Qty"].sum()
    total_received_today = grn_delivered_today["Qty"].sum()
    total_unique_components = df_oob_filtered["Component Code"].nunique()
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
        st.markdown(f'<div class="metric-card"><h2>{total_unique_components}</h2><p>Unique Components Needed</p></div>', unsafe_allow_html=True)
    with c6:
        st.markdown(f'<div class="metric-card shortage"><h2>{total_rejected_today:,.0f}</h2><p>Rejected Today</p></div>', unsafe_allow_html=True)

    st.markdown("---")

    # -- Tabs (GRN mode) -------------------------------------------------------
    tab_g1, tab_g2, tab_g3, tab_g4, tab_g5 = st.tabs([
        "\U0001F3D7\uFE0F Project-wise Summary",
        "\U0001F4E5 GRN by Project Priority",
        "\U0001F50D Component Reconciliation",
        "\U0001F4CA Analytics",
        "\u2699\uFE0F Project Priority Sequence",
    ])

    # TAB G1: Project Summary (GRN only)
    with tab_g1:
        st.subheader("Project-wise Material Requirement & GRN Summary")
        st.caption(f"GRN Date: {selected_date.strftime('%d-%b-%Y')} | Using Orderbook On Hand Qty (no stock file)")

        proj_summary = build_project_summary_grn_only(df_oob_filtered, df_grn_today)

        display_cols = [
            "Project Num", "Project Name", "Unique Orders", "Unique Items (FG)",
            "Unique Components", "Components Available",
            "Required Quantity", "Quantity Issued", "Open Qty", "Fulfillment %",
            "Components Received Today", "Qty Received Today",
        ]
        available_cols = [c for c in display_cols if c in proj_summary.columns]
        proj_display = proj_summary[available_cols].copy()

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
        st.dataframe(styled, width='stretch', height=500)

        csv_proj = proj_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Project Summary", csv_proj, "project_summary_grn.csv", "text/csv")

        # Drill-down
        st.markdown("---")
        st.subheader("\U0001F50E Drill-down: Project Component Detail")
        selected_proj_detail = st.selectbox("Select a project to view details:", all_projects, key="grn_drill")

        if selected_proj_detail:
            proj_detail = df_oob_filtered[df_oob_filtered["Project Name"] == selected_proj_detail]
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
                proj_comp_detail["Net Pending"] <= 0, "\u2705 Fulfilled",
                np.where(proj_comp_detail["Received Today"] > 0, "\U0001F536 Partial", "\u274C Pending")
            )
            proj_comp_detail = proj_comp_detail.sort_values("Net Pending", ascending=False)

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Total Components", len(proj_comp_detail))
            col_b.metric("Received Today", int(proj_comp_detail["Received Today"].gt(0).sum()))
            col_c.metric("Still Pending", int(proj_comp_detail["Net Pending"].gt(0).sum()))
            st.dataframe(proj_comp_detail, width='stretch', height=400)

    # TAB G2: GRN by Project Priority
    with tab_g2:
        st.subheader("Today's GRN Items \u2014 Ordered by Project Priority")
        st.caption(f"GRN Date: {selected_date.strftime('%d-%b-%Y')} | Components cascade to projects by priority. Remaining supply after each allocation can fulfill lower-priority projects")

        grn_by_proj = build_grn_by_project_sequence(df_oob_filtered, df_grn_today, project_sequence)

        if len(grn_by_proj) == 0:
            st.warning("No GRN items matched with Orderbook components for the selected date.")
        else:
            m1, m2, m3 = st.columns(3)
            m1.metric("Projects Impacted", grn_by_proj["Project Name"].nunique())
            m2.metric("Unique Components Received", grn_by_proj["Component Code"].nunique())
            m3.metric("Total GRN Qty", f"{grn_by_proj['GRN Qty Today'].sum():,.0f}")

            st.markdown("---")

            # Add Serial No. starting from 1
            grn_by_proj_display = grn_by_proj.copy()
            grn_by_proj_display = grn_by_proj_display.reset_index(drop=True)
            grn_by_proj_display.insert(0, "Serial No.", range(1, len(grn_by_proj_display) + 1))

            # Add comma-separated fulfillable work orders per (project, component)
            grn_delivered_items = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
            grn_qty_map_wo = grn_delivered_items.groupby("Item")["Qty"].sum().to_dict()
            wo_map_grn, wo_proj_map_grn, wo_proj_order_map_grn = get_fulfillable_wo_map(df_oob_filtered, project_sequence, grn_qty_map=grn_qty_map_wo)
            grn_by_proj_display["Fulfillable Work Orders"] = grn_by_proj_display.apply(
                lambda r: wo_map_grn.get((r["Project Name"], r["Component Code"]), ""), axis=1
            )

            grn_proj_display_cols = [
                "Serial No.", "Priority", "Project Name", "Order Number", "Component Code",
                "Component Desc", "Required Quantity", "Quantity Issued",
                "Open Qty", "GRN Qty Today", "Supplier", "GRN Order Number",
                "Fulfillable Work Orders",
            ]
            available_display_cols = [c for c in grn_proj_display_cols if c in grn_by_proj_display.columns]

            def color_priority(val):
                if val <= 5: return "background-color: #f8d7da"
                elif val <= 15: return "background-color: #fff3cd"
                elif val <= 30: return "background-color: #d4edda"
                else: return ""

            grn_styled = grn_by_proj_display[available_display_cols].copy()
            style_method2 = getattr(grn_styled.style, "map", None) or grn_styled.style.applymap
            grn_styled_df = style_method2(
                color_priority, subset=["Priority"]
            ).format({
                "Required Quantity": "{:,.0f}", "Quantity Issued": "{:,.0f}",
                "Open Qty": "{:,.0f}", "GRN Qty Today": "{:,.0f}",
            })
            display_dataframe_arrow_safe(grn_styled_df, width='stretch', height=600)

            csv_grn_proj = grn_by_proj_display[available_display_cols].to_csv(index=False)
            st.download_button("\U0001F4E5 Download GRN by Project Priority", csv_grn_proj, "grn_by_project_priority.csv", "text/csv")

            # Project-wise Work Order Summary
            if wo_proj_map_grn:
                st.markdown("---")
                st.subheader("\U0001F4CB Fulfillable Work Orders by Project")
                st.caption("Work Orders cascade by priority while GRN supply remains. Partial fulfillment included. Sorted by Job Start Date.")
                wo_proj_norm = {normalize_project_key(k): v for k, v in wo_proj_map_grn.items()}
                wo_proj_order_norm = {normalize_project_key(k): v for k, v in wo_proj_order_map_grn.items()}
                proj_wo_rows = []
                for proj_name in project_sequence["Project Name"].str.strip():
                    proj_key = normalize_project_key(proj_name)
                    if proj_key in wo_proj_norm:
                        proj_wo_rows.append({
                            "Project Name": proj_name,
                            "Fulfillable Work Orders": wo_proj_norm[proj_key],
                            "Order Number": wo_proj_order_norm.get(proj_key, ""),
                        })
                if proj_wo_rows:
                    proj_wo_df = pd.DataFrame(proj_wo_rows)
                    proj_wo_df.insert(0, "Sr.", range(1, len(proj_wo_df) + 1))
                    st.dataframe(proj_wo_df, width='stretch', height=min(400, 35 * len(proj_wo_df) + 38))

            grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
            matched_items = set(grn_by_proj["Component Code"].unique())
            unmatched_grn = grn_delivered[~grn_delivered["Item"].isin(matched_items)]
            if len(unmatched_grn) > 0:
                st.markdown("---")
                st.subheader("\u26A0\uFE0F GRN Items Not in Orderbook")
                grn_detail_cols = ["Item", "Item Description", "Qty", "Supplier", "Order Number"]
                avail_cols = [c for c in grn_detail_cols if c in unmatched_grn.columns]
                display_dataframe_arrow_safe(unmatched_grn[avail_cols], width='stretch', height=300)

    # TAB G3: Component Reconciliation (GRN only)
    with tab_g3:
        st.subheader("Component-Level Reconciliation: Open Qty vs Today's GRN")
        st.caption("Required Quantity \u2212 Quantity Issued = Open Qty \u2192 compared with GRN Received Today")

        reconciliation = build_component_reconciliation_grn_only(df_oob_filtered, df_grn_today)

        reconciliation["Availability Status"] = np.where(
            reconciliation["Stock Qty"] >= reconciliation["Open Qty"],
            "\u2705 Surplus / Available",
            "\u274C Shortage"
        )

        filter_col1, filter_col2 = st.columns(2)
        with filter_col1:
            status_filter = st.multiselect(
                "Filter by GRN Status:",
                options=["\u2705 Fulfilled", "\U0001F536 Partially Received", "\u274C Not Received"],
                default=["\U0001F536 Partially Received", "\u274C Not Received"],
                key="grn_status_filter"
            )
        with filter_col2:
            avail_filter = st.multiselect(
                "Filter by On Hand vs Open Qty:",
                options=["\u2705 Surplus / Available", "\u274C Shortage"],
                default=[],
                key="grn_avail_filter"
            )

        recon_display = reconciliation.copy()
        if status_filter:
            recon_display = recon_display[recon_display["Status"].isin(status_filter)]
        if avail_filter:
            recon_display = recon_display[recon_display["Availability Status"].isin(avail_filter)]

        recon_cols = [
            "Component Code", "Component Desc", "Required Quantity", "Quantity Issued",
            "Open Qty", "Stock Qty", "Received Today", "Still Pending",
            "Total Available", "Supplier", "Status", "Availability Status"
        ]
        available_recon_cols = [c for c in recon_cols if c in recon_display.columns]
        display_dataframe_arrow_safe(
            recon_display[available_recon_cols].sort_values("Still Pending", ascending=False),
            width='stretch', height=500
        )

        csv_recon = recon_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Reconciliation", csv_recon, "reconciliation_grn.csv", "text/csv")

        st.markdown("---")
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("\u2705 Fulfilled", (reconciliation["Status"] == "\u2705 Fulfilled").sum())
        rc2.metric("\U0001F536 Partially Received", (reconciliation["Status"] == "\U0001F536 Partially Received").sum())
        rc3.metric("\u274C Not Received", (reconciliation["Status"] == "\u274C Not Received").sum())

    # TAB G4: Analytics (GRN)
    with tab_g4:
        st.subheader("\U0001F4CA Visual Analytics (GRN Mode)")

        # Build full project summary
        proj_summary_chart = build_project_summary_grn_only(df_oob_filtered, df_grn_today)
        proj_open_components = df_oob_filtered[df_oob_filtered["Open Qty"] > 0].groupby(
            ["Project Num", "Project Name"], as_index=False
        )["Component Code"].nunique().rename(columns={"Component Code": "Open Components"})
        proj_summary_chart = proj_summary_chart.merge(
            proj_open_components, on=["Project Num", "Project Name"], how="left"
        )
        proj_summary_chart["Open Components"] = proj_summary_chart["Open Components"].fillna(0).astype(int)
        proj_summary_chart["Component Fulfillment"] = np.where(
            proj_summary_chart["Unique Components"] > 0,
            (((proj_summary_chart["Unique Components"] - proj_summary_chart["Open Components"]) / proj_summary_chart["Unique Components"]) * 100).round(1),
            100
        )
        proj_summary_chart["Open/Total Label"] = (
            proj_summary_chart["Open Components"].astype(str) + "/" + proj_summary_chart["Unique Components"].astype(str)
        )
        proj_open_qty = df_oob_filtered.groupby("Project Name", as_index=False)["Open Qty"].sum().rename(columns={"Open Qty": "Total Open Qty"})
        proj_summary_chart = proj_summary_chart.merge(proj_open_qty, on="Project Name", how="left")
        proj_summary_chart["Total Open Qty"] = proj_summary_chart["Total Open Qty"].fillna(0)

        # Filters
        st.markdown("##### \U0001F50D Filters & Sorting")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            sort_option_g = st.selectbox("Sort by:", [
                "Open Components (desc)", "Open Components (asc)",
                "Fulfillment % (asc)", "Fulfillment % (desc)",
                "Total Open Qty (desc)", "Total Open Qty (asc)",
                "Project Name (A-Z)",
            ], key="grn_analytics_sort")
        with fc2:
            all_projects_g = proj_summary_chart["Project Name"].unique().tolist()
            selected_projects_g = st.multiselect(
                "Filter by Project:", options=all_projects_g, default=[], key="grn_analytics_proj_filter",
                placeholder="All projects shown"
            )
        with fc3:
            fulfillment_range_g = st.slider(
                "Fulfillment % range:", 0, 100, (0, 100), key="grn_analytics_fulfill_range"
            )

        # Apply filters
        proj_filtered = proj_summary_chart.copy()
        if selected_projects_g:
            proj_filtered = proj_filtered[proj_filtered["Project Name"].isin(selected_projects_g)]
        proj_filtered = proj_filtered[
            (proj_filtered["Component Fulfillment"] >= fulfillment_range_g[0]) &
            (proj_filtered["Component Fulfillment"] <= fulfillment_range_g[1])
        ]

        # Apply sorting
        sort_map = {
            "Open Components (desc)": ("Open Components", False),
            "Open Components (asc)": ("Open Components", True),
            "Fulfillment % (asc)": ("Component Fulfillment", True),
            "Fulfillment % (desc)": ("Component Fulfillment", False),
            "Total Open Qty (desc)": ("Total Open Qty", False),
            "Total Open Qty (asc)": ("Total Open Qty", True),
            "Project Name (A-Z)": ("Project Name", True),
        }
        sort_col, sort_asc = sort_map.get(sort_option_g, ("Open Components", False))
        proj_filtered = proj_filtered.sort_values(sort_col, ascending=sort_asc)

        st.caption(f"Showing {len(proj_filtered)} of {len(proj_summary_chart)} projects")

        chart_height = max(400, len(proj_filtered) * 28 + 100)

        col1, col2 = st.columns(2)
        with col1:
            fig1 = px.bar(
                proj_filtered, x="Open Components", y="Project Name", orientation="h",
                title=f"All Projects by Open Components ({len(proj_filtered)} projects)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Open Components": "Unique Components with Open Qty", "Project Name": "Project"},
                text="Open/Total Label",
            )
            fig1.update_traces(textposition="outside")
            fig1.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig1, width='stretch')

        with col2:
            proj_filtered_text = proj_filtered.copy()
            proj_filtered_text["Label"] = proj_filtered_text["Component Fulfillment"].apply(lambda x: f"{x:.1f}%")
            fig2 = px.bar(
                proj_filtered_text, x="Component Fulfillment", y="Project Name", orientation="h",
                title="Component Fulfillment % (Fulfilled / Total Unique Components)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Component Fulfillment": "Fulfillment %", "Project Name": "Project"},
                text="Label",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig2, width='stretch')

        # Downloadable table
        st.markdown("---")
        st.subheader("Project Analytics Table")
        analytics_table_g = proj_filtered[["Project Name", "Unique Components", "Open Components",
                                           "Component Fulfillment", "Total Open Qty"]].reset_index(drop=True)
        analytics_table_g.index = analytics_table_g.index + 1
        analytics_table_g.index.name = "Sr."
        st.dataframe(analytics_table_g, width='stretch', height=min(500, 40 + len(analytics_table_g) * 35))
        st.download_button("\U0001F4E5 Download Analytics Table", analytics_table_g.to_csv(), "grn_analytics.csv", "text/csv", key="dl_grn_analytics")

        st.markdown("---")
        col3, col4 = st.columns(2)
        with col3:
            recon_full = build_component_reconciliation_grn_only(df_oob_filtered, df_grn_today)
            status_counts = recon_full["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            fig3 = px.pie(
                status_counts, values="Count", names="Status",
                title="Component Reconciliation Status",
                color_discrete_map={
                    "\u2705 Fulfilled": "#2ecc71",
                    "\U0001F536 Partially Received": "#f39c12",
                    "\u274C Not Received": "#e74c3c",
                }
            )
            st.plotly_chart(fig3, width='stretch')

        with col4:
            if len(grn_delivered_today) > 0:
                top_suppliers = grn_delivered_today.groupby("Supplier")["Qty"].sum().nlargest(10).reset_index()
                fig4 = px.bar(
                    top_suppliers, x="Qty", y="Supplier", orientation="h",
                    title=f"Top 10 Suppliers by Qty Delivered ({selected_date.strftime('%d-%b')})",
                    color="Qty", color_continuous_scale="Blues",
                )
                fig4.update_layout(yaxis=dict(autorange="reversed"), height=400)
                st.plotly_chart(fig4, width='stretch')
            else:
                st.info("No deliveries for the selected date.")

        st.markdown("---")
        st.subheader("Component Availability Breakdown")
        avail_counts = df_oob_filtered["Availability"].value_counts().reset_index()
        avail_counts.columns = ["Availability", "Count"]
        fig5 = px.pie(avail_counts, values="Count", names="Availability",
                      title="Availability Status of All Components in Orderbook",
                      color_discrete_map={"Available": "#2ecc71", "Shortage": "#e74c3c"})
        st.plotly_chart(fig5, width='stretch')

    # TAB G5: Project Priority Sequence
    with tab_g5:
        st.subheader("\u2699\uFE0F Project Priority Sequence")
        st.caption("Edit the project priority order. Lower Sr. = higher priority.")
        render_sequence_upload_controls("grn")

        if "Completed" not in st.session_state["project_sequence"].columns:
            st.session_state["project_sequence"]["Completed"] = False

        show_completed = st.checkbox("Show completed projects", value=False, key="grn_show_completed")
        if show_completed:
            display_sequence = st.session_state["project_sequence"].copy()
        else:
            display_sequence = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == False].copy()

        col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4, col_ctrl5 = st.columns([3, 1, 1, 1, 1])
        with col_ctrl1:
            project_names = display_sequence["Project Name"].tolist()
            selected_project = st.selectbox("Select project to move:", project_names, key="grn_proj_select")
        with col_ctrl2:
            if st.button("\u2b06\ufe0f Move Up", key="grn_up"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx > 0:
                        sr_current = seq.at[idx, "Sr."]; sr_above = seq.at[idx-1, "Sr."]
                        seq.at[idx, "Sr."] = sr_above; seq.at[idx-1, "Sr."] = sr_current
                        set_project_sequence(seq.sort_values("Sr.").reset_index(drop=True))
                        st.rerun()
        with col_ctrl3:
            if st.button("\u2b07\ufe0f Move Down", key="grn_down"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx < len(seq) - 1:
                        sr_current = seq.at[idx, "Sr."]; sr_below = seq.at[idx+1, "Sr."]
                        seq.at[idx, "Sr."] = sr_below; seq.at[idx+1, "Sr."] = sr_current
                        set_project_sequence(seq.sort_values("Sr.").reset_index(drop=True))
                        st.rerun()
        with col_ctrl4:
            if st.button("\u2705 Mark Complete", key="grn_complete"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = True
                    set_project_sequence(seq); st.rerun()
        with col_ctrl5:
            if st.button("\u21a9\ufe0f Unmark", key="grn_unmark"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = False
                    set_project_sequence(seq); st.rerun()

        st.markdown("---")
        edited_sequence = st.data_editor(
            display_sequence, num_rows="dynamic", width='stretch', height=500,
            column_config={
                "Sr.": st.column_config.NumberColumn("Sr.", min_value=1, step=1, width="small"),
                "Project Name": st.column_config.TextColumn("Project Name", width="large"),
                "No. Of Cabinet": st.column_config.NumberColumn("No. Of Cabinet", min_value=0, step=1),
                "Prod. Open Dt.": st.column_config.TextColumn("Prod. Open Dt."),
                "SO No.": st.column_config.NumberColumn("SO No.", format="%d"),
                "Completed": st.column_config.CheckboxColumn("Completed", width="small"),
            },
            key="grn_seq_editor",
        )

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("\u2705 Apply Changes", type="primary", key="grn_apply"):
                if show_completed:
                    updated = edited_sequence.sort_values("Sr.").reset_index(drop=True)
                else:
                    completed_items = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == True]
                    updated = pd.concat([edited_sequence, completed_items], ignore_index=True).sort_values("Sr.").reset_index(drop=True)
                set_project_sequence(updated); st.rerun()
        with col_btn2:
            if st.button("\U0001F504 Reset to Default", key="grn_reset"):
                set_project_sequence(DEFAULT_PROJECT_SEQUENCE.copy()); st.rerun()
        with col_btn3:
            if st.button("\U0001F5D1\uFE0F Remove Completed", key="grn_remove"):
                seq = st.session_state["project_sequence"]
                seq = seq[seq["Completed"] == False].copy()
                seq["Sr."] = range(1, len(seq) + 1)
                set_project_sequence(seq); st.rerun()

        seq_projects = set(project_sequence["Project Name"].apply(normalize_project_key))
        oob_projects = [p for p in df_oob_filtered["Project Name"].dropna().unique() if str(p) != "nan"]
        oob_project_lookup = {normalize_project_key(p): p for p in oob_projects if normalize_project_key(p)}
        missing_keys = set(oob_project_lookup.keys()) - seq_projects
        missing_from_seq = sorted([oob_project_lookup[k] for k in missing_keys])
        if missing_from_seq:
            st.markdown("---")
            st.warning(f"**{len(missing_from_seq)} Orderbook project(s) not in priority sequence** (they will appear last in GRN view):")
            for p in sorted(missing_from_seq):
                st.markdown(f"- {p}")


# ##############################################################################
#  MODE: COMBINED VIEW (Stock + GRN)
# ##############################################################################
elif analysis_mode == "\U0001F4CA Combined View (Stock + GRN)":

    # -- KPI Cards (Combined) -------------------------------------------------
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

    # -- Tabs (Combined) -------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "\U0001F3D7\uFE0F Project-wise Summary",
        "\U0001F4E5 GRN by Project Priority",
        "\U0001F50D Component Reconciliation",
        "\U0001F4CA Analytics",
        "\u2699\uFE0F Project Priority Sequence",
    ])

    # TAB 1: Project-wise Summary (Combined)
    with tab1:
        st.subheader("Project-wise Material Requirement & Receipt Summary")
        st.caption(f"GRN Date: {selected_date.strftime('%d-%b-%Y')} | Stock file loaded \u2705")

        proj_summary = build_project_summary(df_oob_filtered, df_grn_today, stock_map)

        display_cols = [
            "Project Num", "Project Name", "Unique Orders", "Unique Items (FG)",
            "Unique Components", "Components Available",
            "Required Quantity", "Quantity Issued", "Open Qty", "Fulfillment %",
            "Stock Qty", "Components Received Today", "Qty Received Today",
        ]
        proj_display = proj_summary[display_cols].copy()

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
            "Stock Qty": "{:,.0f}",
            "Qty Received Today": "{:,.0f}",
        })

        st.dataframe(styled, width='stretch', height=500)

        csv_proj = proj_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Project Summary", csv_proj, "project_summary.csv", "text/csv")

        st.markdown("---")
        st.subheader("\U0001F50E Drill-down: Project Component Detail")
        selected_proj_detail = st.selectbox("Select a project to view details:", all_projects)

        if selected_proj_detail:
            proj_detail = df_oob_filtered[df_oob_filtered["Project Name"] == selected_proj_detail]
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

            proj_comp_detail["Stock Qty"] = proj_comp_detail["Component Code"].map(stock_map).fillna(0)
            proj_comp_detail["Received Today"] = proj_comp_detail["Component Code"].map(grn_qty_map).fillna(0)
            proj_comp_detail["Net Pending"] = proj_comp_detail["Open Qty"] - proj_comp_detail["Received Today"]
            proj_comp_detail["Status"] = np.where(
                proj_comp_detail["Net Pending"] <= 0, "\u2705 Fulfilled",
                np.where(proj_comp_detail["Received Today"] > 0, "\U0001F536 Partial", "\u274C Pending")
            )
            proj_comp_detail = proj_comp_detail.sort_values("Net Pending", ascending=False)

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Total Components", len(proj_comp_detail))
            col_b.metric("Received Today", int(proj_comp_detail["Received Today"].gt(0).sum()))
            col_c.metric("Still Pending", int(proj_comp_detail["Net Pending"].gt(0).sum()))

            st.dataframe(proj_comp_detail, width='stretch', height=400)

    # TAB 2: GRN by Project Priority (Combined)
    with tab2:
        st.subheader("Today's GRN Items \u2014 Ordered by Project Priority")
        st.caption(f"GRN Date: {selected_date.strftime('%d-%b-%Y')} | Components cascade to projects by priority. Remaining supply after each allocation can fulfill lower-priority projects")

        grn_by_proj = build_grn_by_project_sequence(
            df_oob_filtered, df_grn_today, project_sequence, stock_map
        )

        if len(grn_by_proj) == 0:
            st.warning("No GRN items matched with Orderbook components for the selected date.")
        else:
            m1, m2, m3 = st.columns(3)
            m1.metric("Projects Impacted", grn_by_proj["Project Name"].nunique())
            m2.metric("Unique Components Received", grn_by_proj["Component Code"].nunique())
            m3.metric("Total GRN Qty", f"{grn_by_proj['GRN Qty Today'].sum():,.0f}")

            st.markdown("---")

            # Add Serial No. starting from 1
            grn_by_proj_display = grn_by_proj.copy()
            grn_by_proj_display = grn_by_proj_display.reset_index(drop=True)
            grn_by_proj_display.insert(0, "Serial No.", range(1, len(grn_by_proj_display) + 1))

            # Add comma-separated fulfillable work orders per (project, component)
            grn_delivered_items = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
            grn_qty_map_wo = grn_delivered_items.groupby("Item")["Qty"].sum().to_dict()
            wo_map_combined, wo_proj_map_combined, wo_proj_order_map_combined = get_fulfillable_wo_map(
                df_oob_filtered, project_sequence, stock_map=stock_map, grn_qty_map=grn_qty_map_wo
            )
            grn_by_proj_display["Fulfillable Work Orders"] = grn_by_proj_display.apply(
                lambda r: wo_map_combined.get((r["Project Name"], r["Component Code"]), ""), axis=1
            )

            grn_proj_display_cols = [
                "Serial No.", "Priority", "Project Name", "Order Number", "Component Code",
                "Component Desc", "Required Quantity", "Quantity Issued",
                "Open Qty", "Stock Qty", "GRN Qty Today", "Supplier",
                "GRN Order Number", "Fulfillable Work Orders",
            ]
            available_display_cols = [c for c in grn_proj_display_cols if c in grn_by_proj_display.columns]

            def color_priority(val):
                if val <= 5:
                    return "background-color: #f8d7da"
                elif val <= 15:
                    return "background-color: #fff3cd"
                elif val <= 30:
                    return "background-color: #d4edda"
                else:
                    return ""

            grn_styled = grn_by_proj_display[available_display_cols].copy()
            style_method2 = getattr(grn_styled.style, "map", None) or grn_styled.style.applymap
            grn_styled_df = style_method2(
                color_priority, subset=["Priority"]
            ).format({
                "Required Quantity": "{:,.0f}",
                "Quantity Issued": "{:,.0f}",
                "Open Qty": "{:,.0f}",
                "Stock Qty": "{:,.0f}",
                "GRN Qty Today": "{:,.0f}",
            })

            display_dataframe_arrow_safe(grn_styled_df, width='stretch', height=600)

            csv_grn_proj = grn_by_proj_display[available_display_cols].to_csv(index=False)
            st.download_button(
                "\U0001F4E5 Download GRN by Project Priority",
                csv_grn_proj, "grn_by_project_priority.csv", "text/csv"
            )

            # Project-wise Work Order Summary
            if wo_proj_map_combined:
                st.markdown("---")
                st.subheader("\U0001F4CB Fulfillable Work Orders by Project")
                st.caption("Work Orders cascade by priority while stock+GRN supply remains. Partial fulfillment included. Sorted by Job Start Date.")
                wo_proj_norm = {normalize_project_key(k): v for k, v in wo_proj_map_combined.items()}
                wo_proj_order_norm = {normalize_project_key(k): v for k, v in wo_proj_order_map_combined.items()}
                proj_wo_rows = []
                for proj_name in project_sequence["Project Name"].str.strip():
                    proj_key = normalize_project_key(proj_name)
                    if proj_key in wo_proj_norm:
                        proj_wo_rows.append({
                            "Project Name": proj_name,
                            "Fulfillable Work Orders": wo_proj_norm[proj_key],
                            "Order Number": wo_proj_order_norm.get(proj_key, ""),
                        })
                if proj_wo_rows:
                    proj_wo_df = pd.DataFrame(proj_wo_rows)
                    proj_wo_df.insert(0, "Sr.", range(1, len(proj_wo_df) + 1))
                    st.dataframe(proj_wo_df, width='stretch', height=min(400, 35 * len(proj_wo_df) + 38))

            grn_delivered = df_grn_today[df_grn_today["GRN_Status"].str.contains("Deliver", case=False, na=False)]
            matched_items = set(grn_by_proj["Component Code"].unique())
            unmatched_grn = grn_delivered[~grn_delivered["Item"].isin(matched_items)]
            if len(unmatched_grn) > 0:
                st.markdown("---")
                st.subheader("\u26A0\uFE0F GRN Items Not in Orderbook")
                st.caption("These items were received today but don't match any Component Code in the Orderbook")
                grn_detail_cols = ["Item", "Item Description", "Qty", "Supplier", "Order Number"]
                avail_cols = [c for c in grn_detail_cols if c in unmatched_grn.columns]
                display_dataframe_arrow_safe(unmatched_grn[avail_cols], width='stretch', height=300)

    # TAB 3: Component Reconciliation (Combined)
    with tab3:
        st.subheader("Component-Level Reconciliation: Open Qty vs Today's GRN")
        st.caption("Required Quantity \u2212 Quantity Issued = Open Qty \u2192 compared with GRN Received Today | Stock from uploaded file")

        reconciliation = build_component_reconciliation(df_oob_filtered, df_grn_today, stock_map)

        reconciliation["Availability Status"] = np.where(
            reconciliation["Stock Qty"] >= reconciliation["Open Qty"],
            "\u2705 Surplus / Available",
            "\u274C Shortage"
        )

        filter_col1, filter_col2 = st.columns(2)
        with filter_col1:
            status_filter = st.multiselect(
                "Filter by GRN Status:",
                options=["\u2705 Fulfilled", "\U0001F536 Partially Received", "\u274C Not Received"],
                default=["\U0001F536 Partially Received", "\u274C Not Received"]
            )
        with filter_col2:
            avail_filter = st.multiselect(
                "Filter by Stock vs Open Qty:",
                options=["\u2705 Surplus / Available", "\u274C Shortage"],
                default=[],
                help="Surplus/Available = Stock Qty >= Open Qty; Shortage = Stock Qty < Open Qty"
            )

        recon_display = reconciliation.copy()
        if status_filter:
            recon_display = recon_display[recon_display["Status"].isin(status_filter)]
        if avail_filter:
            recon_display = recon_display[recon_display["Availability Status"].isin(avail_filter)]

        recon_cols = [
            "Component Code", "Component Desc", "Required Quantity", "Quantity Issued",
            "Open Qty", "Stock Qty", "Received Today", "Still Pending",
            "Total Available", "Supplier", "Status", "Availability Status"
        ]
        available_recon_cols = [c for c in recon_cols if c in recon_display.columns]

        display_dataframe_arrow_safe(
            recon_display[available_recon_cols].sort_values("Still Pending", ascending=False),
            width='stretch',
            height=500
        )

        csv_recon = recon_display.to_csv(index=False)
        st.download_button("\U0001F4E5 Download Reconciliation", csv_recon, "component_reconciliation.csv", "text/csv")

        st.markdown("---")
        rc1, rc2, rc3, rc4, rc5 = st.columns(5)
        fulfilled_count = (reconciliation["Status"] == "\u2705 Fulfilled").sum()
        partial_count = (reconciliation["Status"] == "\U0001F536 Partially Received").sum()
        pending_count = (reconciliation["Status"] == "\u274C Not Received").sum()
        surplus_count = (reconciliation["Availability Status"] == "\u2705 Surplus / Available").sum()
        shortage_count = (reconciliation["Availability Status"] == "\u274C Shortage").sum()
        rc1.metric("\u2705 Fulfilled", fulfilled_count)
        rc2.metric("\U0001F536 Partially Received", partial_count)
        rc3.metric("\u274C Not Received", pending_count)
        rc4.metric("\u2705 Surplus / Available", surplus_count)
        rc5.metric("\u274C Shortage (Stock < Open)", shortage_count)

    # TAB 4: Analytics (Combined)
    with tab4:
        st.subheader("\U0001F4CA Visual Analytics (Combined View)")

        # Build full project summary
        proj_summary_chart = build_project_summary(df_oob_filtered, df_grn_today, stock_map)
        proj_open_components = df_oob_filtered[df_oob_filtered["Open Qty"] > 0].groupby(
            ["Project Num", "Project Name"], as_index=False
        )["Component Code"].nunique().rename(columns={"Component Code": "Open Components"})
        proj_summary_chart = proj_summary_chart.merge(
            proj_open_components, on=["Project Num", "Project Name"], how="left"
        )
        proj_summary_chart["Open Components"] = proj_summary_chart["Open Components"].fillna(0).astype(int)
        proj_summary_chart["Component Fulfillment"] = np.where(
            proj_summary_chart["Unique Components"] > 0,
            (((proj_summary_chart["Unique Components"] - proj_summary_chart["Open Components"]) / proj_summary_chart["Unique Components"]) * 100).round(1),
            100
        )
        proj_summary_chart["Open/Total Label"] = (
            proj_summary_chart["Open Components"].astype(str) + "/" + proj_summary_chart["Unique Components"].astype(str)
        )
        proj_open_qty = df_oob_filtered.groupby("Project Name", as_index=False)["Open Qty"].sum().rename(columns={"Open Qty": "Total Open Qty"})
        proj_summary_chart = proj_summary_chart.merge(proj_open_qty, on="Project Name", how="left")
        proj_summary_chart["Total Open Qty"] = proj_summary_chart["Total Open Qty"].fillna(0)

        # Filters
        st.markdown("##### \U0001F50D Filters & Sorting")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            sort_option_c = st.selectbox("Sort by:", [
                "Open Components (desc)", "Open Components (asc)",
                "Fulfillment % (asc)", "Fulfillment % (desc)",
                "Total Open Qty (desc)", "Total Open Qty (asc)",
                "Project Name (A-Z)",
            ], key="combined_analytics_sort")
        with fc2:
            all_projects_c = proj_summary_chart["Project Name"].unique().tolist()
            selected_projects_c = st.multiselect(
                "Filter by Project:", options=all_projects_c, default=[], key="combined_analytics_proj_filter",
                placeholder="All projects shown"
            )
        with fc3:
            fulfillment_range_c = st.slider(
                "Fulfillment % range:", 0, 100, (0, 100), key="combined_analytics_fulfill_range"
            )

        # Apply filters
        proj_filtered = proj_summary_chart.copy()
        if selected_projects_c:
            proj_filtered = proj_filtered[proj_filtered["Project Name"].isin(selected_projects_c)]
        proj_filtered = proj_filtered[
            (proj_filtered["Component Fulfillment"] >= fulfillment_range_c[0]) &
            (proj_filtered["Component Fulfillment"] <= fulfillment_range_c[1])
        ]

        # Apply sorting
        sort_map = {
            "Open Components (desc)": ("Open Components", False),
            "Open Components (asc)": ("Open Components", True),
            "Fulfillment % (asc)": ("Component Fulfillment", True),
            "Fulfillment % (desc)": ("Component Fulfillment", False),
            "Total Open Qty (desc)": ("Total Open Qty", False),
            "Total Open Qty (asc)": ("Total Open Qty", True),
            "Project Name (A-Z)": ("Project Name", True),
        }
        sort_col, sort_asc = sort_map.get(sort_option_c, ("Open Components", False))
        proj_filtered = proj_filtered.sort_values(sort_col, ascending=sort_asc)

        st.caption(f"Showing {len(proj_filtered)} of {len(proj_summary_chart)} projects")

        chart_height = max(400, len(proj_filtered) * 28 + 100)

        col1, col2 = st.columns(2)

        with col1:
            fig1 = px.bar(
                proj_filtered, x="Open Components", y="Project Name", orientation="h",
                title=f"All Projects by Open Components ({len(proj_filtered)} projects)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Open Components": "Unique Components with Open Qty", "Project Name": "Project"},
                text="Open/Total Label",
            )
            fig1.update_traces(textposition="outside")
            fig1.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig1, width='stretch')

        with col2:
            proj_filtered_text = proj_filtered.copy()
            proj_filtered_text["Label"] = proj_filtered_text["Component Fulfillment"].apply(lambda x: f"{x:.1f}%")
            fig2 = px.bar(
                proj_filtered_text, x="Component Fulfillment", y="Project Name", orientation="h",
                title="Component Fulfillment % (Fulfilled / Total Unique Components)",
                color="Component Fulfillment", color_continuous_scale="RdYlGn",
                labels={"Component Fulfillment": "Fulfillment %", "Project Name": "Project"},
                text="Label",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(yaxis=dict(autorange="reversed"), height=chart_height)
            st.plotly_chart(fig2, width='stretch')

        # Downloadable table
        st.markdown("---")
        st.subheader("Project Analytics Table")
        analytics_table_c = proj_filtered[["Project Name", "Unique Components", "Open Components",
                                           "Component Fulfillment", "Total Open Qty"]].reset_index(drop=True)
        analytics_table_c.index = analytics_table_c.index + 1
        analytics_table_c.index.name = "Sr."
        st.dataframe(analytics_table_c, width='stretch', height=min(500, 40 + len(analytics_table_c) * 35))
        st.download_button("\U0001F4E5 Download Analytics Table", analytics_table_c.to_csv(), "combined_analytics.csv", "text/csv", key="dl_combined_analytics")

        st.markdown("---")

        col3, col4 = st.columns(2)

        with col3:
            recon_full = build_component_reconciliation(df_oob_filtered, df_grn_today, stock_map)
            status_counts = recon_full["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            fig3 = px.pie(
                status_counts, values="Count", names="Status",
                title="Component Reconciliation Status",
                color_discrete_map={
                    "\u2705 Fulfilled": "#2ecc71",
                    "\U0001F536 Partially Received": "#f39c12",
                    "\u274C Not Received": "#e74c3c",
                }
            )
            st.plotly_chart(fig3, width='stretch')

        with col4:
            if len(grn_delivered_today) > 0:
                top_suppliers = grn_delivered_today.groupby("Supplier")["Qty"].sum().nlargest(10).reset_index()
                fig4 = px.bar(
                    top_suppliers, x="Qty", y="Supplier", orientation="h",
                    title=f"Top 10 Suppliers by Qty Delivered ({selected_date.strftime('%d-%b')})",
                    color="Qty", color_continuous_scale="Blues",
                )
                fig4.update_layout(yaxis=dict(autorange="reversed"), height=400)
                st.plotly_chart(fig4, width='stretch')
            else:
                st.info("No deliveries for the selected date.")

        st.markdown("---")
        st.subheader("Component Availability Breakdown")
        avail_counts = df_oob_filtered["Availability"].value_counts().reset_index()
        avail_counts.columns = ["Availability", "Count"]
        fig5 = px.pie(avail_counts, values="Count", names="Availability",
                      title="Availability Status of All Components in Orderbook",
                      color_discrete_map={"Available": "#2ecc71", "Shortage": "#e74c3c"})
        st.plotly_chart(fig5, width='stretch')

        if stock_map:
            st.markdown("---")
            st.subheader("\U0001F4E6 Stock Summary")
            sc1, sc2 = st.columns(2)
            with sc1:
                st.metric("Unique Items in Stock", len(stock_map))
                st.metric("Total Stock Qty", f"{sum(stock_map.values()):,.0f}")
            with sc2:
                oob_components = set(df_oob_filtered["Component Code"].unique())
                stock_matching = {k: v for k, v in stock_map.items() if k in oob_components}
                st.metric("Stock Items Matching Orderbook", len(stock_matching))
                st.metric("Matching Stock Qty", f"{sum(stock_matching.values()):,.0f}")

    # TAB 5: Project Priority Sequence (Combined)
    with tab5:
        st.subheader("\u2699\uFE0F Project Priority Sequence")
        st.caption(
            "Edit the project priority order below. Lower Sr. number = higher priority. "
            "This sequence determines the order in which GRN items are displayed in the "
            "'GRN by Project Priority' tab. Changes are autosaved and persist after relaunch."
        )
        render_sequence_upload_controls("combined")

        st.markdown("""
    **Instructions:**
    - Select a project below and use \u2b06\ufe0f / \u2b07\ufe0f buttons to move it up/down in priority
    - Edit cells directly in the table (double-click a cell)
    - Add new rows using the **+** button at the bottom
    - Delete rows by selecting and pressing Delete
    - Mark projects as completed to hide them from the list
    """)

        if "Completed" not in st.session_state["project_sequence"].columns:
            st.session_state["project_sequence"]["Completed"] = False

        show_completed = st.checkbox("Show completed projects", value=False)
        if show_completed:
            display_sequence = st.session_state["project_sequence"].copy()
        else:
            display_sequence = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == False].copy()

        col_ctrl1, col_ctrl2, col_ctrl3, col_ctrl4, col_ctrl5 = st.columns([3, 1, 1, 1, 1])
        with col_ctrl1:
            project_names = display_sequence["Project Name"].tolist()
            selected_project = st.selectbox("Select project to move:", project_names, key="proj_select")

        with col_ctrl2:
            if st.button("\u2b06\ufe0f Move Up"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx > 0:
                        sr_current = seq.at[idx, "Sr."]
                        sr_above = seq.at[idx-1, "Sr."]
                        seq.at[idx, "Sr."] = sr_above
                        seq.at[idx-1, "Sr."] = sr_current
                        seq = seq.sort_values("Sr.").reset_index(drop=True)
                        set_project_sequence(seq)
                        st.rerun()

        with col_ctrl3:
            if st.button("\u2b07\ufe0f Move Down"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    if idx < len(seq) - 1:
                        sr_current = seq.at[idx, "Sr."]
                        sr_below = seq.at[idx+1, "Sr."]
                        seq.at[idx, "Sr."] = sr_below
                        seq.at[idx+1, "Sr."] = sr_current
                        seq = seq.sort_values("Sr.").reset_index(drop=True)
                        set_project_sequence(seq)
                        st.rerun()

        with col_ctrl4:
            if st.button("\u2705 Mark Complete"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = True
                    set_project_sequence(seq)
                    st.success(f"Marked '{selected_project}' as completed!")
                    st.rerun()

        with col_ctrl5:
            if st.button("\u21a9\ufe0f Unmark"):
                if selected_project:
                    seq = st.session_state["project_sequence"]
                    idx = seq[seq["Project Name"] == selected_project].index[0]
                    seq.at[idx, "Completed"] = False
                    set_project_sequence(seq)
                    st.success(f"Unmarked '{selected_project}'!")
                    st.rerun()

        st.markdown("---")

        edited_sequence = st.data_editor(
            display_sequence,
            num_rows="dynamic",
            width='stretch',
            height=500,
            column_config={
                "Sr.": st.column_config.NumberColumn("Sr.", min_value=1, step=1, width="small"),
                "Project Name": st.column_config.TextColumn("Project Name", width="large"),
                "No. Of Cabinet": st.column_config.NumberColumn("No. Of Cabinet", min_value=0, step=1),
                "Prod. Open Dt.": st.column_config.TextColumn("Prod. Open Dt."),
                "SO No.": st.column_config.NumberColumn("SO No.", format="%d"),
                "Completed": st.column_config.CheckboxColumn("Completed", width="small"),
            },
            key="seq_editor",
        )

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("\u2705 Apply Changes", type="primary"):
                if show_completed:
                    updated = edited_sequence.sort_values("Sr.").reset_index(drop=True)
                else:
                    completed_items = st.session_state["project_sequence"][st.session_state["project_sequence"]["Completed"] == True]
                    updated = pd.concat([edited_sequence, completed_items], ignore_index=True)
                    updated = updated.sort_values("Sr.").reset_index(drop=True)
                set_project_sequence(updated)
                st.success("Priority sequence updated! Switch to 'GRN by Project Priority' tab to see the new order.")
                st.rerun()

        with col_btn2:
            if st.button("\U0001F504 Reset to Default"):
                set_project_sequence(DEFAULT_PROJECT_SEQUENCE.copy())
                st.success("Reset to default sequence!")
                st.rerun()

        with col_btn3:
            if st.button("\U0001F5D1\uFE0F Remove Completed"):
                seq = st.session_state["project_sequence"]
                seq = seq[seq["Completed"] == False].copy()
                seq["Sr."] = range(1, len(seq) + 1)
                set_project_sequence(seq)
                st.success("Removed all completed projects!")
                st.rerun()

            seq_projects = set(project_sequence["Project Name"].apply(normalize_project_key))
            oob_projects = [p for p in df_oob_filtered["Project Name"].dropna().unique() if str(p) != "nan"]
            oob_project_lookup = {normalize_project_key(p): p for p in oob_projects if normalize_project_key(p)}
            missing_keys = set(oob_project_lookup.keys()) - seq_projects
            missing_from_seq = sorted([oob_project_lookup[k] for k in missing_keys])
        if missing_from_seq:
            st.markdown("---")
            st.warning(
                f"**{len(missing_from_seq)} Orderbook project(s) not in priority sequence** "
                f"(they will appear last in GRN view):"
            )
            for p in sorted(missing_from_seq):
                st.markdown(f"- {p}")


# ##############################################################################
#  MODE: Supply Eligible ANALYSIS
# ##############################################################################
elif analysis_mode == "\U0001F4CB Supply Eligible Analysis":

    if len(df_oob_supply_open_filtered) == 0:
        st.warning("No rows with Sales Status = 'Supply Eligible' found in the Orderbook.")
    else:
        so_data = df_oob_supply_open_filtered.copy()

        # -- KPI Cards (Supply Eligible) -------------------------------------------
        so_total_projects = so_data["Project Name"].nunique()
        so_total_components = so_data["Component Code"].nunique()
        so_total_open_qty = so_data["Open Qty"].sum()
        so_total_stock_qty = sum(stock_map.get(c, 0) for c in so_data["Component Code"].unique())
        so_comp_in_stock = sum(1 for c in so_data["Component Code"].unique() if stock_map.get(c, 0) > 0)
        so_overall_qty_pct = min(100.0, (so_total_stock_qty / so_total_open_qty * 100)) if so_total_open_qty > 0 else 100.0
        so_overall_comp_pct = (so_comp_in_stock / so_total_components * 100) if so_total_components > 0 else 100.0

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.markdown(f'<div class="metric-card project"><h2>{so_total_projects}</h2><p>Supply Eligible Projects</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-card"><h2>{so_total_components}</h2><p>Unique Components</p></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card shortage"><h2>{so_total_open_qty:,.0f}</h2><p>Total Open Qty</p></div>', unsafe_allow_html=True)
        with c4:
            st.markdown(f'<div class="metric-card available"><h2>{so_total_stock_qty:,.0f}</h2><p>Stock Qty (Matching)</p></div>', unsafe_allow_html=True)
        with c5:
            st.markdown(f'<div class="metric-card incoming"><h2>{so_overall_qty_pct:.1f}%</h2><p>Qty Fulfillment</p></div>', unsafe_allow_html=True)
        with c6:
            st.markdown(f'<div class="metric-card incoming"><h2>{so_overall_comp_pct:.1f}%</h2><p>Component Coverage</p></div>', unsafe_allow_html=True)

        st.markdown("---")

        # -- Tabs (Supply Eligible mode) -------------------------------------------
        tab_so1, tab_so2, tab_so3 = st.tabs([
            "\U0001F4CB Fulfillment Summary",
            "\U0001F4CA Visual Analytics",
            "\U0001F50D Project Search",
        ])

        # ==================================================================
        #  TAB SO1: Summary
        # ==================================================================
        with tab_so1:
            st.subheader("\U0001F4CB Supply Eligible — Fulfillment Summary")
            st.caption("Projects with Sales Status = 'Supply Eligible' compared against current stock.")

            so_df = so_data.copy()

            # --- Build project-wise Supply Eligible summary -----------------------
            so_proj = so_df.groupby(["Project Num", "Project Name"], as_index=False).agg({
                "Component Code": "nunique",
                "Required Quantity": "sum",
                "Quantity Issued": "sum",
                "Open Qty": "sum",
                "Order Number": "nunique",
            }).rename(columns={
                "Component Code": "Unique Components",
                "Order Number": "Unique Orders",
            })

            so_proj_comp = so_df.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
            so_proj_comp.columns = ["Project Num", "Project Name", "Component Set"]
            so_proj_comp["Stock Qty"] = so_proj_comp["Component Set"].apply(
                lambda s: sum(stock_map.get(c, 0) for c in s)
            )
            so_proj_comp["Components In Stock"] = so_proj_comp["Component Set"].apply(
                lambda s: sum(1 for c in s if stock_map.get(c, 0) > 0)
            )
            so_proj = so_proj.merge(
                so_proj_comp[["Project Num", "Project Name", "Stock Qty", "Components In Stock"]],
                on=["Project Num", "Project Name"], how="left"
            )
            so_proj["Qty Fulfillment %"] = np.where(
                so_proj["Open Qty"] > 0,
                ((so_proj["Stock Qty"] / so_proj["Open Qty"]).clip(upper=1) * 100).round(1),
                100.0
            )
            so_proj["Component Fulfillment %"] = np.where(
                so_proj["Unique Components"] > 0,
                ((so_proj["Components In Stock"] / so_proj["Unique Components"]) * 100).round(1),
                100.0
            )
            so_proj = so_proj.sort_values("Open Qty", ascending=False).reset_index(drop=True)

            # --- Project-wise table -------------------------------------------
            st.markdown("##### Project-wise Supply Eligible Fulfillment")
            so_display = so_proj.copy()
            so_display.insert(0, "Sr.", range(1, len(so_display) + 1))
            so_display_cols = [
                "Sr.", "Project Num", "Project Name", "Unique Orders",
                "Unique Components", "Components In Stock",
                "Open Qty", "Stock Qty",
                "Qty Fulfillment %", "Component Fulfillment %",
            ]
            avail_cols = [c for c in so_display_cols if c in so_display.columns]

            def color_fulfillment_so(val):
                try:
                    v = float(val)
                except (ValueError, TypeError):
                    return ""
                if v >= 80: return "background-color: #d4edda"
                elif v >= 50: return "background-color: #fff3cd"
                else: return "background-color: #f8d7da"

            so_styled = so_display[avail_cols].copy()
            style_fn = getattr(so_styled.style, "map", None) or so_styled.style.applymap
            so_styled_df = style_fn(
                color_fulfillment_so, subset=["Qty Fulfillment %", "Component Fulfillment %"]
            ).format({
                "Open Qty": "{:,.0f}", "Stock Qty": "{:,.0f}",
                "Qty Fulfillment %": "{:.1f}%", "Component Fulfillment %": "{:.1f}%",
            })
            st.dataframe(so_styled_df, width='stretch', height=min(600, 38 + len(so_display) * 35))

            csv_so = so_display[avail_cols].to_csv(index=False)
            st.download_button(
                "\U0001F4E5 Download Supply Eligible Summary", csv_so,
                "supply_open_summary.csv", "text/csv", key="dl_so_summary"
            )

        # ==================================================================
        #  TAB SO2: Analytics
        # ==================================================================
        with tab_so2:
            st.subheader("\U0001F4CA Supply Eligible — Visual Analytics")

            so_df_a = so_data.copy()

            so_proj_a = so_df_a.groupby(["Project Num", "Project Name"], as_index=False).agg({
                "Component Code": "nunique",
                "Open Qty": "sum",
            }).rename(columns={"Component Code": "Unique Components"})

            so_proj_comp_a = so_df_a.groupby(["Project Num", "Project Name"])["Component Code"].apply(set).reset_index()
            so_proj_comp_a.columns = ["Project Num", "Project Name", "Component Set"]
            so_proj_comp_a["Stock Qty"] = so_proj_comp_a["Component Set"].apply(
                lambda s: sum(stock_map.get(c, 0) for c in s)
            )
            so_proj_comp_a["Components In Stock"] = so_proj_comp_a["Component Set"].apply(
                lambda s: sum(1 for c in s if stock_map.get(c, 0) > 0)
            )
            so_proj_a = so_proj_a.merge(
                so_proj_comp_a[["Project Num", "Project Name", "Stock Qty", "Components In Stock"]],
                on=["Project Num", "Project Name"], how="left"
            )
            so_proj_a["Qty Fulfillment %"] = np.where(
                so_proj_a["Open Qty"] > 0,
                ((so_proj_a["Stock Qty"] / so_proj_a["Open Qty"]).clip(upper=1) * 100).round(1),
                100.0
            )
            so_proj_a["Component Fulfillment %"] = np.where(
                so_proj_a["Unique Components"] > 0,
                ((so_proj_a["Components In Stock"] / so_proj_a["Unique Components"]) * 100).round(1),
                100.0
            )
            so_proj_a["Open Components"] = so_proj_a["Unique Components"] - so_proj_a["Components In Stock"]
            so_proj_a["Open/Total Label"] = (
                so_proj_a["Components In Stock"].astype(int).astype(str) + "/" +
                so_proj_a["Unique Components"].astype(int).astype(str)
            )

            # --- Filters & Sort -----------------------------------------------
            st.markdown("##### \U0001F50D Filters & Sorting")
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                sort_opt_so = st.selectbox("Sort by:", [
                    "Qty Fulfillment % (asc)", "Qty Fulfillment % (desc)",
                    "Component Fulfillment % (asc)", "Component Fulfillment % (desc)",
                    "Open Qty (desc)", "Open Qty (asc)",
                    "Project Name (A-Z)",
                ], key="so_analytics_sort")
            with fc2:
                all_so_projects = so_proj_a["Project Name"].unique().tolist()
                sel_so_proj = st.multiselect(
                    "Filter by Project:", options=all_so_projects, default=[],
                    key="so_analytics_proj_filter", placeholder="All projects shown"
                )
            with fc3:
                so_range = st.slider("Qty Fulfillment % range:", 0, 100, (0, 100), key="so_fulfill_range")

            so_filtered = so_proj_a.copy()
            if sel_so_proj:
                so_filtered = so_filtered[so_filtered["Project Name"].isin(sel_so_proj)]
            so_filtered = so_filtered[
                (so_filtered["Qty Fulfillment %"] >= so_range[0]) &
                (so_filtered["Qty Fulfillment %"] <= so_range[1])
            ]

            sort_map_so = {
                "Qty Fulfillment % (asc)": ("Qty Fulfillment %", True),
                "Qty Fulfillment % (desc)": ("Qty Fulfillment %", False),
                "Component Fulfillment % (asc)": ("Component Fulfillment %", True),
                "Component Fulfillment % (desc)": ("Component Fulfillment %", False),
                "Open Qty (desc)": ("Open Qty", False),
                "Open Qty (asc)": ("Open Qty", True),
                "Project Name (A-Z)": ("Project Name", True),
            }
            s_col, s_asc = sort_map_so.get(sort_opt_so, ("Qty Fulfillment %", True))
            so_filtered = so_filtered.sort_values(s_col, ascending=s_asc)

            st.caption(f"Showing {len(so_filtered)} of {len(so_proj_a)} Supply Eligible projects")

            chart_h = max(400, len(so_filtered) * 28 + 100)

            # --- Charts -------------------------------------------------------
            col_a, col_b = st.columns(2)
            with col_a:
                fig_so1 = px.bar(
                    so_filtered, x="Qty Fulfillment %", y="Project Name", orientation="h",
                    title="Qty Fulfillment % by Project (Stock vs Open Qty)",
                    color="Qty Fulfillment %", color_continuous_scale="RdYlGn",
                    text=so_filtered["Qty Fulfillment %"].apply(lambda x: f"{x:.1f}%"),
                )
                fig_so1.update_traces(textposition="outside")
                fig_so1.update_layout(yaxis=dict(autorange="reversed"), height=chart_h)
                st.plotly_chart(fig_so1, use_container_width=True)

            with col_b:
                fig_so2 = px.bar(
                    so_filtered, x="Component Fulfillment %", y="Project Name", orientation="h",
                    title="Component Coverage % (Unique Components in Stock / Total)",
                    color="Component Fulfillment %", color_continuous_scale="RdYlGn",
                    text="Open/Total Label",
                )
                fig_so2.update_traces(textposition="outside")
                fig_so2.update_layout(yaxis=dict(autorange="reversed"), height=chart_h)
                st.plotly_chart(fig_so2, use_container_width=True)

            st.markdown("---")

            col_c, col_d = st.columns(2)
            with col_c:
                fig_so3 = px.bar(
                    so_filtered, x="Open Qty", y="Project Name", orientation="h",
                    title="Open Qty by Project",
                    color="Qty Fulfillment %", color_continuous_scale="RdYlGn",
                    text=so_filtered["Open Qty"].apply(lambda x: f"{x:,.0f}"),
                )
                fig_so3.update_traces(textposition="outside")
                fig_so3.update_layout(yaxis=dict(autorange="reversed"), height=chart_h)
                st.plotly_chart(fig_so3, use_container_width=True)

            with col_d:
                so_bins = pd.cut(
                    so_filtered["Qty Fulfillment %"],
                    bins=[-1, 25, 50, 75, 100],
                    labels=["0-25%", "26-50%", "51-75%", "76-100%"]
                ).value_counts().reset_index()
                so_bins.columns = ["Range", "Projects"]
                so_bins = so_bins[so_bins["Projects"] > 0]
                fig_so4 = px.pie(
                    so_bins, values="Projects", names="Range",
                    title="Qty Fulfillment Distribution",
                    color_discrete_sequence=["#e74c3c", "#f39c12", "#f1c40f", "#2ecc71"],
                )
                st.plotly_chart(fig_so4, use_container_width=True)

            # Analytics table
            st.markdown("---")
            st.subheader("Supply Eligible Analytics Table")
            so_tbl = so_filtered[[
                "Project Name", "Unique Components", "Components In Stock",
                "Open Components", "Open Qty", "Stock Qty",
                "Qty Fulfillment %", "Component Fulfillment %"
            ]].reset_index(drop=True)
            so_tbl.index = so_tbl.index + 1
            so_tbl.index.name = "Sr."
            st.dataframe(so_tbl, width='stretch', height=min(500, 40 + len(so_tbl) * 35))
            st.download_button(
                "\U0001F4E5 Download Supply Eligible Analytics", so_tbl.to_csv(),
                "supply_open_analytics.csv", "text/csv", key="dl_so_analytics"
            )

        # ==================================================================
        #  TAB SO3: Project Search
        # ==================================================================
        with tab_so3:
            st.subheader("\U0001F50D Supply Eligible — Project Fulfillment Detail")
            st.caption("Search for a Supply Eligible project to see component-wise fulfillment capability from stock.")

            so_projects_list = sorted(so_data["Project Name"].dropna().unique())
            selected_so_project = st.selectbox(
                "Select a Supply Eligible Project:",
                options=so_projects_list,
                key="so_project_search",
                placeholder="Choose a project..."
            )

            if selected_so_project:
                proj_rows = so_data[
                    so_data["Project Name"] == selected_so_project
                ].copy()

                # Component-level aggregation
                comp_detail = proj_rows.groupby(
                    ["Component Code", "Component Desc"], as_index=False
                ).agg({
                    "Required Quantity": "sum",
                    "Quantity Issued": "sum",
                    "Open Qty": "sum",
                })
                comp_detail["Stock Qty"] = comp_detail["Component Code"].map(stock_map).fillna(0)
                comp_detail["Surplus / Deficit"] = comp_detail["Stock Qty"] - comp_detail["Open Qty"]
                comp_detail["Fulfillable Qty"] = comp_detail[["Stock Qty", "Open Qty"]].min(axis=1)
                comp_detail["Fulfillment %"] = np.where(
                    comp_detail["Open Qty"] > 0,
                    ((comp_detail["Fulfillable Qty"] / comp_detail["Open Qty"]) * 100).round(1),
                    100.0
                )
                comp_detail["Status"] = np.where(
                    comp_detail["Stock Qty"] >= comp_detail["Open Qty"],
                    "\u2705 Fully Available",
                    np.where(comp_detail["Stock Qty"] > 0, "\U0001F536 Partial", "\u274C Not Available")
                )
                comp_detail = comp_detail.sort_values("Fulfillment %", ascending=True).reset_index(drop=True)

                # --- Project KPI cards ----------------------------------------
                total_comp = comp_detail["Component Code"].nunique()
                fully_avail = (comp_detail["Status"] == "\u2705 Fully Available").sum()
                partial = (comp_detail["Status"] == "\U0001F536 Partial").sum()
                not_avail = (comp_detail["Status"] == "\u274C Not Available").sum()
                total_open = comp_detail["Open Qty"].sum()
                total_stock = comp_detail["Stock Qty"].sum()
                total_fulfillable = comp_detail["Fulfillable Qty"].sum()
                proj_qty_pct = (total_fulfillable / total_open * 100) if total_open > 0 else 100.0
                proj_comp_pct = (fully_avail / total_comp * 100) if total_comp > 0 else 100.0

                pk1, pk2, pk3, pk4 = st.columns(4)
                pk1.metric("Total Components", total_comp)
                pk2.metric("Fully Available", f"{fully_avail} ({proj_comp_pct:.1f}%)")
                pk3.metric("Partial", partial)
                pk4.metric("Not Available", not_avail)

                pk5, pk6, pk7 = st.columns(3)
                pk5.metric("Total Open Qty", f"{total_open:,.0f}")
                pk6.metric("Fulfillable Qty", f"{total_fulfillable:,.0f}")
                pk7.metric("Qty Fulfillment %", f"{proj_qty_pct:.1f}%")

                st.markdown("---")

                # --- Status filter --------------------------------------------
                status_opts = comp_detail["Status"].unique().tolist()
                sel_status = st.multiselect(
                    "Filter by Status:", options=status_opts, default=status_opts,
                    key="so_detail_status_filter"
                )
                comp_filtered = comp_detail[comp_detail["Status"].isin(sel_status)] if sel_status else comp_detail

                # --- Component table ------------------------------------------
                st.markdown(f"##### Component-wise Fulfillment for **{selected_so_project}**")
                comp_display = comp_filtered.copy()
                comp_display.insert(0, "Sr.", range(1, len(comp_display) + 1))
                display_cols = [
                    "Sr.", "Component Code", "Component Desc",
                    "Required Quantity", "Quantity Issued", "Open Qty",
                    "Stock Qty", "Fulfillable Qty", "Surplus / Deficit",
                    "Fulfillment %", "Status",
                ]
                avail_cols = [c for c in display_cols if c in comp_display.columns]

                def color_status_so(val):
                    if "\u2705" in str(val): return "background-color: #d4edda"
                    elif "\U0001F536" in str(val): return "background-color: #fff3cd"
                    elif "\u274C" in str(val): return "background-color: #f8d7da"
                    return ""

                comp_styled = comp_display[avail_cols].copy()
                style_fn = getattr(comp_styled.style, "map", None) or comp_styled.style.applymap
                comp_styled_df = style_fn(
                    color_status_so, subset=["Status"]
                ).format({
                    "Required Quantity": "{:,.0f}", "Quantity Issued": "{:,.0f}",
                    "Open Qty": "{:,.0f}", "Stock Qty": "{:,.0f}",
                    "Fulfillable Qty": "{:,.0f}", "Surplus / Deficit": "{:,.0f}",
                    "Fulfillment %": "{:.1f}%",
                })
                st.dataframe(comp_styled_df, width='stretch', height=min(600, 38 + len(comp_display) * 35))

                csv_detail = comp_display[avail_cols].to_csv(index=False)
                st.download_button(
                    f"\U0001F4E5 Download {selected_so_project} Detail", csv_detail,
                    f"supply_open_{selected_so_project}.csv", "text/csv", key="dl_so_detail"
                )

                # --- Fulfillment charts for selected project ------------------
                st.markdown("---")
                ch1, ch2 = st.columns(2)
                with ch1:
                    status_counts = comp_detail["Status"].value_counts().reset_index()
                    status_counts.columns = ["Status", "Count"]
                    fig_d1 = px.pie(
                        status_counts, values="Count", names="Status",
                        title=f"{selected_so_project} — Component Availability",
                        color_discrete_map={
                            "\u2705 Fully Available": "#2ecc71",
                            "\U0001F536 Partial": "#f39c12",
                            "\u274C Not Available": "#e74c3c",
                        }
                    )
                    st.plotly_chart(fig_d1, use_container_width=True)

                with ch2:
                    top_shortage = comp_detail[comp_detail["Surplus / Deficit"] < 0].nsmallest(15, "Surplus / Deficit")
                    if len(top_shortage) > 0:
                        fig_d2 = px.bar(
                            top_shortage, x="Surplus / Deficit", y="Component Code", orientation="h",
                            title=f"Top {len(top_shortage)} Shortage Components",
                            color="Surplus / Deficit", color_continuous_scale="Reds_r",
                            text=top_shortage["Surplus / Deficit"].apply(lambda x: f"{x:,.0f}"),
                        )
                        fig_d2.update_traces(textposition="outside")
                        fig_d2.update_layout(yaxis=dict(autorange="reversed"), height=max(350, len(top_shortage) * 28 + 100))
                        st.plotly_chart(fig_d2, use_container_width=True)
                    else:
                        st.success("No shortage components — all fully available!")


# -- Footer -------------------------------------------------------------------
st.markdown("---")
st.caption("Built for Nashik iCenter Supply Chain | Data refreshes on file upload")
