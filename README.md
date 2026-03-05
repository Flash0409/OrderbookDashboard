# Nashik iCenter Orderbook Dashboard

Interactive Streamlit dashboard for project-wise material tracking using Orderbook, GRN, and Stock files.

See [`CHANGELOG.md`](CHANGELOG.md) for a versioned history of updates.

## What's new (latest update)

- Added **Supply Eligible Analysis** mode (Orderbook rows with `Sales Status = Supply Eligible`).
- Enhanced **Project Priority Sequence** management:
   - Edit sequence directly in-app.
   - Upload sequence from `.xlsx`, `.xlsm`, or `.csv`.
   - Apply as **Replace** or **Append / Merge**.
   - Sequence now auto-persists to `project_sequence_state.json`.
- Improved file ingestion with **sheet/header auto-detection** for Orderbook, GRN, and Stock files.
- Added **priority-based cascading allocation views** for stock and GRN by project.
- Expanded analytics and drill-down tables with safer dataframe rendering for mixed data types.

## What this dashboard does

- Tracks **open material demand** from Orderbook.
- Maps **today’s GRN receipts** against pending component demand.
- Shows **stock vs open quantity** shortage/surplus.
- Orders project/component views using a **customizable project priority sequence**.
- Exports key views as CSV for daily reporting.

## Analysis modes

The sidebar provides four modes:

1. **📦 Stock Analysis**
   - Required files: Orderbook + Stock
2. **📥 GRN Analysis**
   - Required files: Orderbook + Daily Material Incoming (GRN)
3. **📊 Combined View (Stock + GRN)**
   - Required files: Orderbook + Stock + Daily Material Incoming (GRN)
4. **📋 Supply Eligible Analysis**
   - Required files: Orderbook + Stock
   - Uses Orderbook rows with `Sales Status = Supply Eligible`

## Key features

- **Project-wise Summary**
- **GRN by Project Priority**
- **Stock by Project Priority**
- **Component Reconciliation**
- **Analytics charts and filters**
- **Editable Project Priority Sequence** (with upload/merge/replace + autosave support)
- **CSV download** on major tables

## Requirements

- Python 3.9+
- Internet connection only for first-time dependency installation

## Quick start

### Windows

1. Run `run_dashboard.bat` (double-click or from terminal).
2. Open http://localhost:8501 if browser does not auto-open.

### macOS / Linux

1. Make script executable (first run only):
   ```bash
   chmod +x run_dashboard.sh
   ```
2. Start dashboard:
   ```bash
   ./run_dashboard.sh
   ```
3. Open http://localhost:8501.

### Manual run

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py --server.port 8501
```

## Input file expectations

The app detects sheets/header rows by required columns, so fixed sheet names are not mandatory.

### 1) Orderbook (`.xlsm` / `.xlsx`) — always required

Minimum required columns:
- `Component Code`
- `Required Quantity`
- `Quantity Issued`
- `Project Num`
- `Project Name`

Commonly used additional columns:
- `Sales Status` (for Production Open / Supply Eligible filtering)
- `Order Number`
- `Work Order Number`
- `Component Desc`
- `Availability`
- `Job Start Date`

### 2) Daily Material Incoming / GRN (`.xlsx` / `.xlsm`) — required in GRN and Combined modes

Minimum required columns:
- `Item`
- `Qty`
- `Date`

Commonly used additional columns:
- `Supplier`
- `Order Number`
- Status field (e.g., Deliver/Reject markers)

### 3) Stock (`.xlsx` / `.xlsm`) — required in Stock, Combined, and Supply Eligible modes

Minimum required columns:
- `Item Number`
- `On Hand Quantity`

## Typical usage flow

1. Select analysis mode in sidebar.
2. Upload required files for that mode.
3. (If GRN is uploaded) choose GRN date.
4. Optionally filter by project(s).
5. Review tabs and export CSVs.
6. Update project priority sequence if needed.
7. Sequence changes are saved automatically for the next run.

## Troubleshooting

**Dashboard does not start**
- Check Python: `python --version`
- Reinstall dependencies: `pip install -r requirements.txt`
- Start manually: `streamlit run app.py --server.port 8501`

**Upload accepted but no data shown**
- Verify required columns exist exactly (names/spaces matter).
- Ensure component keys align:
  - Orderbook `Component Code`
  - GRN `Item`
  - Stock `Item Number`
- Confirm GRN date filter is set to a date with records.

**Unexpected totals**
- Check duplicate rows in source files.
- Validate numeric fields are not text-formatted with symbols.

## Configuration notes

- Default priority sequence is defined in [app.py](app.py) under `DEFAULT_PROJECT_SEQUENCE`.
- Runtime sequence state is stored in `project_sequence_state.json`.
- Project sequence can be edited in-app and uploaded (replace or merge).

## Tech stack

- Streamlit
- Pandas
- NumPy
- Plotly
- OpenPyXL

---
Built for Nashik iCenter Supply Chain Operations.
