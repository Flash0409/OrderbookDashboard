# Nashik iCenter Orderbook Dashboard

Interactive web-based dashboard for analyzing project-wise material requirements, daily GRN receipts, and component shortages.

## Features

- 📦 **Project-wise Summary** - Material requirements and fulfillment status by project
- 🎯 **GRN by Project Priority** - Today's received items ordered by production priority
- 🔍 **Component Reconciliation** - Compare required vs. received quantities
- 📊 **Visual Analytics** - Charts and graphs for quick insights
- ⚙️ **Editable Priority Sequence** - Customize project priority order

## Requirements

- Python 3.7 or higher
- Internet connection (for initial setup only)

## Quick Start

### Windows

1. **Run the dashboard:**
   - Double-click `run_dashboard.bat`
   - Or open Command Prompt and run: `run_dashboard.bat`

2. **Access the dashboard:**
   - Browser opens automatically at http://localhost:8501
   - Or manually navigate to http://localhost:8501

### Mac / Linux

1. **Make the script executable (first time only):**
   ```bash
   chmod +x run_dashboard.sh
   ```

2. **Run the dashboard:**
   ```bash
   ./run_dashboard.sh
   ```
   Or:
   ```bash
   bash run_dashboard.sh
   ```

3. **Access the dashboard:**
   - Open your browser and go to http://localhost:8501

### Manual Installation

If the automatic scripts don't work:

```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# Mac/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the dashboard
streamlit run app.py --server.port 8501
```

## Usage

1. **Upload Files** (via sidebar):
   - **Orderbook (.xlsm)** - Required
   - **Daily Material Incoming (.xlsx)** - Required  
   - **Stock (.xlsx)** - Optional (for real-time inventory data)

2. **Explore Tabs:**
   - **Project-wise Summary** - See all projects and their material status
   - **GRN by Project Priority** - View today's receipts organized by priority
   - **Component Reconciliation** - Detailed component-level analysis
   - **Analytics** - Visual charts and insights
   - **Project Priority Sequence** - Edit the project priority order

3. **Filter & Analyze:**
   - Select specific GRN dates
   - Filter by project
   - Download CSV reports

## File Requirements

### Orderbook File
- Format: `.xlsm` or `.xlsx`
- Required Sheet: `OpenOrdersBOM`
- Key Columns: Component Code, Required Quantity, Quantity Issued, Project Name

### Daily Material Incoming File
- Format: `.xlsx`
- Required Sheet: `Daily GRN`
- Key Columns: Item, Qty, Date, Supplier

### Stock File (Optional)
- Format: `.xlsx`
- Required Sheet: `Stock`
- Key Columns: Item Number, On Hand Quantity

## Troubleshooting

**Dashboard won't start:**
- Ensure Python is installed: `python --version`
- Check internet connection (needed for first-time package installation)
- Try manual installation steps above

**Files won't upload:**
- Verify file formats (.xlsm, .xlsx)
- Check that required sheets exist in files
- Ensure files are not corrupted

**Data not displaying:**
- Check that Component Code (Orderbook) matches Item (GRN)
- Verify date formats in GRN file
- Look for error messages in the browser

## Customization

### Project Priority Sequence
Edit the default priority in the "Project Priority Sequence" tab, or modify `DEFAULT_PROJECT_SEQUENCE` in [app.py](app.py) (lines 58-108).

### Styling
Modify CSS in [app.py](app.py) (lines 24-40) to change colors and layout.

## Technical Details

**Built with:**
- Streamlit (web framework)
- Pandas (data processing)
- Plotly (visualizations)
- OpenPyXL (Excel file handling)

**Data Flow:**
1. Files uploaded → Cached in memory
2. Data cleaned and processed
3. Reconciliation performed
4. Results displayed in tabs
5. Session resets on browser refresh

## Support

For issues or questions, contact the Nashik iCenter Supply Chain team.

---
*Built for Nashik iCenter Supply Chain Operations*
