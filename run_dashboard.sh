#!/bin/bash

echo "=========================================="
echo " Nashik iCenter Orderbook Dashboard"
echo "=========================================="
echo

# Check if virtual environment exists, create if not
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    echo
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate

# Check if streamlit is installed in venv
if ! .venv/bin/pip show streamlit > /dev/null 2>&1; then
    echo "Installing required packages..."
    .venv/bin/pip install -r requirements.txt
    echo
fi

echo "Starting Dashboard..."
echo "Open your browser at http://localhost:8501"
echo "Press Ctrl+C to stop the server"
echo

.venv/bin/streamlit run app.py --server.port 8501
