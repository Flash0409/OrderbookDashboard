"""
Launcher script for Streamlit dashboard.
This gets bundled into an exe by PyInstaller.
It starts the Streamlit server and opens the browser automatically.
"""

import sys
import os
import subprocess
import webbrowser
import time
import socket


def get_free_port():
    """Find a free port on localhost."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        return s.getsockname()[1]


def main():
    # Determine paths
    if getattr(sys, '_MEIPASS', None):
        # Running from PyInstaller bundle
        bundle_dir = sys._MEIPASS
    else:
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    app_path = os.path.join(bundle_dir, "app.py")
    port = get_free_port()

    print(f"Starting iCenter Orderbook Dashboard on port {port}...")
    print(f"App path: {app_path}")
    print("Please wait, the browser will open automatically...")

    # Use streamlit's CLI directly
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit", "run",
        app_path,
        f"--server.port={port}",
        "--server.headless=true",
        f"--browser.serverAddress=localhost",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
        "--global.developmentMode=false",
    ]

    # Open browser after a short delay
    def open_browser():
        time.sleep(3)
        webbrowser.open(f"http://localhost:{port}")

    import threading
    threading.Thread(target=open_browser, daemon=True).start()

    stcli.main()


if __name__ == "__main__":
    main()
