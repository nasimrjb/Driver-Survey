"""
Dashboard Web Server
====================
Serves both Driver Survey dashboards via HTTP and auto-regenerates them
from SQL Server on a schedule or on-demand.

Endpoints:
  /                         → Index page with links to both dashboards
  /driver-survey            → DriverSurvey Dashboard (charts)
  /routine-analysis         → Routine Analysis Dashboard (heatmap tables)
  /refresh                  → Regenerate both dashboards from SQL Server
  /refresh/driver-survey    → Regenerate DriverSurvey only
  /refresh/routine          → Regenerate Routine Analysis only

Usage:
  python serve_dashboards.py                   # Start server on port 8765
  python serve_dashboards.py --port 9000       # Custom port
  python serve_dashboards.py --refresh         # Refresh + start server
"""

import http.server
import socketserver
import os
import sys
import json
import threading
import time
import traceback
from datetime import datetime
from urllib.parse import urlparse

PORT = 8765
DASHBOARD_DIR = r"D:\Work\Driver Survey\PowerBI"
DRIVER_SURVEY_HTML = "DriverSurvey_Dashboard.html"
ROUTINE_HTML = "RoutineAnalysis_Dashboard.html"

# Track last refresh times
last_refresh = {
    "driver_survey": None,
    "routine_analysis": None,
}
refresh_lock = threading.Lock()
refresh_in_progress = False


def regenerate_driver_survey():
    """Regenerate DriverSurvey Dashboard from SQL Server views."""
    global refresh_in_progress
    print(f"[{datetime.now():%H:%M:%S}] Regenerating Driver Survey Dashboard...")
    try:
        import subprocess
        result = subprocess.run(
            [sys.executable, os.path.join(DASHBOARD_DIR, "build_dashboard.py")],
            capture_output=True, text=True, timeout=300,
            cwd=DASHBOARD_DIR
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr[-500:] if result.stderr else "Unknown error")

        last_refresh["driver_survey"] = datetime.now().isoformat()
        output_path = os.path.join(DASHBOARD_DIR, DRIVER_SURVEY_HTML)
        size_kb = os.path.getsize(output_path) / 1024
        print(f"  Done! {size_kb:.0f} KB written.")
        return True, f"Driver Survey refreshed ({size_kb:.0f} KB)"
    except Exception as e:
        msg = f"Driver Survey refresh failed: {e}"
        print(f"  ERROR: {msg}")
        traceback.print_exc()
        return False, msg


def regenerate_routine_analysis():
    """Regenerate Routine Analysis Dashboard from the latest Excel output."""
    print(f"[{datetime.now():%H:%M:%S}] Regenerating Routine Analysis Dashboard...")
    try:
        import subprocess
        # Find latest Excel file
        from pathlib import Path
        import re as _re
        files = list(Path(r"D:\Work\Driver Survey\processed").glob(
            "routine_analysis_week_*.xlsx"))
        if not files:
            raise FileNotFoundError("No routine_analysis_week_*.xlsx found")
        files.sort(key=lambda f: int(
            _re.search(r"week_(\d+)", f.name).group(1)
            if _re.search(r"week_(\d+)", f.name) else 0))
        excel_path = str(files[-1])

        result = subprocess.run(
            [sys.executable, os.path.join(DASHBOARD_DIR,
             "build_routine_dashboard.py"), excel_path],
            capture_output=True, text=True, timeout=120,
            cwd=DASHBOARD_DIR
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr[-500:] if result.stderr else "Unknown error")

        last_refresh["routine_analysis"] = datetime.now().isoformat()
        output_path = os.path.join(DASHBOARD_DIR, ROUTINE_HTML)
        size_kb = os.path.getsize(output_path) / 1024
        print(f"  Done! {size_kb:.0f} KB written.")
        return True, f"Routine Analysis refreshed ({size_kb:.0f} KB)"
    except Exception as e:
        msg = f"Routine Analysis refresh failed: {e}"
        print(f"  ERROR: {msg}")
        traceback.print_exc()
        return False, msg


def build_index_page():
    """Build the index page with links to both dashboards."""
    ds_exists = os.path.exists(os.path.join(DASHBOARD_DIR, DRIVER_SURVEY_HTML))
    ra_exists = os.path.exists(os.path.join(DASHBOARD_DIR, ROUTINE_HTML))
    ds_size = f"{os.path.getsize(os.path.join(DASHBOARD_DIR, DRIVER_SURVEY_HTML))/1024:.0f} KB" if ds_exists else "Not generated"
    ra_size = f"{os.path.getsize(os.path.join(DASHBOARD_DIR, ROUTINE_HTML))/1024:.0f} KB" if ra_exists else "Not generated"
    ds_time = last_refresh["driver_survey"] or ("File exists" if ds_exists else "Never")
    ra_time = last_refresh["routine_analysis"] or ("File exists" if ra_exists else "Never")

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Driver Survey Dashboards</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
    font-family: 'Segoe UI', Tahoma, sans-serif;
    background: linear-gradient(135deg, #1e3a5f 0%, #2c5282 50%, #1e3a5f 100%);
    min-height: 100vh;
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 40px;
}}
.container {{
    max-width: 800px;
    width: 100%;
}}
h1 {{
    color: white;
    text-align: center;
    font-size: 28px;
    margin-bottom: 8px;
}}
.subtitle {{
    color: rgba(255,255,255,0.7);
    text-align: center;
    font-size: 14px;
    margin-bottom: 40px;
}}
.card {{
    background: white;
    border-radius: 12px;
    padding: 28px;
    margin-bottom: 20px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    display: flex;
    align-items: center;
    gap: 24px;
    text-decoration: none;
    color: #333;
    transition: transform 0.15s, box-shadow 0.15s;
    cursor: pointer;
}}
.card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 30px rgba(0,0,0,0.2);
}}
.card-icon {{
    width: 64px;
    height: 64px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    flex-shrink: 0;
}}
.icon-charts {{ background: linear-gradient(135deg, #00C853, #00a844); }}
.icon-tables {{ background: linear-gradient(135deg, #3498db, #2980b9); }}
.icon-refresh {{ background: linear-gradient(135deg, #FF6D00, #e65100); }}
.card-body h2 {{
    font-size: 18px;
    margin-bottom: 4px;
    color: #1e3a5f;
}}
.card-body p {{
    font-size: 13px;
    color: #666;
    line-height: 1.4;
}}
.card-meta {{
    font-size: 11px;
    color: #999;
    margin-top: 6px;
}}
.refresh-row {{
    display: flex;
    gap: 12px;
    margin-top: 10px;
}}
.btn {{
    display: inline-block;
    padding: 8px 16px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 500;
    text-decoration: none;
    color: white;
    transition: opacity 0.15s;
}}
.btn:hover {{ opacity: 0.85; }}
.btn-green {{ background: #00C853; }}
.btn-blue {{ background: #3498db; }}
.btn-orange {{ background: #FF6D00; }}
.status {{ color: rgba(255,255,255,0.6); text-align: center; font-size: 12px; margin-top: 20px; }}
</style>
</head>
<body>
<div class="container">
    <h1>Driver Survey Dashboards</h1>
    <div class="subtitle">Cab Studies | SQL Server: 192.168.18.37 | Cab_Studies</div>

    <a class="card" href="/driver-survey">
        <div class="card-icon icon-charts">&#128202;</div>
        <div class="card-body">
            <h2>Driver Survey Dashboard</h2>
            <p>Interactive charts: satisfaction trends, NPS, incentives, ride share, demographics, and survey explorer. 46 plotly charts across 5 pages.</p>
            <div class="card-meta">Size: {ds_size} | Last refresh: {ds_time}</div>
        </div>
    </a>

    <a class="card" href="/routine-analysis">
        <div class="card-icon icon-tables">&#128203;</div>
        <div class="card-body">
            <h2>Routine Analysis Dashboard</h2>
            <p>Weekly heatmap tables: incentive amounts, satisfaction review, ride share, demographics, CS support, NPS — 45 sheets across 8 tabs with color scales.</p>
            <div class="card-meta">Size: {ra_size} | Last refresh: {ra_time}</div>
        </div>
    </a>

    <div class="refresh-row" style="justify-content:center;">
        <a class="btn btn-green" href="/refresh/driver-survey">Refresh Charts Dashboard</a>
        <a class="btn btn-blue" href="/refresh/routine">Refresh Routine Dashboard</a>
        <a class="btn btn-orange" href="/refresh">Refresh Both</a>
    </div>

    <div class="status">
        Server running on port {PORT} |
        Data source: SQL Server views (20 views, 250K+ rows)
    </div>
</div>
</body>
</html>'''


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    """Custom HTTP handler with dashboard routing and refresh endpoints."""

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        if path == "" or path == "/index":
            self._serve_html(build_index_page())

        elif path == "/driver-survey":
            self._serve_file(DRIVER_SURVEY_HTML)

        elif path == "/routine-analysis":
            self._serve_file(ROUTINE_HTML)

        elif path == "/refresh":
            self._do_refresh("both")

        elif path == "/refresh/driver-survey":
            self._do_refresh("driver_survey")

        elif path == "/refresh/routine":
            self._do_refresh("routine")

        elif path == "/api/status":
            self._serve_json({
                "status": "ok",
                "refresh_in_progress": refresh_in_progress,
                "last_refresh": last_refresh,
                "dashboards": {
                    "driver_survey": os.path.exists(
                        os.path.join(DASHBOARD_DIR, DRIVER_SURVEY_HTML)),
                    "routine_analysis": os.path.exists(
                        os.path.join(DASHBOARD_DIR, ROUTINE_HTML)),
                }
            })

        else:
            # Serve static files (CSS, JS, images, etc.)
            super().do_GET()

    def _serve_file(self, filename):
        filepath = os.path.join(DASHBOARD_DIR, filename)
        if os.path.exists(filepath):
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            with open(filepath, "rb") as f:
                self.wfile.write(f.read())
        else:
            self._serve_html(
                f"<html><body><h1>{filename} not found</h1>"
                f"<p>Run a refresh first: <a href='/refresh'>/refresh</a></p>"
                f"</body></html>",
                status=404
            )

    def _serve_html(self, html, status=200):
        self.send_response(status)
        self.send_header("Content-type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html.encode("utf-8"))

    def _serve_json(self, data):
        self.send_response(200)
        self.send_header("Content-type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode("utf-8"))

    def _do_refresh(self, target):
        global refresh_in_progress
        if refresh_in_progress:
            self._serve_html(
                "<html><body><h1>Refresh already in progress...</h1>"
                "<p><a href='/'>Back to index</a></p></body></html>"
            )
            return

        refresh_in_progress = True
        results = []
        try:
            if target in ("both", "driver_survey"):
                ok, msg = regenerate_driver_survey()
                results.append(f"{'OK' if ok else 'FAIL'}: {msg}")
            if target in ("both", "routine"):
                ok, msg = regenerate_routine_analysis()
                results.append(f"{'OK' if ok else 'FAIL'}: {msg}")
        finally:
            refresh_in_progress = False

        result_html = "<br>".join(results)
        self._serve_html(
            f"<html><body style='font-family:Segoe UI; padding:40px;'>"
            f"<h1>Refresh Complete</h1>"
            f"<p>{result_html}</p>"
            f"<p style='margin-top:20px;'>"
            f"<a href='/'>Back to index</a> | "
            f"<a href='/driver-survey'>View Charts Dashboard</a> | "
            f"<a href='/routine-analysis'>View Routine Dashboard</a>"
            f"</p></body></html>"
        )

    def log_message(self, format, *args):
        # Quieter logging
        if "/api/" not in str(args[0]) if args else True:
            print(f"[{datetime.now():%H:%M:%S}] {args[0]}" if args else "")


def main():
    global PORT

    # Parse args
    args = sys.argv[1:]
    auto_refresh = False
    for i, arg in enumerate(args):
        if arg == "--port" and i + 1 < len(args):
            PORT = int(args[i + 1])
        elif arg == "--refresh":
            auto_refresh = True

    os.chdir(DASHBOARD_DIR)

    if auto_refresh:
        print("Auto-refreshing dashboards from SQL Server...")
        regenerate_driver_survey()
        regenerate_routine_analysis()

    # Start server
    handler = DashboardHandler
    with socketserver.TCPServer(("", PORT), handler) as httpd:
        print(f"\n{'='*50}")
        print(f"  Dashboard Server running on port {PORT}")
        print(f"  http://localhost:{PORT}/")
        print(f"  http://localhost:{PORT}/driver-survey")
        print(f"  http://localhost:{PORT}/routine-analysis")
        print(f"  http://localhost:{PORT}/refresh")
        print(f"{'='*50}\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server...")
            httpd.shutdown()


if __name__ == "__main__":
    main()
