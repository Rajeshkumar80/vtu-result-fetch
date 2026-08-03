import os
import re
import sys
import time
import subprocess
import pandas as pd
from flask import Flask, render_template, send_from_directory
from flask_socketio import SocketIO, emit
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_SCRIPT = os.path.join(BASE_DIR, "bulk_fetcher_6.py")
DEFAULT_CSV = os.path.join(BASE_DIR, "students.csv")
RAW_DATA = os.path.join(BASE_DIR, "raw_results.csv")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "vtu_results.xlsx")
PY = sys.executable
USN_PATTERN = re.compile(r"^[A-Z0-9]{10}$")

current_process = None
fetch_stats = {"total": 0, "success": 0, "fail": 0}

app = Flask(__name__)
socketio = SocketIO(app, async_mode="threading", cors_allowed_origins="*")


# -----------------------------------------
# FRONTEND ROUTES
# -----------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/download")
def download():
    try:
        return send_from_directory(BASE_DIR, os.path.basename(OUTPUT_EXCEL), as_attachment=True)
    except FileNotFoundError:
        return "Excel not generated yet", 404


# -----------------------------------------
# IMPORT CSV
# -----------------------------------------
@socketio.on("import-csv")
def import_csv():
    if not os.path.exists(DEFAULT_CSV):
        emit("log-message", {"data": "[Error] students.csv missing\n"})
        return

    with open(DEFAULT_CSV) as f:
        lines = f.readlines()[1:]

    emit("csv-data", {"usns": "".join(lines)})
    emit("log-message", {"data": f"Imported {len(lines)} USNs\n"})


# -----------------------------------------
# START FETCHING
# -----------------------------------------
@socketio.on("start-fetch")
def start_fetch(msg):
    global fetch_stats, current_process

    usns = msg.get("usns", [])
    vtu_url = msg.get("url", "")

    if current_process and current_process.poll() is None:
        emit("log-message", {"data": "[Error] A fetch is already running. Stop it first.\n"})
        return

    if not usns:
        emit("log-message", {"data": "[Error] No USNs provided.\n"})
        return

    if not vtu_url:
        emit("log-message", {"data": "[Error] No VTU URL provided.\n"})
        return

    # Validate & clean USNs server-side
    clean = []
    skipped = []
    for u in usns:
        u = str(u).strip().upper()
        if USN_PATTERN.match(u):
            clean.append(u)
        else:
            skipped.append(u)

    if skipped:
        emit("log-message", {"data": f"[Warn] Skipped invalid USN entries: {', '.join(skipped)}\n"})
    if not clean:
        emit("log-message", {"data": "[Error] No valid USNs. Format e.g. 1GD23CS001\n"})
        return

    with open(DEFAULT_CSV, "w") as f:
        f.write("USN\n")
        for u in clean:
            f.write(u + "\n")

    fetch_stats = {"total": len(clean), "success": 0, "fail": 0}

    socketio.start_background_task(target=run_scraper, vtu_url=vtu_url, total_usns=len(clean))
    emit("fetch-started", {"total": len(clean)})


# -----------------------------------------
# STOP FETCHING
# -----------------------------------------
@socketio.on("stop-fetch")
def stop_fetch():
    global current_process
    if current_process and current_process.poll() is None:
        current_process.terminate()
        emit("log-message", {"data": "\n[Stopped] Fetch terminated by user.\n"})
    else:
        emit("log-message", {"data": "\n[Info] No running process to stop.\n"})


# -----------------------------------------
# RUN SCRAPER WITH URL ARG
# -----------------------------------------
def run_scraper(vtu_url, total_usns):
    global current_process, fetch_stats
    cmd = [PY, BACKEND_SCRIPT, vtu_url]
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=BASE_DIR
    )
    current_process = process

    for line in iter(process.stdout.readline, ""):
        socketio.emit("log-message", {"data": line})
        if "[Success] Scraped" in line:
            fetch_stats["success"] += 1
            socketio.emit("fetch-progress", dict(fetch_stats))
        elif "FAILED to fetch" in line:
            fetch_stats["fail"] += 1
            socketio.emit("fetch-progress", dict(fetch_stats))

    process.wait()
    current_process = None

    socketio.emit("fetch-complete", dict(fetch_stats))

    if os.path.exists(OUTPUT_EXCEL):
        socketio.emit("download-ready")


# -----------------------------------------
# SGPA CALCULATION
# -----------------------------------------
def marks_to_gp(m):
    try:
        m = int(m)
    except:
        return 0

    if m >= 90: return 10
    if m >= 80: return 9
    if m >= 70: return 8
    if m >= 60: return 7
    if m >= 50: return 6
    if m >= 45: return 5
    if m >= 40: return 4
    return 0


@socketio.on("sgpa-credits")
def sgpa_calc(data):
    credits = data.get("credits", {})

    if not os.path.exists(RAW_DATA):
        emit("log-message", {"data": "[SGPA ERROR] raw_results.csv not found\n"})
        return

    df = pd.read_csv(RAW_DATA)
    sgpa_map = {}

    for usn, group in df.groupby("USN"):
        total_pts = 0
        total_cr = 0

        for _, row in group.iterrows():
            sub = row["Subject Code"]
            if sub not in credits:
                continue

            cr = credits[sub]
            gp = marks_to_gp(row["Total Marks"])

            total_pts += gp * cr
            total_cr += cr

        sgpa = round(total_pts / total_cr, 2) if total_cr else 0
        sgpa_map[usn] = sgpa

    # Update Excel file
    excel = pd.read_excel(OUTPUT_EXCEL)
    excel["SGPA"] = excel["USN"].apply(lambda u: sgpa_map.get(u, 0))
    excel.to_excel(OUTPUT_EXCEL, index=False)

    emit("log-message", {"data": "\n[SGPA] SGPA added to Excel.\n"})
    emit("download-ready")


# -----------------------------------------
if __name__ == "__main__":
    socketio.run(app, host="127.0.0.1", port=5000, allow_unsafe_werkzeug=True, use_reloader=False)
