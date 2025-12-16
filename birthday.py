import os
import sys
import time
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# -------------------------------
# Startup delay (important)
# -------------------------------
time.sleep(15)

# -------------------------------
# Logging (silent troubleshooting)
# -------------------------------
def log(msg):
    try:
        with open("startup_log.txt", "a") as f:
            f.write(f"{datetime.now()} : {msg}\n")
    except:
        pass

# -------------------------------
# Get base path (EXE or PY)
# -------------------------------
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(base_path, "data.xlsx")

# -------------------------------
# Load Excel safely
# -------------------------------
try:
    df = pd.read_excel(excel_path)
except Exception as e:
    log(f"Excel load failed: {e}")
    sys.exit(0)   # Exit silently (no crash popup)

# -------------------------------
# Date check
# -------------------------------
today_md = datetime.now().strftime("%m-%d")
alert_messages = []

for _, row in df.iterrows():

    # Birthday
    if "Birthday" in df.columns and pd.notnull(row["Birthday"]):
        if today_md == row["Birthday"].strftime("%m-%d"):
            alert_messages.append(f"🎉 Birthday Today: {row['Name']}")

    # Anniversary (optional column)
    if "Anniversary" in df.columns and pd.notnull(row["Anniversary"]):
        if today_md == row["Anniversary"].strftime("%m-%d"):
            alert_messages.append(f"💍 Anniversary Today: {row['Name']}")

# -------------------------------
# Show alert if needed
# -------------------------------
if alert_messages:
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Today's Alerts",
        "\n".join(alert_messages)
    )
    root.destroy()

log("Script completed successfully")
