#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_for_date.py — date picker launcher for daily_report.py
=============================================================
v3.7+v3.9: Tkinter calendar dialog with a 30-second auto-run countdown.

v3.9 change: default selected date = YESTERDAY (not today). maxdate stays
as today so users can still pick today for end-of-day runs.

Flow:
  1. Open a small dialog with a calendar widget (default = yesterday, max = today)
  2. Bottom-right button labelled "Generate & Send (30)" counts down
  3. When the user clicks Generate, OR when the countdown hits 0,
     this script calls:  daily_report.py --fetch-odoo --date YYYY-MM-DD
  4. Cancel or close-window → exit without running the pipeline

The countdown does NOT reset when the user clicks on a date; that's
intentional so that an unattended run still proceeds with today as the
default (or whatever date is currently highlighted).
"""
import os, sys, subprocess
from datetime import date, timedelta

COUNTDOWN_SEC = 30

# ── Import GUI libs; if any fail, fall back to running for today ─────────────
try:
    import tkinter as tk
    from tkcalendar import Calendar
except Exception as e:
    print(f"[run_for_date] tkinter/tkcalendar not available ({e}); "
          f"falling back to today's date.")
    here = os.path.dirname(os.path.abspath(__file__))
    sys.exit(subprocess.call([sys.executable, os.path.join(here, 'daily_report.py'),
                              '--fetch-odoo']))

# ── Build the dialog ─────────────────────────────────────────────────────────
root = tk.Tk()
root.title("Sale Report — เลือกวันที่")
root.geometry("360x400")
root.resizable(False, False)

# Center on screen
root.update_idletasks()
w, h = 360, 400
sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

# Try to bring the window to the front (helpful when launched from a .bat)
try:
    root.attributes('-topmost', True)
    root.after(500, lambda: root.attributes('-topmost', False))
except Exception:
    pass

tk.Label(root, text="📅 เลือกวันที่ของ Sale Report",
         font=('Segoe UI', 12, 'bold')).pack(pady=(12, 6))

today = date.today()
yesterday = today - timedelta(days=1)
cal = Calendar(root, selectmode='day', date_pattern='yyyy-mm-dd',
               year=yesterday.year, month=yesterday.month, day=yesterday.day,
               maxdate=today,           # max = today (can still pick today)
               showweeknumbers=False)
cal.pack(padx=14, pady=6)

tk.Label(root,
         text='ถ้าไม่ทำอะไรภายใน 30 วินาที จะรันด้วย "เมื่อวาน" (default)',
         fg='gray', font=('Segoe UI', 8)).pack(pady=2)

# ── State + buttons ──────────────────────────────────────────────────────────
state = {'cancelled': False}

def on_send():
    """User pressed Generate & Send — or countdown hit 0."""
    root.destroy()

def on_cancel():
    state['cancelled'] = True
    root.destroy()

btn_frame = tk.Frame(root)
btn_frame.pack(pady=12)
tk.Button(btn_frame, text="Cancel", command=on_cancel, width=10
          ).pack(side='left', padx=6)
send_btn = tk.Button(btn_frame,
                     text=f"Generate & Send ({COUNTDOWN_SEC})",
                     command=on_send, width=24,
                     bg='#0078d4', fg='white',
                     font=('Segoe UI', 10, 'bold'),
                     activebackground='#005a9e', activeforeground='white')
send_btn.pack(side='left', padx=6)

# Treat window-close (X) as Cancel
root.protocol("WM_DELETE_WINDOW", on_cancel)

# ── Countdown — ticks every second ───────────────────────────────────────────
remaining = [COUNTDOWN_SEC]

def tick():
    if state.get('done'):
        return
    remaining[0] -= 1
    if remaining[0] <= 0:
        on_send()
        return
    try:
        send_btn.config(text=f"Generate & Send ({remaining[0]})")
    except tk.TclError:
        # Window already destroyed
        return
    root.after(1000, tick)

root.after(1000, tick)
root.mainloop()

# ── After the dialog closes ─────────────────────────────────────────────────
if state['cancelled']:
    print("[run_for_date] User cancelled — pipeline NOT run.")
    sys.exit(0)

chosen = cal.get_date()  # 'YYYY-MM-DD' (string per date_pattern)
print(f"[run_for_date] Running pipeline for: {chosen}")

here = os.path.dirname(os.path.abspath(__file__))
result = subprocess.run([sys.executable,
                         os.path.join(here, 'daily_report.py'),
                         '--fetch-odoo',
                         '--date', chosen])
sys.exit(result.returncode)
