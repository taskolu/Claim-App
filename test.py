
import os
import re
import shutil
import threading
import queue
import tkinter as tk
from tkinter import scrolledtext

# ===================== CONFIG ======================
ARCHIVE_PATH = r"\\cctprodfsx.converaprod.com\share\CHGFE-FILES\Treasury\WIP - Delia\Trading\WSFX\Watchlist Files\Trade Uploads PROD\Archive"

# Your real OneDrive Desktop folder
DEST_FOLDER = r"C:\Users\AbduTas\OneDrive - Convera\Desktop\WSFX-FILES"

log_q = queue.Queue()

def gui_log(msg):
    log_q.put(msg)

def pump_log():
    try:
        while True:
            msg = log_q.get_nowait()
            log_box.configure(state="normal")
            log_box.insert(tk.END, msg + "\n")
            log_box.see(tk.END)
            log_box.configure(state="disabled")
    except queue.Empty:
        pass
    root.after(80, pump_log)

# ================== File Normalization ==================
# Accept ANY timestamp suffix of 10–20 digits (covers WinSCP variations)
TS_SUFFIX = re.compile(r"^(.*)\.\d{10,20}$")

def normalize_name(raw):
    # Remove leading path piece shown by WinSCP
    if "/interface/" in raw:
        raw = raw.split("/interface/")[-1]

    # Clean whitespace and slashes
    raw = raw.strip().replace("/", "").replace("\\", "")

    # Remove double quotes if pasted with quotes
    if raw.startswith('"') and raw.endswith('"'):
        raw = raw[1:-1]

    # Strip .txt if present in the pasted value
    if raw.lower().endswith(".txt"):
        raw = raw[:-4]

    # Strip any trailing timestamp suffix like .20260116075505 or .202601160655
    m = TS_SUFFIX.match(raw)
    if m:
        raw = m.group(1)

    return raw


def parse_names(pasted):
    names = []
    for line in pasted.splitlines():
        line = line.strip()
        if not line:
            continue

        if "/interface/" in line:
            line = line.split("/interface/")[-1]

        base = normalize_name(line)
        if base:
            names.append(base)

    # Unique (case-insensitive)
    seen = set()
    uniq = []
    for n in names:
        low = n.lower()
        if low not in seen:
            seen.add(low)
            uniq.append(n)
    return uniq


# ===================== Instant Exact Finder (v9) ======================
def find_file_fast(base_name):
    """
    ZERO scanning. ZERO delay.
    Only direct exact checks:
        1) base_name + ".txt"
        2) base_name + ".TXT"
        3) base_name (no extension)
    """
    candidate1 = os.path.join(ARCHIVE_PATH, base_name + ".txt")
    if os.path.isfile(candidate1):
        return candidate1

    candidate2 = os.path.join(ARCHIVE_PATH, base_name + ".TXT")
    if os.path.isfile(candidate2):
        return candidate2

    candidate3 = os.path.join(ARCHIVE_PATH, base_name)
    if os.path.isfile(candidate3):
        return candidate3

    # No scanning in v9 — prevents long waits on huge UNC dirs
    return None


# ===================== Worker Thread ======================
def worker(pasted):
    try:
        names = parse_names(pasted)

        if not names:
            gui_log("⚠️ No filenames detected.")
            btn_run.config(state="normal")
            return

        os.makedirs(DEST_FOLDER, exist_ok=True)

        gui_log(f"Destination: {DEST_FOLDER}")
        gui_log(f"Files to process: {len(names)}\n")

        copied = 0
        missing = 0

        for base in names:
            gui_log(f"→ Searching for: {base}")

            found = find_file_fast(base)

            if found:
                fname = os.path.basename(found)
                dst = os.path.join(DEST_FOLDER, fname)
                try:
                    shutil.copy2(found, dst)
                    gui_log(f"   ✅ Copied: {fname}")
                    copied += 1
                except Exception as e:
                    gui_log(f"   ❌ Copy error: {e}")
            else:
                gui_log("   ⚠️ Not found in Archive")
                missing += 1

        gui_log("\n====== SUMMARY ======")
        gui_log(f"Copied:  {copied}")
        gui_log(f"Missing: {missing}")
        gui_log("Done.")
    finally:
        btn_run.config(state="normal")


# ===================== GUI ======================

root = tk.Tk()
root.title("WSFX Failed File Retriever (v9 - Instant + Flexible Timestamp)")
root.geometry("900x680")

tk.Label(root, text="Archive folder:", anchor="w").pack(fill="x", padx=10)
tk.Label(root, text=ARCHIVE_PATH, fg="#333", wraplength=840).pack(fill="x", padx=10, pady=(0,10))

tk.Label(
    root,
    text="Paste WinSCP failed file paths here (one per line):",
    anchor="w"
).pack(fill="x", padx=10)

input_box = scrolledtext.ScrolledText(root, width=120, height=10)
input_box.pack(fill="x", padx=10)

btn_frame = tk.Frame(root)
btn_frame.pack(fill="x", padx=10, pady=10)

btn_run = tk.Button(
    btn_frame,
    text="FIND & COPY",
    command=lambda: threading.Thread(
        target=worker,
        args=(input_box.get('1.0', tk.END),),
        daemon=True
    ).start(),
    bg="#2563EB",
    fg="white"
)
btn_run.pack(side="left")

def open_dest():
    os.makedirs(DEST_FOLDER, exist_ok=True)
    os.startfile(DEST_FOLDER)

tk.Button(btn_frame, text="Open Destination", command=open_dest).pack(side="left", padx=10)

tk.Label(root, text="Log:").pack(anchor="w", padx=10)
log_box = scrolledtext.ScrolledText(root, width=120, height=25, state="disabled")
log_box.pack(fill="both", expand=True, padx=10, pady=(0,10))

root.after(80, pump_log)
root.mainloop()