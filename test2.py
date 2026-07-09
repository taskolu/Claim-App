import os
import re
import shutil
import threading
import queue
import time  # <--- For the 2-second sleep
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import scrolledtext, messagebox

# ===================== CONFIG ======================
# 1. Archive (Where we get files FROM)
ARCHIVE_PATH = r"\\cctprodfsx.converaprod.com\share\CHGFE-FILES\Treasury\WIP - Delia\Trading\WSFX\Watchlist Files\Trade Uploads PROD\Archive"

# 2. PROD (Where we REDROP files TO)
PROD_PATH = r"\\cctprodfsx.converaprod.com\share\CHGFE-FILES\Treasury\WIP - Delia\Trading\WSFX\Watchlist Files\Trade Uploads PROD"

# 3. Local (Where we check files safely first)
USER_HOME = os.path.expanduser("~") 
onedrive_desktop = os.path.join(USER_HOME, "OneDrive - Convera", "Desktop", "WSFX-FILES")
local_desktop = os.path.join(USER_HOME, "Desktop", "WSFX-FILES")

if os.path.exists(os.path.join(USER_HOME, "OneDrive - Convera")):
    DEST_FOLDER = onedrive_desktop
else:
    DEST_FOLDER = local_desktop

# ===================== LOGGING ======================
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

# ================== Parsers ==================
VALID_PATTERN = re.compile(r"\d{4}-\d{2}-\d{2}")

def clean_name(raw):
    raw = raw.strip()
    if not raw: return None
    parts = raw.split()
    if not parts: return None
    raw = parts[0]
    if "/interface/" in raw: raw = raw.split("/interface/")[-1]
    raw = re.sub(r"\.\d{10,20}$", "", raw)
    if raw.lower().endswith(".txt"): raw = raw[:-4]
    return raw

def parse_xml_data(filepath):
    """Returns list of deal info for verification."""
    deals_found = []
    try:
        tree = ET.parse(filepath)
        root = tree.getroot()
        trades = root.findall(".//FXTRADE")
        if not trades:
            if root.find("external_deal_number") is not None: trades = [root]

        for trade in trades:
            d = trade.findtext("external_deal_number")
            a = trade.findtext("purchase_amount")
            c = trade.findtext("purchase_ccy")
            
            amt_fmt = "0.00"
            if a:
                try:
                    val = float(a)
                    amt_fmt = "{:,.2f}".format(val)
                except: amt_fmt = a
            
            deals_found.append(f"{d} ({amt_fmt} {c})")
    except:
        deals_found.append("Error parsing XML")
    return deals_found

def find_file_fast(base_name):
    candidates = [
        os.path.join(ARCHIVE_PATH, base_name + ".txt"),
        os.path.join(ARCHIVE_PATH, base_name + ".TXT"),
        os.path.join(ARCHIVE_PATH, base_name)
    ]
    for path in candidates:
        if os.path.isfile(path): return path
    return None

# ================== WORKER 1: RETRIEVE ==================
def run_retrieval():
    selected_indices = file_listbox.curselection()
    if not selected_indices:
        file_listbox.select_set(0, tk.END)
        selected_indices = file_listbox.curselection()
    selected_names = [file_listbox.get(i) for i in selected_indices]
    
    btn_retrieve.config(state="disabled")
    threading.Thread(target=retrieval_worker, args=(selected_names,), daemon=True).start()

def retrieval_worker(names):
    try:
        os.makedirs(DEST_FOLDER, exist_ok=True)
        gui_log(f"\n--- Retrieving {len(names)} files ---")
        
        for base in names:
            found = find_file_fast(base)
            if found:
                fname = os.path.basename(found)
                dst = os.path.join(DEST_FOLDER, fname)
                try:
                    shutil.copy2(found, dst)
                    gui_log(f"✅ Copied: {fname}")
                    # Log details
                    info = parse_xml_data(dst)
                    for i in info: gui_log(f"   ↳ {i}")
                except Exception as e:
                    gui_log(f"❌ Error: {e}")
            else:
                gui_log(f"⚠️ Not found: {base}")
        gui_log("====== DONE ======")
    finally:
        btn_retrieve.config(state="normal")

# ================== WORKER 2: REDROP (The New Feature) ==================
def confirm_and_redrop():
    selected_indices = file_listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("Select Files", "Please select files to redrop.")
        return

    selected_names = [file_listbox.get(i) for i in selected_indices]
    
    # 1. Build Verification String
    verify_msg = "Are you sure you want to upload these to PROD?\n\n"
    files_to_upload = []

    for base in selected_names:
        # Look for file in LOCAL folder (we only redrop what we already retrieved)
        # Try finding with/without extension in local folder
        local_path = None
        for ext in [".txt", ".TXT", ""]:
            p = os.path.join(DEST_FOLDER, base + ext)
            if os.path.exists(p):
                local_path = p
                break
        
        if local_path:
            details = parse_xml_data(local_path)
            details_str = ", ".join(details)
            verify_msg += f"File: {base}\nInfo: {details_str}\n\n"
            files_to_upload.append(local_path)
        else:
            verify_msg += f"File: {base} (NOT FOUND IN LOCAL FOLDER - Retrieve first!)\n\n"

    if not files_to_upload:
        messagebox.showerror("Error", "No valid local files found. Please 'Retrieve' them first.")
        return

    # 2. Ask User
    confirm = messagebox.askyesno("Confirm Redrop", verify_msg)
    if confirm:
        btn_redrop.config(state="disabled")
        threading.Thread(target=redrop_worker, args=(files_to_upload,), daemon=True).start()

def redrop_worker(file_paths):
    try:
        gui_log("\n--- 🚀 STARTING UPLOAD TO PROD ---")
        count = len(file_paths)
        
        for i, src in enumerate(file_paths):
            fname = os.path.basename(src)
            dst = os.path.join(PROD_PATH, fname)
            
            # --- THE SLEEP LOGIC ---
            # If we are doing more than 1 file, sleep BEFORE every copy (except maybe the very first one, 
            # but sleeping before all is safer to ensure gap from previous manual actions)
            if count > 1 and i > 0:
                gui_log("   ⏳ Sleeping 2s to prevent collision...")
                time.sleep(2)
            # -----------------------

            try:
                shutil.copy2(src, dst)
                gui_log(f"   🚀 UPLOADED: {fname}")
            except Exception as e:
                gui_log(f"   ❌ FAILED TO UPLOAD: {e}")

        gui_log("====== UPLOAD COMPLETE ======")
    finally:
        btn_redrop.config(state="normal")

# ===================== GUI Functions ======================
def analyze_paste():
    raw_text = paste_box.get("1.0", tk.END)
    found_names = []
    seen = set()
    ignored_count = 0
    for line in raw_text.splitlines():
        clean = clean_name(line)
        if clean:
            if VALID_PATTERN.search(clean):
                if clean.lower() not in seen:
                    seen.add(clean.lower())
                    found_names.append(clean)
            else:
                ignored_count += 1
    file_listbox.delete(0, tk.END)
    if not found_names:
        gui_log(f"⚠️ No files found (Ignored {ignored_count}).")
        return
    for name in found_names:
        file_listbox.insert(tk.END, name)
    gui_log(f"Analyzed: {len(found_names)} valid files.")

# ===================== LAYOUT ======================
root = tk.Tk()
root.title("WSFX Manager (v18 - Redrop)")
root.geometry("1000x800")

tk.Label(root, text=f"Archive: {ARCHIVE_PATH}", fg="gray", font=("Arial", 8)).pack(anchor="w", padx=10, pady=5)

main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=5)

# Left
left_frame = tk.Frame(main_frame)
left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
tk.Label(left_frame, text="1. Paste WinSCP Results:", font=("Arial", 9, "bold")).pack(anchor="w")
paste_box = scrolledtext.ScrolledText(left_frame, width=40, height=18)
paste_box.pack(fill="both", expand=True)
tk.Button(left_frame, text="Analyze List >>", command=analyze_paste, bg="#DDDDDD").pack(fill="x", pady=5)

# Right
right_frame = tk.Frame(main_frame)
right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
tk.Label(right_frame, text="2. Select Files:", font=("Arial", 9, "bold")).pack(anchor="w")
scrollbar = tk.Scrollbar(right_frame)
scrollbar.pack(side="right", fill="y")
file_listbox = tk.Listbox(right_frame, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set, width=40)
file_listbox.pack(fill="both", expand=True)
scrollbar.config(command=file_listbox.yview)

# Bottom
action_frame = tk.Frame(root)
action_frame.pack(fill="x", padx=10, pady=10)

btn_retrieve = tk.Button(action_frame, text="3. RETRIEVE LOCALLY (Verify Logs)", command=run_retrieval, bg="#007ACC", fg="white", font=("Arial", 11, "bold"), height=2)
btn_retrieve.pack(fill="x", pady=(0, 5))

# REDROP BUTTON (Different Color)
btn_redrop = tk.Button(action_frame, text="4. REDROP SELECTED TO PROD (Admin)", command=confirm_and_redrop, bg="#D32F2F", fg="white", font=("Arial", 11, "bold"), height=2)
btn_redrop.pack(fill="x")

tk.Button(root, text="Open Local Folder", command=lambda: os.startfile(DEST_FOLDER)).pack(pady=5)

tk.Label(root, text="Log:", anchor="w").pack(fill="x", padx=10)
log_box = scrolledtext.ScrolledText(root, height=12, state="disabled", bg="#F0F0F0")
log_box.pack(fill="x", padx=10, pady=(0,10))

root.after(80, pump_log)
root.mainloop()