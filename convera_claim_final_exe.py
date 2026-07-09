import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from fpdf import FPDF
from datetime import datetime
import json
import os
import re
import sys # Required for EXE resource path
import pdfplumber

# --- TRY IMPORTING OUTLOOK LIBRARY ---
try:
    import win32com.client as win32
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False

# --- EXE RESOURCE PATH HELPER ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Configuration: Paths ---
USER_BASE = os.path.expanduser("~") 
SSI_FOLDER_PATH = r"C:\Users\AbduTas\OneDrive - Convera\Documents - Trade Confirmations Team site\TRADE CONFIRMATION\SSI"
SAVE_FOLDER_PATH = r"C:\Users\AbduTas\OneDrive - Convera\Documents - Trade Confirmations Team site\TRADE CONFIRMATION\CLAIMS\claim letters"

if not os.path.exists(SSI_FOLDER_PATH): SSI_FOLDER_PATH = USER_BASE
if not os.path.exists(SAVE_FOLDER_PATH): SAVE_FOLDER_PATH = USER_BASE

# --- Configuration: Company ---
COMPANY_NAME = "CONVERA UK LIMITED"
COMPANY_ADDR = [
    "ALPHABETA BUILDING",
    "14-18 FINSBURY SQUARE",
    "LONDON EC2A 1AH",
    "UNITED KINGDOM"
]
CONTACT_EMAIL = "TreasuryConfirms@Convera.com"

# --- COUNTERPARTY STORE ---
# Counterparties live in a JSON file so they can be added/edited/removed from
# inside the app (Manage Counterparties button) and survive app updates.
# The defaults below are only used to seed the file on first run.
APP_DATA_DIR = os.path.join(USER_BASE, ".convera_claim_manager")
COUNTERPARTY_FILE = os.path.join(APP_DATA_DIR, "counterparties.json")

DEFAULT_COUNTERPARTIES = {
    "BANK OF AMERICA, N.A., New Castle": "usfxcomp@bofa.com; emeaservicingfi@bankofamerica.com; usfxcomp@bofa.com",
    "Bank of Montreal Toronto, Montreal": "bmo.investigation@bmo.com",
    "BARCLAYS BANK PLC, Washington": "xrasgptsyinterestcla@barclays.com",
    "CITI EUROPE, DUBLIN": "branchfx.presettlements@citi.com",
    "CITIBANK, NATIONAL ASSOCIATION, New Castle": "dl.ops.nam.foreignexchangeldn@citi.com; michael.kania@citi.com",
    "CROWN AGENTS BANK LTD, SUTTON": "TreasuryOperations@crownagentsbank.com; CAB-FX@crownagentsbank.com",
    "Deutsche Bank AG, FRANKFURT": "fxcash.jax@db.com",
    "Fifth Third Bank, National Associat, CINCINNATI": "OpsCMInvestigations.Bancorp@53.com",
    "Macquarie Group Ltd, SYDNEY": "COGMODANZFICOps@macquarie.com; ficnylonops@macquarie.com; fx@macquarie.com",
    "Merrill Lynch International, LONDON": "usfxcomp@bofa.com; emeaservicingfi@bankofamerica.com; usfxcomp@bofa.com",
    "Mizuho Capital Markets Corporation, NEW YORK": "ControlTowerInvestigations@mizuhogroup.com",
    "MORGAN STANLEY & CO. INTERNATIONAL, LONDON": "NA.Claims@morganstanley.com",
    "Nomura International PLC, LONDON": "fxopssettlements@nomura.com",
    "Royal Bank of Canada, TORONTO": "fxmminvestigations@rbccm.com; GMFXCORPConfirmations@rbccm.com",
    "Societe Generale SA, PARIS": "BLR-OTCPOST-SETTLEMENTS@SOCGEN.COM",
    "State Street Bank and Trust Company, BOSTON": "SSGM_NAFXConfirmations@StateStreet.com; SSGM_EMEAFXConfirmations@StateStreet.com",
    "U.S. BANK N.A., MINNEAPOLIS": "FXConfirmations@usbank.com; fxsettlementinstructions@usbank.com",
    "UBS AG, Washington": "sh-fxmm-globalclaims@ubs.com",
    "WELLS FARGO BANK, NATIONAL ASSOCIAT, NEW YORK": "SFFXCorporate@wellsfargo.com; Neethu.Abraham@wellsfargo.com"
}


def load_counterparties():
    """Load counterparties from JSON, seeding with defaults on first run.
    A corrupt file is backed up (.bak) instead of being silently overwritten."""
    try:
        with open(COUNTERPARTY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and data:
            return {str(k).strip(): str(v).strip() for k, v in data.items() if str(k).strip()}
    except FileNotFoundError:
        pass
    except Exception:
        try:
            os.replace(COUNTERPARTY_FILE, COUNTERPARTY_FILE + ".bak")
        except Exception:
            pass
    counterparties = dict(DEFAULT_COUNTERPARTIES)
    save_counterparties(counterparties)
    return counterparties


def save_counterparties(counterparties):
    """Atomically persist the counterparty dict to JSON."""
    os.makedirs(APP_DATA_DIR, exist_ok=True)
    tmp_path = COUNTERPARTY_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(counterparties, f, indent=2, ensure_ascii=False, sort_keys=True)
    os.replace(tmp_path, COUNTERPARTY_FILE)


def validate_emails(email_str):
    """Check a 'a@b.com; c@d.com' string. Returns (ok, cleaned_string_or_error).
    An empty string is allowed (counterparty with no email yet)."""
    parts = [p.strip() for p in email_str.replace(",", ";").split(";") if p.strip()]
    for part in parts:
        if part.count("@") != 1 or "." not in part.split("@")[1] or " " in part:
            return False, f"'{part}' does not look like a valid email address."
    return True, "; ".join(parts)

# --- Data Lists ---
CURRENCY_LIST = sorted([
    "AED", "AUD", "BGN", "BHD", "BWP", "CAD", "CHF", "CNH", "CZK", "DKK", 
    "EUR", "FJD", "GBP", "GHS", "GMD", "HKD", "HUF", "ILS", "ISK", "JOD", 
    "JPY", "KES", "KWD", "MAD", "MGA", "MUR", "MWK", "MXN", "MZN", "NAD", 
    "NGN", "NOK", "NZD", "OMR", "PGK", "PLN", "QAR", "RON", "SAR", "SBD", 
    "SEK", "SGD", "THB", "TND", "TRY", "TZS", "UGX", "USD", "VUV", "WST", 
    "XAF", "XOF", "XPF", "ZAR", "ZMW"
])

class CounterpartyEditDialog(tk.Toplevel):
    """Small modal form to add a new counterparty or edit an existing one."""

    def __init__(self, parent, title, name="", emails=""):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.configure(bg="#f5f5f5")
        self.result = None  # (name, emails) when saved

        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Name:", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.name_entry = ttk.Entry(frame, width=55)
        self.name_entry.insert(0, name)
        self.name_entry.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=(0, 8))

        ttk.Label(frame, text="Email(s):", style="Header.TLabel").grid(row=1, column=0, sticky="w")
        self.email_entry = ttk.Entry(frame, width=55)
        self.email_entry.insert(0, emails)
        self.email_entry.grid(row=1, column=1, sticky="ew", padx=(10, 0))
        ttk.Label(frame, text="Separate multiple addresses with ; (semicolon)",
                  font=("Segoe UI", 8), foreground="gray").grid(row=2, column=1, sticky="w", padx=(10, 0))

        btns = ttk.Frame(frame)
        btns.grid(row=3, column=0, columnspan=2, pady=(20, 0))
        tk.Button(btns, text="SAVE", command=self._save, bg="#0096C8", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(btns, text="CANCEL", command=self.destroy, bg="#999", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=5).pack(side=tk.LEFT, padx=5)

        self.name_entry.focus_set()
        self.bind("<Return>", lambda e: self._save())
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.grab_set()

    def _save(self):
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Error", "Counterparty name is required.", parent=self)
            return
        ok, emails_or_err = validate_emails(self.email_entry.get())
        if not ok:
            messagebox.showerror("Error", emails_or_err, parent=self)
            return
        self.result = (name, emails_or_err)
        self.destroy()


class CounterpartyManager(tk.Toplevel):
    """Manage counterparties: add new ones, edit names/emails, remove old ones.
    Changes are saved to disk immediately and pushed back to the main window."""

    def __init__(self, parent, counterparties, on_change):
        super().__init__(parent)
        self.title("Manage Counterparties")
        self.geometry("820x480")
        self.configure(bg="#f5f5f5")
        self.counterparties = counterparties
        self.on_change = on_change

        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(frame)
        top.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(top, text="Counterparties", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self._refresh_tree())
        search_entry = ttk.Entry(top, textvariable=self.search_var, width=30)
        search_entry.pack(side=tk.RIGHT)
        ttk.Label(top, text="Search:").pack(side=tk.RIGHT, padx=(0, 5))

        # Table
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=("name", "emails"), show="headings", selectmode="browse")
        self.tree.heading("name", text="Counterparty")
        self.tree.heading("emails", text="Email Address(es)")
        self.tree.column("name", width=320, anchor="w")
        self.tree.column("emails", width=440, anchor="w")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        # Buttons
        btns = ttk.Frame(frame)
        btns.pack(fill=tk.X, pady=(15, 0))

        def styled_btn(text, cmd, color):
            return tk.Button(btns, text=text, command=cmd, bg=color, fg="white",
                             font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6)

        styled_btn("+ ADD", self.add_new, "#28a745").pack(side=tk.LEFT, padx=(0, 8))
        styled_btn("EDIT", self.edit_selected, "#0096C8").pack(side=tk.LEFT, padx=8)
        styled_btn("REMOVE", self.remove_selected, "#d9534f").pack(side=tk.LEFT, padx=8)
        styled_btn("CLOSE", self.destroy, "#999").pack(side=tk.RIGHT)

        self._refresh_tree()
        self.transient(parent)
        self.grab_set()

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        query = self.search_var.get().strip().lower()
        for name in sorted(self.counterparties, key=str.lower):
            emails = self.counterparties[name]
            if query and query not in name.lower() and query not in emails.lower():
                continue
            self.tree.insert("", tk.END, iid=name, values=(name, emails or "(no email on file)"))

    def _selected_name(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select a counterparty first.", parent=self)
            return None
        return sel[0]

    def _persist(self):
        try:
            save_counterparties(self.counterparties)
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save counterparties:\n{e}", parent=self)
        self.on_change()
        self._refresh_tree()

    def add_new(self):
        dlg = CounterpartyEditDialog(self, "Add Counterparty")
        self.wait_window(dlg)
        if not dlg.result:
            return
        name, emails = dlg.result
        if name in self.counterparties:
            messagebox.showerror("Error", f"'{name}' already exists. Use EDIT to change it.", parent=self)
            return
        self.counterparties[name] = emails
        self._persist()

    def edit_selected(self):
        old_name = self._selected_name()
        if not old_name:
            return
        dlg = CounterpartyEditDialog(self, "Edit Counterparty", old_name, self.counterparties[old_name])
        self.wait_window(dlg)
        if not dlg.result:
            return
        new_name, emails = dlg.result
        if new_name != old_name and new_name in self.counterparties:
            messagebox.showerror("Error", f"'{new_name}' already exists.", parent=self)
            return
        del self.counterparties[old_name]
        self.counterparties[new_name] = emails
        self._persist()

    def remove_selected(self):
        name = self._selected_name()
        if not name:
            return
        if messagebox.askyesno("Confirm Removal", f"Remove counterparty?\n\n{name}", parent=self):
            del self.counterparties[name]
            self._persist()

class MinimalClaimPDF(FPDF):
    def header(self):
        # Colors
        convera_blue = (0, 150, 200)
        self.set_draw_color(*convera_blue) 
        self.set_fill_color(*convera_blue) 
        
        # 1. Logo Logic (USING RESOURCE PATH FOR EXE)
        logo_path = resource_path("convera_logo.png")
        
        if os.path.exists(logo_path):
            self.image(logo_path, 10, 10, 35)
        else:
            self.set_xy(10, 10)
            self.set_font('Helvetica', 'B', 20)
            self.set_text_color(0, 0, 0)
            self.cell(0, 10, 'Convera', ln=False)

        # 2. "COMPENSATION CLAIM" Text (Right)
        self.set_xy(10, 11) 
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'COMPENSATION CLAIM', align='R', ln=True)

        # 3. Horizontal Blue Line 
        self.set_line_width(1.5) 
        self.line(10, 26, 200, 26) 
        
        # Spacer
        self.ln(12)

    def footer(self):
        self.set_y(-12)
        self.set_font('Helvetica', 'I', 7)
        self.set_text_color(128, 128, 128)
        self.cell(0, 5, f'{COMPANY_NAME}', align='C', ln=1)

class ClaimApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convera Claims Manager")
        self.root.geometry("640x800")
        self.root.minsize(600, 720)
        self.root.configure(bg="#f5f5f5")
        
        self.ssi_data = {}
        self.ssi_pdf_path = None
        self.last_generated_pdf = None
        self.counterparties = load_counterparties()

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TFrame", background="#f5f5f5")
        style.configure("TLabel", background="#f5f5f5", foreground="#333", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#000")
        style.configure("TButton", font=("Segoe UI", 10), background="#ddd", borderwidth=1)
        style.map("TButton", background=[('active', '#ccc')])
        style.configure("TCombobox", padding=5)
        style.configure("TEntry", padding=5)

        # --- Layout ---
        # Header bar (Convera blue)
        header = tk.Frame(root, bg="#0096C8")
        header.pack(fill=tk.X)
        tk.Label(header, text="Convera Claims Manager", bg="#0096C8", fg="white",
                 font=("Segoe UI", 15, "bold"), padx=25, pady=12).pack(side=tk.LEFT)
        tk.Button(header, text="⚙ Manage Counterparties", command=self.open_counterparty_manager,
                  bg="#007AA3", fg="white", activebackground="#00658A", activeforeground="white",
                  font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
                  padx=12, pady=6).pack(side=tk.RIGHT, padx=20)

        main_frame = ttk.Frame(root, padding="30 20 30 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Generate Claim Letter", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 15))

        # 1. SSI Section
        ssi_frame = ttk.Frame(main_frame)
        ssi_frame.pack(fill=tk.X, pady=(0, 20))
        self.lbl_ssi_status = ttk.Label(ssi_frame, text="Checking SSI...", foreground="gray")
        self.lbl_ssi_status.pack(side=tk.LEFT)
        self.btn_manual_upload = ttk.Button(ssi_frame, text="Upload SSI (Manual)", command=self.manual_upload_ssi)
        
        self.startup_ssi_check()
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(0, 20))

        # 2. Form Grid
        grid_frame = ttk.Frame(main_frame)
        grid_frame.pack(fill=tk.BOTH, expand=True)
        grid_frame.columnconfigure(1, weight=1)

        # Counterparty
        ttk.Label(grid_frame, text="Counterparty:", style="Header.TLabel").grid(row=0, column=0, sticky="w", pady=10)
        self.cp_var = tk.StringVar()
        self.cp_combo = ttk.Combobox(grid_frame, textvariable=self.cp_var, values=self.counterparty_names())
        self.cp_combo.grid(row=0, column=1, sticky="ew", padx=(20, 0))
        self.lbl_cp_email = ttk.Label(grid_frame, text="(Select or type to search)", font=("Segoe UI", 8), foreground="gray")
        self.lbl_cp_email.grid(row=1, column=1, sticky="w", padx=(20, 0))
        self.cp_combo.bind("<<ComboboxSelected>>", self._on_counterparty_selected)
        self.cp_combo.bind("<KeyRelease>", self._filter_counterparty_dropdown)
        self.cp_var.trace_add("write", lambda *a: self._on_counterparty_selected())

        # Claim Ref
        ttk.Label(grid_frame, text="Claim Ref:", style="Header.TLabel").grid(row=2, column=0, sticky="w", pady=10)
        ref_container = ttk.Frame(grid_frame)
        ref_container.grid(row=2, column=1, sticky="w", padx=(20, 0))
        
        today_str = datetime.now().strftime("%Y%m%d")
        self.ref_prefix = f"CONV-{today_str}-"
        
        ttk.Label(ref_container, text=self.ref_prefix, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        self.ref_suffix_combo = ttk.Combobox(ref_container, values=[f"{i:02d}" for i in range(1, 21)], width=3)
        self.ref_suffix_combo.current(0)
        self.ref_suffix_combo.pack(side=tk.LEFT)

        # Amount
        ttk.Label(grid_frame, text="Amount:", style="Header.TLabel").grid(row=3, column=0, sticky="w", pady=10)
        amt_frame = ttk.Frame(grid_frame)
        amt_frame.grid(row=3, column=1, sticky="ew", padx=(20, 0))
        self.curr_combo = ttk.Combobox(amt_frame, values=CURRENCY_LIST, width=5)
        self.curr_combo.current(CURRENCY_LIST.index("USD"))
        self.curr_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.curr_combo.bind("<<ComboboxSelected>>", lambda e: self._update_preview())
        self.amount_var = tk.StringVar()
        self.amount_var.trace_add("write", lambda *a: self._update_preview())
        self.amount_entry = ttk.Entry(amt_frame, width=20, textvariable=self.amount_var)
        self.amount_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Dates
        ttk.Label(grid_frame, text="Dates:", style="Header.TLabel").grid(row=4, column=0, sticky="w", pady=10)
        date_frame = ttk.Frame(grid_frame)
        date_frame.grid(row=4, column=1, sticky="w", padx=(20, 0))
        ttk.Label(date_frame, text="Due:").pack(side=tk.LEFT)
        self.date_due = DateEntry(date_frame, width=12, date_pattern='dd/mm/yyyy', background='#444', foreground='white')
        self.date_due.pack(side=tk.LEFT, padx=(5, 20))
        ttk.Label(date_frame, text="Received:").pack(side=tk.LEFT)
        self.date_rec = DateEntry(date_frame, width=12, date_pattern='dd/mm/yyyy', background='#444', foreground='white')
        self.date_rec.pack(side=tk.LEFT, padx=(5, 0))
        self.date_due.bind("<<DateEntrySelected>>", lambda e: self._update_preview())
        self.date_rec.bind("<<DateEntrySelected>>", lambda e: self._update_preview())

        # Rate
        ttk.Label(grid_frame, text="Rate (%):", style="Header.TLabel").grid(row=5, column=0, sticky="w", pady=10)
        self.rate_var = tk.StringVar()
        self.rate_var.trace_add("write", lambda *a: self._update_preview())
        self.rate_entry = ttk.Entry(grid_frame, width=10, textvariable=self.rate_var)
        self.rate_entry.grid(row=5, column=1, sticky="w", padx=(20, 0))

        # Live claim preview
        preview_box = tk.Frame(main_frame, bg="white", highlightbackground="#d0d0d0",
                               highlightthickness=1)
        preview_box.pack(fill=tk.X, pady=(20, 0))
        tk.Label(preview_box, text="CLAIM PREVIEW", bg="white", fg="#888",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        self.lbl_preview = tk.Label(preview_box, text="Enter amount, rate and dates to see the claim...",
                                    bg="white", fg="#888", font=("Segoe UI", 12), anchor="w",
                                    justify=tk.LEFT, padx=12, pady=8)
        self.lbl_preview.pack(fill=tk.X)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=25, fill=tk.X)

        def action_btn(text, cmd, color, active):
            return tk.Button(btn_frame, text=text, command=cmd, bg=color, fg="white",
                             activebackground=active, activeforeground="white", cursor="hand2",
                             font=("Segoe UI", 11, "bold"), relief="flat", padx=20, pady=10)

        action_btn("GENERATE PDF", self.generate_pdf, "#0096C8", "#007AA3").pack(side=tk.LEFT, padx=(50, 10))
        action_btn("DRAFT EMAIL", self.draft_email, "#28a745", "#1F8637").pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="CLEAR", command=self.clear_form, bg="#f5f5f5", fg="#666",
                  activebackground="#e0e0e0", cursor="hand2", font=("Segoe UI", 10),
                  relief="flat", padx=15, pady=10).pack(side=tk.RIGHT)

        # Status bar
        self.lbl_status = tk.Label(root, text="Ready", bg="#e8e8e8", fg="#555", anchor="w",
                                   font=("Segoe UI", 8), padx=12, pady=4)
        self.lbl_status.pack(side=tk.BOTTOM, fill=tk.X)

    # --- GUI helpers ---
    def set_status(self, text):
        self.lbl_status.config(text=text)

    def _filter_counterparty_dropdown(self, event):
        """Type-to-search: narrow the dropdown list to names containing the typed text."""
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        text = self.cp_var.get().strip().lower()
        names = self.counterparty_names()
        matches = [n for n in names if text in n.lower()] if text else names
        self.cp_combo.config(values=matches or names)

    def _update_preview(self):
        data = self.get_data()
        if not data:
            self.lbl_preview.config(text="Enter amount, rate and dates to see the claim...",
                                    fg="#888", font=("Segoe UI", 12))
            return
        amount, rate, days, d1, d2 = data
        if amount <= 0 or not self.rate_var.get().strip():
            self.lbl_preview.config(text="Enter amount, rate and dates to see the claim...",
                                    fg="#888", font=("Segoe UI", 12))
            return
        currency = self.curr_combo.get()
        interest = (amount * days * (rate / 100)) / 360
        if days <= 0:
            self.lbl_preview.config(
                text=f"Days late: {days}  -  claim would be {currency} 0.00\nCheck the dates: received date is not after the due date.",
                fg="#d9534f", font=("Segoe UI", 11))
        else:
            self.lbl_preview.config(
                text=f"{days} day(s) late  ·  Claim: {currency} {interest:,.2f}",
                fg="#0096C8", font=("Segoe UI", 14, "bold"))

    def clear_form(self):
        self.cp_var.set("")
        self.amount_var.set("")
        self.rate_var.set("")
        self.ref_suffix_combo.current(0)
        today = datetime.now().date()
        self.date_due.set_date(today)
        self.date_rec.set_date(today)
        self.cp_combo.config(values=self.counterparty_names())
        self._update_preview()
        self.set_status("Form cleared")

    # --- Counterparty management ---
    def counterparty_names(self):
        return sorted(self.counterparties, key=str.lower)

    def _on_counterparty_selected(self, event=None):
        name = self.cp_var.get()
        if not name:
            self.lbl_cp_email.config(text="(Select or Type)", foreground="gray")
            return
        emails = self.counterparties.get(name)
        if emails:
            self.lbl_cp_email.config(text=f"Email: {emails}", foreground="green")
        elif name in self.counterparties:
            self.lbl_cp_email.config(text="No email on file - add one via Manage Counterparties", foreground="#d9534f")
        else:
            self.lbl_cp_email.config(text="Not in counterparty list - email must be entered manually", foreground="#e68a00")

    def refresh_counterparty_combo(self):
        self.cp_combo.config(values=self.counterparty_names())
        self._on_counterparty_selected()

    def open_counterparty_manager(self):
        CounterpartyManager(self.root, self.counterparties, self.refresh_counterparty_combo)

    def startup_ssi_check(self):
        target_dir = SSI_FOLDER_PATH
        if os.path.exists(target_dir):
            try:
                pdf_files = [f for f in os.listdir(target_dir) if f.lower().endswith(".pdf")]
                if pdf_files:
                    full_paths = [os.path.join(target_dir, f) for f in pdf_files]
                    largest_pdf = max(full_paths, key=os.path.getsize)
                    self.ssi_pdf_path = largest_pdf
                    file_name = os.path.basename(largest_pdf)
                    self.lbl_ssi_status.config(text=f"SSI Loaded: {file_name}", foreground="green")
                    return
            except Exception:
                pass
        self.lbl_ssi_status.config(text="SSI Not Found (Please Upload)", foreground="#d9534f")
        self.btn_manual_upload.pack(side=tk.RIGHT)

    def manual_upload_ssi(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")], title="Select SSI Document")
        if filepath:
            self.ssi_pdf_path = filepath
            self.lbl_ssi_status.config(text=f"Loaded: {os.path.basename(filepath)}", foreground="green")

    def extract_ssi_for_currency(self, target_ccy):
        if not self.ssi_pdf_path: return None
        target_ccy = target_ccy.upper().strip()
        try:
            with pdfplumber.open(self.ssi_pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if not table: continue
                    for row in table:
                        clean_row = [cell.strip() if cell else "" for cell in row]
                        if len(clean_row) >= 3:
                            if clean_row[1].upper() == target_ccy:
                                return {
                                    "Beneficiary": clean_row[2],
                                    "AccountWith": clean_row[3],
                                    "Intermediary": clean_row[4] if len(clean_row) > 4 else ""
                                }
        except Exception:
            return None
        return None

    def get_data(self):
        try:
            amt_str = self.amount_entry.get().replace(',', '')
            amount = float(amt_str) if amt_str else 0.0
            rate = float(self.rate_entry.get())
            d1 = self.date_due.get_date()
            d2 = self.date_rec.get_date()
            days = (d2 - d1).days
            return amount, rate, days, d1, d2
        except ValueError:
            return None

    def generate_pdf(self):
        data = self.get_data()
        if not data:
            messagebox.showerror("Error", "Check Amount and Rate fields.")
            return
        
        amount, rate, days, d1, d2 = data
        
        if days <= 0:
            proceed = messagebox.askyesno("Warning", f"Days Late is {days}. The claim amount will be 0.\n\nAre you sure the dates are correct?")
            if not proceed: return

        interest = (amount * days * (rate/100)) / 360
        
        cp_name = self.cp_var.get()
        currency = self.curr_combo.get()
        ref_suffix = self.ref_suffix_combo.get()
        claim_ref = f"{self.ref_prefix}{ref_suffix}"

        if not cp_name:
            messagebox.showerror("Error", "Counterparty Name required.")
            return

        ssi_info = self.extract_ssi_for_currency(currency)
        
        fname = f"Claim_{cp_name}_{claim_ref}.pdf"
        filepath = filedialog.asksaveasfilename(
            initialdir=SAVE_FOLDER_PATH,
            defaultextension=".pdf", 
            initialfile=fname,
            title="Save Claim Letter"
        )
        if not filepath: return

        # --- PDF Generation ---
        pdf = MinimalClaimPDF()
        pdf.set_display_mode(zoom=100, layout='SinglePage')
        pdf.add_page()
        
        # 1. Company Address (Compact)
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(80, 80, 80)
        for line in COMPANY_ADDR:
            pdf.cell(0, 3.5, line, ln=True)
        pdf.ln(8)

        # 2. TO and DATE
        date_str = datetime.now().strftime("%d %B %Y")
        
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(0, 0, 0)
        
        pdf.cell(100, 5, "TO:", ln=0)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 5, date_str, align='R', ln=1)
        
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 5, cp_name, ln=True)
        pdf.ln(8)

        # Subject
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, f"SUBJECT: Compensation Claim for Late Settlement - Ref: {claim_ref}", ln=True)
        pdf.ln(3)

        # Body
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 5, "Dear Settlements Team,\n\nWe hereby claim compensation for the late settlement of the transaction detailed below. The funds due to us were not received on the agreed value date, and we have incurred a corresponding cost of funds. Please arrange payment of the claim amount calculated as follows:")
        pdf.ln(5)

        # Table
        pdf.set_fill_color(250, 250, 250)
        pdf.set_draw_color(230, 230, 230) 
        
        col1 = 45
        col2 = 145
        h = 6.5

        def row(label, val, is_bold=False):
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(col1, h, f" {label}", border="B", fill=True)
            pdf.set_font("Helvetica", "B" if is_bold else "", 9)
            pdf.cell(col2, h, f" {val}", border="B", ln=1)

        row("Claim Reference", claim_ref)
        row("Notional Amount", f"{currency} {amount:,.2f}")
        row("Value Date (Due)", d1.strftime('%d/%m/%Y'))
        row("Date Received", d2.strftime('%d/%m/%Y'))
        row("Days Late", str(days))
        row("Interest Rate", f"{rate}%")
        
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(col1, h, " Calculation", border="B", fill=True)
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(col2, h, f" ({amount:,.2f} * {days} days * {rate}%) / 360", border="B", ln=1)
        
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(col1, 9, " TOTAL CLAIM", border="B", fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(col2, 9, f" {currency} {interest:,.2f}", border="B", ln=1)
        
        pdf.ln(8)

        # Payment Box
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, "PAYMENT INSTRUCTIONS", ln=True)
        
        box_start_y = pdf.get_y()
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.2)
        
        pdf.set_xy(12, box_start_y + 3)
        pdf.set_font("Helvetica", "", 9)
        
        if ssi_info:
            def print_field(label, value, y_pos):
                pdf.set_xy(12, y_pos)
                pdf.set_font("Helvetica", "B", 9)
                pdf.cell(30, 5, label)
                pdf.set_font("Helvetica", "", 9)
                x_val = pdf.get_x()
                pdf.set_xy(x_val, y_pos)
                pdf.multi_cell(0, 5, value)
                return pdf.get_y() + 1

            curr_y = box_start_y + 3
            curr_y = print_field("Beneficiary:", ssi_info.get("Beneficiary", ""), curr_y)
            curr_y = print_field("Account With:", ssi_info.get("AccountWith", ""), curr_y)
            
            if ssi_info.get("Intermediary"):
                curr_y = print_field("Intermediary:", ssi_info.get("Intermediary", ""), curr_y)
                
            pdf.set_xy(12, curr_y)
            pdf.set_font("Helvetica", "B", 9)
            pdf.cell(30, 5, "Payment Ref:")
            pdf.set_font("Helvetica", "", 9)
            pdf.cell(0, 5, f"{claim_ref} / CLAIM", ln=True)
            box_h = pdf.get_y() - box_start_y + 3
        else:
            pdf.set_text_color(200, 0, 0)
            pdf.multi_cell(0, 5, f"No SSI instructions found for {currency}. Please refer to standard SSI.")
            pdf.set_text_color(0, 0, 0)
            box_h = 15

        pdf.rect(10, box_start_y, 190, box_h)
        
        pdf.set_y(box_start_y + box_h + 5)

        pdf.set_font("Helvetica", "", 8)
        pdf.multi_cell(0, 4,
            f"Please acknowledge this claim within 5 business days of the date of this letter and remit the total "
            f"claim amount to the account details provided above, quoting the payment reference. If the claim "
            f"remains unacknowledged, we will follow up accordingly. If you have any queries regarding this "
            f"calculation, please contact: {CONTACT_EMAIL}")
        pdf.ln(3)
        pdf.multi_cell(0, 4, "This claim is made without prejudice to any other rights or remedies available to Convera.")
        pdf.ln(5)

        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 4, "Sincerely,", ln=True)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 4, "Convera Treasury Confirmation Team", ln=True)

        try:
            pdf.output(filepath)
            self.last_generated_pdf = filepath
            self.set_status(f"PDF saved: {filepath}")
            messagebox.showinfo("Success", "PDF Generated!")
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def draft_email(self):
        if not OUTLOOK_AVAILABLE:
            messagebox.showerror("Error", "Outlook library not installed.")
            return
        
        if not self.last_generated_pdf or not os.path.exists(self.last_generated_pdf):
            messagebox.showwarning("Warning", "Please GENERATE PDF first.")
            return

        data = self.get_data()
        if not data: return
        amount, rate, days, d1, d2 = data
        interest = (amount * days * (rate/100)) / 360
        
        cp_name = self.cp_var.get()
        currency = self.curr_combo.get()
        ref_suffix = self.ref_suffix_combo.get()
        claim_ref = f"{self.ref_prefix}{ref_suffix}"
        
        date_due_str = d1.strftime("%d %b %Y")
        date_rec_str = d2.strftime("%d %b %Y")
        recipient_email = self.counterparties.get(cp_name, "")
        if not recipient_email:
            proceed = messagebox.askyesno(
                "No Email On File",
                f"No email address is saved for:\n\n{cp_name}\n\n"
                "You can add one via 'Manage Counterparties'.\n"
                "Draft the email anyway (recipient left blank)?"
            )
            if not proceed:
                return

        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = CONTACT_EMAIL
            if recipient_email: mail.To = recipient_email
            
            mail.Subject = f"Interest Claim for Notional Amount {currency} {amount:,.2f} / VD {date_due_str} / {claim_ref}"
            mail.Display()
            
            paragraphs = [
                "Dear All,",
                f"We were due to receive {currency} {amount:,.2f} for value {date_due_str}. However, funds were received on {date_rec_str}.",
                f"We hereby claim compensation of <b>{currency} {interest:,.2f}</b>, calculated at {rate}% for {days} day(s) "
                f"({currency} {amount:,.2f} × {days} × {rate}% / 360). This represents the liquidity cost / cost of funds "
                f"incurred by Convera due to the late settlement.",
                f"Please acknowledge this claim within <b>5 business days</b> and arrange settlement to the account detailed "
                f"in the attached claim letter, quoting payment reference <b>{claim_ref} / CLAIM</b>.",
                "Kindly let us know if you have any questions regarding this claim. If you aren’t the intended department, "
                "please help to route this to the concerned team.",
            ]
            body_html = "".join(
                f'<p style="font-family: Calibri, sans-serif; font-size: 11pt; margin: 0 0 12px 0;">{p}</p>'
                for p in paragraphs
            )

            # Outlook's default HTMLBody starts with empty "cursor" paragraphs
            # before the signature; prepending text above them creates a large
            # gap between the body and the signature. Strip those empty
            # paragraphs and inject our text inside the <body> tag instead of
            # concatenating two full HTML documents.
            signature_html = mail.HTMLBody
            signature_html = re.sub(
                r"<p\b[^>]*>(?:\s|&nbsp;|<o:p>|</o:p>|<br[^>]*>)*</p>",
                "", signature_html, flags=re.IGNORECASE)
            body_open = re.search(r"<body[^>]*>", signature_html, re.IGNORECASE)
            if body_open:
                idx = body_open.end()
                mail.HTMLBody = signature_html[:idx] + body_html + signature_html[idx:]
            else:
                mail.HTMLBody = body_html + signature_html
            mail.Attachments.Add(self.last_generated_pdf)
            if self.ssi_pdf_path and os.path.exists(self.ssi_pdf_path):
                mail.Attachments.Add(self.ssi_pdf_path)
            self.set_status(f"Email drafted for {cp_name}")
                
        except Exception as e:
            messagebox.showerror("Outlook Error", f"Failed to draft email:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ClaimApp(root)
    root.mainloop()