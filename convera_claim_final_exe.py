import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from fpdf import FPDF
from datetime import datetime
import os
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

# --- EMAIL MAPPING ---
COUNTERPARTY_EMAILS = {
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

# --- Data Lists ---
CURRENCY_LIST = sorted([
    "AED", "AUD", "BGN", "BHD", "BWP", "CAD", "CHF", "CNH", "CZK", "DKK", 
    "EUR", "FJD", "GBP", "GHS", "GMD", "HKD", "HUF", "ILS", "ISK", "JOD", 
    "JPY", "KES", "KWD", "MAD", "MGA", "MUR", "MWK", "MXN", "MZN", "NAD", 
    "NGN", "NOK", "NZD", "OMR", "PGK", "PLN", "QAR", "RON", "SAR", "SBD", 
    "SEK", "SGD", "THB", "TND", "TRY", "TZS", "UGX", "USD", "VUV", "WST", 
    "XAF", "XOF", "XPF", "ZAR", "ZMW"
])

COUNTERPARTY_LIST = sorted([
    "BANK OF AMERICA, N.A., New Castle",
    "Bank of Montreal Toronto, Montreal",
    "BARCLAYS BANK PLC, Washington",
    "CITI EUROPE, DUBLIN",
    "CITIBANK, NATIONAL ASSOCIATION, New Castle",
    "CROWN AGENTS BANK LTD, SUTTON",
    "Deutsche Bank AG, FRANKFURT",
    "Fifth Third Bank, National Associat, CINCINNATI",
    "Macquarie Group Ltd, SYDNEY",
    "Merrill Lynch International, LONDON",
    "Mizuho Capital Markets Corporation, NEW YORK",
    "MORGAN STANLEY & CO. INTERNATIONAL, LONDON",
    "Nomura International PLC, LONDON",
    "Royal Bank of Canada, TORONTO",
    "Societe Generale SA, PARIS",
    "State Street Bank and Trust Company, BOSTON",
    "U.S. BANK N.A., MINNEAPOLIS",
    "UBS AG, Washington",
    "WELLS FARGO BANK, NATIONAL ASSOCIAT, NEW YORK"
])

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
        self.root.geometry("600x750")
        self.root.configure(bg="#f5f5f5")
        
        self.ssi_data = {} 
        self.ssi_pdf_path = None
        self.last_generated_pdf = None

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
        main_frame = ttk.Frame(root, padding="30")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Generate Claim Letter", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 20))

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
        self.cp_combo = ttk.Combobox(grid_frame, textvariable=self.cp_var, values=COUNTERPARTY_LIST)
        self.cp_combo.grid(row=0, column=1, sticky="ew", padx=(20, 0))
        ttk.Label(grid_frame, text="(Select or Type)", font=("Segoe UI", 8), foreground="gray").grid(row=1, column=1, sticky="w", padx=(20, 0))

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
        self.amount_entry = ttk.Entry(amt_frame, width=20)
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

        # Rate
        ttk.Label(grid_frame, text="Rate (%):", style="Header.TLabel").grid(row=5, column=0, sticky="w", pady=10)
        self.rate_entry = ttk.Entry(grid_frame, width=10)
        self.rate_entry.grid(row=5, column=1, sticky="w", padx=(20, 0))

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=40, fill=tk.X)
        
        gen_btn = tk.Button(btn_frame, text="GENERATE PDF", command=self.generate_pdf, 
                           bg="#0096C8", fg="white", font=("Segoe UI", 11, "bold"), 
                           relief="flat", padx=20, pady=10)
        gen_btn.pack(side=tk.LEFT, padx=(50, 10))

        email_btn = tk.Button(btn_frame, text="DRAFT EMAIL", command=self.draft_email, 
                           bg="#28a745", fg="white", font=("Segoe UI", 11, "bold"), 
                           relief="flat", padx=20, pady=10)
        email_btn.pack(side=tk.LEFT, padx=10)

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
        pdf.cell(0, 6, f"SUBJECT: Compensation Claim for Late Payment", ln=True)
        pdf.ln(3)
        
        # Body
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 5, "Dear Customer,\n\nWe are writing to formally claim interest regarding the late settlement of the transaction detailed below. Please arrange payment of the compensation cost calculated as follows:")
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
        pdf.multi_cell(0, 4, f"Please remit the total claim amount to the account details provided above. If you have any queries regarding this calculation, please contact: {CONTACT_EMAIL}")
        pdf.ln(5)
        
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 4, "Sincerely,", ln=True)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 4, "Convera Settlements Team", ln=True)

        try:
            pdf.output(filepath)
            self.last_generated_pdf = filepath
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
        recipient_email = COUNTERPARTY_EMAILS.get(cp_name, "")
        
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.SentOnBehalfOfName = CONTACT_EMAIL
            if recipient_email: mail.To = recipient_email
            
            mail.Subject = f"Interest Claim for Notional Amount {currency} {amount:,.2f} / VD {date_due_str} / {claim_ref}"
            mail.Display()
            
            body_html = f"""
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">Dear All,</p>
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">We were due to receive {currency} {amount:,.2f} for value {date_due_str}. However, funds were received on {date_rec_str}.</p>
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">Hence, we have incurred a cost of {currency} {interest:,.2f} calculated at {rate}% for {days} day(s).</p>
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">This claim represents the liquidity cost / cost of funds incurred by Convera due to the late settlement.</p>
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">Please acknowledge our claim and pay our cost at your earliest convenience.</p>
            <p style="font-family: Calibri, sans-serif; font-size: 11pt;">Kindly let us know if you have any questions regarding this claim. If you aren’t the intended department, please help to route this to the concerned team.</p>
            <br>
            """
            mail.HTMLBody = body_html + mail.HTMLBody
            mail.Attachments.Add(self.last_generated_pdf)
            if self.ssi_pdf_path and os.path.exists(self.ssi_pdf_path):
                mail.Attachments.Add(self.ssi_pdf_path)
                
        except Exception as e:
            messagebox.showerror("Outlook Error", f"Failed to draft email:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ClaimApp(root)
    root.mainloop()