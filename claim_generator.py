import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from fpdf import FPDF
from datetime import datetime
import os

# --- Configuration: Convera's Details ---
COMPANY_NAME = "CONVERA UK LIMITED"
COMPANY_ADDR = [
    "ALPHABETA BUILDING",
    "14-18 FINSBURY SQUARE",
    "LONDON EC2A 1AH",
    "UNITED KINGDOM"
]
COMPANY_EMAIL = "claims@convera.com"
COMPANY_PHONE = "+44 20 7000 0000"

# --- Bank Details for Payment ---
BANK_DETAILS = {
    "Bank Name": "FIRST ABU DHABI BANK",
    "Swift Code": "NBADAEAAXXX",
    "Account Number": "AE920354022003572526010",
    "Beneficiary": "CONVERA UK LIMITED",
}

class ModernClaimPDF(FPDF):
    def header(self):
        # Top Bar Line
        self.set_draw_color(200, 200, 200) # Light Gray
        self.set_line_width(1)
        self.line(10, 25, 200, 25)
        
        # Company Name (Left)
        self.set_font('Helvetica', 'B', 20)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, 'Convera.', ln=True)
        
        # Document Title (Right - manually positioned)
        self.set_xy(10, 10)
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, 'INTEREST CLAIM STATEMENT', align='R', ln=True)
        self.ln(10)

    def footer(self):
        self.set_y(-20)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 5, f'{COMPANY_NAME} - Registered in England & Wales', align='C', ln=1)
        self.cell(0, 5, f'Page {self.page_no()}', align='C')

class ClaimApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convera Claims Manager")
        self.root.geometry("550x700")
        
        # Styling
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabel", font=("Arial", 10))
        style.configure("TButton", font=("Arial", 10, "bold"))
        
        # --- Data ---
        self.counterparties = ["SOCIETE GENERALE", "BARCLAYS", "JP MORGAN", "HSBC", "CITIBANK"]

        # --- Layout ---
        main_frame = ttk.Frame(root, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main_frame, text="Generate Claim Letter", font=("Arial", 16, "bold")).pack(pady=(0, 20))

        # Input Grid
        grid_frame = ttk.Frame(main_frame)
        grid_frame.pack(fill=tk.BOTH, expand=True)

        # Row 1: Counterparty
        ttk.Label(grid_frame, text="Counterparty Name:").grid(row=0, column=0, sticky="w", pady=8)
        self.cp_var = tk.StringVar()
        self.cp_combo = ttk.Combobox(grid_frame, textvariable=self.cp_var, values=self.counterparties, width=30)
        self.cp_combo.grid(row=0, column=1, sticky="w", pady=8)

        # Row 2: Address (Multiline)
        ttk.Label(grid_frame, text="Counterparty Address:").grid(row=1, column=0, sticky="nw", pady=8)
        self.cp_addr_text = tk.Text(grid_frame, height=3, width=32, font=("Arial", 9))
        self.cp_addr_text.grid(row=1, column=1, sticky="w", pady=8)

        # Row 3: Reference
        ttk.Label(grid_frame, text="Our Reference (Deal ID):").grid(row=2, column=0, sticky="w", pady=8)
        self.ref_entry = ttk.Entry(grid_frame, width=32)
        self.ref_entry.grid(row=2, column=1, sticky="w", pady=8)

        # Row 4: Currency & Amount
        ttk.Label(grid_frame, text="Currency & Notional:").grid(row=3, column=0, sticky="w", pady=8)
        amt_frame = ttk.Frame(grid_frame)
        amt_frame.grid(row=3, column=1, sticky="w", pady=8)
        
        self.curr_combo = ttk.Combobox(amt_frame, values=["USD", "EUR", "GBP", "AED", "JPY", "CAD"], width=5)
        self.curr_combo.current(0)
        self.curr_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        self.amount_entry = ttk.Entry(amt_frame, width=23)
        self.amount_entry.pack(side=tk.LEFT)

        # Row 5: Due Date
        ttk.Label(grid_frame, text="Value Date (Due):").grid(row=4, column=0, sticky="w", pady=8)
        self.date_due = DateEntry(grid_frame, width=15, background='#333', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_due.grid(row=4, column=1, sticky="w", pady=8)

        # Row 6: Receive Date
        ttk.Label(grid_frame, text="Date Received:").grid(row=5, column=0, sticky="w", pady=8)
        self.date_rec = DateEntry(grid_frame, width=15, background='#333', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.date_rec.grid(row=5, column=1, sticky="w", pady=8)

        # Row 7: Rate
        ttk.Label(grid_frame, text="Claim Rate (%):").grid(row=6, column=0, sticky="w", pady=8)
        self.rate_entry = ttk.Entry(grid_frame, width=10)
        self.rate_entry.grid(row=6, column=1, sticky="w", pady=8)

        # Output Display
        self.result_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10")
        self.result_frame.pack(fill=tk.X, pady=20)
        
        self.lbl_preview = ttk.Label(self.result_frame, text="Enter details to calculate...", foreground="#555")
        self.lbl_preview.pack(anchor="w")

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Calculate", command=self.calculate_logic).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Generate PDF Letter", command=self.generate_pdf).pack(side=tk.LEFT, padx=10)

    def get_data(self):
        try:
            amt_str = self.amount_entry.get().replace(',', '')
            amount = float(amt_str) if amt_str else 0.0
            
            rate_str = self.rate_entry.get()
            rate = float(rate_str) if rate_str else 0.0
            
            d1 = self.date_due.get_date()
            d2 = self.date_rec.get_date()
            days = (d2 - d1).days
            
            return amount, rate, days, d1, d2
        except ValueError:
            return None

    def calculate_logic(self):
        data = self.get_data()
        if data:
            amount, rate, days, _, _ = data
            currency = self.curr_combo.get()
            
            # Formula
            interest = (amount * days * (rate/100)) / 360
            
            self.lbl_preview.config(
                text=f"Days Late: {days}\n"
                     f"Claim Amount: {currency} {interest:,.2f}",
                foreground="black"
            )
            return interest
        else:
            self.lbl_preview.config(text="Invalid numbers entered.", foreground="red")
            return None

    def generate_pdf(self):
        interest = self.calculate_logic()
        if interest is None:
            messagebox.showerror("Error", "Please correct the input fields first.")
            return
            
        amount, rate, days, d1, d2 = self.get_data()
        cp_name = self.cp_var.get()
        if not cp_name:
            messagebox.showwarning("Missing Info", "Please enter a Counterparty name.")
            return

        # Prepare PDF Data
        currency = self.curr_combo.get()
        ref = self.ref_entry.get()
        if not ref: ref = "N/A"
        
        # Save Dialog
        fname = f"Claim_{cp_name}_{datetime.now().strftime('%Y%m%d')}.pdf"
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=fname)
        if not filepath: return

        # --- PDF GENERATION START ---
        pdf = ModernClaimPDF()
        pdf.add_page()
        
        # 1. Sender Info (Small text under header)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(80, 80, 80)
        for line in COMPANY_ADDR:
            pdf.cell(0, 4, line, ln=True)
        pdf.cell(0, 4, f"Tel: {COMPANY_PHONE} | Email: {COMPANY_EMAIL}", ln=True)
        pdf.ln(10)

        # 2. Recipient & Date Block
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(0, 0, 0)
        
        # Date on Right
        date_str = datetime.now().strftime("%d %B %Y")
        pdf.cell(0, 5, date_str, align='R', ln=True)
        
        # Recipient on Left
        pdf.set_y(pdf.get_y() - 5) # Go back up to align
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 5, "TO:", ln=True)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 5, cp_name, ln=True)
        
        # Address lines
        addr_lines = self.cp_addr_text.get("1.0", tk.END).strip().split('\n')
        for line in addr_lines:
            if line: pdf.cell(0, 5, line, ln=True)
            
        pdf.ln(15)

        # 3. Letter Body
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 6, f"SUBJECT: Compensation Claim for Late Payment - Ref: {ref}", ln=True)
        pdf.ln(5)
        
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 5, "Dear Customer,\n\nWe are writing to formally claim interest regarding the late settlement of the transaction detailed below. Please find the calculation and payment details enclosed.")
        pdf.ln(10)

        # 4. The "Table" (Using Cells with borders)
        # Header Row
        pdf.set_fill_color(240, 240, 240) # Light Gray fill
        pdf.set_font("Helvetica", "B", 10)
        
        col_w_1 = 60
        col_w_2 = 130
        
        pdf.cell(col_w_1, 8, "Description", border=1, fill=True)
        pdf.cell(col_w_2, 8, "Details", border=1, ln=1, fill=True)
        
        # Data Rows
        pdf.set_font("Helvetica", "", 10)
        
        def add_table_row(label, value):
            pdf.cell(col_w_1, 8, f" {label}", border='LBR') # Left, Bottom, Right border
            pdf.cell(col_w_2, 8, f" {value}", border='LBR', ln=1)

        add_table_row("Transaction Reference", ref)
        add_table_row("Notional Amount", f"{currency} {amount:,.2f}")
        add_table_row("Value Date (Due)", d1.strftime('%d/%m/%Y'))
        add_table_row("Date Received", d2.strftime('%d/%m/%Y'))
        add_table_row("Days Late", str(days))
        add_table_row("Interest Rate", f"{rate}%")
        
        # Calculation Row (formula)
        pdf.cell(col_w_1, 8, " Formula Used", border='LBR')
        pdf.set_font("Courier", "", 9) # Monospace for math looks technical
        pdf.cell(col_w_2, 8, f" ( {amount:,.2f} * {days} days * {rate}% ) / 360", border='LBR', ln=1)
        
        # Total Row (Highlighted)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(col_w_1, 10, " TOTAL CLAIM", border='LBR')
        pdf.cell(col_w_2, 10, f" {currency} {interest:,.2f}", border='LBR', ln=1)
        
        pdf.ln(10)

        # 5. Payment Box
        pdf.set_draw_color(0, 0, 0)
        pdf.rect(pdf.get_x(), pdf.get_y(), 190, 45) # Draw a rectangle box
        
        pdf.set_xy(pdf.get_x() + 5, pdf.get_y() + 5)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 5, "PAYMENT INSTRUCTIONS", ln=True)
        pdf.ln(2)
        
        pdf.set_font("Helvetica", "", 10)
        # Create quote reference
        quote_ref = f"{ref} / CLAIM"
        
        # Payment details inside box
        pdf.cell(40, 6, "Beneficiary:")
        pdf.cell(0, 6, BANK_DETAILS["Beneficiary"], ln=True)
        pdf.set_x(15)
        pdf.cell(40, 6, "Bank Name:")
        pdf.cell(0, 6, BANK_DETAILS["Bank Name"], ln=True)
        pdf.set_x(15)
        pdf.cell(40, 6, "Swift/BIC:")
        pdf.cell(0, 6, BANK_DETAILS["Swift Code"], ln=True)
        pdf.set_x(15)
        pdf.cell(40, 6, "Account No:")
        pdf.cell(0, 6, BANK_DETAILS["Account Number"], ln=True)
        pdf.set_x(15)
        pdf.cell(40, 6, "Payment Ref:")
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, quote_ref, ln=True)

        pdf.set_y(pdf.get_y() + 10) # Move below box

        # 6. Closing
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 5, "Please remit the total claim amount to the account details provided above. If you have any queries regarding this calculation, please contact our settlements team.")
        pdf.ln(10)
        pdf.cell(0, 5, "Sincerely,", ln=True)
        pdf.cell(0, 5, "Convera Settlements Team", ln=True)

        try:
            pdf.output(filepath)
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = ClaimApp(root)
    root.mainloop()