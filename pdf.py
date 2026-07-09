import re
import html
import sys
from datetime import datetime
from tkinter import Tk, Frame, Label, Button, Checkbutton, IntVar, END, BOTH, X, LEFT, RIGHT
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog, messagebox

# ReportLab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics

# ----------------------------
# Helpers
# ----------------------------

def unescape_maybe_html(text: str) -> str:
    """
    Convert HTML-escaped content (e.g., &lt;Document&gt;) to real XML.
    Also normalize non-breaking spaces to normal spaces.
    """
    unescaped = html.unescape(text or "")
    # Replace non-breaking spaces with regular spaces
    return unescaped.replace("\xa0", " ")

def extract_xml_body(text: str) -> str:
    """
    Heuristic to extract the XML from pasted content.
    - If user pasted HTML with <table><pre><span>... &lt;Document&gt; ...</span></pre>,
      this will unescape and then take everything from the first '<' to the last '>'.
    - Otherwise returns the input after unescaping.
    """
    s = unescape_maybe_html(text).strip()

    # If there is any XML start tag, slice from the first '<' through last '>'
    first_lt = s.find("<")
    last_gt = s.rfind(">")
    if first_lt != -1 and last_gt != -1 and last_gt > first_lt:
        s = s[first_lt:last_gt + 1].strip()

    return s

def find_uetr(xml_text: str) -> str | None:
    """
    Extract UETR from XML. Works regardless of namespaces or whitespace.
    Examples it matches:
      <UETR>...</UETR>
      <ns:UETR>...</ns:UETR>
    """
    m = re.search(r"<(?:[A-Za-z0-9_]+:)?UETR>\s*([^<\s]+)\s*</(?:[A-Za-z0-9_]+:)?UETR>", xml_text, re.IGNORECASE)
    return m.group(1) if m else None

def pretty_xml(xml_text: str) -> str:
    """
    Try to pretty-print XML. If parsing fails, return original text unchanged.
    """
    try:
        from xml.dom import minidom
        # minidom needs a bytes-like object or a clean string
        parsed = minidom.parseString(xml_text.encode("utf-8"))
        return parsed.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
    except Exception:
        return xml_text

# ------------- PDF text wrapping for monospaced output -------------

def _wrap_line_to_width(line: str, max_width_pts: float, font_name: str, font_size: float):
    """
    Break a single line into chunks that fit into max_width_pts, measuring actual width.
    Falls back to hard splits when there are no good whitespace breakpoints.
    """
    chunks = []
    i = 0
    n = len(line)

    # Quick escape: empty line
    if n == 0:
        return [""]

    while i < n:
        # Expand as much as fits
        lo, hi = 1, n - i
        best = 1
        # First try to find the largest fitting slice
        while lo <= hi:
            mid = (lo + hi) // 2
            segment = line[i:i + mid]
            width = pdfmetrics.stringWidth(segment, font_name, font_size)
            if width <= max_width_pts:
                best = mid
                lo = mid + 1
            else:
                hi = mid - 1

        # best is the largest count that fits from position i
        segment = line[i:i + best]

        # If we didn't reach the end and we broke mid-word, try to step back to last whitespace
        if i + best < n and " " in segment:
            last_space = segment.rfind(" ")
            if last_space > 0:
                segment = segment[:last_space]
                best = last_space

        if not segment:
            # Nothing fits (extremely narrow width); hard-split one char to avoid infinite loop
            segment = line[i:i+1]
            best = 1

        chunks.append(segment)
        i += best

        # Skip a single space right after a wrap to avoid leading spaces at line start
        if i < n and line[i] == " ":
            i += 1

    return chunks

def xml_to_pdf(xml_text: str, out_path: str, title: str | None = None):
    """
    Create a copy-pastable PDF with monospaced text and soft wrapping.
    """
    page_w, page_h = A4
    margin_left = 36   # 0.5 inch
    margin_right = 36
    margin_top = 48
    margin_bottom = 36

    font_name = "Courier"     # built-in, monospaced (good for XML)
    font_size = 9
    line_gap = 12             # leading
    max_text_width = page_w - margin_left - margin_right

    c = canvas.Canvas(out_path, pagesize=A4)
    if title:
        try:
            c.setTitle(title)
        except Exception:
            pass

    # Use a textobject for selectable text
    def new_textobject():
        t = c.beginText()
        t.setTextOrigin(margin_left, page_h - margin_top)
        t.setFont(font_name, font_size)
        return t

    text_obj = new_textobject()
    y_min = margin_bottom

    lines = xml_text.splitlines()

    for raw_line in lines:
        # Keep original indentation and characters
        if raw_line.strip() == "":
            # Blank line
            if text_obj.getY() - line_gap < y_min:
                c.drawText(text_obj)
                c.showPage()
                text_obj = new_textobject()
            text_obj.textLine("")  # advance one line
            continue

        # Wrap line by actual width
        wrapped = _wrap_line_to_width(raw_line, max_text_width, font_name, font_size)
        for seg in wrapped:
            if text_obj.getY() - line_gap < y_min:
                c.drawText(text_obj)
                c.showPage()
                text_obj = new_textobject()
            text_obj.textLine(seg)

    # Flush last page
    c.drawText(text_obj)
    c.save()

# ----------------------------
# GUI
# ----------------------------

class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("PACS XML → PDF (UETR Named)")

        top = Frame(root)
        top.pack(fill=X, padx=8, pady=(8, 4))

        Label(top, text="Paste pacs.* XML (raw or HTML-escaped) below:").pack(side=LEFT)

        self.pretty_var = IntVar(value=1)
        Checkbutton(top, text="Pretty-print XML", variable=self.pretty_var).pack(side=RIGHT)

        # Editor
        self.editor = ScrolledText(root, wrap="none", undo=True, height=24)
        self.editor.pack(fill=BOTH, expand=True, padx=8, pady=4)

        # Buttons
        bottom = Frame(root)
        bottom.pack(fill=X, padx=8, pady=(4, 8))

        Button(bottom, text="Create PDF", command=self.create_pdf).pack(side=LEFT)
        Button(bottom, text="Clear", command=self.clear).pack(side=LEFT, padx=(8, 0))
        Button(bottom, text="Quit", command=root.quit).pack(side=RIGHT)

    def clear(self):
        self.editor.delete("1.0", END)

    def create_pdf(self):
        raw = self.editor.get("1.0", END).strip()
        if not raw:
            messagebox.showwarning("Empty", "Please paste a pacs.* XML first.")
            return

        # Extract XML
        xml_text = extract_xml_body(raw)

        # Optional pretty-print
        if self.pretty_var.get() == 1:
            xml_text = pretty_xml(xml_text)

        # Determine filename from UETR
        uetr = find_uetr(xml_text)
        if uetr:
            default_name = f"{uetr}.pdf"
            pdf_title = f"UETR {uetr}"
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"pacs_{stamp}.pdf"
            pdf_title = "PACS Document"

        # Ask where to save
        out_path = filedialog.asksaveasfilename(
            title="Save PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not out_path:
            return

        # Create the PDF
        try:
            xml_to_pdf(xml_text, out_path, title=pdf_title)
        except Exception as e:
            messagebox.showerror("PDF Error", f"Failed to create PDF:\n{e}")
            return

        messagebox.showinfo("Done", f"Saved:\n{out_path}")

def main():
    root = Tk()
    # Reasonable default geometry
    root.geometry("900x650")
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
    