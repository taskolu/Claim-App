# Convera Claims Manager

Desktop app (Tkinter) for generating late-settlement compensation claim letters (PDF) and drafting the corresponding Outlook emails to counterparties.

**Main application:** `convera_claim_final_exe.py` (packaged with `Convera Claim Manager.spec` via PyInstaller).

## Features

- **Claim letter PDF generation** — branded letter with claim reference, notional, dates, days late, rate, calculated interest, and payment instructions pulled from the SSI PDF for the selected currency.
- **Outlook email drafting** — pre-filled subject/body with the claim PDF and SSI attached, addressed to the counterparty's saved email(s).
- **Counterparty management** — click **⚙ Manage Counterparties** in the top-right to:
  - **Add** new counterparties with their email address(es)
  - **Edit** names and email addresses (use `;` to separate multiple addresses)
  - **Remove** counterparties you no longer deal with
  - **Search** the list by name or email

  Changes are saved instantly to `~/.convera_claim_manager/counterparties.json`, so they persist across restarts **and app updates** — no more editing the code when an email address changes. On first run the file is seeded with the built-in default list. If the file ever becomes corrupted it is backed up as `counterparties.json.bak` and the defaults are restored.

- When you pick a counterparty in the main form, its on-file email is shown underneath so you can verify it before drafting. If none is saved, you'll be warned before an email is drafted.

## Setup

```bash
pip install tkcalendar fpdf pdfplumber pywin32
python convera_claim_final_exe.py
```

`pywin32` (Outlook drafting) and `os.startfile` are Windows-only; PDF generation works without Outlook.

## Building the EXE

```bash
pyinstaller "Convera Claim Manager.spec"
```

## Other scripts in this repo

- `app.py` — Window Rescue, a separate multi-monitor window manager utility
- `claim_generator.py` — earlier prototype of the claim letter generator
- `htmltopdf.py`, `pdf.py`, `test.py`, `test2.py`, `timetray.py` — misc utilities/experiments
