# L'AQUILA Invoice Creator

Creates L'AQUILA Company invoices (Word `.docx`) from a simple form, using
`invoice_template.docx` as the layout.

## Using it on any computer (no Python needed)

Copy **`Invoice Creator.exe`** into a folder on that computer (e.g. a folder on the Desktop)
and double-click it. Python is built in, and so is the invoice template. Invoices, the log
and the client list are created next to the `.exe`.

The first time, Windows may show *"Windows protected your PC"* because the program isn't
from a known publisher: click **More info → Run anyway**.

To use your own edited template, put `invoice_template.docx` next to the `.exe` — it is
used instead of the built-in one.

## Running from the code

Needs Python 3.9+ from [python.org](https://www.python.org/downloads/). Double-click
**`Create Invoice.bat`** — it uses `Invoice Creator.exe` if it's in the folder, otherwise it
finds Python on its own.

To make a new `.exe` after changing the code, double-click **`build_exe.bat`**.

## Making an invoice

1. Open the program (`Invoice Creator.exe` or `Create Invoice.bat`).
2. Pick a saved client (or type a new one), check the date and invoice number, and choose the
   invoice language at the top: **Shqip**, **English** or **Srpski**. Each client remembers the
   language last used for them. Item descriptions are printed exactly as typed.
3. Add the items: description, quantity, unit price. Totals and 18% VAT update as you type.
4. Click **Create invoice** (or press Ctrl+Enter), then **Open** to see it in Word.

Invoices are saved to `Faturat/<year>/` as `135. 0135-2026 CLIENT.docx` and numbered
automatically. Every invoice is also logged to `invoices.csv` (opens in Excel), and new
clients are remembered in `clients.json`. These three stay on the computer and are not
uploaded to GitHub.

## Without the form

```
python invoice_generator.py --json example_invoice.json
python invoice_generator.py --json example_invoice.json --out test.docx   # preview only, no numbering/log
```

Add `"language": "en"` (or `"sq"`, `"sr"`) to the JSON to pick the invoice language.
The labels for each language are in `LABELS` near the top of `invoice_generator.py`.

## Changing company details

Open `invoice_template.docx` in Word and edit the text (address, bank account, logo).