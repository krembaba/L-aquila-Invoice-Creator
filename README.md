# L'AQUILA Invoice Creator

Creates L'AQUILA Company invoices (Word `.docx`) from a simple form, using
`invoice_template.docx` as the layout.

## Requirements

Python 3.9+ for Windows from [python.org](https://www.python.org/downloads/) — tick
**"Add python.exe to PATH"** when installing. Nothing else to install.

## Making an invoice

1. Double-click **`Create Invoice.bat`**.
2. Pick a saved client (or type a new one), check the date and invoice number.
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

## Changing company details

Open `invoice_template.docx` in Word and edit the text (address, bank account, logo).