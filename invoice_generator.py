"""
L'AQUILA Company - invoice generator.

Fills invoice_template.docx (the 0134/2026 invoice turned into a template) and saves
the result to Faturat/<year>/, numbering invoices automatically and logging each one
to invoices.csv (opens in Excel).

    python invoice_generator.py                     -> opens the form
    python invoice_generator.py --json data.json    -> creates an invoice without the form
    python invoice_generator.py --json data.json --out test.docx
                                                    -> writes exactly there, no numbering/log
"""
import argparse
import csv
import json
import os
import re
import sys
import zipfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "invoice_template.docx"
CLIENTS_FILE = HERE / "clients.json"
LOG_FILE = HERE / "invoices.csv"
OUTPUT_DIR = HERE / "Faturat"

DEFAULT_VAT = Decimal("18")
# Last invoice number made by hand, per year. Numbering continues from here.
LAST_MANUAL_NUMBER = {2026: 134}

LOG_FIELDS = ["created", "date", "year", "number", "invoice_no", "client",
              "subtotal", "vat", "total", "file"]
CLIENT_FIELDS = ["name", "short_name", "nui", "address", "city", "email", "language"]

# Invoice languages. Each label fills a {{L_*}} slot in the template; the Albanian set is
# the original invoice word for word (including its English payment/total labels).
LANGUAGES = {"sq": "Shqip", "en": "English", "sr": "Srpski"}
DEFAULT_LANGUAGE = "sq"
LABELS = {
    "sq": {
        "L_COMPANY_NUI": "NUI:", "L_COMPANY_VAT_NO": "Nr. TVSH: ", "L_COMPANY_ADDRESS": "Adresa: ",
        "L_TITLE": " FATURË", "L_BILL_TO": "FATURË PËR:", "L_DATE": "Data :", "L_INVOICE_NO": "Fatura Nr :",
        "L_CLIENT_NUI": "NUI: ", "L_ADDRESS": "Adresa: ",
        "L_COL_DESC": "PËRSHKRIMI", "L_COL_QTY": "SASIA", "L_COL_PRICE": "ÇMIMI", "L_COL_VAT": "TVSH",
        "L_COL_TOTAL": "TOTALI ",
        "L_PAYMENT": "Payment method", "L_SWIFT": "SWIFT code: ", "L_ACCOUNT": "   Account no:     ",
        "L_SUBTOTAL": "Sub-total:", "L_VAT": " VAT:", "L_TOTAL": "Total :",
        "L_SIGNATURE": "Signature:", "L_RECEIVED": "Received: ",
    },
    "en": {
        "L_COMPANY_NUI": "Business ID: ", "L_COMPANY_VAT_NO": "VAT No.: ", "L_COMPANY_ADDRESS": "Address: ",
        "L_TITLE": "INVOICE", "L_BILL_TO": "BILL TO:", "L_DATE": "Date: ", "L_INVOICE_NO": "Invoice No: ",
        "L_CLIENT_NUI": "Business ID: ", "L_ADDRESS": "Address: ",
        "L_COL_DESC": "DESCRIPTION", "L_COL_QTY": "QTY", "L_COL_PRICE": "PRICE", "L_COL_VAT": "VAT",
        "L_COL_TOTAL": "TOTAL",
        "L_PAYMENT": "Payment method", "L_SWIFT": "SWIFT code: ", "L_ACCOUNT": "   Account no:     ",
        "L_SUBTOTAL": "Sub-total:", "L_VAT": " VAT:", "L_TOTAL": "Total: ",
        "L_SIGNATURE": "Signature:", "L_RECEIVED": "Received: ",
    },
    "sr": {
        "L_COMPANY_NUI": "JIB: ", "L_COMPANY_VAT_NO": "PDV br.: ", "L_COMPANY_ADDRESS": "Adresa: ",
        "L_TITLE": "FAKTURA", "L_BILL_TO": "KUPAC:", "L_DATE": "Datum: ", "L_INVOICE_NO": "Faktura br.: ",
        "L_CLIENT_NUI": "JIB: ", "L_ADDRESS": "Adresa: ",
        "L_COL_DESC": "OPIS", "L_COL_QTY": "KOL.", "L_COL_PRICE": "CENA", "L_COL_VAT": "PDV",
        "L_COL_TOTAL": "UKUPNO",
        "L_PAYMENT": "Način plaćanja", "L_SWIFT": "SWIFT kod: ", "L_ACCOUNT": "   Broj računa:     ",
        "L_SUBTOTAL": "Osnovica:", "L_VAT": " PDV:", "L_TOTAL": "Ukupno: ",
        "L_SIGNATURE": "Potpis:", "L_RECEIVED": "Primio: ",
    },
}
# Client lines that disappear from the invoice when left empty.
OPTIONAL_PARAGRAPHS = ["CLIENT_NUI", "CLIENT_ADDRESS", "CLIENT_CITY", "CLIENT_EMAIL"]
# Empty paragraphs under the items table; one is dropped per extra item so the page holds.
SPACER = '<w:p><w:pPr><w:rPr><w:sz w:val="27"/></w:rPr></w:pPr></w:p>'

CENT = Decimal("0.01")


# ---------------------------------------------------------------- numbers & money

def parse_number(text):
    """Accepts 800, 800.50, 800,50, 1,200.50, '1 200,50 €'."""
    t = str(text).strip().replace("€", "").replace(" ", "")
    if "," in t and "." in t:
        t = t.replace(",", "") if t.rindex(",") < t.rindex(".") else t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        value = Decimal(t)
    except InvalidOperation:
        raise ValueError(f"'{text}' is not a number")
    if not value.is_finite() or value < 0:
        raise ValueError(f"'{text}' is not a valid amount")
    return value


def money(value):
    return f"{value.quantize(CENT, ROUND_HALF_UP):,.2f}€"


def qty_text(value):
    text = f"{value.normalize():f}"
    return text


def parse_date(text):
    return datetime.strptime(text.strip(), "%d/%m/%Y").date()


def compute(items, vat_rate):
    """items: [{description, qty, unit_price}] -> (lines, subtotal, vat, total)."""
    lines = []
    for item in items:
        qty = parse_number(item["qty"])
        price = parse_number(item["unit_price"])
        net = (qty * price).quantize(CENT, ROUND_HALF_UP)
        vat = (net * vat_rate / 100).quantize(CENT, ROUND_HALF_UP)
        lines.append({"description": item["description"].strip(), "qty": qty,
                      "net": net, "vat": vat, "total": net + vat})
    subtotal = sum((l["net"] for l in lines), Decimal(0))
    vat = sum((l["vat"] for l in lines), Decimal(0))
    return lines, subtotal, vat, subtotal + vat


# ---------------------------------------------------------------- docx rendering

def text_xml(value):
    """Escape for <w:t>; line breaks become Word line breaks."""
    lines = escape(str(value)).split("\n")
    return '</w:t><w:br/><w:t xml:space="preserve">'.join(lines)


def fill(xml, values):
    return re.sub(r"\{\{(\w+)\}\}", lambda m: text_xml(values[m.group(1)]), xml)


def paragraph_span(xml, anchor):
    i = xml.index(anchor)
    return xml.rindex("<w:p>", 0, i), xml.index("</w:p>", i) + len("</w:p>")


def render(invoice, out_path):
    """invoice: {client{...}, invoice_no, date (str), vat_rate, items[...]}. Returns totals."""
    lines, subtotal, vat_total, total = compute(invoice["items"], invoice["vat_rate"])
    if not lines:
        raise ValueError("The invoice needs at least one item.")
    client = invoice["client"]

    with zipfile.ZipFile(TEMPLATE) as z:
        entries = [(info, z.read(info.filename)) for info in z.infolist()]
    xml = dict((info.filename, data) for info, data in entries)["word/document.xml"].decode("utf-8")

    # Split around the item row so it can be repeated per line.
    anchor = xml.index("{{ITEM_DESC}}")
    row_start = xml.rindex("<w:tr>", 0, anchor)
    row_end = xml.index("</w:tr>", anchor) + len("</w:tr>")
    head, row, tail = xml[:row_start], xml[row_start:row_end], xml[row_end:]

    for key in OPTIONAL_PARAGRAPHS:
        field = key.removeprefix("CLIENT_").lower()
        if not str(client.get(field, "")).strip():
            s, e = paragraph_span(head, "{{%s}}" % key)
            head = head[:s] + head[e:]
    for _ in range(min(len(lines) - 1, tail.count(SPACER))):
        tail = tail.replace(SPACER, "", 1)

    values = {
        "DATE": invoice["date"],
        "INVOICE_NO": invoice["invoice_no"],
        "CLIENT_NAME": client["name"].strip(),
        "CLIENT_NUI": client.get("nui", "").strip(),
        "CLIENT_ADDRESS": client.get("address", "").strip(),
        "CLIENT_CITY": client.get("city", "").strip(),
        "CLIENT_EMAIL": client.get("email", "").strip(),
        "SUBTOTAL": money(subtotal),
        "VAT_TOTAL": money(vat_total),
        "GRAND_TOTAL": money(total),
        **LABELS[invoice.get("language") or DEFAULT_LANGUAGE],
    }
    rows ="".join(fill(row, {
        "ITEM_NO": f"{n}.",
        "ITEM_DESC": line["description"],
        "ITEM_QTY": qty_text(line["qty"]),
        "ITEM_NET": money(line["net"]),
        "ITEM_VAT": money(line["vat"]),
        "ITEM_TOTAL": money(line["total"]),
    }) for n, line in enumerate(lines, start=1))
    xml = fill(head, values) + rows + fill(tail, values)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = out_path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            if info.filename == "word/document.xml":
                data = xml.encode("utf-8")
            out.writestr(info, data)
    os.replace(tmp, out_path)
    return subtotal, vat_total, total


# ---------------------------------------------------------------- numbering, log, clients

def read_log():
    if not LOG_FILE.exists():
        return []
    with open(LOG_FILE, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def next_number(year):
    used = [int(r["number"]) for r in read_log() if r["year"] == str(year)]
    return max(used, default=LAST_MANUAL_NUMBER.get(year, 0)) + 1


def number_used(year, number):
    return any(r["year"] == str(year) and r["number"] == str(number) for r in read_log())


def output_path(number, year, short_name):
    safe = re.sub(r'[<>:"/\\|?*]', "", short_name).strip() or "KLIENT"
    return OUTPUT_DIR / str(year) / f"{number}. {number:04d}-{year} {safe}.docx"


def short_name_for(client):
    return client.get("short_name", "").strip() or " ".join(client["name"].split()[:2])


def create_invoice(client, date_text, number, items, vat_rate=DEFAULT_VAT, language=DEFAULT_LANGUAGE):
    """Numbered, logged invoice in Faturat/. Returns the saved path."""
    year = parse_date(date_text).year
    path = output_path(number, year, short_name_for(client))
    client = dict(client, language=language)  # remembered as this client's language
    invoice = {"client": client, "invoice_no": f"{number:04d}/{year}", "date": date_text,
               "vat_rate": vat_rate, "items": items, "language": language}
    subtotal, vat, total = render(invoice, path)

    new_log = not LOG_FILE.exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, LOG_FIELDS)
        if new_log:
            writer.writeheader()
        writer.writerow({"created": datetime.now().strftime("%d/%m/%Y %H:%M"), "date": date_text,
                         "year": year, "number": number, "invoice_no": invoice["invoice_no"],
                         "client": client["name"], "subtotal": f"{subtotal:.2f}", "vat": f"{vat:.2f}",
                         "total": f"{total:.2f}", "file": str(path.relative_to(HERE))})
    save_client(client)
    return path


def load_clients():
    if not CLIENTS_FILE.exists():
        return []
    return json.loads(CLIENTS_FILE.read_text(encoding="utf-8"))


def save_client(client):
    clients = [c for c in load_clients() if c["name"].strip().lower() != client["name"].strip().lower()]
    clients.append({k: client.get(k, "").strip() for k in CLIENT_FIELDS})
    clients.sort(key=lambda c: c["name"].lower())
    CLIENTS_FILE.write_text(json.dumps(clients, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- form

def load_logo_png():
    """Company logo from the template, shown in the window header."""
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            name = next(n for n in z.namelist() if n.startswith("word/media/") and n.endswith(".png"))
            return z.read(name)
    except (OSError, StopIteration, zipfile.BadZipFile):
        return None


def run_gui():
    import base64
    import tkinter as tk
    from tkinter import messagebox, ttk

    try:  # sharp text on high-DPI screens
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass

    BG, CARD, BORDER, FIELD_BORDER = "#F3F4F6", "#FFFFFF", "#E4E7EC", "#D0D5DD"
    TEXT, MUTED = "#1F2328", "#667085"
    GREEN, GREEN_DARK, GREEN_SOFT = "#2EAD4B", "#238C3B", "#ECF8EF"
    RED, RED_SOFT = "#D92D20", "#FEF0EE"
    FONT = "Segoe UI"

    root = tk.Tk()
    root.title("L'AQUILA - Invoices")
    root.configure(bg=BG)
    scale = root.winfo_fpixels("1i") / 96

    def px(n):
        return int(n * scale)

    root.geometry(f"{px(1140)}x{px(780)}")
    root.minsize(px(1000), px(680))
    root.option_add("*TCombobox*Listbox.font", (FONT, 10))
    root.option_add("*TCombobox*Listbox.selectBackground", GREEN)

    # ---- styles
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=(FONT, 10), background=CARD, foreground=TEXT)
    field_look = {"fieldbackground": CARD, "bordercolor": FIELD_BORDER, "lightcolor": CARD,
                  "darkcolor": CARD, "padding": px(6)}
    focus_look = {"bordercolor": [("focus", GREEN)], "lightcolor": [("focus", GREEN)]}
    style.configure("TEntry", **field_look)
    style.map("TEntry", **focus_look)
    style.configure("Error.TEntry", **dict(field_look, bordercolor=RED, fieldbackground=RED_SOFT))
    style.configure("TCombobox", arrowcolor=MUTED, background=CARD, **field_look)
    style.map("TCombobox", fieldbackground=[("readonly", CARD)], **focus_look)
    style.configure("Vertical.TScrollbar", background=BORDER, troughcolor=CARD, bordercolor=CARD,
                    lightcolor=BORDER, darkcolor=BORDER, arrowcolor=MUTED, gripcount=0)

    def button_style(name, bg, fg, active_bg, active_fg=None, border=None, **kw):
        style.configure(name, background=bg, foreground=fg, bordercolor=border or bg, lightcolor=bg,
                        darkcolor=bg, focuscolor=bg, **kw)
        style.map(name, background=[("active", active_bg)], lightcolor=[("active", active_bg)],
                  darkcolor=[("active", active_bg)], bordercolor=[("active", border or active_bg)],
                  foreground=[("disabled", BORDER), ("active", active_fg or fg)])

    button_style("Primary.TButton", GREEN, "white", GREEN_DARK, font=(FONT, 11, "bold"), padding=(px(16), px(10)))
    button_style("Ghost.TButton", CARD, TEXT, BG, border=FIELD_BORDER, padding=(px(12), px(6)))
    button_style("Link.TButton", CARD, GREEN, GREEN_SOFT, GREEN_DARK, font=(FONT, 10, "bold"),
                 padding=(px(8), px(3)))
    button_style("Remove.TButton", CARD, MUTED, RED_SOFT, RED, padding=(px(6), px(4)), width=2)

    # ---- small builders
    def label(parent, text="", size=10, color=TEXT, bold=False, bg=CARD, **kw):
        return tk.Label(parent, text=text, bg=bg, fg=color, font=(FONT, size, "bold" if bold else "normal"), **kw)

    def card(parent, title, action=None):
        outer = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER)
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="both", expand=True, padx=px(18), pady=px(14))
        head = tk.Frame(inner, bg=CARD)
        head.pack(fill="x", pady=(0, px(8)))
        label(head, title, 12, bold=True).pack(side="left")
        if action:
            ttk.Button(head, text=action[0], style="Link.TButton", command=action[1]).pack(side="right")
        return outer, inner

    def field(parent, text, widget, row, col, colspan=1):
        label(parent, text, 9, MUTED).grid(row=row * 2, column=col, columnspan=colspan, sticky="w", padx=px(5))
        widget.grid(row=row * 2 + 1, column=col, columnspan=colspan, sticky="we", padx=px(5), pady=(px(2), px(10)))
        return widget

    def focus_next(event):
        event.widget.tk_focusNext().focus_set()
        return "break"

    def focus_prev(event):
        event.widget.tk_focusPrev().focus_set()
        return "break"

    # ---- header
    header = tk.Frame(root, bg=CARD)
    header.pack(fill="x")
    tk.Frame(root, bg=BORDER, height=1).pack(fill="x")
    logo = None
    logo_png = load_logo_png()
    if logo_png:
        try:
            logo = tk.PhotoImage(data=base64.b64encode(logo_png))
            logo = logo.subsample(max(1, round(logo.height() / px(46))))
            label(header, image=logo).pack(side="left", padx=(px(24), px(16)), pady=px(12))
        except tk.TclError:
            logo = None
    titles = tk.Frame(header, bg=CARD)
    titles.pack(side="left", padx=(0 if logo else px(24), 0), pady=px(12))
    label(titles, "New invoice", 16, bold=True).pack(anchor="w")
    subtitle = label(titles, "", 10, MUTED)
    subtitle.pack(anchor="w")
    ttk.Button(header, text="Open invoices folder", style="Ghost.TButton",
               command=lambda: open_folder()).pack(side="right", padx=px(24))

    # language switch: Shqip | English | Srpski
    language_var = tk.StringVar(value=DEFAULT_LANGUAGE)
    language_box = tk.Frame(header, bg=CARD)
    language_box.pack(side="right", padx=(0, px(8)))
    label(language_box, "Invoice language", 9, MUTED).pack(anchor="w")
    segments_frame = tk.Frame(language_box, bg=CARD, highlightthickness=1,
                              highlightbackground=FIELD_BORDER, highlightcolor=FIELD_BORDER)
    segments_frame.pack(anchor="w", pady=(px(2), 0))
    segments = {}
    for code, name in LANGUAGES.items():
        segment = label(segments_frame, name, 10, cursor="hand2", width=8, pady=px(5))
        segment.pack(side="left")
        segment.bind("<Button-1>", lambda _e, c=code: language_var.set(c))
        segment.bind("<Enter>", lambda _e, c=code: c != language_var.get() and segments[c].config(bg=GREEN_SOFT))
        segment.bind("<Leave>", lambda _e: paint_language())
        segments[code] = segment

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=px(20), pady=px(18))
    body.columnconfigure(0, weight=1)
    body.columnconfigure(1, minsize=px(320))
    body.rowconfigure(1, weight=1)
    side = tk.Frame(body, bg=BG)
    side.grid(row=0, column=1, rowspan=2, sticky="nsew")

    # ---- client card
    client_vars = {k: tk.StringVar() for k in CLIENT_FIELDS if k != "language"}
    client_vars["language"] = language_var  # remembered per client
    client_card, client_inner = card(body, "Client", ("Clear", lambda: new_client()))
    client_card.grid(row=0, column=0, sticky="nsew", padx=(0, px(16)), pady=(0, px(16)))
    client_grid = tk.Frame(client_inner, bg=CARD)
    client_grid.pack(fill="x")
    client_grid.columnconfigure((0, 1), weight=1, uniform="client")
    name_box = field(client_grid, "Company name - pick a saved client or type a new one",
                     ttk.Combobox(client_grid, textvariable=client_vars["name"], font=(FONT, 11)), 0, 0, 2)
    for key, text, row, col in (("nui", "NUI", 1, 0), ("email", "Email", 1, 1),
                                ("address", "Address", 2, 0), ("city", "City / country", 2, 1),
                                ("short_name", "Short name for the file name (optional)", 3, 0)):
        field(client_grid, text, ttk.Entry(client_grid, textvariable=client_vars[key]), row, col)

    # ---- items card
    items_card, items_inner = card(body, "Items", ("+ Add item", lambda: add_row()))
    items_card.grid(row=1, column=0, sticky="nsew", padx=(0, px(16)))
    scroller = tk.Frame(items_inner, bg=CARD)
    scroller.pack(fill="both", expand=True)
    canvas = tk.Canvas(scroller, bg=CARD, highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(scroller, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    table = tk.Frame(canvas, bg=CARD)
    table_window = canvas.create_window(0, 0, window=table, anchor="nw")
    table.columnconfigure(1, weight=1)
    for col, width in ((0, 26), (2, 74), (3, 104), (4, 110), (5, 36)):
        table.columnconfigure(col, minsize=px(width))
    for col, (text, anchor) in enumerate((("#", "w"), ("Description", "w"), ("Qty", "e"),
                                          ("Unit price €", "e"), ("Amount", "e"), ("", "w"))):
        label(table, text, 9, MUTED, anchor=anchor).grid(row=0, column=col, sticky="we", padx=px(4), pady=(0, px(6)))
    tk.Frame(table, bg=BORDER, height=1).grid(row=1, column=0, columnspan=6, sticky="we", pady=(0, px(4)))

    # ---- invoice details card
    date_var = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
    number_var = tk.StringVar(value=str(next_number(date.today().year)))
    vat_var = tk.StringVar(value=str(DEFAULT_VAT))
    details_card, details_inner = card(side, "Invoice details")
    details_card.pack(fill="x", pady=(0, px(16)))
    details = tk.Frame(details_inner, bg=CARD)
    details.pack(fill="x")
    details.columnconfigure((0, 1), weight=1, uniform="details")
    number_box = tk.Frame(details, bg=CARD)
    number_entry = ttk.Entry(number_box, textvariable=number_var, width=6)
    number_entry.pack(side="left", fill="x", expand=True)
    year_label = label(number_box, f"/{date.today().year}", 10, MUTED)
    year_label.pack(side="left", padx=(px(6), 0))
    field(details, "Invoice no.", number_box, 0, 0)
    date_entry = field(details, "Date (dd/mm/yyyy)", ttk.Entry(details, textvariable=date_var, width=11), 0, 1)
    vat_entry = field(details, "VAT %", ttk.Entry(details, textvariable=vat_var, width=6), 1, 0)

    # ---- summary card
    summary_card, summary_inner = card(side, "Summary")
    summary_card.pack(fill="x", pady=(0, px(16)))

    def summary_line(text, big=False):
        line = tk.Frame(summary_inner, bg=CARD)
        line.pack(fill="x", pady=px(2))
        name = label(line, text, 12 if big else 10, TEXT if big else MUTED, bold=big)
        name.pack(side="left", anchor="s")
        value = label(line, money(Decimal(0)), 18 if big else 10, bold=big)
        value.pack(side="right")
        return name, value

    _, subtotal_value = summary_line("Sub-total")
    vat_name, vat_value = summary_line("VAT")
    tk.Frame(summary_inner, bg=BORDER, height=1).pack(fill="x", pady=px(8))
    _, total_value = summary_line("Total", big=True)
    ttk.Button(summary_inner, text="Create invoice", style="Primary.TButton",
               command=lambda: create()).pack(fill="x", pady=(px(14), 0))
    status = tk.Frame(summary_inner, bg=CARD)
    status_text = label(status, "", 9, MUTED, wraplength=px(230), justify="left", anchor="w")
    status_text.pack(side="left", fill="x", expand=True, padx=px(10), pady=px(8))
    status_open = label(status, "Open", 10, GREEN_DARK, bold=True, cursor="hand2")

    # ---- recent invoices card
    recent_card, recent_inner = card(side, "Recent invoices")
    recent_card.pack(fill="both", expand=True)
    recent_list = tk.Frame(recent_inner, bg=CARD, height=1)
    recent_list.pack(fill="both", expand=True)
    recent_list.pack_propagate(False)
    recent_items = []

    # ---- behaviour
    rows = []

    def open_folder():
        OUTPUT_DIR.mkdir(exist_ok=True)
        os.startfile(OUTPUT_DIR)

    def open_file(path):
        if path.exists():
            os.startfile(path)
        else:
            show_status("That file was moved or deleted.", "error")

    def show_status(text, kind, path=None, focus=None):
        bg, fg = {"success": (GREEN_SOFT, GREEN_DARK), "error": (RED_SOFT, RED)}[kind]
        status.pack(fill="x", pady=(px(10), 0))
        status.config(bg=bg)
        status_text.config(text=text, bg=bg, fg=fg)
        status_open.config(bg=bg)
        status_open.pack_forget()
        if path:
            status_open.pack(side="right", padx=px(10))
            status_open.bind("<Button-1>", lambda _e: open_file(path))
        if focus is not None:
            focus.focus_set()

    def client_names():
        return [c["name"] for c in load_clients()]

    def pick_client(_event=None):
        typed = client_vars["name"].get().strip().lower()
        match = next((c for c in load_clients() if c["name"].lower() == typed), None)
        if match:
            for key in CLIENT_FIELDS:
                if key != "language" or match.get(key):  # older clients have no language saved
                    client_vars[key].set(match.get(key, ""))

    def paint_language(*_):
        if language_var.get() not in LANGUAGES:
            language_var.set(DEFAULT_LANGUAGE)
            return
        for code, segment in segments.items():
            chosen = code == language_var.get()
            segment.config(bg=GREEN if chosen else CARD, fg="white" if chosen else TEXT,
                           font=(FONT, 10, "bold" if chosen else "normal"))

    def filter_clients(event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        typed = client_vars["name"].get().strip().lower()
        name_box.config(values=[n for n in client_names() if typed in n.lower()])

    def new_client():
        for var in client_vars.values():
            var.set("")
        name_box.config(values=client_names())
        name_box.focus_set()

    def sync_scroll(_event=None):
        canvas.configure(scrollregion=(0, 0, table.winfo_reqwidth(), table.winfo_reqheight()))
        if table.winfo_reqheight() > canvas.winfo_height():
            if not scrollbar.winfo_ismapped():
                scrollbar.pack(side="right", fill="y", before=canvas)
        else:
            scrollbar.pack_forget()
            canvas.yview_moveto(0)

    def wheel(event):
        if scrollbar.winfo_ismapped():
            canvas.yview_scroll(int(-event.delta / 120), "units")

    def recalc(*_):
        try:
            vat_rate = parse_number(vat_var.get())
            vat_entry.configure(style="TEntry")
        except ValueError:
            vat_rate = Decimal(0)
            vat_entry.configure(style="Error.TEntry")
        valid = []
        for row in rows:
            ok = True
            for var, entry in ((row["qty"], row["qty_entry"]), (row["price"], row["price_entry"])):
                try:
                    parse_number(var.get())
                    entry.configure(style="TEntry")
                except ValueError:
                    ok = False
                    entry.configure(style="Error.TEntry" if var.get().strip() else "TEntry")
            if ok:
                (line,), *_ = compute([{"description": "", "qty": row["qty"].get(),
                                        "unit_price": row["price"].get()}], vat_rate)
                row["amount"].config(text=money(line["net"]))
                valid.append({"description": "", "qty": row["qty"].get(), "unit_price": row["price"].get()})
            else:
                row["amount"].config(text="—")
        _, subtotal, vat, total = compute(valid, vat_rate)
        vat_name.config(text=f"VAT {vat_var.get().strip() or 0}%")
        subtotal_value.config(text=money(subtotal))
        vat_value.config(text=money(vat))
        total_value.config(text=money(total))

    def update_subtitle(*_):
        try:
            year = parse_date(date_var.get()).year
        except ValueError:
            subtitle.config(text="Check the date (dd/mm/yyyy)")
            return
        if year_label.cget("text") != f"/{year}":
            year_label.config(text=f"/{year}")
            number_var.set(str(next_number(year)))
        number = number_var.get().strip()
        number = f"{int(number):04d}" if number.isdigit() else number
        subtitle.config(text=f"No. {number}/{year}   ·   {date_var.get().strip()}")

    def layout_rows():
        for i, row in enumerate(rows):
            r = i + 2
            row["num"].config(text=f"{i + 1}.")
            row["num"].grid(row=r, column=0, sticky="nw", padx=px(4), pady=(px(11), px(5)))
            row["desc"].grid(row=r, column=1, sticky="we", padx=px(4), pady=px(5))
            row["qty_entry"].grid(row=r, column=2, sticky="nwe", padx=px(4), pady=px(5))
            row["price_entry"].grid(row=r, column=3, sticky="nwe", padx=px(4), pady=px(5))
            row["amount"].grid(row=r, column=4, sticky="ne", padx=px(4), pady=(px(11), px(5)))
            row["remove"].grid(row=r, column=5, sticky="n", padx=px(2), pady=(px(7), px(5)))
            row["remove"].state(["disabled"] if len(rows) == 1 else ["!disabled"])
        recalc()

    def remove_row(row):
        if len(rows) == 1:
            return
        for key in ("num", "desc", "qty_entry", "price_entry", "amount", "remove"):
            row[key].destroy()
        rows.remove(row)
        layout_rows()

    def add_row():
        row = {"qty": tk.StringVar(value="1"), "price": tk.StringVar()}
        row["num"] = label(table, "", 10, MUTED)
        row["desc"] = tk.Text(table, height=2, width=10, wrap="word", font=(FONT, 10), fg=TEXT, relief="flat",
                              bd=0, highlightthickness=1, highlightbackground=FIELD_BORDER, highlightcolor=GREEN,
                              padx=px(8), pady=px(6), insertbackground=TEXT)
        row["desc"].bind("<Tab>", focus_next)
        row["desc"].bind("<Shift-Tab>", focus_prev)
        row["qty_entry"] = ttk.Entry(table, textvariable=row["qty"], width=6, justify="right")
        row["price_entry"] = ttk.Entry(table, textvariable=row["price"], width=10, justify="right")
        row["amount"] = label(table, "—", 10, bold=True, anchor="e")
        row["remove"] = ttk.Button(table, text="✕", style="Remove.TButton", command=lambda: remove_row(row))
        row["qty"].trace_add("write", recalc)
        row["price"].trace_add("write", recalc)
        rows.append(row)
        layout_rows()
        row["desc"].focus_set()
        root.after_idle(lambda: canvas.yview_moveto(1))

    def refresh_recent():
        for widget in recent_list.winfo_children():
            widget.destroy()
        recent_items.clear()
        entries = read_log()[-8:][::-1]
        if not entries:
            label(recent_list, "No invoices yet - your first one will show up here.", 9, MUTED,
                  wraplength=px(260), justify="left").pack(anchor="w")
        for entry in entries:
            item = tk.Frame(recent_list, bg=CARD, cursor="hand2")
            recent_items.append(item)
            item.columnconfigure(0, weight=1)
            try:
                amount = money(Decimal(entry["total"]))
            except InvalidOperation:
                amount = entry["total"]
            client = entry["client"] if len(entry["client"]) <= 30 else entry["client"][:29] + "…"
            parts = [item,
                     label(item, entry["invoice_no"], 10, bold=True, cursor="hand2"),
                     label(item, amount, 10, bold=True, cursor="hand2"),
                     label(item, f"{client}  ·  {entry['date']}", 9, MUTED, cursor="hand2")]
            parts[1].grid(row=0, column=0, sticky="w", padx=px(8), pady=(px(5), 0))
            parts[2].grid(row=0, column=1, sticky="e", padx=px(8), pady=(px(5), 0))
            parts[3].grid(row=1, column=0, columnspan=2, sticky="w", padx=px(8), pady=(0, px(5)))
            path = HERE / entry["file"]
            for part in parts:
                part.bind("<Button-1>", lambda _e, p=path: open_file(p))
                part.bind("<Enter>", lambda _e, ps=parts: [p.config(bg=GREEN_SOFT) for p in ps])
                part.bind("<Leave>", lambda _e, ps=parts: [p.config(bg=CARD) for p in ps])
        recent_list.update_idletasks()  # measure the new rows before deciding which fit
        fit_recent()

    def fit_recent(_event=None):
        room = recent_list.winfo_height()
        for item in recent_items:
            item.pack_forget()
        for item in recent_items:
            room -= item.winfo_reqheight() + px(2)
            if room < 0:
                break
            item.pack(fill="x", pady=px(1))

    def create(_event=None):
        client = {k: v.get().strip() for k, v in client_vars.items()}
        if not client["name"]:
            return show_status("Enter the client's company name.", "error", focus=name_box)
        try:
            invoice_date = parse_date(date_var.get())
        except ValueError:
            return show_status("Write the date as dd/mm/yyyy, e.g. 23/10/2026.", "error", focus=date_entry)
        if not number_var.get().strip().isdigit():
            return show_status("The invoice number must be a whole number.", "error", focus=number_entry)
        number = int(number_var.get())
        try:
            vat_rate = parse_number(vat_var.get())
        except ValueError:
            return show_status("VAT % must be a number, e.g. 18.", "error", focus=vat_entry)
        items = []
        for i, row in enumerate(rows, start=1):
            description = row["desc"].get("1.0", "end").strip()
            if not description:
                return show_status(f"Item {i} needs a description.", "error", focus=row["desc"])
            for var, entry, what in ((row["qty"], row["qty_entry"], "quantity"),
                                     (row["price"], row["price_entry"], "unit price")):
                try:
                    parse_number(var.get())
                except ValueError:
                    return show_status(f"Item {i}: the {what} isn't a number.", "error", focus=entry)
            items.append({"description": description, "qty": row["qty"].get(), "unit_price": row["price"].get()})
        year = invoice_date.year
        if number_used(year, number) and not messagebox.askyesno(
                "Number already used", f"Invoice {number:04d}/{year} was already created.\nCreate it again anyway?"):
            return
        try:
            path = create_invoice(client, date_var.get().strip(), number, items, vat_rate, language_var.get())
        except PermissionError:
            return show_status("That invoice is open in Word - close it and try again.", "error")
        name_box.config(values=client_names())
        number_var.set(str(next_number(year)))
        for row in rows[1:]:
            remove_row(row)
        rows[0]["desc"].delete("1.0", "end")
        rows[0]["qty"].set("1")
        rows[0]["price"].set("")
        refresh_recent()
        show_status(f"Invoice {number:04d}/{year} saved.", "success", path=path)

    # ---- wiring
    name_box.config(values=client_names())
    name_box.bind("<<ComboboxSelected>>", pick_client)
    name_box.bind("<Return>", pick_client)
    name_box.bind("<KeyRelease>", filter_clients, add="+")
    table.bind("<Configure>", sync_scroll)
    recent_list.bind("<Configure>", fit_recent)
    canvas.bind("<Configure>", lambda e: (canvas.itemconfigure(table_window, width=e.width), sync_scroll()))
    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", wheel))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
    vat_var.trace_add("write", recalc)
    language_var.trace_add("write", paint_language)
    date_var.trace_add("write", update_subtitle)
    number_var.trace_add("write", update_subtitle)
    root.bind("<Control-Return>", create)

    add_row()
    paint_language()
    update_subtitle()
    refresh_recent()
    name_box.focus_set()
    root.mainloop()


# ---------------------------------------------------------------- command line

def run_json(json_path, out=None):
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    client = data.get("client")
    if isinstance(client, str):
        client = next((c for c in load_clients() if c["name"].lower() == client.lower()), None)
        if client is None:
            sys.exit(f"Unknown client '{data['client']}' - add it to clients.json or pass the full details.")
    date_text = data.get("date") or date.today().strftime("%d/%m/%Y")
    year = parse_date(date_text).year
    number = int(data.get("number") or next_number(year))
    vat_rate = parse_number(data.get("vat_rate", DEFAULT_VAT))
    language = data.get("language") or client.get("language") or DEFAULT_LANGUAGE
    if language not in LABELS:
        sys.exit(f"Unknown language '{language}' - use one of: {', '.join(LABELS)}.")
    if out:
        invoice = {"client": client, "invoice_no": f"{number:04d}/{year}", "date": date_text,
                   "vat_rate": vat_rate, "items": data["items"], "language": language}
        render(invoice, out)
        print(out)
    elif number_used(year, number):
        sys.exit(f"Invoice {number:04d}/{year} was already created - see invoices.csv.")
    else:
        print(create_invoice(client, date_text, number, data["items"], vat_rate, language))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--json", help="create an invoice from a JSON file instead of opening the form")
    parser.add_argument("--out", help="with --json: write to this file only (no numbering, no log)")
    args = parser.parse_args()
    if args.json:
        run_json(args.json, args.out)
    else:
        run_gui()
