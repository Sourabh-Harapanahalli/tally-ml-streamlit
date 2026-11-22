"""GST portal Excel (e.g. GSTR-2B)  ->  Purchase & Sales template: the checks
that run before the conversion.

The portal workbook has one sheet per section (B2B, B2B-CDNR, …). For each
sheet the user converts, this module finds the tax columns under the header
row they chose and works out the GST rate of every row:

    rate % = (Integrated Tax + Central Tax + State/UT Tax) x 100 / Taxable Value

A row belongs to the 5 / 12 / 18 / 28 group when its rate lands on one of
those numbers; every other row (mixed-rate invoices, nil-rated, blanks) is
flagged so it can be highlighted and dealt with by hand.

A bill is taxed either with Integrated Tax (IGST) or with Central + State/UT
Tax, never both: IGST bills go to the GST_<rate>_IGST group, the others to
GST_<rate>. A row carrying both kinds of tax is flagged too.

Pure pandas/openpyxl — no Streamlit — so it can be tested on its own.
"""

import io
import re

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

GST_RATES = (5, 12, 18, 28)

# A rate matches when it is within this many percentage points of 5/12/18/28,
# or when the tax is within this many rupees of taxable x rate. The second
# test stops paise rounding on small bills from being flagged.
RATE_TOLERANCE_PCT = 0.1
RATE_TOLERANCE_RUPEES = 1.0

SKIP = "Don't convert"
CONVERSIONS = [SKIP, "Purchase", "Sales", "Purchase credit / debit notes",
               "Sales credit / debit notes"]

# The columns the rate check needs, with the normalised portal heading for
# each (lower case, letters and digits only — so "Taxable Value (₹)" and
# "Taxable Value" are the same heading).
TAX_COLUMNS = {
    "taxable": ("Taxable Value", ("taxablevalue",)),
    "igst": ("Integrated Tax", ("integratedtax", "igst")),
    "cgst": ("Central Tax", ("centraltax", "cgst")),
    "sgst": ("State/UT Tax", ("stateuttax", "statetax", "sgst", "sgstutgst")),
}

HEADER_SCAN_ROWS = 30
ADDED_COLUMNS = ["Tax_Total", "Tax_Rate_%", "GST_Group"]
NOT_MATCHED = "Check"


def _blank(value):
    if value is None:
        return True
    if isinstance(value, float) and value != value:      # NaN
        return True
    return isinstance(value, str) and not value.strip()


def _norm(value):
    return "" if _blank(value) else re.sub(r"[^a-z0-9]", "", str(value).lower())


def load_sheets(file_bytes):
    """{sheet name: DataFrame} of the raw cells, with no header applied
    (row 1 of the sheet is index 0)."""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None,
                         dtype=object)


def data_row_count(raw, header_row):
    """Non-blank rows below ``header_row`` (1-based)."""
    body = raw.iloc[header_row:]
    return int((~body.map(_blank)).any(axis=1).sum()) if len(body) else 0


def find_tax_columns(header_values):
    """Locate the tax columns in a header row.

    Returns ``(positions, missing)``: ``positions`` maps taxable/igst/cgst/sgst
    to a column position, ``missing`` lists the headings not found. When a
    heading repeats (B2B-CDNR repeats the tax columns for ITC reduction) the
    first one — the note's own tax — is used.
    """
    norm = [_norm(v) for v in header_values]
    positions, missing = {}, []
    for key, (label, names) in TAX_COLUMNS.items():
        hit = next((i for i, n in enumerate(norm) if n in names), None)
        if hit is None:
            missing.append(label)
        else:
            positions[key] = hit
    return positions, missing


def detect_header_row(raw):
    """1-based row holding the tax headings, or None if no top row has them."""
    for i in range(min(HEADER_SCAN_ROWS, len(raw))):
        if not find_tax_columns(raw.iloc[i].tolist())[1]:
            return i + 1
    return None


def suggest_conversion(sheet_name):
    """Default dropdown choice for a portal sheet, from its name."""
    name = _norm(sheet_name)
    if name == "b2b":
        return "Purchase"
    if name == "b2bcdnr":
        return "Purchase credit / debit notes"
    return SKIP


def ledger_side(conversion):
    """'Purchase' or 'Sales' — which set of ledgers a conversion posts to."""
    return "Sales" if conversion.startswith("Sales") else "Purchase"


def _unique_headers(values):
    """Header names made unique and non-blank, in column order."""
    seen, out = {}, []
    for i, value in enumerate(values):
        name = f"Column {get_column_letter(i + 1)}" if _blank(value) else str(value).strip()
        seen[name] = seen.get(name, 0) + 1
        out.append(name if seen[name] == 1 else f"{name} ({seen[name]})")
    return out


def group_name(rate, igst=False):
    """Template column for a rate: GST_5, or GST_5_IGST for IGST bills."""
    return f"GST_{rate}_IGST" if igst else f"GST_{rate}"


_HALF_RATE = {5: "2.5", 12: "6", 18: "9", 28: "14"}


def link_rows(igst=False, rates=GST_RATES):
    """The MASTER_LEDGER_NAME_LINK rows a kind of bill needs, as a list of
    ``(row title, [(link column, rate), …])``: for local bills the purchase /
    sales ledger and the CGST and SGST tax ledgers of each rate, for
    interstate bills the purchase / sales ledger and the IGST tax ledger."""
    if igst:
        return [
            ("Purchase / sales ledger", [(group_name(r, True), r) for r in rates]),
            ("IGST tax ledger", [(f"{r}_IGST", r) for r in rates]),
        ]
    return [
        ("Purchase / sales ledger", [(group_name(r), r) for r in rates]),
        ("CGST tax ledger", [(f"{_HALF_RATE[r]}_CGST", r) for r in rates]),
        ("SGST tax ledger", [(f"{_HALF_RATE[r]}_SGST", r) for r in rates]),
    ]


def _match_rate(taxable, tax, rates=GST_RATES):
    """The rate (one of ``rates``) a row belongs to, or None."""
    if not rates or not taxable or taxable != taxable or tax != tax:
        return None
    rate = tax * 100 / taxable
    best = min(rates, key=lambda r: abs(rate - r))
    if (abs(rate - best) <= RATE_TOLERANCE_PCT
            or abs(tax - taxable * best / 100) <= RATE_TOLERANCE_RUPEES):
        return best
    return None


def classify_sheet(raw, header_row, rates=GST_RATES):
    """Work out the GST group of every data row under ``header_row`` (1-based).

    ``rates`` are the GST rates that apply; a row at any other rate — also
    one of 5/12/18/28 that was left out — is a "Check" row.

    Returns a dict:

    * ``missing`` — tax headings not found on the header row (when non-empty
      nothing else is filled in: the header row is wrong for this sheet).
    * ``data`` — the sheet's rows with its own headings, indexed by Excel row
      number, plus Tax_Total, Tax_Rate_% and GST_Group (GST_5 … GST_28,
      GST_5_IGST … GST_28_IGST, or "Check" for rows that fit none of them
      or carry both IGST and Central/State tax).
    * ``counts`` — rows per rate, {5: n, 12: n, 18: n, 28: n}.
    * ``group_counts`` — rows per group, {"GST_5": n, "GST_5_IGST": n, …}.
    * ``unmatched`` — the Excel row numbers flagged "Check".
    """
    if header_row < 1 or header_row > len(raw):
        return {"missing": [label for label, _ in TAX_COLUMNS.values()]}
    positions, missing = find_tax_columns(raw.iloc[header_row - 1].tolist())
    if missing:
        return {"missing": missing}

    body = raw.iloc[header_row:]
    body = body[(~body.map(_blank)).any(axis=1)]
    data = body.copy()
    data.columns = _unique_headers(raw.iloc[header_row - 1].tolist())
    data.index = [i + 1 for i in body.index]              # Excel row numbers

    def number(key):
        return pd.to_numeric(body.iloc[:, positions[key]], errors="coerce")

    taxable = number("taxable")
    igst = number("igst").fillna(0)
    local = number("cgst").fillna(0) + number("sgst").fillna(0)
    tax = igst + local
    # IGST and Central/State tax never go together: such a row fits no group.
    rates = [None if i and c else _match_rate(t, x, rates)
             for t, x, i, c in zip(taxable.tolist(), tax.tolist(),
                                   igst.tolist(), local.tolist())]
    groups = [group_name(r, bool(i)) if r else NOT_MATCHED
              for r, i in zip(rates, igst.tolist())]

    data["Tax_Total"] = tax.round(2).tolist()
    data["Tax_Rate_%"] = [
        round(x * 100 / t, 2) if t and t == t else None
        for t, x in zip(taxable.tolist(), tax.tolist())]
    data["GST_Group"] = groups
    return {
        "missing": [],
        "data": data,
        "counts": {r: rates.count(r) for r in GST_RATES},
        "group_counts": {group_name(r, i): groups.count(group_name(r, i))
                         for i in (False, True) for r in GST_RATES},
        "unmatched": [row for row, r in zip(data.index, rates) if r is None],
    }


_HEADER_FONT = Font(bold=True)
_HIGHLIGHT = PatternFill("solid", fgColor="FFF2A8")


def build_review_workbook(results):
    """Excel copy of the checked sheets, with the rows that fit no GST group
    highlighted. ``results`` maps sheet name -> a ``classify_sheet`` result."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet, result in results.items():
        data = result["data"]
        ws = wb.create_sheet(sheet[:31])
        ws.append(["Portal_Row"] + list(data.columns))
        for cell in ws[1]:
            cell.font = _HEADER_FONT
        flagged = set(result["unmatched"])
        for row_no, values in zip(data.index, data.itertuples(index=False, name=None)):
            ws.append([row_no] + [None if _blank(v) else v for v in values])
            if row_no in flagged:
                for cell in ws[ws.max_row]:
                    cell.fill = _HIGHLIGHT
        ws.freeze_panes = "A2"
        for i in range(1, len(data.columns) + 2):
            ws.column_dimensions[get_column_letter(i)].width = 18
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------
# Portal rows  ->  the Purchase & Sales Bills template.
# --------------------------------------------------------------------------
# Bill details read from the portal sheet: normalised headings for each.
BILL_COLUMNS = {
    "number": ("Invoice number", ("invoicenumber", "notenumber")),
    "date": ("Invoice Date", ("invoicedate", "notedate")),
    "party": ("Trade/Legal name", ("tradelegalname", "legalname", "tradename")),
    "value": ("Invoice Value", ("invoicevalue", "notevalue")),
}
_RED = PatternFill("solid", fgColor="F8B4B4")


def find_bill_columns(columns):
    """``(names, missing)`` — the sheet's own column name for each bill
    detail, and the labels of the details it has no column for."""
    names, missing = {}, []
    for key, (label, wanted) in BILL_COLUMNS.items():
        hit = next((c for c in columns if _norm(c) in wanted), None)
        if hit is None:
            missing.append(label)
        else:
            names[key] = hit
    return names, missing


def _date_text(value):
    """A portal date as dd/mm/yyyy text (the portal already writes it so)."""
    if _blank(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def _amount(value):
    number = pd.to_numeric(value, errors="coerce")
    return None if number != number else round(float(number), 2)


def template_rows(result, vch_type):
    """Purchase & Sales template rows for one checked sheet.

    Returns ``(rows, flags)``: ``rows`` are dicts keyed by template column,
    ``flags`` says for each row whether it is a "Check" row. The taxable value
    goes in the row's group column: GST_<rate> for a bill taxed with Central
    + State/UT tax, GST_<rate>_IGST for one taxed with Integrated tax. "Check"
    rows keep every GST column blank.
    """
    data = result["data"]
    names, _ = find_bill_columns(data.columns)
    tax_pos, _ = find_tax_columns(data.columns)
    taxable_col = data.columns[tax_pos["taxable"]]

    def cell(record, key):
        return record[names[key]] if key in names else None

    rows, flags = [], []
    for record in data.to_dict("records"):
        party = cell(record, "party")
        party = "" if _blank(party) else str(party).strip()
        number = cell(record, "number")
        row = {
            "Supplier_Invoice": "" if _blank(number) else str(number).strip(),
            "Datetime": _date_text(cell(record, "date")),
            "Vch_Type": vch_type,
            "PartyLedgerName": party,
            "Dr_LedgerName": party,
            "Total_Amount": _amount(cell(record, "value")),
        }
        group = record["GST_Group"]
        if group != NOT_MATCHED:
            row[group] = _amount(record[taxable_col])
        rows.append(row)
        flags.append(group == NOT_MATCHED)
    return rows, flags


def build_template(parts, ledgers):
    """A filled Purchase & Sales Bills workbook.

    ``parts`` is a list of ``(classify_sheet result, Vch_Type)``; ``ledgers``
    maps a MASTER_LEDGER_NAME_LINK column (GST_5, GST_5_IGST, 2.5_CGST,
    2.5_SGST, 5_IGST, …) to its Tally ledger name, written on that sheet.
    "Check" rows are filled red with their GST columns left blank.

    Returns ``(workbook bytes, rows written, rows to check)``.
    """
    import tally_core  # here, not at the top: only this function needs it

    wb = load_workbook(io.BytesIO(tally_core.template_purchase()))
    ws = wb["TEMPLATE"]
    headers = [c.value for c in ws[1]]
    date_col = headers.index("Datetime") + 1
    # Rows are placed by number, not appended: the blank template pre-formats
    # thousands of empty rows, so an append would land below all of them.
    written = to_check = 0
    for result, vch_type in parts:
        rows, flags = template_rows(result, vch_type)
        for row, flagged in zip(rows, flags):
            line = written + 2
            for c, header in enumerate(headers, start=1):
                target = ws.cell(line, c, row.get(header))
                if flagged:
                    target.fill = _RED
            ws.cell(line, date_col).number_format = "@"
            written += 1
            to_check += flagged

    link = wb["MASTER_LEDGER_NAME_LINK"]
    for key_cell, name_cell in link.iter_rows(min_row=2, max_col=2):
        if ledgers.get(key_cell.value):
            name_cell.value = ledgers[key_cell.value]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), written, to_check
