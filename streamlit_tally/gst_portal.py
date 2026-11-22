"""GST portal Excel (e.g. GSTR-2B)  ->  Purchase & Sales template: the checks
that run before the conversion.

The portal workbook has one sheet per section (B2B, B2B-CDNR, …). For each
sheet the user converts, this module finds the tax columns under the header
row they chose and works out the GST rate of every row:

    rate % = (Integrated Tax + Central Tax + State/UT Tax) x 100 / Taxable Value

A row belongs to the 5 / 12 / 18 / 28 group when its rate lands on one of
those numbers; every other row (mixed-rate invoices, nil-rated, blanks) is
flagged so it can be highlighted and dealt with by hand.

Pure pandas/openpyxl — no Streamlit — so it can be tested on its own.
"""

import io
import re

import pandas as pd
from openpyxl import Workbook
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


def _match_rate(taxable, tax):
    """The 5/12/18/28 rate a row belongs to, or None."""
    if not taxable or taxable != taxable or tax != tax:
        return None
    rate = tax * 100 / taxable
    best = min(GST_RATES, key=lambda r: abs(rate - r))
    if (abs(rate - best) <= RATE_TOLERANCE_PCT
            or abs(tax - taxable * best / 100) <= RATE_TOLERANCE_RUPEES):
        return best
    return None


def classify_sheet(raw, header_row):
    """Work out the GST group of every data row under ``header_row`` (1-based).

    Returns a dict:

    * ``missing`` — tax headings not found on the header row (when non-empty
      nothing else is filled in: the header row is wrong for this sheet).
    * ``data`` — the sheet's rows with its own headings, indexed by Excel row
      number, plus Tax_Total, Tax_Rate_% and GST_Group (GST_5 … GST_28, or
      "Check" for rows that fit none of them).
    * ``counts`` — {5: n, 12: n, 18: n, 28: n}.
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
    tax = sum(number(k).fillna(0) for k in ("igst", "cgst", "sgst"))
    rates = [_match_rate(t, x) for t, x in zip(taxable.tolist(), tax.tolist())]

    data["Tax_Total"] = tax.round(2).tolist()
    data["Tax_Rate_%"] = [
        round(x * 100 / t, 2) if t and t == t else None
        for t, x in zip(taxable.tolist(), tax.tolist())]
    data["GST_Group"] = [f"GST_{r}" if r else NOT_MATCHED for r in rates]
    return {
        "missing": [],
        "data": data,
        "counts": {r: rates.count(r) for r in GST_RATES},
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
