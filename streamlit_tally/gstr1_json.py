"""GSTR-1 JSON (GST portal / offline tool download)  ->  Excel.

Each section of the return becomes one or two sheets:

    B2B   -> b2b        one row per invoice + tax rate
             b2b_rate   the same, totalled per rate
    CDNR  -> cdnr       one row per credit / debit note + tax rate
             cdnr_rate  the same, totalled per rate
    B2CS  -> b2cs       totalled per rate
    NIL   -> nil        nil-rated / exempted / non-GST supplies per supply type
    HSN   -> hsn_b2b, hsn_b2c (or hsn for the older single-table layout)

The app converts each JSON to its own workbook; returns are never combined.

A return file rarely holds every section (the portal splits B2B into its own
file, and a taxpayer may simply have no B2CS or NIL supplies), so sections that
are absent are skipped and reported rather than treated as an error.

Pure pandas/openpyxl — no Streamlit — so it can be tested on its own.
"""

import io
import json
import zipfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

SECTIONS = ["B2B", "CDNR", "B2CS", "NIL", "HSN"]
ALL = "ALL"
# JSON key behind each section.
SECTION_KEYS = {"B2B": "b2b", "CDNR": "cdnr", "B2CS": "b2cs", "NIL": "nil",
                "HSN": "hsn"}
# Top-level keys that are return metadata rather than a section.
META_KEYS = {"gstin", "fp", "filing_typ", "gt", "cur_gt", "fil_dt", "version",
             "hash"}

AMOUNTS = ["Taxable_amt", "IGST", "CGST", "SGST"]
B2B_COLUMNS = ["GSTN", "Invoice_no", "Invoice Date", "Taxable_amt", "Rate",
               "IGST", "CGST", "SGST"]
CDNR_COLUMNS = ["GSTN", "Invoice_no", "Invoice Date", "Note Type",
                "Taxable_amt", "Rate", "IGST", "CGST", "SGST"]
RATE_COLUMNS = ["Rate"] + AMOUNTS
NIL_COLUMNS = ["Supply Type", "Nil_Rated_amt", "Exempted_amt", "Non_GST_amt"]
HSN_COLUMNS = ["HSN", "Description", "UQC", "Qty", "Rate"] + AMOUNTS

NIL_SUPPLY_TYPES = {
    "INTRB2B": "Inter-State supplies to registered persons",
    "INTRAB2B": "Intra-State supplies to registered persons",
    "INTRB2C": "Inter-State supplies to unregistered persons",
    "INTRAB2C": "Intra-State supplies to unregistered persons",
}


def _num(value):
    """A JSON amount as a number (ints stay ints; blanks become 0)."""
    if value in (None, ""):
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        number = value
    else:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return 0
    if isinstance(number, float) and number.is_integer():
        return int(number)
    return number


def _clean(number):
    """Round away float noise from summing (7863.870000000001 -> 7863.87)."""
    number = round(float(number), 2)
    return int(number) if number.is_integer() else number


def load_returns(filename, data):
    """Parse an uploaded file into a list of (name, return-dict).

    Accepts a .json file or the .zip the portal hands out (every .json inside).
    """
    if filename.lower().endswith(".zip") or data[:2] == b"PK":
        out = []
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for member in zf.namelist():
                if member.lower().endswith(".json"):
                    out.append((member.rsplit("/", 1)[-1],
                                _parse(zf.read(member))))
        if not out:
            raise ValueError("the zip contains no .json file")
        return out
    return [(filename, _parse(data))]


def _parse(data):
    obj = json.loads(data.decode("utf-8-sig"))
    if not isinstance(obj, dict):
        raise ValueError("not a GSTR-1 return (expected a JSON object)")
    return obj


def _has(ret, section):
    value = ret.get(SECTION_KEYS[section])
    if section == "HSN":
        return isinstance(value, dict) and any(
            value.get(k) for k in ("hsn_b2b", "hsn_b2c", "data"))
    if section == "NIL":
        return isinstance(value, dict) and bool(value.get("inv"))
    return bool(value)


def available_sections(returns):
    """Sections (in menu order) that at least one of the returns contains."""
    return [s for s in SECTIONS if any(_has(r, s) for r in returns)]


def other_sections(returns):
    """Keys present in the returns that this converter does not export."""
    known = set(SECTION_KEYS.values()) | META_KEYS
    seen = []
    for ret in returns:
        for key, value in ret.items():
            if key not in known and value and key not in seen:
                seen.append(key)
    return seen


# ---- Section -> rows ------------------------------------------------------

def _item_rows(parties, doc_key, num_key, date_key, with_note_type):
    rows = []
    for party in parties or []:
        for doc in party.get(doc_key) or []:
            for item in doc.get("itms") or []:
                det = item.get("itm_det") or {}
                row = {"GSTN": party.get("ctin", ""),
                       "Invoice_no": doc.get(num_key, ""),
                       "Invoice Date": doc.get(date_key, "")}
                if with_note_type:
                    row["Note Type"] = doc.get("ntty", "")
                row.update({"Taxable_amt": _num(det.get("txval")),
                            "Rate": _num(det.get("rt")),
                            "IGST": _num(det.get("iamt")),
                            "CGST": _num(det.get("camt")),
                            "SGST": _num(det.get("samt"))})
                rows.append(row)
    return rows


def _rate_summary(rows):
    """Totals per tax rate, lowest rate first, with a Total row."""
    df = pd.DataFrame(rows, columns=RATE_COLUMNS)
    if df.empty:
        return df
    out = df.groupby("Rate", sort=True)[AMOUNTS].sum().reset_index()
    total = {"Rate": "Total", **{c: out[c].sum() for c in AMOUNTS}}
    out = pd.concat([out.astype(object), pd.DataFrame([total])],
                    ignore_index=True)
    for col in AMOUNTS:
        out[col] = out[col].map(_clean)
    return out


def _with_total(df, label_col, sum_cols):
    if df.empty:
        return df
    total = {c: "" for c in df.columns}
    total[label_col] = "Total"
    total.update({c: df[c].sum() for c in sum_cols})
    out = pd.concat([df.astype(object), pd.DataFrame([total])],
                    ignore_index=True)
    for col in sum_cols:
        out[col] = out[col].map(_clean)
    return out


def _b2b(returns):
    rows = [r for ret in returns
            for r in _item_rows(ret.get("b2b"), "inv", "inum", "idt", False)]
    return {"b2b": pd.DataFrame(rows, columns=B2B_COLUMNS),
            "b2b_rate": _rate_summary(rows)}


def _cdnr(returns):
    rows = [r for ret in returns
            for r in _item_rows(ret.get("cdnr"), "nt", "nt_num", "nt_dt", True)]
    return {"cdnr": pd.DataFrame(rows, columns=CDNR_COLUMNS),
            "cdnr_rate": _rate_summary(rows)}


def _b2cs(returns):
    rows = [{"Rate": _num(e.get("rt")), "Taxable_amt": _num(e.get("txval")),
             "IGST": _num(e.get("iamt")), "CGST": _num(e.get("camt")),
             "SGST": _num(e.get("samt"))}
            for ret in returns for e in ret.get("b2cs") or []]
    return {"b2cs": _rate_summary(rows)}


def _nil(returns):
    rows = []
    for ret in returns:
        for e in (ret.get("nil") or {}).get("inv") or []:
            code = e.get("sply_ty", "")
            rows.append({"Supply Type": NIL_SUPPLY_TYPES.get(code, code),
                         "Nil_Rated_amt": _num(e.get("nil_amt")),
                         "Exempted_amt": _num(e.get("expt_amt")),
                         "Non_GST_amt": _num(e.get("ngsup_amt"))})
    df = pd.DataFrame(rows, columns=NIL_COLUMNS)
    if len(returns) > 1 and not df.empty:
        df = df.groupby("Supply Type", sort=False)[NIL_COLUMNS[1:]] \
               .sum().reset_index()
    return {"nil": _with_total(df, "Supply Type", NIL_COLUMNS[1:])}


def _hsn(returns):
    tables = {}
    for ret in returns:
        hsn = ret.get("hsn") or {}
        for key, sheet in (("hsn_b2b", "hsn_b2b"), ("hsn_b2c", "hsn_b2c"),
                           ("data", "hsn")):
            for e in hsn.get(key) or []:
                tables.setdefault(sheet, []).append({
                    "HSN": str(e.get("hsn_sc", "")),
                    "Description": e.get("desc", ""),
                    "UQC": e.get("uqc", ""),
                    "Qty": _num(e.get("qty")),
                    "Rate": _num(e.get("rt")),
                    "Taxable_amt": _num(e.get("txval")),
                    "IGST": _num(e.get("iamt")),
                    "CGST": _num(e.get("camt")),
                    "SGST": _num(e.get("samt"))})
    return {sheet: _with_total(pd.DataFrame(rows, columns=HSN_COLUMNS),
                               "HSN", AMOUNTS)
            for sheet, rows in tables.items()}


BUILDERS = {"B2B": _b2b, "CDNR": _cdnr, "B2CS": _b2cs, "NIL": _nil,
            "HSN": _hsn}


def build_sheets(returns, choice):
    """Sheets for one section, or for every section present when choice is ALL.

    Returns (sheets, missing): an ordered {sheet name: DataFrame} and the list
    of requested sections that none of the returns contain.
    """
    wanted = SECTIONS if choice == ALL else [choice]
    sheets, missing = {}, []
    for section in wanted:
        if not any(_has(r, section) for r in returns):
            missing.append(section)
            continue
        for name, df in BUILDERS[section](returns).items():
            if not df.empty:
                sheets[name] = df
    return sheets, missing


def sheets_to_xlsx(sheets):
    """Write the sheets to one workbook: bold centred headers, bold labels on
    the rate / total rows, sensible column widths."""
    wb = Workbook()
    wb.remove(wb.active)
    bold, centre = Font(bold=True), Alignment(horizontal="center")
    for name, df in sheets.items():
        ws = wb.create_sheet(title=name[:31])
        ws.append(list(df.columns))
        for cell in ws[1]:
            cell.font, cell.alignment = bold, centre
        for record in df.itertuples(index=False):
            ws.append([v.item() if hasattr(v, "item") else v for v in record])
        if df.columns[0] == "Rate":           # rate summary: label column
            for row in ws.iter_rows(min_row=2, max_col=1):
                row[0].font, row[0].alignment = bold, centre
        elif len(df) and df.iloc[-1, 0] == "Total":
            for cell in ws[ws.max_row]:
                cell.font = bold
        for i, col in enumerate(df.columns, start=1):
            longest = max([len(str(col))] + [len(str(v)) for v in df[col]])
            ws.column_dimensions[get_column_letter(i)].width = \
                min(max(longest + 2, 10), 45)
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
