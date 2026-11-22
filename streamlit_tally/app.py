"""
Tally AI — Streamlit app.

A self-contained Excel <-> Tally-XML converter. All processing logic lives in
``tally_core`` (pure pandas/openpyxl, no web framework); this module is the
entire UI. The conversion output is byte-for-byte identical to the original
Django implementation this was ported from.

Run:  streamlit run app.py
"""

import difflib
import hashlib
import html
import io
import os
import re
import threading
import time
import zipfile
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx

import tally_core
import bank_pdf
import excel_compare
import excel_merge
import gst_portal
import gstr1_json
import marg_dbf
import marg_xml
import tally_log
import ui

# Load OPENAI_API_KEY (and any other vars) from a .env file if present.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Must be the first Streamlit command in the script.
st.set_page_config(page_title="Tally AI", page_icon="📒", layout="wide",
                   initial_sidebar_state="expanded")
ui.inject_style()

# --------------------------------------------------------------------------
# Tool catalogue — maps each converter to its upload key, conversion function
# and blank-template generator (all in tally_core).
# --------------------------------------------------------------------------
TOOLS = {
    "Purchase / Sales": {
        "upload_key": "upload_file",
        "convert": tally_core.convert_purchase_sales,
        "template": tally_core.template_purchase,
        "template_name": "Template_Purchase.xlsx",
        "help": "Purchase & Sales vouchers with GST (CGST/SGST/IGST) breakup. "
                "The template has two sheets: TEMPLATE (your data) and "
                "MASTER_LEDGER_NAME_LINK (map GST columns to your Tally ledger names).",
    },
    "Payment / Contra / Receipt": {
        "upload_key": "upload_file_pay_con_rec",
        "convert": tally_core.convert_pay_con_rec,
        "template": tally_core.template_pay_con_rec,
        "template_name": "Template_Pay_Con_Rec.xlsx",
        "help": "Payment, Contra and Receipt vouchers (Dr/Cr ledger entries).",
    },
    "Master — Ledger": {
        "upload_key": "upload_file_master_ledger",
        "convert": tally_core.convert_master_ledger,
        "template": tally_core.template_master_ledger,
        "template_name": "Template_Master_Ledger.xlsx",
        "help": "Create ledger masters (groups, GST no., opening balance, address). "
                "The REFERENCE sheet lists the valid Tally group names.",
    },
    "Master — Duties & Taxes": {
        "upload_key": "upload_file_master_duties",
        "convert": tally_core.convert_master_duties,
        "template": tally_core.template_master_duties,
        "template_name": "Template_Master_DPS.xlsx",
        "help": "Create Duties & Taxes ledgers (rate of tax, tax type).",
    },
    "Master — Purchase/Sales Ledgers": {
        "upload_key": "upload_file_master_ps",
        "convert": tally_core.convert_master_ps,
        "template": tally_core.template_master_ps,
        "template_name": "Template_Master_PS.xlsx",
        "help": "Create Purchase/Sales ledgers with nature of transaction and GST rates.",
    },
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def build_template(tool_name):
    """Generate a blank template workbook and return its bytes."""
    return TOOLS[tool_name]["template"]()


def sheet_header_set(ws):
    """Non-blank header names from the first row of a worksheet."""
    row = next(ws.iter_rows(max_row=1, values_only=True), ())
    return {str(v).strip() for v in row if v is not None and str(v).strip()}


@st.cache_data(show_spinner=False)
def template_headers(tool_name):
    """Header names on the TEMPLATE sheet of a tool's blank template."""
    wb = load_workbook(io.BytesIO(build_template(tool_name)), read_only=True)
    return frozenset(sheet_header_set(wb["TEMPLATE"]))


def template_coverage(tool_name, headers):
    """Share of the tool's template columns present in ``headers`` (0–1)."""
    expected = template_headers(tool_name)
    return len(expected & headers) / len(expected) if expected else 0.0


def detect_template_tool(headers):
    """Return ``(tool_name, coverage)`` for the template that best matches
    ``headers``, or None when nothing matches at all."""
    best = max(((name, template_coverage(name, headers)) for name in TOOLS
                if "template" in TOOLS[name]), key=lambda x: x[1], default=None)
    return best if best and best[1] > 0 else None


# --------------------------------------------------------------------------
# Example data for each converter. These are sample rows that show the
# expected format. EXAMPLES fills the TEMPLATE sheet; EXAMPLE_AUX fills the
# secondary sheet (only Purchase/Sales uses the ledger-name mapping sheet).
# --------------------------------------------------------------------------
EXAMPLES = {
    "Purchase / Sales": [
        {"Supplier_Invoice": "INV-001", "Datetime": "01/04/2024", "Vch_Type": "Purchase",
         "PartyLedgerName": "ABC Traders", "Dr_LedgerName": "Purchase @ 18%",
         "Total_Amount": 11800, "GST_18": 10000, "Narration": "Goods purchased - 18% GST (intra-state)"},
        {"Supplier_Invoice": "INV-002", "Datetime": "02/04/2024", "Vch_Type": "Sales",
         "PartyLedgerName": "XYZ Enterprises", "Dr_LedgerName": "Sales @ 12%",
         "Total_Amount": 5600, "GST_12": 5000, "Narration": "Goods sold - 12% GST (intra-state)"},
        {"Supplier_Invoice": "INV-003", "Datetime": "03/04/2024", "Vch_Type": "Sales",
         "PartyLedgerName": "PQR Pvt Ltd", "Dr_LedgerName": "Sales @ 18% (IGST)",
         "Total_Amount": 11800, "GST_18_IGST": 10000, "Narration": "Interstate sale - 18% IGST"},
    ],
    "Payment / Contra / Receipt": [
        {"DATE_TIME": "01/04/2024", "PartyLedgerName": "ABC Traders", "Dr_LedgerName": "ABC Traders",
         "Dr_Amount": 10000, "Cr_LedgerName": "HDFC Bank", "Cr_Amount": 10000,
         "Narration": "Payment to supplier", "Vch_Type": "Payment"},
        {"DATE_TIME": "02/04/2024", "PartyLedgerName": "XYZ Enterprises", "Dr_LedgerName": "HDFC Bank",
         "Dr_Amount": 15000, "Cr_LedgerName": "XYZ Enterprises", "Cr_Amount": 15000,
         "Narration": "Receipt from customer", "Vch_Type": "Receipt"},
        {"DATE_TIME": "03/04/2024", "PartyLedgerName": "Cash", "Dr_LedgerName": "Cash",
         "Dr_Amount": 5000, "Cr_LedgerName": "HDFC Bank", "Cr_Amount": 5000,
         "Narration": "Cash withdrawn from bank", "Vch_Type": "Contra"},
    ],
    "Master — Ledger": [
        {"Ledger_Name": "ABC Traders", "Alias": "ABC", "Group_Name": "Sundry Creditors",
         "Country": "India", "State_Name": "Karnataka", "Pincode": "560001",
         "Registration_Type": "Regular", "GST_NO": "29ABCDE1234F1Z5", "Opening_Balance": 0,
         "Dr/Cr": "Cr", "Address": "123 MG Road, Bangalore"},
        {"Ledger_Name": "XYZ Enterprises", "Alias": "XYZ", "Group_Name": "Sundry Debtors",
         "Country": "India", "State_Name": "Maharashtra", "Pincode": "400001",
         "Registration_Type": "Regular", "GST_NO": "27XYZAB5678G1Z3", "Opening_Balance": 25000,
         "Dr/Cr": "Dr", "Address": "45 Marine Drive, Mumbai"},
    ],
    "Master — Duties & Taxes": [
        {"Ledger_Name": "CGST", "Group_Name": "Duties & Taxes", "Rate_of_Tax": 9, "Tax_Type": "Central Tax"},
        {"Ledger_Name": "SGST", "Group_Name": "Duties & Taxes", "Rate_of_Tax": 9, "Tax_Type": "State Tax"},
        {"Ledger_Name": "IGST", "Group_Name": "Duties & Taxes", "Rate_of_Tax": 18, "Tax_Type": "Integrated Tax"},
    ],
    "Master — Purchase/Sales Ledgers": [
        {"Ledger_Name": "Sales @ 18%", "Group_Name": "Sales Accounts",
         "Nature_of_transaction": "Sales Taxable", "RATE_OF_CGST_SGST": 18, "RATE_OF_IGST": 18},
        {"Ledger_Name": "Purchase @ 12%", "Group_Name": "Purchase Accounts",
         "Nature_of_transaction": "Purchase Taxable", "RATE_OF_CGST_SGST": 12, "RATE_OF_IGST": 12},
    ],
}

# Maps each GST column on the TEMPLATE sheet to a sample Tally ledger name on
# the MASTER_LEDGER_NAME_LINK sheet (Purchase / Sales only).
EXAMPLE_AUX = {
    "Purchase / Sales": {
        "sheet": "MASTER_LEDGER_NAME_LINK",
        "values": {
            "GST_5": "Purchase/Sales @ 5%", "GST_12": "Purchase/Sales @ 12%",
            "GST_18": "Purchase/Sales @ 18%", "GST_28": "Purchase/Sales @ 28%",
            "GST_5_IGST": "Purchase/Sales @ 5% (IGST)", "GST_12_IGST": "Purchase/Sales @ 12% (IGST)",
            "GST_18_IGST": "Purchase/Sales @ 18% (IGST)", "GST_28_IGST": "Purchase/Sales @ 28% (IGST)",
            "2.5_CGST": "CGST", "6_CGST": "CGST", "9_CGST": "CGST", "14_CGST": "CGST",
            "2.5_SGST": "SGST", "6_SGST": "SGST", "9_SGST": "SGST", "14_SGST": "SGST",
            "5_IGST": "IGST", "12_IGST": "IGST", "18_IGST": "IGST", "28_IGST": "IGST",
        },
    },
}


@st.cache_data(show_spinner=False)
def build_example(tool_name):
    """Return bytes of the blank template pre-filled with the example rows."""
    wb = load_workbook(io.BytesIO(build_template(tool_name)))

    # Fill the TEMPLATE sheet, matching each example value to its header.
    ws = wb["TEMPLATE"]
    headers = [cell.value for cell in ws[1]]
    for r, row in enumerate(EXAMPLES[tool_name], start=2):
        for header, value in row.items():
            if header in headers:
                ws.cell(row=r, column=headers.index(header) + 1, value=value)

    # Fill the secondary mapping sheet, if this tool has one.
    aux = EXAMPLE_AUX.get(tool_name)
    if aux:
        ws2 = wb[aux["sheet"]]
        for r in range(2, ws2.max_row + 1):
            key = ws2.cell(row=r, column=1).value
            if key in aux["values"]:
                ws2.cell(row=r, column=2, value=aux["values"][key])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------
# Reverse direction: Tally "All Masters" XML export  ->  Master-Ledger Excel.
#
# This mirrors the Master_Ledger view in reverse: it reads each <LEDGER> block
# from a Tally export and emits the same columns the Master — Ledger template
# uses (Ledger_Name, Alias, Group_Name, Country, State_Name, Pincode,
# Registration_Type, GST_NO, Opening_Balance, Dr/Cr, Address).
# --------------------------------------------------------------------------
MASTER_XML_TOOL = "Master XML → Excel (reverse)"
BANK_PDF_TOOL = "Bank Statement PDF → Pay/Contra/Receipt"
CONSOLIDATE_TOOL = "Consolidate Excel Files"
GSTR1_JSON_TOOL = "GSTR-1 JSON → Excel"
COMPARE_TOOL = "Find Missing Values (Excel A vs B)"
ERROR_LOG_TOOL = "Tally Error Log Viewer"
MARG_MASTER_TOOL = "MARG Master (Tally Export)"
MARG_VOUCHER_TOOL = "MARG Voucher (Tally Export)"
MARG_LEDGER_TOOL = "MARG Ledger DBF → Master — Ledger"
GST_PORTAL_TOOL = "GST Portal Excel → Purchase / Sales"


@st.cache_data(show_spinner=False)
def convert_statement_pdf(pdf_bytes, model):
    """Run the PDF → PCR conversion, cached on (pdf_bytes, model).

    Caching means a Streamlit rerun — e.g. clicking a download button — returns
    the previous result instantly instead of re-calling the (paid) LLM.
    """
    return bank_pdf.pcr_from_pdf(pdf_bytes, model=model)

# Column order must match the Master — Ledger TEMPLATE sheet.
LEDGER_COLUMNS = [
    "Ledger_Name", "Alias", "Group_Name", "Country", "State_Name", "Pincode",
    "Registration_Type", "GST_NO", "Opening_Balance", "Dr/Cr", "Address",
]


@st.cache_data(show_spinner=False)
def consolidate_excel(files, add_source, drop_duplicates, keep_identical_once,
                      header_rows):
    """Merge same-named sheets across the uploaded workbooks (cached)."""
    return excel_merge.consolidate_workbooks(
        list(files), add_source=add_source, drop_duplicates=drop_duplicates,
        keep_identical_once=keep_identical_once, header_rows=header_rows)


@st.cache_data(show_spinner=False)
def compare_load_sheets(file_bytes):
    """Every non-empty sheet of a workbook as raw rows (cached)."""
    return excel_compare.load_sheets(file_bytes)


def compare_pick_column(side, upload):
    """Preview one workbook and let the user click a single column.

    Returns (description, Series) for the chosen column, or None while no
    column is selected. Only one sheet is shown at a time and the table
    allows a single column, so each workbook can have just one selection.
    """
    st.subheader(f"Excel {side} — {upload.name}")
    try:
        sheets = compare_load_sheets(upload.getvalue())
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read **{upload.name}**: {exc}")
        return None
    if not sheets:
        st.warning("This workbook has no data.")
        return None

    names = list(sheets)
    c1, c2 = st.columns([3, 1])
    if len(names) <= 6:
        sheet = c1.radio("Sheet", names, horizontal=True,
                         key=f"cmp_sheet_{side}")
    else:
        sheet = c1.selectbox("Sheet", names, key=f"cmp_sheet_{side}")
    rows = sheets[sheet]
    has_header = st.checkbox(
        "This sheet has a header — exclude it from the values",
        value=True, key=f"cmp_has_header_{side}_{sheet}",
        help="Untick if the sheet has no headings, or if the heading row "
             "should be compared like any other row. Columns are then "
             "named by their Excel letter and every row counts as a value.")
    header_row = int(c2.number_input(
        "Header row", min_value=1, max_value=len(rows),
        value=min(excel_compare.guess_header_row(rows), len(rows)), step=1,
        key=f"cmp_header_{side}_{sheet}", disabled=not has_header,
        help="The Excel row that holds the column headings. Change it if "
             "the headings below look wrong. Rows above it are ignored."))
    if not has_header:
        header_row = 0                        # every row is data
    table = excel_compare.to_table(rows, header_row)
    if table.empty:
        st.warning("There are no rows below the header row.")
        return None

    preview = table.head(500)
    preview = preview.where(preview.notna(), "").astype(str)
    # The key carries the file, sheet and header row, so changing any of them
    # starts a fresh table with nothing selected.
    event = st.dataframe(
        preview, use_container_width=True, height=300,
        on_select="rerun", selection_mode="single-column",
        key=f"cmp_table_{side}_{upload.name}_{upload.size}_{sheet}_{header_row}")
    if len(table) > 500:
        st.caption(f"Preview shows the first 500 of {len(table):,} rows — "
                   "the comparison uses all of them.")

    picked = [c for c in event.selection.columns if c in table.columns]
    if not picked:
        st.info(f"👆 Click a column heading above to select the column of "
                f"Excel {side} to compare.")
        return None
    column = picked[0]
    st.success(f"Excel {side} column selected: **{column}** (sheet "
               f"**{sheet}**)")
    description = f"column '{column}' of sheet '{sheet}' in {upload.name}"
    return description, table[column]


def compare_pasted_column(side, text):
    """Pasted values as a column: (description, Series indexed by line no.)."""
    series, extra_cells = excel_compare.pasted_series(text)
    if extra_cells:
        st.warning(f"**{side}:** {extra_cells} pasted line(s) have more than "
                   "one tab-separated cell — only the first cell of each "
                   "line is used.")
    filled = int(series.map(lambda v: bool(str(v).strip())).sum())
    st.success(f"{side}: **{filled:,}** pasted value(s)")
    return f"the values pasted for {side}", series


def _tag_text(block, tag):
    """Return the unescaped, trimmed text of the first <tag>…</tag> in block."""
    m = re.search(r"<" + tag + r">(.*?)</" + tag + r">", block, re.S)
    return html.unescape(m.group(1)).strip() if m else ""


def parse_master_ledgers(xml_text):
    """Parse every <LEDGER> in a Tally All-Masters export into ledger rows."""
    rows = []
    for block in re.findall(r"<LEDGER\b.*?</LEDGER>", xml_text, re.S):
        m = re.search(r'<LEDGER\s+NAME="(.*?)"', block, re.S)
        name = html.unescape(m.group(1)).strip() if m else ""

        # Alias = 2nd <NAME> in LANGUAGENAME.LIST, when distinct from the name.
        names = [html.unescape(n).strip()
                 for n in re.findall(r"<NAME>(.*?)</NAME>", block, re.S)]
        alias = names[1] if len(names) >= 2 and names[1] and names[1] != names[0] else ""

        # Address = all <ADDRESS> lines inside ADDRESS.LIST, joined.
        addr = ""
        am = re.search(r"<ADDRESS\.LIST.*?</ADDRESS\.LIST>", block, re.S)
        if am:
            parts = [html.unescape(a).strip()
                     for a in re.findall(r"<ADDRESS>(.*?)</ADDRESS>", am.group(0), re.S)]
            addr = ", ".join(p for p in parts if p)

        state = (_tag_text(block, "LEDSTATENAME")
                 or _tag_text(block, "STATENAME")
                 or _tag_text(block, "PRIORSTATENAME"))

        # Opening balance: Tally stores Dr as negative, Cr as positive (this is
        # the inverse of the Dr/Cr -> sign mapping in the Master_Ledger view).
        ob_raw = _tag_text(block, "OPENINGBALANCE")
        opening, drcr = "", ""
        if ob_raw:
            try:
                v = float(ob_raw)
                if v != 0:
                    opening = abs(v)
                    drcr = "Dr" if v < 0 else "Cr"
            except ValueError:
                pass

        rows.append({
            "Ledger_Name": name,
            "Alias": alias,
            "Group_Name": _tag_text(block, "PARENT"),
            "Country": _tag_text(block, "COUNTRYNAME"),
            "State_Name": state,
            "Pincode": _tag_text(block, "PINCODE"),
            "Registration_Type": _tag_text(block, "GSTREGISTRATIONTYPE"),
            "GST_NO": _tag_text(block, "PARTYGSTIN"),
            "Opening_Balance": opening,
            "Dr/Cr": drcr,
            "Address": addr,
        })
    return rows


def read_tally_xml(raw_bytes):
    """Decode a Tally XML export, which is usually UTF-16, sometimes UTF-8."""
    for encoding in ("utf-16", "utf-8-sig", "utf-8", "latin-1"):
        try:
            return raw_bytes.decode(encoding)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return raw_bytes.decode("utf-8", errors="replace")


def ledgers_to_excel(rows):
    """Write ledger rows into the Master — Ledger template workbook (bytes)."""
    wb = load_workbook(io.BytesIO(build_template("Master — Ledger")))
    ws = wb["TEMPLATE"]

    # Clear any rows below the header, then write parsed data.
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    headers = [cell.value for cell in ws[1]]
    for r, row in enumerate(rows, start=2):
        for header, value in row.items():
            if header in headers:
                ws.cell(row=r, column=headers.index(header) + 1, value=value)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------
# Validation: the Master — Ledger sheet must not contain duplicate GST numbers.
# --------------------------------------------------------------------------
def find_duplicate_gst(file_bytes):
    """Read the TEMPLATE sheet and locate duplicate GST_NO values.

    Returns (dataframe, duplicate_mask, duplicate_values). The mask is a
    boolean Series over the rows flagged as duplicate GST numbers; blank GST
    numbers are ignored. Returns (None, None, []) if there's no GST_NO column.
    """
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE").fillna("")
    if "GST_NO" not in df.columns:
        return None, None, []

    # Normalise for comparison (trim + uppercase) without mutating the data.
    norm = df["GST_NO"].astype(str).str.strip().str.upper()
    non_blank = norm != ""
    dup_mask = norm.duplicated(keep=False) & non_blank
    dup_values = sorted(norm[dup_mask].unique())
    return df, dup_mask, dup_values


# Columns that are mandatory when a ledger's registration type is "Regular".
REGULAR_REQUIRED_FIELDS = ["GST_NO", "Country", "State_Name"]


def find_missing_regular_fields(df):
    """Flag rows where Registration_Type is Regular but required fields are blank.

    Returns (row_mask, detail_df) where row_mask marks the offending rows and
    detail_df lists each offending ledger with the fields it is missing.
    """
    if df is None or "Registration_Type" not in df.columns:
        return None, None

    reg = df["Registration_Type"].astype(str).str.strip().str.lower() == "regular"
    present = {f: df[f].astype(str).str.strip() != "" if f in df.columns
               else pd.Series(False, index=df.index)
               for f in REGULAR_REQUIRED_FIELDS}

    row_mask = pd.Series(False, index=df.index)
    records = []
    for idx in df.index[reg]:
        missing = [f for f in REGULAR_REQUIRED_FIELDS if not present[f][idx]]
        if missing:
            row_mask[idx] = True
            name = df.at[idx, "Ledger_Name"] if "Ledger_Name" in df.columns else ""
            records.append({
                "Row": int(idx) + 2,  # +2: 1-based + header row
                "Ledger_Name": name,
                "Missing_Fields": ", ".join(missing),
            })
    detail_df = pd.DataFrame(records) if records else None
    return row_mask, detail_df


# --------------------------------------------------------------------------
# Live data: connect to a running Tally instance over ODBC and browse it.
#
# Tally exposes an ODBC server (default port 9000) when "ODBC Server" is
# enabled. This requires (a) the Tally ODBC driver installed on the machine
# running this app and (b) Tally running with ODBC on. It therefore only works
# locally next to Tally — not on a remote/cloud host — so pyodbc is imported
# lazily and any failure degrades to setup instructions instead of crashing.
# --------------------------------------------------------------------------
ODBC_TOOL = "Tally ODBC — Live Data"

# Preset Tally collections, as Tally-dialect SQL ($field names, FROM <type>).
# The Ledgers query mirrors the Master — Ledger Excel template columns, using
# the field names verified against Tally's column metadata:
#   Ledger_Name -> $Name            Group_Name        -> $Parent
#   Country     -> $CountryName      State_Name        -> $LedStateName
#   Pincode     -> $PINCode          Registration_Type -> $GSTRegistrationType
#   GST_NO      -> $PartyGSTIN       Opening_Balance   -> $OpeningBalance
#   Address     -> $Address
# Alias and Dr/Cr have no ODBC column: aliases live in Tally's language-name
# list (not exposed via SQL), and Dr/Cr is just the sign of the opening balance.
ODBC_PRESETS = {
    "Ledgers": "SELECT $Name, $Parent, $CountryName, $LedStateName, $PINCode, "
               "$GSTRegistrationType, $PartyGSTIN, $OpeningBalance, "
               "$_ClosingBalance, $Address FROM Ledger",
    "Groups": "SELECT $Name, $Parent FROM Group",
    "Stock Items": "SELECT $Name, $Parent, $BaseUnits, $ClosingBalance, "
                   "$ClosingValue FROM StockItem",
    "Cost Centres": "SELECT $Name, $Parent FROM CostCentre",
    "Voucher Types": "SELECT $Name, $Parent, $NumberingMethod FROM VoucherType",
    "Vouchers": "SELECT $Date, $VoucherTypeName, $VoucherNumber, "
                "$PartyLedgerName, $Amount FROM Voucher",
}


def list_odbc_drivers():
    """Return the ODBC driver names registered on this machine (or raise)."""
    import pyodbc  # lazy: optional dependency
    return list(pyodbc.drivers())


def build_odbc_conn_str(mode, dsn, driver, host, port):
    """Build a pyodbc connection string from the UI inputs."""
    if mode == "DSN":
        return f"DSN={dsn}"
    return f"DRIVER={{{driver}}};SERVER={host};PORT={port}"


def run_odbc_query(conn_str, sql):
    """Run a query against Tally over ODBC and return a DataFrame.

    Raises ImportError if pyodbc is unavailable, or pyodbc.Error on connection/
    query failure — the caller surfaces these as friendly messages.
    """
    import pyodbc  # lazy: optional dependency, only needed for this section

    conn = pyodbc.connect(conn_str, autocommit=True, timeout=10)
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [d[0] for d in cursor.description] if cursor.description else []
        rows = [tuple(r) for r in cursor.fetchall()]
        return pd.DataFrame.from_records(rows, columns=columns)
    finally:
        conn.close()


def list_collection_columns(conn_str, table):
    """Return the available column names for a Tally collection (no data).

    Uses the ODBC metadata call cursor.columns(); falls back to reading the
    cursor description of a 'SELECT * FROM <table>' if metadata is empty.
    """
    import pyodbc  # lazy: optional dependency

    conn = pyodbc.connect(conn_str, autocommit=True, timeout=10)
    try:
        cursor = conn.cursor()
        cols = [row.column_name for row in cursor.columns(table=table)]
        if not cols:
            cursor.execute(f"SELECT * FROM {table}")
            cols = [d[0] for d in cursor.description] if cursor.description else []
        return cols
    finally:
        conn.close()


# --------------------------------------------------------------------------
# Ledger-name validation for Purchase / Sales.
#
# Tally imports are case-sensitive and need EXACT ledger names — a single typo
# or wrong case makes the whole XML import fail. These helpers fetch the live
# ledger names from Tally (ODBC) and compare them against the PartyLedgerName /
# Dr_LedgerName columns, suggesting the closest real name for any mismatch.
# --------------------------------------------------------------------------
PS_LEDGER_COLUMNS = ["PartyLedgerName", "Dr_LedgerName"]
# Payment / Contra / Receipt also has a credit-side ledger.
PCR_LEDGER_COLUMNS = ["PartyLedgerName", "Dr_LedgerName", "Cr_LedgerName"]


def fetch_tally_ledger_names(conn_str):
    """Return the list of ledger names from a running Tally instance (ODBC)."""
    df = run_odbc_query(conn_str, "SELECT $Name FROM Ledger")
    if df.empty:
        return []
    return [str(v).strip() for v in df.iloc[:, 0].tolist() if str(v).strip()]


def collect_ledger_names(file_bytes, columns):
    """Sorted unique non-blank ledger-name values used in the given columns."""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE").fillna("")
    names = set()
    for col in columns:
        if col in df.columns:
            for v in df[col].astype(str):
                v = v.strip()
                if v:
                    names.add(v)
    return sorted(names)


# Dr/Cr ledger columns that must never be blank, per converter.
REQUIRED_LEDGER_COLUMNS = {
    "Purchase / Sales": ["Dr_LedgerName"],
    "Payment / Contra / Receipt": ["Dr_LedgerName", "Cr_LedgerName"],
}


def find_empty_ledger_rows(file_bytes, columns):
    """Find rows where any of the given Dr/Cr ledger columns is blank.

    An empty Dr/Cr ledger name produces ``<LEDGERNAME></LEDGERNAME>`` in the
    XML, which Tally rejects. Fully-blank rows (every checked column empty) are
    ignored — they are just unused template rows, not partial data. Returns
    ``(detail_df, cols_checked)``: detail_df lists the offending rows (Row +
    each checked column) or None if none are blank.
    """
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE").fillna("")
    cols = [c for c in columns if c in df.columns]
    if not cols:
        return None, cols

    # A blank in any required ledger column is a problem...
    any_blank = pd.Series(False, index=df.index)
    for c in cols:
        any_blank = any_blank | (df[c].astype(str).str.strip() == "")

    # ...unless the ENTIRE row is empty (an unused template row, not a voucher).
    row_all_blank = (df.astype(str).apply(lambda s: s.str.strip())
                     == "").all(axis=1)

    bad_mask = any_blank & ~row_all_blank
    if not bad_mask.any():
        return None, cols

    records = []
    for idx in df.index[bad_mask]:
        rec = {"Row": int(idx) + 2}  # +2: 1-based + header row
        for c in cols:
            rec[c] = df.at[idx, c]
        records.append(rec)
    return pd.DataFrame(records), cols


def _norm_name(s):
    """Lower-case and collapse whitespace for case-insensitive comparison."""
    return " ".join(str(s).lower().split())


def find_dates_out_of_range(file_bytes, date_from, date_to,
                            date_col="Datetime", id_col="Supplier_Invoice"):
    """Find voucher dates outside the [date_from, date_to] range.

    ``date_col`` is the date column to check (``Datetime`` for Purchase/Sales,
    ``DATE_TIME`` for Payment/Contra/Receipt) and ``id_col`` an optional
    identifier column shown in the detail table. Dates are parsed day-first
    (DD/MM/YYYY) with the same robust parser the converter uses, so string /
    datetime / other cell types all work. Returns
    ``(detail_df, unparsed_df)``: detail_df lists rows whose date is outside the
    range (or None); unparsed_df lists non-blank rows whose date could not be
    parsed at all (or None).
    """
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE")
    if date_col not in df.columns:
        return None, None

    raw = df[date_col]
    non_blank = raw.notna() & (raw.astype(str).str.strip() != "")
    # Day-first (DD/MM/YYYY), matching the converter exactly.
    parsed = tally_core.parse_ddmmyyyy(raw)
    lo, hi = pd.Timestamp(date_from), pd.Timestamp(date_to)

    out_mask = non_blank & parsed.notna() & ((parsed < lo) | (parsed > hi))
    unparsed_mask = non_blank & parsed.isna()

    def _rows(mask):
        records = []
        for idx in df.index[mask]:
            rec = {"Row": int(idx) + 2}  # +2: 1-based + header row
            if id_col and id_col in df.columns:
                rec[id_col] = df.at[idx, id_col]
            rec[date_col] = str(df.at[idx, date_col])
            records.append(rec)
        return pd.DataFrame(records) if records else None

    return _rows(out_mask), _rows(unparsed_mask)


def find_excel_converted_dates(file_bytes, date_col, id_col=None):
    """Detect date cells Excel stored as real dates instead of literal text.

    When a file is filled using the current template the date column is Text, so
    ``01/04/2025`` stays the exact string. If cells come back as real datetime
    values it means Excel auto-converted them on entry — and for ambiguous days
    (<= 12) it may have swapped day/month (dd/mm -> mm/dd). Those values can no
    longer be trusted. Returns ``(detail_df, n_total)`` where detail_df samples
    the affected rows, or ``(None, 0)`` if the column is clean (all text/blank).
    """
    import datetime as _dt
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE")
    if date_col not in df.columns:
        return None, 0

    def _is_real_date(v):
        return isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)) and not pd.isna(v)

    mask = df[date_col].map(_is_real_date)
    if not mask.any():
        return None, 0

    records = []
    for idx in df.index[mask][:50]:  # sample: enough to show the problem
        rec = {"Row": int(idx) + 2}  # +2: 1-based + header row
        if id_col and id_col in df.columns:
            rec[id_col] = df.at[idx, id_col]
        rec[date_col] = str(df.at[idx, date_col])
        records.append(rec)
    return pd.DataFrame(records), int(mask.sum())


def find_unparsable_dates(file_bytes, date_col, id_col=None):
    """List non-blank date rows the strict parser can't read, with a suggestion.

    Each item is a dict: ``row_index`` (0-based), ``Row`` (Excel row), ``raw``
    (the original value), ``suggestion`` (a recovered pandas Timestamp or None)
    and, if available, ``id`` (the identifier column value).
    """
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE")
    if date_col not in df.columns:
        return []
    raw = df[date_col]
    non_blank = raw.notna() & (raw.astype(str).str.strip() != "")
    parsed = tally_core.parse_ddmmyyyy(raw)
    mask = non_blank & parsed.isna()

    rows = []
    for idx in df.index[mask]:
        rec = {
            "row_index": int(idx),
            "Row": int(idx) + 2,  # +2: 1-based + header row
            "raw": str(df.at[idx, date_col]),
            "suggestion": tally_core.suggest_date(df.at[idx, date_col]),
        }
        if id_col and id_col in df.columns:
            rec["id"] = df.at[idx, id_col]
        rows.append(rec)
    return rows


def apply_date_corrections(file_bytes, date_col, corrections):
    """Rewrite date cells (keyed by 0-based row index) to 'dd/mm/yyyy' text.

    Only the date column on the TEMPLATE sheet is touched; the cell is stored as
    Text so the corrected value can't be re-mangled by Excel.
    """
    if not corrections:
        return file_bytes
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb["TEMPLATE"]
    headers = [c.value for c in ws[1]]
    if date_col not in headers:
        return file_bytes
    ci = headers.index(date_col) + 1
    for row_index, value in corrections.items():
        cell = ws.cell(row=int(row_index) + 2, column=ci)
        cell.value = value
        cell.number_format = "@"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def date_resolver_ui(file_bytes, date_col, id_col, key_prefix):
    """Confirm/repair unreadable dates before conversion (blocks until clean).

    Suggests a compatible DD/MM/YYYY value for each unreadable date and offers a
    calendar picker to override it. Returns ``(convert_bytes, blocked)``;
    ``blocked`` stays True until every unreadable date has been resolved.
    """
    rows = find_unparsable_dates(file_bytes, date_col, id_col)
    if not rows:
        return file_bytes, False

    corr_key = f"{key_prefix}_datecorr"
    corrections = st.session_state.setdefault(corr_key, {})
    n_sugg = sum(1 for r in rows if r["suggestion"] is not None)

    ui.step("Step 2b — Fix unreadable dates")
    st.warning(
        f"⚠️ **{len(rows)}** date(s) can't be read as DD/MM/YYYY. Confirm the "
        f"suggested fix or pick a date from the calendar. **Conversion is "
        "paused until every date is resolved.**"
    )

    # Fast path: accept every suggestion at once.
    if n_sugg and st.button(f"✨ Apply all {n_sugg} suggested date(s)",
                            key=f"{key_prefix}_date_all", type="primary"):
        for r in rows:
            if r["suggestion"] is not None:
                corrections[r["row_index"]] = r["suggestion"].strftime("%d/%m/%Y")
        st.rerun()

    # Manual review / override, one calendar per row.
    with st.form(f"{key_prefix}_dateform"):
        st.caption("Original value → confirm or change the date, then Apply.")
        picks = {}
        for i, r in enumerate(rows):
            c0, c1, c2 = st.columns([1, 3, 2])
            c0.markdown(f"Row **{r['Row']}**")
            ident = f"**{r['id']}**  " if r.get("id") else ""
            c1.markdown(f"{ident}`{r['raw']}`")
            staged = corrections.get(r["row_index"])
            if staged:
                default = pd.to_datetime(staged, dayfirst=True).date()
            elif r["suggestion"] is not None:
                default = r["suggestion"].date()
            else:
                default = None
            picks[r["row_index"]] = c2.date_input(
                "Date", value=default, format="DD/MM/YYYY",
                min_value=date(2000, 1, 1), max_value=date(2100, 12, 31),
                key=f"{key_prefix}_dp_{i}", label_visibility="collapsed")
        if st.form_submit_button("✅ Apply these dates"):
            for ri, d in picks.items():
                if d is not None:
                    corrections[ri] = d.strftime("%d/%m/%Y")
            st.rerun()

    convert_bytes = apply_date_corrections(file_bytes, date_col, corrections)
    remaining = find_unparsable_dates(convert_bytes, date_col, id_col)
    if remaining:
        st.info(f"⏸️ {len(remaining)} date(s) still unresolved — apply a "
                "suggestion or pick a date above to continue.")
        return convert_bytes, True
    st.success("✅ All dates are now readable.")
    return convert_bytes, False


def closest_ledger_names(name, tally_names, n=6):
    """Return up to n closest Tally ledger names to `name`, best first.

    Case-INSENSITIVE (Tally names are often title-case while uploads are
    upper-case) and word-aware: the score blends a character-level similarity
    with a token-overlap (Jaccard) score, so reordered or extra words (e.g. an
    extra "AND") still rank the right ledger highly. Always returns the best
    candidates so the user has options to pick from.
    """
    q = _norm_name(name)
    q_tokens = set(q.split())
    scored = []
    for cand in tally_names:
        c = _norm_name(cand)
        seq = difflib.SequenceMatcher(None, q, c).ratio()
        c_tokens = set(c.split())
        union = q_tokens | c_tokens
        jacc = len(q_tokens & c_tokens) / len(union) if union else 0.0
        score = 0.6 * seq + 0.4 * jacc
        scored.append((score, cand))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [cand for _score, cand in scored[:n]]


def apply_ledger_corrections(file_bytes, corrections, columns=PS_LEDGER_COLUMNS):
    """Return new .xlsx bytes with the given ledger columns' cells renamed.

    Only the named ledger columns on the TEMPLATE sheet are touched; everything
    else (including any MASTER_LEDGER_NAME_LINK sheet) is preserved verbatim.
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb["TEMPLATE"]
    headers = [c.value for c in ws[1]]
    for col in columns:
        if col in headers:
            ci = headers.index(col) + 1
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(row=r, column=ci)
                key = str(cell.value).strip() if cell.value is not None else ""
                if key in corrections:
                    cell.value = corrections[key]
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def odbc_connection_form(prefix):
    """Render Tally ODBC connection inputs and return a pyodbc connection string."""
    c1, c2 = st.columns(2)
    host = c1.text_input("Host", value="localhost", key=f"{prefix}_host")
    port = c2.number_input("Port", value=9000, step=1, key=f"{prefix}_port")
    mode = st.radio("Connection style", ["DSN", "Driver name"],
                    horizontal=True, key=f"{prefix}_mode")
    if mode == "DSN":
        dsn = st.text_input("ODBC DSN name", value="TallyODBC64_9000", key=f"{prefix}_dsn")
        return build_odbc_conn_str("DSN", dsn, "", host, int(port))
    driver = st.text_input("Driver name (exact)", value="Tally ODBC Driver64",
                           key=f"{prefix}_driver")
    return build_odbc_conn_str("Driver", "", driver, host, int(port))


def date_range_ui(key_prefix):
    """Optional voucher-date range gate (before upload).

    ``key_prefix`` namespaces the Streamlit widget keys so each converter keeps
    its own state. Returns ``(enabled, date_from, date_to)``; when enabled,
    every voucher's date must fall within the range or conversion is blocked.
    """
    ui.step("Step 1b — Voucher date range (optional)")
    enabled = st.checkbox(
        "Restrict voucher dates to a range (your Tally data period)",
        key=f"{key_prefix}_date_enable",
        help="Tally rejects vouchers dated outside the company's open period. "
             "Turn this on to catch out-of-range dates before converting.",
    )
    if not enabled:
        return False, None, None

    c1, c2 = st.columns(2)
    date_from = c1.date_input("From date", key=f"{key_prefix}_date_from",
                              format="DD/MM/YYYY")
    date_to = c2.date_input("To date", key=f"{key_prefix}_date_to",
                            format="DD/MM/YYYY")
    if date_from and date_to and date_from > date_to:
        st.warning("‘From date’ is after ‘To date’ — please fix the range.")
    return True, date_from, date_to


def ledger_check_ui(file_bytes, columns, key_prefix, columns_label):
    """Verify a converter's ledger names against Tally (ODBC), before conversion.

    ``columns`` are the TEMPLATE columns to validate, ``key_prefix`` namespaces
    the Streamlit widget/session keys so multiple converters don't collide, and
    ``columns_label`` is the human-readable list shown in the expander header.

    Returns ``(convert_bytes, blocked)``:

    * ``convert_bytes`` — the (possibly corrected) workbook bytes to convert.
    * ``blocked`` — True while conversion must wait: the Tally check hasn't
      been run yet (and not skipped), or there are unresolved mismatches the
      user hasn't chosen to ignore. The caller must not convert when True.
    """
    convert_bytes = file_bytes
    ledgers_key = f"{key_prefix}_tally_ledgers"
    corr_key = f"{key_prefix}_corrections"
    ui.step("Step 2a — Verify ledger names against Tally")
    with st.expander(f"🔗 Check {columns_label} against your Tally company",
                     expanded=True):
        st.caption(
            "Tally imports are **case-sensitive** and need the exact ledger "
            "names — a small typo or wrong case makes the whole import fail. "
            "Connect to a running Tally (ODBC, same machine) to catch problems "
            "before converting."
        )
        conn_str = odbc_connection_form(f"{key_prefix}_chk")
        st.caption(f"Connection string: `{conn_str}`")

        if st.button("🔍 Fetch ledger names from Tally & check",
                     key=f"{key_prefix}_chk_fetch"):
            try:
                names = fetch_tally_ledger_names(conn_str)
                st.session_state[ledgers_key] = names
                st.session_state[corr_key] = {}
                if not names:
                    st.warning("Connected, but Tally returned no ledgers.")
            except ImportError:
                st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop(ledgers_key, None)
                st.error(f"Could not fetch ledgers from Tally: {exc}")
                st.caption(
                    "Check that Tally is open with ODBC enabled and the port "
                    "matches, and that the driver name / DSN and bitness are right."
                )

        tally_names = st.session_state.get(ledgers_key)
        if not tally_names:
            # Conversion waits until the check is run — unless explicitly skipped.
            skip = st.checkbox(
                "Skip Tally check (convert without validating ledger names)",
                key=f"{key_prefix}_chk_skip")
            if skip:
                st.caption("⚠️ Ledger names will NOT be validated against Tally.")
                return convert_bytes, False
            st.info("⏸️ Conversion is paused — fetch ledger names from Tally above "
                    "to validate, or tick “Skip Tally check”.")
            return convert_bytes, True

        corrections = st.session_state.setdefault(corr_key, {})
        tally_set = set(tally_names)
        used = collect_ledger_names(file_bytes, columns)

        unresolved = []
        for name in used:
            effective = corrections.get(name, name)
            if effective in tally_set:
                continue
            unresolved.append((name, effective, closest_ledger_names(effective, tally_names)))

        st.caption(f"Loaded **{len(tally_names)}** ledger(s) from Tally · "
                   f"checking **{len(used)}** name(s) used in your file.")

        if corrections:
            st.caption(f"✏️ {len(corrections)} fix(es) staged — they'll be applied "
                       "to the file before conversion.")

        if not unresolved:
            st.success("✅ All ledger names match Tally — converting below.")
            corrected = (apply_ledger_corrections(file_bytes, corrections, columns)
                         if corrections else convert_bytes)
            return corrected, False

        # Offer a one-click "apply all suggestions" for names that have one.
        fixable = [u for u in unresolved if u[2]]
        if fixable and st.button(f"✨ Apply all {len(fixable)} suggested fix(es)",
                                 key=f"{key_prefix}_chk_fix_all", type="primary"):
            for name, _eff, sugg in fixable:
                corrections[name] = sugg[0]
            st.rerun()

        st.warning(f"⚠️ **{len(unresolved)}** ledger name(s) don't match Tally. "
                   f"Pick the right ledger (search across all {len(tally_names)} "
                   "Tally ledgers) and click Fix.")

        sorted_names = sorted(tally_names, key=str.lower)
        for i, (name, effective, suggestions) in enumerate(unresolved):
            shown = name if effective == name else f"{name}  →  {effective}"
            st.markdown(f"❌ `{shown}` — pick the correct ledger:")

            # Suggestions first (best match pre-selected), then every other
            # ledger, all in their original Tally case.
            seen = set(suggestions)
            options = suggestions + [n for n in sorted_names if n not in seen]

            sc1, sc2, sc3 = st.columns([3, 4, 1])
            # Case-INSENSITIVE substring search; the list keeps original case
            # and narrows as you type (press Enter to apply the filter).
            query = sc1.text_input(
                "Search", key=f"{key_prefix}_chk_search_{i}",
                placeholder="🔎 search (case-insensitive)…",
                label_visibility="collapsed")
            if query:
                ql = query.lower()
                filtered = [n for n in options if ql in n.lower()]
            else:
                filtered = options
            if not filtered:
                sc1.caption("No ledger contains that text.")
                filtered = options

            choice = sc2.selectbox(
                "Closest Tally ledger", filtered,
                key=f"{key_prefix}_chk_sugg_{i}", label_visibility="collapsed")
            if sc3.button("Fix", key=f"{key_prefix}_chk_fix_{i}"):
                corrections[name] = choice
                st.rerun()

        # Apply whatever has been staged so far so partial fixes still take effect.
        convert_bytes = (apply_ledger_corrections(file_bytes, corrections, columns)
                         if corrections else file_bytes)

        # Conversion stays paused while names are unresolved, unless the user
        # explicitly opts to convert anyway (e.g. ledgers created later in Tally).
        convert_anyway = st.checkbox(
            "Convert anyway, ignoring the unmatched names above",
            key=f"{key_prefix}_chk_convert_anyway")
        if not convert_anyway:
            st.info("⏸️ Conversion is paused until every name is fixed "
                    "(or tick “Convert anyway”).")

    return convert_bytes, (not convert_anyway)


def ps_ledger_check_ui(file_bytes):
    """Ledger-name check for Purchase / Sales (PartyLedgerName / Dr_LedgerName)."""
    return ledger_check_ui(
        file_bytes, PS_LEDGER_COLUMNS, "ps",
        "PartyLedgerName / Dr_LedgerName")


def pcr_ledger_check_ui(file_bytes):
    """Ledger-name check for Payment / Contra / Receipt (Party / Dr / Cr)."""
    return ledger_check_ui(
        file_bytes, PCR_LEDGER_COLUMNS, "pcr",
        "PartyLedgerName / Dr_LedgerName / Cr_LedgerName")


# --------------------------------------------------------------------------
# Master — Ledger: GST numbers that already exist in Tally under another name.
#
# Importing a ledger whose GSTIN is already on a differently-named Tally
# ledger creates a second ledger for the same party. These helpers fetch the
# Tally ledgers with their GSTIN (ODBC) and flag such rows so the user can
# either rename the row to the existing Tally ledger (Fix) or keep it (Ignore).
# --------------------------------------------------------------------------
def _norm_gstin(value):
    """Trim + uppercase a GSTIN; blanks and NaN become ''."""
    s = str(value).strip().upper()
    return "" if s in ("", "NAN", "NONE") else s


def fetch_tally_ledger_gstins(conn_str):
    """Return ``{GSTIN: [ledger names]}`` for Tally ledgers that have a GSTIN."""
    df = run_odbc_query(conn_str, "SELECT $Name, $PartyGSTIN FROM Ledger")
    by_gst = {}
    for name, gst in df.itertuples(index=False, name=None):
        name, gst = str(name or "").strip(), _norm_gstin(gst or "")
        if name and gst:
            by_gst.setdefault(gst, []).append(name)
    return by_gst


def find_gst_name_conflicts(file_bytes, tally_by_gst):
    """Rows whose GST_NO exists in Tally only under a different ledger name.

    A row is fine when any Tally ledger with that GSTIN has the same name
    (case/whitespace-insensitive) — the import then just updates that ledger.
    Returns a list of dicts: row_index, Row, Ledger_Name, GST_NO, tally_names.
    """
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="TEMPLATE").fillna("")
    if "GST_NO" not in df.columns or "Ledger_Name" not in df.columns:
        return []
    conflicts = []
    for idx, row in df.iterrows():
        gst = _norm_gstin(row["GST_NO"])
        names = tally_by_gst.get(gst) if gst else None
        if not names:
            continue
        name = str(row["Ledger_Name"]).strip()
        if _norm_name(name) in {_norm_name(n) for n in names}:
            continue
        conflicts.append({"row_index": int(idx), "Row": int(idx) + 2,
                          "Ledger_Name": name, "GST_NO": gst,
                          "tally_names": sorted(names, key=str.lower)})
    return conflicts


def apply_row_ledger_names(file_bytes, renames, drop_rows=()):
    """Return new .xlsx bytes with ``{row_index: name}`` written to Ledger_Name
    and the ``drop_rows`` (row indexes) removed from the TEMPLATE sheet."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb["TEMPLATE"]
    headers = [c.value for c in ws[1]]
    if "Ledger_Name" not in headers:
        return file_bytes
    ci = headers.index("Ledger_Name") + 1
    for row_index, name in renames.items():
        ws.cell(row=int(row_index) + 2, column=ci).value = name
    # Delete bottom-up so earlier row numbers don't shift.
    for row_index in sorted({int(r) for r in drop_rows}, reverse=True):
        ws.delete_rows(row_index + 2)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def ml_gst_check_ui(file_bytes, key_prefix="ml", drop_rows=()):
    """Step 2a for Master — Ledger: flag GSTINs already in Tally under another
    ledger name, with Fix (use the Tally name) or Ignore (drop the row) per row.

    ``drop_rows`` are row indexes already marked for removal (Step 2 duplicate
    resolver); they are skipped here and removed together with ignored rows,
    so row numbers shown in every step match the uploaded file.
    Returns ``(convert_bytes, blocked)`` like :func:`ledger_check_ui`.
    """
    drop_rows = {int(r) for r in drop_rows}

    def finish(renames=None, ignored=()):
        rows = drop_rows | set(ignored)
        if not renames and not rows:
            return file_bytes
        return apply_row_ledger_names(file_bytes, renames or {}, rows)

    gst_key = f"{key_prefix}_tally_gst"
    fix_key = f"{key_prefix}_corrections"  # {row_index: Tally ledger name}
    ui.step("Step 2a — Check GST numbers against Tally")
    with st.expander("🔗 Find GST numbers already in Tally under a different "
                     "ledger name", expanded=True):
        st.caption(
            "If a GST number in your file already belongs to a Tally ledger with "
            "a **different name**, importing it creates a **duplicate ledger** "
            "for the same party. Connect to a running Tally (ODBC, same machine) "
            "to catch these before converting."
        )
        conn_str = odbc_connection_form(f"{key_prefix}_chk")
        st.caption(f"Connection string: `{conn_str}`")

        if st.button("🔍 Fetch ledgers & GST numbers from Tally & check",
                     key=f"{key_prefix}_chk_fetch"):
            try:
                by_gst = fetch_tally_ledger_gstins(conn_str)
                st.session_state[gst_key] = by_gst
                st.session_state[fix_key] = {}
                if not by_gst:
                    st.warning("Connected, but no Tally ledger has a GST number.")
            except ImportError:
                st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop(gst_key, None)
                st.error(f"Could not fetch ledgers from Tally: {exc}")
                st.caption(
                    "Check that Tally is open with ODBC enabled and the port "
                    "matches, and that the driver name / DSN and bitness are right."
                )

        by_gst = st.session_state.get(gst_key)
        if by_gst is None:
            skip = st.checkbox(
                "Skip Tally check (convert without checking GST numbers)",
                key=f"{key_prefix}_chk_skip")
            if skip:
                st.caption("⚠️ GST numbers will NOT be checked against Tally.")
                return finish(), False
            st.info("⏸️ Conversion is paused — fetch from Tally above to check, "
                    "or tick “Skip Tally check”.")
            return finish(), True

        fixes = st.session_state.setdefault(fix_key, {})
        conflicts = [c for c in find_gst_name_conflicts(file_bytes, by_gst)
                     if c["row_index"] not in drop_rows]
        st.caption(f"Loaded **{sum(map(len, by_gst.values()))}** Tally ledger(s) "
                   f"with a GST number · **{len(by_gst)}** distinct GSTIN(s).")

        if not conflicts:
            st.success("✅ No GST number in your file is already in Tally under "
                       "a different ledger name.")
            return finish(), False

        def status(c):
            if c["row_index"] in fixes:
                return f"✅ Fixed → {fixes[c['row_index']]}"
            if st.session_state.get(f"{key_prefix}_chk_ignore_{c['row_index']}"):
                return "➖ Ignored (row removed)"
            return "❌ Needs action"

        table = pd.DataFrame([{
            "Row": c["Row"], "Status": status(c),
            "Ledger_Name (file)": c["Ledger_Name"], "GST_NO": c["GST_NO"],
            "Existing Tally ledger(s)": ", ".join(c["tally_names"]),
            } for c in conflicts])
        st.warning(
            f"⚠️ **{len(conflicts)}** ledger(s) have a GST number that already "
            "exists in Tally under a **different name**. **Fix** renames the row "
            "to the existing Tally ledger (the import then updates it instead "
            "of creating a duplicate). **Ignore** removes the row from the file, "
            "so that ledger is not imported."
        )
        st.dataframe(
            table.style.apply(
                lambda r: ["background-color: #ffd6d6"] * len(r)
                if r["Status"].startswith("❌") else [""] * len(r), axis=1),
            use_container_width=True, hide_index=True)

        open_rows = [c for c in conflicts if c["row_index"] not in fixes
                     and not st.session_state.get(
                         f"{key_prefix}_chk_ignore_{c['row_index']}")]
        if open_rows:
            b1, b2 = st.columns(2)
            if b1.button(
                    f"✨ Fix all {len(open_rows)} — use the existing Tally names",
                    key=f"{key_prefix}_chk_fix_all", type="primary",
                    use_container_width=True):
                for c in open_rows:
                    fixes[c["row_index"]] = c["tally_names"][0]
                st.rerun()
            if b2.button(
                    f"➖ Ignore all {len(open_rows)} — remove these rows",
                    key=f"{key_prefix}_chk_ignore_all",
                    use_container_width=True):
                # Set the per-row Ignore checkboxes before they render below.
                for c in open_rows:
                    st.session_state[f"{key_prefix}_chk_ignore_{c['row_index']}"] = True
                st.rerun()

        for c in conflicts:
            ri = c["row_index"]
            c0, c1, c2, c3, c4 = st.columns([1, 3, 3, 1.5, 1.5],
                                            vertical_alignment="center")
            c0.markdown(f"Row **{c['Row']}**")
            c1.markdown(f"`{c['Ledger_Name']}`  \n{c['GST_NO']}")
            if ri in fixes:
                c2.markdown(f"→ **{fixes[ri]}**")
                if c3.button("Undo", key=f"{key_prefix}_chk_fix_{ri}"):
                    fixes.pop(ri, None)
                    st.rerun()
                continue
            ignored = c4.checkbox("Ignore", key=f"{key_prefix}_chk_ignore_{ri}")
            if ignored:
                c2.caption("This row will be removed — it won't be imported.")
                continue
            choice = c2.selectbox(
                "Tally ledger", c["tally_names"],
                key=f"{key_prefix}_chk_sugg_{ri}", label_visibility="collapsed")
            if c3.button("Fix", key=f"{key_prefix}_chk_fix_{ri}"):
                fixes[ri] = choice
                st.rerun()

        ignored = [c["row_index"] for c in conflicts
                   if c["row_index"] not in fixes and st.session_state.get(
                       f"{key_prefix}_chk_ignore_{c['row_index']}")]
        convert_bytes = finish(fixes, ignored)
        remaining = [c for c in conflicts if c["row_index"] not in fixes
                     and c["row_index"] not in ignored]
        if remaining:
            st.info(f"⏸️ Conversion is paused — Fix or Ignore the "
                    f"{len(remaining)} remaining row(s) above.")
            return convert_bytes, True
        left = pd.read_excel(io.BytesIO(convert_bytes), sheet_name="TEMPLATE")
        if left.dropna(how="all").empty:
            st.warning("Every row has been ignored — there is nothing left to "
                       "convert. Untick Ignore on a row or Restart.")
            return convert_bytes, True
        st.success("✅ All GST conflicts resolved."
                   + (f" **{len(ignored)}** ignored row(s) will be removed "
                      "before converting." if ignored else ""))
        return convert_bytes, False


# --------------------------------------------------------------------------
# Run / progress / restart helpers shared by the Tally tools.
# --------------------------------------------------------------------------
def run_with_progress(fn, *args, label="Running"):
    """Run ``fn(*args)`` in a worker thread while a progress bar slides along.

    The converters don't report real progress, so the bar eases toward 95%
    while the work runs, then slides to 100% once it finishes. Exceptions
    from ``fn`` are re-raised after the bar is cleared.
    """
    bar = st.progress(0, text=f"{label}… 0%")
    box = {}

    def work():
        try:
            box["result"] = fn(*args)
        except BaseException as exc:  # noqa: BLE001 - re-raised below
            box["error"] = exc

    worker = threading.Thread(target=work, daemon=True)
    add_script_run_ctx(worker, get_script_run_ctx())
    worker.start()
    pct = 0.0
    while worker.is_alive():
        pct += (95 - pct) * 0.06
        bar.progress(int(pct), text=f"{label}… {int(pct)}%")
        worker.join(0.1)
    if "error" in box:
        bar.empty()
        raise box["error"]
    for p in range(int(pct), 101, 4):
        bar.progress(p, text=f"{label}… {p}%")
        time.sleep(0.03)
    bar.empty()  # the caller shows the persistent "Completed — 100%" bar
    return box["result"]


def file_signature(*parts):
    """Stable hash of the run inputs, used to tell if a stored result is current."""
    h = hashlib.sha256()
    for part in parts:
        h.update(part if isinstance(part, bytes) else str(part).encode())
    return h.hexdigest()


def uploader_nonce(state_key):
    """Suffix for an uploader's widget key; bumping it clears the upload."""
    return st.session_state.get(f"{state_key}_nonce", 0)


# Per-file state from the checks (fixes, skip choice, …) cleared on Restart.
# Date-range settings and the Tally connection/ledger list are kept.
_RESTART_STATE = ("_datecorr", "_date_all", "_dp_", "_corrections",
                  "_chk_skip", "_chk_search_", "_chk_sugg_", "_chk_fix_",
                  "_chk_convert_anyway", "_chk_ignore_", "_dedupe")


def restart_tool(state_key, check_prefix=None):
    """on_click for Restart: drop the result, uploaded file and file fixes."""
    st.session_state[f"{state_key}_nonce"] = uploader_nonce(state_key) + 1
    st.session_state.pop(f"{state_key}_result", None)
    if check_prefix:
        for key in list(st.session_state.keys()):
            if any(key.startswith(check_prefix + s) for s in _RESTART_STATE):
                del st.session_state[key]


def restart_button(state_key, check_prefix=None):
    st.button("🔄 Restart", key=f"{state_key}_restart", on_click=restart_tool,
              args=(state_key, check_prefix),
              help="Clear the uploaded file and its results to start fresh.")


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
# Sidebar sections, in the order they appear in the menu and on Home.
TOOL_SECTIONS = {
    "Excel to Tally": list(TOOLS.keys()) + [GST_PORTAL_TOOL, BANK_PDF_TOOL,
                                                ERROR_LOG_TOOL],
    "Tally to Excel": [MASTER_XML_TOOL, ODBC_TOOL],
    "Marg to Tally": [MARG_MASTER_TOOL, MARG_VOUCHER_TOOL,
                                MARG_LEDGER_TOOL],
    "Excel & GST Tools": [CONSOLIDATE_TOOL, GSTR1_JSON_TOOL, COMPARE_TOOL],
}
SECTION_NOTES = {
    "Excel to Tally":
        "Turn your Excel sheets and statements into files Tally can import.",
    "Tally to Excel":
        "Bring data out of Tally and back into Excel.",
    "Marg to Tally":
        "Move a company's books from Marg ERP 9+ into Tally.",
    "Excel & GST Tools":
        "Everyday helpers for Excel and GST files.",
}
SECTION_ICONS = {
    "Excel to Tally": "📥",
    "Tally to Excel": "📤",
    "Marg to Tally": "🔁",
    "Excel & GST Tools": "🧰",
}
# Per tool: icon, the plain-language name shown on screen, a line saying what
# it is for, and extra search words for the Home screen. The dictionary keys
# are the internal tool names used throughout this file.
TOOL_META = {
    "Purchase / Sales": {
        "title": "Purchase & Sales Bills",
        "icon": "🧾",
        "desc": "Turn purchase and sales bills in Excel into Tally vouchers, "
                "with GST worked out for you.",
        "tags": "invoice bill voucher gst cgst sgst igst xml"},
    "Payment / Contra / Receipt": {
        "title": "Payments & Receipts",
        "icon": "💸",
        "desc": "Turn payments, receipts and bank transfers (contra) in "
                "Excel into Tally vouchers.",
        "tags": "voucher bank cash xml"},
    "Master — Ledger": {
        "title": "Ledgers",
        "icon": "📇",
        "desc": "Create customers, suppliers and other ledger masters in "
                "Tally from an Excel list.",
        "tags": "party customer supplier gstin opening balance group xml"},
    "Master — Duties & Taxes": {
        "title": "GST Tax Ledgers",
        "icon": "🏛️",
        "desc": "Create GST tax ledgers (CGST, SGST, IGST) in Tally.",
        "tags": "gst tax ledger xml"},
    "Master — Purchase/Sales Ledgers": {
        "title": "Purchase & Sales Ledgers",
        "icon": "🏷️",
        "desc": "Create purchase and sales ledgers with their GST rates.",
        "tags": "gst rate ledger xml"},
    GST_PORTAL_TOOL: {
        "title": "GST Portal Bills (GSTR-2B)",
        "icon": "🧮",
        "desc": "Turn the Excel downloaded from the GST portal into the "
                "Purchase & Sales Bills format, sorted by GST rate.",
        "tags": "gst portal gstr2b 2b b2b cdnr credit debit note purchase "
                "sales itc rate ledger"},
    BANK_PDF_TOOL: {
        "title": "Bank Statement (PDF)",
        "icon": "🏦",
        "desc": "Read a bank or loan statement PDF and turn it into payment "
                "and receipt entries.",
        "tags": "bank loan statement pdf ai payment receipt"},
    ERROR_LOG_TOOL: {
        "title": "Tally Import Errors",
        "icon": "🩺",
        "desc": "See why a Tally import failed, explained in plain language.",
        "tags": "tally.imp import error log failed"},
    MASTER_XML_TOOL: {
        "title": "Tally Ledgers to Excel",
        "icon": "↩️",
        "desc": "Turn a masters export from Tally back into an Excel list of "
                "ledgers.",
        "tags": "tally export xml excel ledger reverse"},
    ODBC_TOOL: {
        "title": "Live Tally Data",
        "icon": "🔌",
        "desc": "Look at live data from a Tally company that is open on this "
                "computer.",
        "tags": "odbc live sql ledger connect"},
    MARG_MASTER_TOOL: {
        "title": "Marg Masters",
        "icon": "📦",
        "desc": "Move ledgers and other masters from Marg into Tally, or "
                "into Excel.",
        "tags": "marg erp master ledger xml export"},
    MARG_VOUCHER_TOOL: {
        "title": "Marg Vouchers",
        "icon": "📑",
        "desc": "Move sales, purchases, payments and receipts from Marg into "
                "Tally, or into Excel.",
        "tags": "marg erp daybook voucher xml export"},
    MARG_LEDGER_TOOL: {
        "title": "Marg Ledger File (DBF)",
        "icon": "🗂️",
        "desc": "Convert Marg's ledger export (.dbf) into the Master — "
                "Ledger Excel format.",
        "tags": "marg erp dbf ledger master"},
    CONSOLIDATE_TOOL: {
        "title": "Merge Excel Files",
        "icon": "🧩",
        "desc": "Merge several Excel files into one, sheet by sheet.",
        "tags": "merge combine workbook excel"},
    GSTR1_JSON_TOOL: {
        "title": "GSTR-1 JSON to Excel",
        "icon": "📊",
        "desc": "Convert a GSTR-1 JSON file from the GST portal into Excel.",
        "tags": "gst gstr1 json b2b hsn portal"},
    COMPARE_TOOL: {
        "title": "Compare Two Excel Files",
        "icon": "🔍",
        "desc": "Find values that are in one Excel file but missing from "
                "another.",
        "tags": "compare excel missing difference"},
}
ALL_TOOLS = [t for tools in TOOL_SECTIONS.values() for t in tools]
SECTION_OF = {t: s for s, tools in TOOL_SECTIONS.items() for t in tools}

# The active page is kept in session state; ?tool=… in the URL opens a tool
# directly, so a page can be bookmarked or shared.
if "active_tool" not in st.session_state:
    st.session_state["active_tool"] = ui.initial_tool(ALL_TOOLS)

tool_name = st.session_state["active_tool"]
ui.sidebar_nav(TOOL_SECTIONS, TOOL_META, tool_name, SECTION_ICONS)

if tool_name == ui.HOME:
    ui.home(TOOL_SECTIONS, TOOL_META, SECTION_NOTES)
    st.stop()

# ==========================================================================
# Tally error log viewer: reads tally.imp, the log Tally appends to on every
# import, and lists the imports newest first with their errors.
# ==========================================================================
@st.cache_data(show_spinner=False)
def parse_tally_log(raw_bytes):
    return tally_log.parse_import_log(tally_log.decode_log(raw_bytes))


if tool_name == ERROR_LOG_TOOL:
    ui.page_header(ERROR_LOG_TOOL, TOOL_META, SECTION_OF.get(ERROR_LOG_TOOL, ""))
    st.info(
        "Shows what happened on each Tally import, newest first. Tally "
        "writes this log to `tally.imp` in its program folder every time "
        "you import a file.")

    log_path = st.text_input(
        "Log file path", value=tally_log.DEFAULT_LOG_PATH,
        key="tally_log_path",
        help="Where Tally keeps tally.imp on this computer. Change it if "
             "Tally is installed in another folder.")
    c1, c2 = st.columns([3, 1])
    log_upload = c1.file_uploader(
        "Or upload a tally.imp file (optional)", type=["imp", "txt", "log"],
        key="uploader_tally_log")
    c2.write("")
    c2.write("")
    c2.button("🔄 Reload", key="tally_log_reload", use_container_width=True,
              help="Read the log file again after a new import.")

    log_bytes, log_source = None, ""
    if log_upload is not None:
        log_bytes, log_source = log_upload.getvalue(), log_upload.name
    else:
        path = log_path.strip().strip('"')
        try:
            with open(path, "rb") as fh:
                log_bytes, log_source = fh.read(), path
        except FileNotFoundError:
            st.warning(
                f"No log file at `{path}`. Check the path, or upload the "
                "file above. The path only works when this app runs on the "
                "same computer as Tally.")
        except OSError as exc:
            st.error(f"Could not read `{path}`: {exc}")
    if log_bytes is None:
        st.stop()

    log_runs = parse_tally_log(log_bytes)
    if not log_runs:
        st.warning("No import entries found in this file. Make sure it is "
                   "Tally's `tally.imp` log.")
        st.stop()

    failed_runs = [r for r in log_runs if r["errors"] or r["counts"]["Errors"]]
    total_errors = sum(max(len(r["errors"]), r["counts"]["Errors"])
                       for r in log_runs)
    st.caption(f"Reading `{log_source}`")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Imports", len(log_runs))
    m2.metric("With errors", len(failed_runs))
    m3.metric("Total errors", total_errors)
    m4.metric("Latest import", log_runs[0]["when"].strftime("%d %b %y")
              if log_runs[0]["when"] else log_runs[0]["when_text"])

    f1, f2 = st.columns([1, 2])
    only_errors = f1.toggle("Only imports with errors", value=True,
                            key="tally_log_only_errors")
    companies = sorted({r["company"] for r in log_runs if r["company"]})
    chosen = f2.multiselect("Company", companies, key="tally_log_companies",
                            placeholder="All companies")
    search = st.text_input(
        "Search", key="tally_log_search",
        placeholder="Ledger name, file name or error text")

    def _matches(run):
        if only_errors and run not in failed_runs:
            return False
        if chosen and run["company"] not in chosen:
            return False
        if search.strip():
            haystack = " ".join(
                [run["file"], run["company"]]
                + [f"{e['kind']} {e['subject']} {e['detail']}"
                   for e in run["errors"]]).lower()
            return search.strip().lower() in haystack
        return True

    shown_runs = [r for r in log_runs if _matches(r)]
    if not shown_runs:
        st.success("No errors — every import in this log went through."
                   if only_errors and not failed_runs
                   else "No imports match these filters.")
        st.stop()
    st.caption(f"Showing {len(shown_runs)} of {len(log_runs)} import(s), "
               "newest first.")

    for card, run in enumerate(shown_runs):
        counts = run["counts"]
        n_errors = max(len(run["errors"]), counts["Errors"])
        when = (run["when"].strftime("%d %b %Y, %I:%M %p")
                if run["when"] else run["when_text"])
        with st.container(key=f"{'logerr' if n_errors else 'logok'}_{card}"):
            top, badge = st.columns([4, 1])
            top.markdown(f"**{run['file_name']}**  \n{when}")
            if n_errors:
                badge.error(f"{n_errors} error{'s' if n_errors != 1 else ''}")
            else:
                badge.success("OK")
            st.caption(f"Company: {run['company'] or '—'}  ·  "
                       f"File: `{run['file']}`")
            st.markdown("  ·  ".join(
                f"{key} **{counts[key]}**"
                for key in tally_log.COUNT_KEYS if key != "Errors"))
            for err in run["errors"]:
                st.markdown(
                    f"🔴 **{err['kind']}: {err['subject']}**  \n"
                    f"{err['detail']}")
                if err["hint"]:
                    st.caption("💡 " + err["hint"])
            unlisted = counts["Errors"] - len(run["errors"])
            if unlisted > 0:
                st.markdown(
                    f"🔴 **{unlisted} error{'s' if unlisted != 1 else ''} "
                    "without details**  \nTally counted "
                    f"{'them' if unlisted != 1 else 'it'} in the summary but "
                    "did not write the reason to the log.")

    st.stop()


# ==========================================================================
# Marg ERP 9+  ->  Tally, using Marg's own "Tally Export". These two pages are
# instructions only: Marg writes Tally-importable files itself. The steps and
# labels follow Marg's help pages (care.margcompusoft.com) for ERP 9+.
# ==========================================================================
MARG_TALLY_EXPORT_PATH = (
    "**Reports > Management Report > Utilities > Data Import/Export > "
    "Tally/Busy/ Export/Import**")


def marg_tally_export_prerequisites():
    """The checks Marg asks for before its Tally Export will open."""
    st.subheader("Before you start")
    st.markdown(
        """
- **Close Marg on every other node / client machine.** The export must run
  with only one system logged in.
- **Export rights must be on.**
    - For the company: **Masters > Marg Setups > Control Room**, search
      **Export Import**, set *Option of Data Import/Export/Both/Not Required*
      to **Both** and *Data Import/Export/Y-Both/No* to **Y**.
    - For your user: **Masters > Passwords & Powers > Operator's**, select the
      user, search **Export** and enable it.
    - Press **Esc** and choose **Yes** to save, then log in again (**Ctrl+U**).
- **Update Marg's Tally patch** if the Tally/Busy option does not open.
- **Tally** must be a licensed copy, on the same machine or one you can copy
  the exported files to.
"""
    )


@st.cache_data(show_spinner=False)
def marg_masters_from_xml(file_bytes, standard_groups):
    """Masters XML -> (parsed tables, {file name: xlsx bytes}), cached."""
    parsed = marg_xml.parse_masters(marg_xml.decode_xml(file_bytes),
                                    standard_groups=standard_groups)
    return parsed, marg_xml.master_workbooks(parsed)


@st.cache_data(show_spinner=False)
def marg_vouchers_from_xml(file_bytes):
    """Daybook XML -> (converted tables, {file name: xlsx bytes}), cached."""
    converted = marg_xml.convert_vouchers(
        marg_xml.parse_vouchers(marg_xml.decode_xml(file_bytes)))
    return converted, marg_xml.voucher_workbooks(converted)


def marg_workbook_downloads(books, zip_name, key_prefix):
    """One download button per workbook, plus everything as a .zip."""
    for name, data in books.items():
        st.download_button(f"⬇️ {name}", data=data, file_name=name,
                           mime=XLSX_MIME, key=f"{key_prefix}_{name}",
                           use_container_width=True)
    if len(books) > 1:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in books.items():
                zf.writestr(name, data)
        st.download_button("⬇️ Download all (.zip)", data=buf.getvalue(),
                           file_name=zip_name, mime="application/zip",
                           key=f"{key_prefix}_zip", type="primary",
                           use_container_width=True)


def marg_tally_export_footer():
    st.caption(
        "Steps follow Marg's help pages for Marg ERP 9+. Option names can "
        "differ slightly between Marg versions — press **F1** on any Marg "
        "screen for its built-in help.")


if tool_name == MARG_MASTER_TOOL:
    ui.page_header(MARG_MASTER_TOOL, TOOL_META, SECTION_OF.get(MARG_MASTER_TOOL, ""))
    st.info(
        "Move your **masters** (ledgers, and stock items if you keep "
        "inventory) from Marg ERP 9+ into Tally using Marg's built-in "
        "**Tally Export**. Do this before exporting vouchers — Tally needs "
        "the masters to exist first."
    )
    marg_tally_export_prerequisites()

    st.divider()
    ui.step("Step 1 — Export the masters from Marg")
    st.markdown(
        f"""
1. In Marg, go to {MARG_TALLY_EXPORT_PATH}.
2. Select **Tally 9 Import/Export** and click **Okay**.
3. In the **Tally Import/Export** window, set **Import/Export** to
   **Export Marg to Tally**.
4. In **Export Options**, select **Masters**.
5. Set **Include Inventory** to **Yes** if you want stock items as well as
   ledgers; otherwise **No**.
6. In **Master File**, browse to the folder where the masters file should be
   saved, and note that folder.
7. Click **Proceed** and wait for the export to finish.
"""
    )

    st.divider()
    ui.step("Step 2 — Import the masters into Tally.ERP 9")
    st.markdown(
        """
1. Open the company (**Alt + F3** to create it, with the same financial year
   and books-beginning date as Marg).
2. Press **F12 > General** and set **Import/Export Configuration** to **Yes**.
3. Press **F11 > Statutory & Taxation** and enable **GST**.
4. From **Gateway of Tally**, choose **Import Data > Masters**.
5. Enter the path of the masters file Marg saved in Step 1 and press
   **Enter**.
6. Choose **Modify with New Data**.
"""
    )

    st.divider()
    ui.step("Step 3 — Or convert the masters file to Excel")
    st.caption(
        "Instead of importing Marg's file straight into Tally, upload it "
        "here to get it in this app's **Master** templates, where you can "
        "review and edit it before converting to Tally XML.")
    masters_file = st.file_uploader(
        "Upload the masters XML written by Marg's Tally Export",
        type=["xml"], key="uploader_marg_master_xml")
    if masters_file is not None:
        group_style = st.radio(
            "Group names", ["Standard Tally groups", "Marg's own groups"],
            horizontal=True, key="marg_master_group_style",
            help="Marg has its own sub-groups, such as SUNDRY CREDITORS "
                 "(SUPPLIERS). Standard moves each ledger up to the Tally "
                 "group above it, which already exists in every company.")
        try:
            parsed, books = marg_masters_from_xml(
                masters_file.getvalue(), group_style == "Standard Tally groups")
        except marg_xml.MargXmlError as exc:
            st.error(str(exc))
            st.stop()

        st.success(
            f"{len(parsed['ledger'])} ledger(s), {len(parsed['duties'])} "
            f"duties & taxes ledger(s) and {len(parsed['ps'])} purchase/sales "
            f"ledger(s) converted; {len(parsed['review'])} need a look.")
        stock_items = parsed["counts"].get("STOCKITEM", 0)
        if stock_items:
            st.caption(
                f"The file also holds {stock_items} stock item(s). This app "
                "has no stock-item template, so they are not converted.")
        dupes = marg_xml.duplicate_names(parsed["ledger"])
        if dupes:
            st.warning(
                f"{len(dupes)} ledger name(s) appear more than once. Tally "
                "keeps one ledger per name, so rename or remove the extras: "
                + ", ".join(f"**{n}**" for n in dupes))

        tabs = st.tabs(["Ledgers", "Duties & Taxes", "Purchase/Sales ledgers",
                        f"Needs a look ({len(parsed['review'])})"])
        for tab, key in zip(tabs, ("ledger", "duties", "ps", "review")):
            tab.dataframe(parsed[key], use_container_width=True,
                          hide_index=True)
        tabs[3].caption(
            "These ledgers are not in the Excel files, apart from any marked "
            "'Kept'. The same list is on the REVIEW sheet of "
            "Master_Ledger.xlsx.")

        marg_workbook_downloads(books, "Marg_Masters.zip", "marg_master_dl")
        st.caption(
            "Each file matches the template of the tool with the same name "
            "under **Excel to Tally** — upload it there to get the "
            "Tally XML.")

    st.divider()
    st.subheader("If something goes wrong")
    st.markdown(
        """
- **A master is rejected** — a special character in a ledger or item name is
  the usual cause. Rename it in Marg and export again.
- **Marg writes an `EXP` log and Tally an `IMP` log** listing what failed.
- **You only need the ledgers, or want to review them first** — use
  **Marg Ledger File (DBF)** in this section instead.
"""
    )
    marg_tally_export_footer()
    st.stop()


if tool_name == MARG_VOUCHER_TOOL:
    ui.page_header(MARG_VOUCHER_TOOL, TOOL_META, SECTION_OF.get(MARG_VOUCHER_TOOL, ""))
    st.info(
        "Move your **vouchers** (the day book: sales, purchases, receipts, "
        "payments…) from Marg ERP 9+ into Tally using Marg's built-in "
        "**Tally Export**. Import the **masters** into Tally first — see "
        "**Marg Masters**."
    )
    marg_tally_export_prerequisites()

    st.divider()
    ui.step("Step 1 — Export the vouchers from Marg")
    st.markdown(
        f"""
1. In Marg, go to {MARG_TALLY_EXPORT_PATH}.
2. Select **Tally 9 Import/Export** and click **Okay**.
3. In the **Tally Import/Export** window, set **Import/Export** to
   **Export Marg to Tally**.
4. In **Export Options**, select **Transactions**. (Select both **Masters**
   and **Transactions** to export everything in one go.)
5. Set **Include Inventory** to **Yes** if your bills carry stock items.
6. In **Daybook File**, browse to the folder where the vouchers file should
   be saved, and note that folder.
7. In **Voucher Definition**, select the **Sales Return** and
   **Purchase Return** ledgers. Marg asks for these because Tally has no
   such ledgers of its own; skipping this leaves return vouchers without a
   ledger.
8. In **Voucher From**, enter the from and to dates of the vouchers to
   export.
9. Click **Proceed** and wait for the export to finish.
"""
    )

    st.divider()
    ui.step("Step 2 — Import the vouchers into Tally.ERP 9")
    st.markdown(
        """
1. Open the company and make sure the masters are already imported.
2. Press **Alt + F2** and set the from and to dates to cover the vouchers.
3. From **Gateway of Tally**, choose **Import Data > Vouchers**.
4. Enter the path of the daybook file Marg saved in Step 1 and press
   **Enter**.
5. Choose **Modify with New Data**.
"""
    )

    st.divider()
    ui.step("Step 3 — Check the totals")
    st.markdown(
        """
Compare **Profit & Loss** in Tally with Marg — the purchase and sale totals
should match.

- **"Ledger not found" in Tally** — import the masters again, then the
  vouchers.
- **"Stock item not found"** — export again with **Include Inventory** set
  to **Yes**.
- **A total is off by a small amount** — in Marg open
  **Books > Entry Books > Day Book**, go to the voucher and press **F8** to
  see what is wrong with it.
- **Marg writes an `EXP` log and Tally an `IMP` log** listing what failed.
"""
    )

    st.divider()
    ui.step("Step 4 — Or convert the daybook file to Excel")
    st.caption(
        "Instead of importing Marg's file straight into Tally, upload it "
        "here to get the vouchers in this app's **Purchase & Sales Bills** and "
        "**Payments & Receipts** templates.")
    daybook_file = st.file_uploader(
        "Upload the daybook XML written by Marg's Tally Export",
        type=["xml"], key="uploader_marg_daybook_xml")
    if daybook_file is not None:
        try:
            with st.spinner("Reading vouchers…"):
                converted, books = marg_vouchers_from_xml(
                    daybook_file.getvalue())
        except marg_xml.MargXmlError as exc:
            st.error(str(exc))
            st.stop()

        summary = converted["summary"]
        st.success(f"{int(summary['Vouchers'].sum())} voucher(s) read.")
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.markdown(
            """
- **Sales.xlsx / Purchase.xlsx** — Purchase / Sales template. A bill is put
  here only when the template can reproduce it exactly. Sales and purchases
  are separate files because each needs its own ledger names on the
  `MASTER_LEDGER_NAME_LINK` sheet, which is already filled in with Marg's
  ledgers.
- **Pay_Con_Rec.xlsx** — Payment / Contra / Receipt template.
- **Other_Vouchers.xlsx** — everything else (journals, credit notes, returns,
  and bills the Purchase / Sales template can't hold), as plain Dr/Cr rows in
  the Payment / Contra / Receipt layout. Every ledger and amount is kept as
  Marg posted it; a voucher with several lines becomes several rows.
"""
        )
        st.caption(
            "The templates carry ledger amounts only, so stock items, "
            "quantities and batches are not converted. The Purchase / Sales "
            "template works tax out again from the taxable value, so a tax "
            "ledger can differ from Marg by a few paise per bill, with the "
            "difference going to ROUND OFF. Sale Return and Purchase Return "
            "are Marg voucher types: they must exist in Tally before "
            "Other_Vouchers is imported.")

        review = converted["review"]
        tabs = st.tabs(["Sales", "Purchase", "Pay / Con / Rec",
                        "Other vouchers", f"Moved or split ({len(review)})"])
        for tab, key in zip(tabs, ("sales", "purchase", "pcr", "other",
                                   "review")):
            tab.dataframe(converted[key], use_container_width=True,
                          hide_index=True)
        tabs[4].caption(
            "Vouchers that were not placed in their usual template, or were "
            "split into several rows, and why. The same list is on the "
            "REVIEW sheet of Pay_Con_Rec.xlsx and Other_Vouchers.xlsx.")

        marg_workbook_downloads(books, "Marg_Vouchers.zip", "marg_voucher_dl")
        st.caption(
            "Upload Sales.xlsx and Purchase.xlsx in **Purchase & Sales Bills**, "
            "and the other two in **Payments & Receipts**, under **Excel to "
            "Tally**, to get the Tally XML.")

    marg_tally_export_footer()
    st.stop()


# ==========================================================================
# Marg ERP 9+ ledger export (.dbf)  ->  Master — Ledger Excel / Tally XML.
# ==========================================================================
@st.cache_data(show_spinner=False)
def marg_read_ledgers(file_bytes):
    """Parse the Marg ledger DBF (cached per uploaded file)."""
    return marg_dbf.read_ledger_dbf(file_bytes)


if tool_name == MARG_LEDGER_TOOL:
    ui.page_header(MARG_LEDGER_TOOL, TOOL_META, SECTION_OF.get(MARG_LEDGER_TOOL, ""))
    st.info(
        "Upload the ledger master exported from **Marg ERP 9+** as a `.dbf` "
        "file. Each ledger is converted into the **Ledgers** Excel format, "
        "ready to download as Excel or as Tally XML."
    )

    ui.step("Step 1 — Export the ledgers from Marg as a DBF")
    st.markdown(
        """
1. In Marg ERP 9+, go to **Masters > Ledger Master > Ledger Export**.
   (The same option is also under **Reports > Management Report > Utilities
   > Data Import/Export > Ledger Export**.)
2. Select the ledgers to export — select all of them for a full migration.
3. The export window shows the **folder path** where the file will be
   saved (by default, your Marg folder). Note it, then confirm the export.
4. Open that folder and pick up the exported **`.dbf`** file.

If the option is greyed out or says you are not authorised, switch on
export rights first: **Masters > Marg Setups > Control Room**, search
**Export Import**, set it to **Both** / **Y**, then log in again.
"""
    )

    st.divider()
    ui.step("Step 2 — Upload the Marg ledger DBF")
    dbf_file = st.file_uploader(
        "Upload the Marg ledger export (.dbf)", type=["dbf"],
        key="uploader_marg_ledger_dbf")

    if dbf_file is None:
        st.stop()

    dbf_bytes = dbf_file.getvalue()
    try:
        marg_records = marg_read_ledgers(dbf_bytes)
    except marg_dbf.MargDbfError as exc:
        st.error(str(exc))
        st.stop()
    st.success(f"Read {len(marg_records)} ledger(s) from Marg.")

    st.divider()
    ui.step("Step 3 — Map Marg groups to Tally groups")
    st.caption(
        "Marg stores a group **code** on each ledger, not a group name. The "
        "Tally groups below are suggestions worked out from the ledgers under "
        "each code — check them, change any that are wrong, and untick the "
        "codes you don't want. Tax, purchase and sales ledgers are unticked "
        "by default: create those with the **GST Tax Ledgers** and "
        "**Purchase & Sales Ledgers** tools, which set the tax rates."
    )
    group_table = st.data_editor(
        marg_dbf.group_summary(marg_records),
        key=f"marg_groups_{file_signature(dbf_bytes)[:12]}",
        hide_index=True, use_container_width=True,
        disabled=["Marg_Code", "Ledgers", "Examples"],
        column_config={
            "Include": st.column_config.CheckboxColumn("Include", width="small"),
            "Marg_Code": st.column_config.TextColumn("Marg code", width="small"),
            "Ledgers": st.column_config.NumberColumn("Ledgers", width="small"),
            "Tally_Group": st.column_config.SelectboxColumn(
                "Tally group", options=marg_dbf.TALLY_GROUPS, width="medium"),
            "Examples": st.column_config.TextColumn("Example ledgers"),
        })
    group_mapping = {
        row["Marg_Code"]: (row["Tally_Group"] if isinstance(row["Tally_Group"], str)
                           else "", bool(row["Include"]))
        for _, row in group_table.iterrows()}
    ungrouped = [code for code, (group, include) in group_mapping.items()
                 if include and not group]
    if ungrouped:
        st.warning("Pick a Tally group for the ticked code(s) "
                   + ", ".join(f"`{c}`" for c in ungrouped)
                   + " — their ledgers are left out until you do.")

    st.divider()
    ui.step("Step 4 — Options")
    name_style = st.radio(
        "Ledger name", ["Name + city", "Name only"], horizontal=True,
        key="marg_name_style",
        help="Marg adds the party's city to the end of its name. Keeping it "
             "tells apart parties with the same name in different cities.")
    state_choice = st.selectbox(
        "State for customers/suppliers that have no GST number",
        ["(leave blank)"] + marg_dbf.STATE_NAMES, key="marg_default_state",
        help="For ledgers with a GST number the state is taken from its "
             "first two digits.")

    ledger_df, marg_notes = marg_dbf.convert_ledgers(
        marg_records, group_mapping,
        keep_city=name_style == "Name + city",
        default_state="" if state_choice == "(leave blank)" else state_choice)

    st.divider()
    ui.step("Step 5 — Review and download")
    if ledger_df.empty:
        st.error("No ledgers selected — tick at least one group in Step 3.")
        st.stop()
    st.success(f"{len(ledger_df)} of {len(marg_records)} ledger(s) converted "
               f"({len(marg_records) - len(ledger_df)} left out by Step 3).")

    marg_dupes = marg_dbf.duplicate_names(ledger_df)
    if marg_dupes:
        st.warning(
            f"{len(marg_dupes)} ledger name(s) appear more than once. Tally "
            "keeps one ledger per name, so rename or remove the extras in the "
            "Excel before importing: " + ", ".join(f"**{n}**" for n in marg_dupes))
    if marg_notes:
        with st.expander(f"⚠️ {len(marg_notes)} GST number(s) need a look",
                         expanded=True):
            for note in marg_notes:
                st.markdown("- " + note)

    st.dataframe(ledger_df, use_container_width=True, hide_index=True)

    marg_xlsx = ledgers_to_excel(ledger_df.to_dict("records"))
    base_name = os.path.splitext(dbf_file.name)[0]
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Download Excel (Master — Ledger format)", data=marg_xlsx,
        file_name=f"Master_Ledger_from_{base_name}.xlsx", mime=XLSX_MIME,
        use_container_width=True)
    try:
        marg_xml = tally_core.result_xml_bytes(
            tally_core.convert_master_ledger(io.BytesIO(marg_xlsx)))
        d2.download_button(
            "⬇️ Download Tally XML", data=marg_xml,
            file_name=f"Master_Ledger_from_{base_name}.xml",
            mime="application/xml", use_container_width=True)
    except Exception as exc:  # noqa: BLE001
        d2.error(f"Could not build the Tally XML: {exc}")
    st.caption(
        "To run the duplicate-GST and Tally ledger checks first, download the "
        "Excel and upload it in **Excel to Tally → Ledgers**.")

    st.stop()

# ==========================================================================
# Bank / loan statement PDF  ->  Payment/Contra/Receipt template (LLM).
# ==========================================================================
if tool_name == BANK_PDF_TOOL:
    ui.page_header(BANK_PDF_TOOL, TOOL_META, SECTION_OF.get(BANK_PDF_TOOL, ""))
    st.info(
        "Upload a **bank or loan statement PDF**. Every page is rendered to an "
        "image and read by a vision model, which formats each transaction into "
        "the Payment / Contra / Receipt template — **debits → Payment, credits "
        "→ Receipt, cash → Contra** — inferring the ledger names from the "
        "statement."
    )

    model = st.selectbox("Model (OpenAI)", bank_pdf.MODEL_CHOICES,
                         index=bank_pdf.MODEL_CHOICES.index(bank_pdf.DEFAULT_MODEL))
    custom_model = st.text_input(
        "…or enter a custom model name", value="",
        placeholder="leave blank to use the selection above").strip()
    model = custom_model or model

    if bank_pdf.provider_for(model) == "gemini":
        has_key = bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"))
        key_name = "GEMINI_API_KEY"
    else:
        has_key = bool(os.getenv("OPENAI_API_KEY"))
        key_name = "OPENAI_API_KEY"
    if not has_key:
        st.warning(f"`{key_name}` is not set. Add it to a `.env` file in the "
                   "app folder (or the environment) and reload.")

    pdf_file = st.file_uploader("Upload the statement PDF", type=["pdf"],
                                key=f"uploader_bank_pdf_{uploader_nonce('bank_pdf')}")

    if pdf_file is not None:
        sig = file_signature(pdf_file.getvalue(), model)
        stored = st.session_state.get("bank_pdf_result")
        if not (stored and stored["sig"] == sig):
            stored = None
            run_slot = st.empty()
            if run_slot.button("▶️ Run", key="bank_pdf_run", type="primary"):
                run_slot.empty()
                try:
                    pcr_df, usage = run_with_progress(
                        convert_statement_pdf, pdf_file.getvalue(), model,
                        label=f"Reading the statement pages with {model}")
                    stored = {"sig": sig, "pcr_df": pcr_df, "usage": usage}
                    st.session_state["bank_pdf_result"] = stored
                except ImportError as exc:
                    st.error("A required package is missing. Run "
                             "`pip install openai google-genai pypdfium2`. "
                             f"({exc})")
                    restart_button("bank_pdf")
                except Exception as exc:  # noqa: BLE001
                    st.error(f"Could not convert the statement: {exc}")
                    restart_button("bank_pdf")
            else:
                st.caption("Click **Run** to read the statement.")

        pcr_df, usage = (stored["pcr_df"], stored["usage"]) if stored else (None, None)
        if stored:
            st.progress(100, text="✅ Completed — 100%")

        if pcr_df is not None and pcr_df.empty:
            st.warning(
                f"**{model}** read the pages but returned no vouchers — its "
                "vision is likely too weak for this statement's table. Switch "
                "to **gpt-5-mini** (or gpt-5) above and try again.")
        elif pcr_df is not None:
            counts = pcr_df["Vch_Type"].value_counts().to_dict()
            st.success(f"Formatted **{len(pcr_df)}** voucher(s) — "
                       + " · ".join(f"**{k}**: {v}" for k, v in counts.items()))
            st.dataframe(pcr_df, use_container_width=True, hide_index=True)

            # ---- Token usage & cost --------------------------------------
            if usage:
                st.subheader("Token usage & cost")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Input tokens", f"{usage['input_tokens']:,}")
                m2.metric("Output tokens", f"{usage['output_tokens']:,}")
                m3.metric("Total tokens", f"{usage['total_tokens']:,}")
                if usage.get("cost_usd") is not None:
                    m4.metric("Est. cost", f"${usage['cost_usd']:.4f}")
                    st.caption(
                        f"Model `{usage['model']}` · rates "
                        f"${usage['input_rate']:.2f}/1M in, "
                        f"${usage['output_rate']:.2f}/1M out · "
                        f"input ${usage['input_cost']:.4f} + "
                        f"output ${usage['output_cost']:.4f}. "
                        "Rates are list prices configured in `bank_pdf.py`."
                    )
                else:
                    m4.metric("Est. cost", "n/a")
                    st.caption(f"No pricing configured for `{usage['model']}` in "
                               "`bank_pdf.py` — showing token counts only.")

            xlsx_bytes = bank_pdf.pcr_to_xlsx(pcr_df)
            c1, c2 = st.columns(2)
            c1.download_button(
                "⬇️ Pay/Contra/Receipt template (.xlsx)",
                data=xlsx_bytes, file_name="Pay_Con_Rec_from_statement.xlsx",
                mime=XLSX_MIME)
            try:
                result = tally_core.convert_pay_con_rec(io.BytesIO(xlsx_bytes))
                c2.download_button(
                    "⬇️ Tally XML", data=tally_core.result_xml_bytes(result),
                    file_name="Pay_Con_Rec_from_statement.xml",
                    mime="application/xml")
            except Exception as exc:  # noqa: BLE001
                c2.warning(f"Could not build XML: {exc}")

            st.caption("Tip: upload the .xlsx under **Payments & "
                       "Receipts** to run the ledger, date and empty-cell checks "
                       "before importing to Tally.")

        if stored:
            restart_button("bank_pdf")

    st.stop()

# ==========================================================================
# Consolidate Excel files: same-named sheets are merged into one.
# ==========================================================================
if tool_name == CONSOLIDATE_TOOL:
    ui.page_header(CONSOLIDATE_TOOL, TOOL_META, SECTION_OF.get(CONSOLIDATE_TOOL, ""))
    st.info(
        "Upload **two or more Excel files**. Sheets with the **same name** are "
        "merged automatically into one sheet — the header is kept **once** "
        "(from the first file) and only the data rows of the other files are "
        "stacked below it, in upload order. Headers that span **several rows** "
        "(title lines, grouped / merged headings) are detected automatically. "
        "Sheets found in only one file are carried over as they are."
    )

    merge_files = st.file_uploader(
        "Upload the Excel files to consolidate",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="uploader_consolidate",
    )

    with st.expander("Options"):
        manual_header = st.checkbox(
            "Set the number of header rows myself",
            value=False, key="consolidate_manual_header",
            help="By default the header is the block of top rows that is "
                 "identical in every file. Tick this if the detected count in "
                 "the summary is wrong.")
        header_rows = None
        if manual_header:
            header_rows = int(st.number_input(
                "Header rows at the top of every sheet", min_value=0,
                max_value=100, value=1, step=1, key="consolidate_header_rows"))
        keep_once = st.checkbox(
            "Keep a single copy of sheets that are identical in every file",
            value=True, key="consolidate_keep_once",
            help="Stops reference / mapping sheets such as "
                 "MASTER_LEDGER_NAME_LINK from being repeated once per file.")
        drop_dupes = st.checkbox(
            "Remove duplicate rows within each merged sheet",
            value=False, key="consolidate_drop_dupes")
        add_source = st.checkbox(
            "Add a Source_File column showing which file each row came from",
            value=False, key="consolidate_add_source",
            help="Leave this off if the result will be uploaded to one of the "
                 "Tally converters.")
        out_name = st.text_input("Output file name", value="Consolidated.xlsx",
                                 key="consolidate_out_name").strip()
    out_name = out_name or "Consolidated.xlsx"
    if not out_name.lower().endswith(".xlsx"):
        out_name += ".xlsx"

    if merge_files:
        if len(merge_files) < 2:
            st.caption("Only one file uploaded — add more to merge their sheets.")
        with st.spinner("Consolidating…"):
            merged_bytes, merge_summary, merge_errors = consolidate_excel(
                tuple((f.name, f.getvalue()) for f in merge_files),
                add_source, drop_dupes, keep_once, header_rows)

        for fname, msg in merge_errors:
            st.error(f"Could not read **{fname}**: {msg}")

        if merged_bytes is None:
            st.error("No sheets could be read from the uploaded file(s).")
        else:
            n_merged = sum(1 for s in merge_summary if s["files"] > 1)
            st.success(
                f"Consolidated **{len(merge_files) - len(merge_errors)}** "
                f"file(s) into **{len(merge_summary)}** sheet(s) — "
                f"**{n_merged}** merged from more than one file.")

            st.dataframe(
                pd.DataFrame([{
                    "Sheet": s["sheet"],
                    "Files": s["files"],
                    "Header rows": ("—" if s["header_rows"] is None
                                    else str(s["header_rows"])),
                    "Data rows": s["rows"],
                    "Rows per file": " + ".join(str(n) for _, n in s["sources"]),
                    "Notes": "; ".join(s["notes"]),
                } for s in merge_summary]),
                use_container_width=True, hide_index=True)

            for s in merge_summary:
                for note in s["notes"]:
                    if note.startswith("Columns not in every file"):
                        st.warning(f"**{s['sheet']}** — {note}. Those cells "
                                   "are left blank for the files without them.")

            st.download_button(
                f"⬇️ Download {out_name}", data=merged_bytes,
                file_name=out_name, mime=XLSX_MIME, type="primary")

            st.subheader("Preview")
            preview_sheet = st.selectbox(
                "Sheet", [s["sheet"] for s in merge_summary],
                key="consolidate_preview_sheet")
            preview_df = pd.read_excel(io.BytesIO(merged_bytes),
                                       sheet_name=preview_sheet, dtype=object,
                                       header=None)
            preview_df.columns = [str(i + 1) for i in range(preview_df.shape[1])]
            st.dataframe(preview_df.head(500).fillna("").astype(str),
                         use_container_width=True)
            if len(preview_df) > 500:
                st.caption(f"Showing the first 500 of {len(preview_df):,} rows.")

            st.caption("Formulas are saved as their calculated values. "
                       "Empty sheets are skipped.")

    st.stop()

# ==========================================================================
# GSTR-1 JSON  ->  Excel (B2B / CDNR / B2CS / NIL / HSN / ALL).
# ==========================================================================
if tool_name == GSTR1_JSON_TOOL:
    ui.page_header(GSTR1_JSON_TOOL, TOOL_META, SECTION_OF.get(GSTR1_JSON_TOOL, ""))
    st.info(
        "Upload one or more **GSTR-1 JSON** files downloaded from the GST "
        "portal, then choose what to export: **B2B, CDNR, B2CS, NIL, HSN** or "
        "**ALL**. Every JSON becomes **its own workbook** — files are never "
        "combined. Sections a file does not contain are skipped."
    )

    json_files = st.file_uploader(
        "Upload the GSTR-1 JSON file(s)",
        type=["json", "zip"],
        accept_multiple_files=True,
        key="uploader_gstr1_json",
    )

    if json_files:
        gstr_returns = []                     # (json name, return dict)
        for f in json_files:
            try:
                gstr_returns.extend(
                    gstr1_json.load_returns(f.name, f.getvalue()))
            except Exception as exc:  # noqa: BLE001
                st.error(f"Could not read **{f.name}**: {exc}")

        if gstr_returns:
            choice = st.radio(
                "What do you want to export?",
                gstr1_json.SECTIONS + [gstr1_json.ALL],
                index=len(gstr1_json.SECTIONS), horizontal=True,
                key="gstr1_json_choice")

            # One workbook per JSON.
            workbooks, used_names = [], set()
            for json_name, ret in gstr_returns:
                sheets, missing = gstr1_json.build_sheets([ret], choice)
                out_file = None
                if sheets:
                    stem = os.path.splitext(json_name)[0]
                    out_file, n = f"{stem}_{choice.lower()}.xlsx", 2
                    while out_file in used_names:
                        out_file = f"{stem}_{choice.lower()}_{n}.xlsx"
                        n += 1
                    used_names.add(out_file)
                workbooks.append({
                    "json": json_name, "ret": ret, "sheets": sheets,
                    "missing": missing, "file": out_file,
                    "data": gstr1_json.sheets_to_xlsx(sheets) if sheets else None,
                })

            made = [w for w in workbooks if w["data"]]
            if not made:
                st.warning(f"None of the uploaded files contains **{choice}** "
                           "data.")
            else:
                st.success(f"Created **{len(made)}** workbook(s) from "
                           f"**{len(workbooks)}** JSON file(s).")
            if len(made) > 1:
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    for w in made:
                        zf.writestr(w["file"], w["data"])
                st.download_button(
                    f"⬇️ Download all {len(made)} workbooks (.zip)",
                    data=zip_buf.getvalue(),
                    file_name=f"GSTR1_{choice.lower()}_workbooks.zip",
                    mime="application/zip", type="primary",
                    key="gstr1_json_zip")

            for i, w in enumerate(workbooks):
                ret = w["ret"]
                with st.expander(f"📄 {w['json']}", expanded=len(workbooks) == 1):
                    present = gstr1_json.available_sections([ret])
                    st.caption(
                        f"GSTIN: **{ret.get('gstin') or '—'}** · Period: "
                        f"**{ret.get('fp') or '—'}** · Sections found: "
                        f"**{', '.join(present) or 'none'}**")
                    if w["missing"]:
                        msg = ("Not in this file, so skipped: "
                               f"**{', '.join(w['missing'])}**.")
                        (st.info if w["sheets"] else st.warning)(msg)
                    if w["data"]:
                        st.download_button(
                            f"⬇️ Download {w['file']}", data=w["data"],
                            file_name=w["file"], mime=XLSX_MIME,
                            type="primary" if len(made) == 1 else "secondary",
                            key=f"gstr1_json_dl_{i}")
                        for tab, (sheet_name, sheet_df) in zip(
                                st.tabs(list(w["sheets"])), w["sheets"].items()):
                            with tab:
                                st.caption(f"{len(sheet_df):,} row(s)")
                                st.dataframe(
                                    sheet_df.head(1000).astype(str),
                                    use_container_width=True, hide_index=True)
                                if len(sheet_df) > 1000:
                                    st.caption("Showing the first 1,000 rows.")
                    others = gstr1_json.other_sections([ret])
                    if others:
                        st.caption("Also in the file but not exported: "
                                   + ", ".join(others) + ".")

    st.stop()

# ==========================================================================
# Find missing values: one column of Excel A against one column of Excel B.
# ==========================================================================
if tool_name == COMPARE_TOOL:
    ui.page_header(COMPARE_TOOL, TOOL_META, SECTION_OF.get(COMPARE_TOOL, ""))
    st.info(
        "Give two lists, **A** and **B** — for each one either upload an "
        "Excel file and **click one column heading in its preview**, or "
        "**paste the values** (one per line). The tool lists the values of "
        "each list that are missing from the other."
    )

    upload_a = upload_b = None
    pasted_a = pasted_b = ""
    for side, number, box in zip("AB", "12", st.columns(2)):
        source = box.radio(
            f"Source {number} — {side}", ["Upload Excel", "Paste values"],
            horizontal=True, key=f"cmp_source_{side}")
        if source == "Upload Excel":
            upload = box.file_uploader(
                f"Upload {number} — Excel {side}", type=["xlsx", "xlsm"],
                accept_multiple_files=False,
                key=f"uploader_compare_{side.lower()}")
            text = ""
        else:
            upload = None
            text = box.text_area(
                f"Paste the values for {side}", height=180,
                key=f"cmp_paste_{side}",
                placeholder="One value per line, e.g. a column copied "
                            "from Excel",
                help="Every line is one value — do not include a heading. "
                     "If a line has several tab-separated cells, only the "
                     "first is used.")
        if side == "A":
            upload_a, pasted_a = upload, text
        else:
            upload_b, pasted_b = upload, text

    have_a = bool(upload_a or pasted_a.strip())
    have_b = bool(upload_b or pasted_b.strip())

    def compare_side(side, upload, text):
        if upload:
            return compare_pick_column(side, upload)
        if text.strip():
            return compare_pasted_column(side, text)
        return None

    picked_a = compare_side("A", upload_a, pasted_a)
    if upload_a and upload_b:
        st.divider()
    picked_b = compare_side("B", upload_b, pasted_b)

    if have_a and have_b:
        st.divider()
        st.subheader("Result")
        with st.expander("Matching options"):
            cmp_case = st.checkbox("Ignore upper / lower case", value=True,
                                   key="cmp_ignore_case")
            cmp_space = st.checkbox("Ignore extra spaces", value=True,
                                    key="cmp_ignore_spaces")

        if not (picked_a and picked_b):
            waiting = [side for side, p in (("A", picked_a), ("B", picked_b))
                       if not p]
            st.warning("Select a column for " + " and ".join(waiting)
                       + " to see the result.")
        else:
            (label_a, series_a), (label_b, series_b) = picked_a, picked_b
            result = excel_compare.compare(series_a, series_b,
                                           ignore_case=cmp_case,
                                           ignore_spaces=cmp_space)
            only_a, only_b = result["only_a"], result["only_b"]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Distinct in A", f"{result['unique_a']:,}")
            m2.metric("Distinct in B", f"{result['unique_b']:,}")
            m3.metric("In both", f"{result['common']:,}")
            m4.metric("Missing in total", f"{len(only_a) + len(only_b):,}")

            if only_a.empty and only_b.empty:
                st.success("Nothing is missing — every value of A "
                           f"({label_a}) is in B ({label_b}), and every "
                           "value of B is in A.")

            for side, other, frame, mine, theirs in (
                    ("A", "B", only_a, label_a, label_b),
                    ("B", "A", only_b, label_b, label_a)):
                if frame.empty:
                    if not (only_a.empty and only_b.empty):
                        st.success(f"**Missing from {other}: none.** "
                                   f"Every value of {mine} is also in "
                                   f"{theirs}.")
                    continue
                st.error(f"**Missing from {other}: {len(frame):,} "
                         f"value(s).** These are in {mine} but not in "
                         f"{theirs}.")
                st.dataframe(frame.astype({"Value": str}),
                             use_container_width=True, hide_index=True)

            st.download_button(
                "⬇️ Download the result (.xlsx)",
                data=excel_compare.result_xlsx(result, label_a, label_b),
                file_name="Missing_values_A_vs_B.xlsx", mime=XLSX_MIME,
                type="primary")
            st.caption("Blank cells are ignored and repeated values are "
                       "counted once. “Row(s)” is where each value sits in "
                       "its own list: the Excel row for an uploaded file, the "
                       "line number for pasted values.")

    st.stop()

# ==========================================================================
# Live tool: browse a running Tally instance over ODBC.
# ==========================================================================
if tool_name == ODBC_TOOL:
    ui.page_header(ODBC_TOOL, TOOL_META, SECTION_OF.get(ODBC_TOOL, ""))
    st.info(
        "Connect to a **running Tally** instance over its ODBC server and view "
        "the data live in this dashboard. This works only when the app runs on "
        "the same machine as Tally (with the Tally ODBC driver installed)."
    )

    with st.expander("ℹ️ How to enable ODBC in Tally", expanded=False):
        st.markdown(
            """
1. Open **Tally** (TallyPrime / Tally.ERP 9) and load your company.
2. Go to **F1: Help → Settings → Connectivity** (or **F12: Configure → Advanced**).
3. Set **ODBC Server** to **Yes** and note the **port** (default **9000**).
4. Keep Tally open while you use this section.
5. `pip install pyodbc` if it isn't already available.

If you see **“HY000 — the driver did not supply an error”**, it's usually a
wrong **driver name** or a **32-bit/64-bit mismatch** — use *List installed
ODBC drivers* below to find the exact name.
"""
        )

    ui.step("Step 1 — Connect")
    col_a, col_b = st.columns(2)
    host = col_a.text_input("Host", value="localhost", key="tally_host")
    port = col_b.number_input("Port", value=9000, step=1, key="tally_port")

    if st.button("🔍 List installed ODBC drivers"):
        try:
            drivers = list_odbc_drivers()
            if drivers:
                st.write("Installed ODBC drivers on this machine:")
                st.code("\n".join(drivers))
                tally_like = [d for d in drivers if "tally" in d.lower()]
                if tally_like:
                    st.success("Tally driver(s) found: " +
                               ", ".join(f"`{d}`" for d in tally_like) +
                               " — copy the exact name into the field below.")
                else:
                    st.warning(
                        "No Tally driver found. Install the Tally ODBC driver "
                        "matching your Python bitness."
                    )
            else:
                st.warning("No ODBC drivers are registered on this machine.")
        except ImportError:
            st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not list drivers: {exc}")

    odbc_mode = st.radio("ODBC connection style", ["DSN", "Driver name"],
                         horizontal=True, key="odbc_mode")
    if odbc_mode == "DSN":
        dsn = st.text_input("ODBC DSN name", value="TallyODBC64_9000",
                            key="odbc_dsn")
        conn_str = build_odbc_conn_str("DSN", dsn, "", host, int(port))
    else:
        driver = st.text_input("Driver name (exact)",
                               value="Tally ODBC Driver64", key="odbc_driver")
        conn_str = build_odbc_conn_str("Driver", "", driver, host, int(port))
    st.caption(f"Connection string: `{conn_str}`")

    if st.button("🔌 Test connection"):
        try:
            df = run_odbc_query(conn_str, "SELECT $Name FROM Company")
            names = [str(v) for v in df.iloc[:, 0].tolist()] if not df.empty else []
            st.success(f"Connected. Open compan(y/ies): {', '.join(names) or '—'}")
        except ImportError:
            st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not connect: {exc}")
            st.caption(
                "Check that Tally is open with ODBC enabled and the port "
                "matches. Confirm the driver name / DSN and bitness."
            )

    st.divider()

    # ---- Browse data ----------------------------------------------------
    ui.step("Step 2 — Browse data")
    preset = st.selectbox("Pick a collection", list(ODBC_PRESETS.keys()))

    # Table name = the word after FROM in the preset SQL (Ledger, Group, …).
    table = ODBC_PRESETS[preset].split("FROM")[-1].strip().split()[0]
    if st.button(f"🧾 Show available columns for {table}"):
        try:
            cols = list_collection_columns(conn_str, table)
            if cols:
                st.write(f"{len(cols)} column(s) available on `{table}`:")
                st.code(", ".join(cols))
            else:
                st.warning("The driver returned no column metadata for this table.")
        except ImportError:
            st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not list columns: {exc}")

    sql = st.text_area("SQL query (Tally dialect)", value=ODBC_PRESETS[preset],
                       height=100, key=f"odbc_sql_{preset}")

    if st.button("▶️ Run query", type="primary"):
        with st.spinner("Querying Tally…"):
            try:
                df = run_odbc_query(conn_str, sql)
                if df.empty:
                    st.warning("Query ran but returned no rows.")
                else:
                    st.success(f"{len(df)} row(s).")
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    out = io.BytesIO()
                    df.to_excel(out, index=False)
                    slug = preset.lower().replace(" ", "_")
                    c1, c2 = st.columns(2)
                    c1.download_button(
                        "⬇️ Excel (.xlsx)", data=out.getvalue(),
                        file_name=f"tally_{slug}.xlsx", mime=XLSX_MIME,
                    )
                    c2.download_button(
                        "⬇️ CSV", data=df.to_csv(index=False).encode("utf-8"),
                        file_name=f"tally_{slug}.csv", mime="text/csv",
                    )
            except ImportError:
                st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
            except Exception as exc:  # noqa: BLE001
                st.error(f"Query failed: {exc}")

    st.stop()

# ==========================================================================
# Reverse tool: Tally Master XML  ->  Master-Ledger Excel.
# ==========================================================================
if tool_name == MASTER_XML_TOOL:
    ui.page_header(MASTER_XML_TOOL, TOOL_META, SECTION_OF.get(MASTER_XML_TOOL, ""))
    st.info(
        "Upload a Tally **All Masters** XML export (e.g. `Master.xml`). Each "
        "ledger is extracted into the **Ledgers** Excel format "
        "(Ledger_Name, Alias, Group_Name, Country, State_Name, Pincode, "
        "Registration_Type, GST_NO, Opening_Balance, Dr/Cr, Address)."
    )

    ui.step("Step 1 — Export your masters from Tally as XML")
    st.markdown(
        """
1. In Tally, navigate to **Display**.
2. Open **List of Accounts**.
3. Press **Alt + E** to **Export**.
4. Change the **Format** to **XML**.
5. Check the **output file location** (note the folder where the file is saved).
6. Click **Export**.
7. **Upload the XML file** below.
"""
    )

    st.divider()

    ui.step("Step 2 — Upload your Tally Master XML")
    xml_file = st.file_uploader(
        "Upload the Tally All-Masters XML export",
        type=["xml"],
        key="uploader_master_xml",
    )

    if xml_file is not None:
        with st.spinner("Parsing ledgers…"):
            try:
                xml_text = read_tally_xml(xml_file.getvalue())
                rows = parse_master_ledgers(xml_text)

                if not rows:
                    st.error(
                        "No <LEDGER> records were found. Make sure this is a "
                        "Tally 'All Masters' XML export."
                    )
                else:
                    st.success(f"Done! Found {len(rows)} ledger(s).")
                    df = pd.DataFrame(rows, columns=LEDGER_COLUMNS)
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    st.download_button(
                        "⬇️ Download Excel (Master — Ledger format)",
                        data=ledgers_to_excel(rows),
                        file_name="Master_Ledger_from_XML.xlsx",
                        mime=XLSX_MIME,
                    )
            except Exception as exc:  # noqa: BLE001
                st.error(f"Conversion failed: {exc}")
                st.exception(exc)

    st.stop()

# ==========================================================================
# GST portal Excel (GSTR-2B: B2B, B2B-CDNR, …)  ->  Purchase & Sales format.
# Header row, what each sheet becomes, GST-rate grouping of every row, the
# Tally ledger for each rate, then the filled Purchase & Sales template.
# Credit / debit note sheets are checked but not written out yet.
# ==========================================================================
@st.cache_data(show_spinner=False)
def gp_load_sheets(file_bytes):
    return gst_portal.load_sheets(file_bytes)


@st.cache_data(show_spinner=False)
def gp_classify(file_bytes, sheet, header_row, rates=gst_portal.GST_RATES):
    return gst_portal.classify_sheet(gp_load_sheets(file_bytes)[sheet],
                                     header_row, rates)


@st.cache_data(show_spinner=False)
def gp_review_workbook(file_bytes, sheets, header_row, rates):
    return gst_portal.build_review_workbook(
        {s: gp_classify(file_bytes, s, header_row, rates) for s in sheets})


@st.cache_data(show_spinner=False)
def gp_template(file_bytes, header_row, rates, sheets, vch_type, ledgers):
    """Filled Purchase & Sales template for ``sheets``; ``ledgers`` is a
    tuple of (GST column, Tally ledger name)."""
    return gst_portal.build_template(
        [(gp_classify(file_bytes, s, header_row, rates), vch_type) for s in sheets],
        dict(ledgers))


if tool_name == GST_PORTAL_TOOL:
    ui.page_header(GST_PORTAL_TOOL, TOOL_META, SECTION_OF.get(GST_PORTAL_TOOL, ""))
    st.info(
        "Upload the Excel you downloaded from the **GST portal** (for example "
        "GSTR-2B). Each sheet you pick — **B2B**, **B2B-CDNR**, … — is sorted "
        "into the **5 / 12 / 18 / 28 %** GST groups and matched to your Tally "
        "ledgers, ready for the Purchase & Sales Bills format."
    )

    ui.step("Step 1 — Upload the Excel from the GST portal")
    gp_file = st.file_uploader(
        "Upload the GST portal Excel", type=["xlsx", "xlsm"],
        key=f"uploader_gst_portal_{uploader_nonce('gp')}")
    if gp_file is None:
        st.stop()
    restart_button("gp")

    gp_bytes = gp_file.getvalue()
    try:
        gp_sheets = gp_load_sheets(gp_bytes)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read this Excel file: {exc}")
        st.stop()
    gp_filled = [n for n, raw in gp_sheets.items() if not raw.empty]
    if not gp_filled:
        st.error("This Excel file has no data in any sheet.")
        st.stop()
    # Widget keys carry the file's signature, so a new file starts from its
    # own defaults instead of the previous file's answers.
    gp_key = f"gp_{uploader_nonce('gp')}_{file_signature(gp_bytes)[:10]}"

    # ---- Step 2: the header row ------------------------------------------
    ui.step("Step 2 — Which row has the column headings?")
    st.info(
        "Each sheet must have its column headings on **one row only**. If the "
        "headings are spread over two or more rows (grouped or merged "
        "headings), open the file in Excel, bring them into a single row, "
        "save it and upload it again."
    )
    gp_detected = {n: gst_portal.detect_header_row(gp_sheets[n]) for n in gp_filled}
    gp_first_hit = next((n for n in gp_filled if gp_detected[n]), gp_filled[0])
    c_row, c_sheet = st.columns([1, 2])
    gp_header_row = int(c_row.number_input(
        "Header row number (as in Excel)", min_value=1, max_value=1000,
        value=gp_detected[gp_first_hit] or 1, step=1, key=f"{gp_key}_header",
        help="The same row is used for every sheet you convert."))
    gp_peek = c_sheet.selectbox(
        "Preview the top of sheet", gp_filled,
        index=gp_filled.index(gp_first_hit), key=f"{gp_key}_peek")
    gp_top = gp_sheets[gp_peek].head(max(8, gp_header_row + 4)).fillna("").astype(str)
    gp_top.index = range(1, len(gp_top) + 1)
    gp_top.columns = [get_column_letter(i + 1) for i in range(gp_top.shape[1])]
    st.dataframe(
        gp_top.style.apply(
            lambda r: ["background-color: #dcdefa; font-weight: 600"
                       if r.name == gp_header_row else ""] * len(r), axis=1),
        use_container_width=True)
    st.caption(f"Row **{gp_header_row}** (shaded) is read as the headings; "
               "the bills start on the row below it.")

    # ---- Step 3: what each sheet becomes ---------------------------------
    ui.step("Step 3 — Choose what each sheet becomes")
    gp_chosen, gp_bad = {}, []
    for n in gp_sheets:
        c_name, c_pick, c_note = st.columns([3, 3, 4], vertical_alignment="center")
        if n not in gp_filled:
            c_name.markdown(f"**{n}**")
            c_note.caption("Empty sheet — nothing to convert.")
            continue
        rows = gst_portal.data_row_count(gp_sheets[n], gp_header_row)
        c_name.markdown(f"**{n}**  \n:gray[{rows:,} rows]")
        choice = c_pick.selectbox(
            f"Convert {n} as", gst_portal.CONVERSIONS,
            index=gst_portal.CONVERSIONS.index(gst_portal.suggest_conversion(n)),
            key=f"{gp_key}_conv_{n}", label_visibility="collapsed")
        if choice == gst_portal.SKIP:
            c_note.caption("Left out.")
            continue
        missing = gp_classify(gp_bytes, n, gp_header_row)["missing"]
        if missing:
            gp_bad.append(n)
            c_note.error(f"Row {gp_header_row} has no heading for: "
                         + ", ".join(missing))
        else:
            gp_chosen[n] = choice
            c_note.success("Headings found.")

    if gp_bad:
        st.warning(
            "Fix the sheets marked in red before going on: check the header "
            "row in Step 2, and that the sheet has a single row of headings. "
            "Or set them to “Don't convert”.")
        st.stop()
    if not gp_chosen:
        st.info("Choose at least one sheet to convert.")
        st.stop()

    # ---- Step 4: GST rate of every row -----------------------------------
    ui.step("Step 4 — Sort the bills into GST rates")
    st.caption(
        "GST rate = (Integrated Tax + Central Tax + State/UT Tax) × 100 ÷ "
        "Taxable Value. A bill that lands on one of the rates ticked below "
        "goes into that group. Bills that land anywhere else — or that carry "
        "both IGST and CGST/SGST — are highlighted for you to check.")
    st.markdown("**GST rates that apply to this file**")
    gp_rates = tuple(
        rate for col, rate in zip(st.columns(len(gst_portal.GST_RATES) + 1),
                                  gst_portal.GST_RATES)
        if col.checkbox(f"GST {rate}%", value=True, key=f"{gp_key}_rate_{rate}"))
    st.caption("Untick a rate you don't use: its bills then go under “To "
               "check” and no ledger is asked for it.")
    if not gp_rates:
        st.warning("Tick at least one GST rate.")
        st.stop()
    gp_rates_text = " / ".join(str(r) for r in gp_rates)
    gp_results = {n: gp_classify(gp_bytes, n, gp_header_row, gp_rates)
                  for n in gp_chosen}
    for n, result in gp_results.items():
        st.markdown(f"**{n}** → {gp_chosen[n]}")
        cols = st.columns(len(gst_portal.GST_RATES) + 1)
        for col, rate in zip(cols, gp_rates):
            col.metric(f"GST {rate}%", f"{result['counts'][rate]:,}")
        flagged = result["unmatched"]
        cols[len(gp_rates)].metric("To check", f"{len(flagged):,}")
        if flagged:
            with st.expander(f"⚠️ {len(flagged):,} highlighted rows in {n} — "
                             f"not {gp_rates_text} %"):
                shown = result["data"].loc[flagged[:500]]
                lead = gst_portal.ADDED_COLUMNS
                shown = shown[lead + [c for c in shown.columns if c not in lead]]
                st.dataframe(
                    shown.fillna("").astype(str).style.apply(
                        lambda r: ["background-color: #fff2a8"] * len(r), axis=1),
                    use_container_width=True)
                note = "The row numbers are the rows in the portal Excel."
                if len(flagged) > 500:
                    note += f" Showing the first 500 of {len(flagged):,}."
                st.caption(note)
    with st.spinner("Preparing the checked Excel…"):
        gp_review = gp_review_workbook(gp_bytes, tuple(gp_chosen), gp_header_row,
                                       gp_rates)
    st.download_button(
        "⬇️ Download the checked Excel (rows to check highlighted)",
        data=gp_review, file_name="GST_Portal_Rate_Check.xlsx", mime=XLSX_MIME,
        key=f"{gp_key}_review")

    # ---- Step 5: Tally ledger for each rate ------------------------------
    ui.step("Step 5 — Pick the Tally ledger for each GST rate")
    st.caption("Connect to the Tally company that is open on this computer, "
               "then pick the ledgers for each GST rate: the purchase / sales "
               "ledger and the CGST and SGST tax ledgers for local bills, and "
               "the purchase / sales ledger and the IGST tax ledger for "
               "interstate bills.")
    with st.expander("🔌 Tally connection", expanded="gp_tally_ledgers" not in st.session_state):
        gp_conn = odbc_connection_form("gp")
        st.caption(f"Connection string: `{gp_conn}`")
        if st.button("🔍 Get the ledger list from Tally", key="gp_fetch"):
            try:
                with st.spinner("Reading ledgers from Tally…"):
                    st.session_state["gp_tally_ledgers"] = sorted(
                        set(fetch_tally_ledger_names(gp_conn)), key=str.lower)
                if not st.session_state["gp_tally_ledgers"]:
                    st.warning("Connected, but Tally returned no ledgers.")
            except ImportError:
                st.error("`pyodbc` is not installed. Run `pip install pyodbc`.")
            except Exception as exc:  # noqa: BLE001
                st.session_state.pop("gp_tally_ledgers", None)
                st.error(f"Could not get the ledgers from Tally: {exc}")
                st.caption(
                    "Check that Tally is open with ODBC enabled and the port "
                    "matches, and that the driver name / DSN and bitness are right.")
        gp_manual = st.checkbox(
            "Tally is not on this computer — type the ledger names instead",
            key="gp_manual")

    gp_ledgers = st.session_state.get("gp_tally_ledgers")
    if gp_ledgers:
        st.caption(f"✅ {len(gp_ledgers):,} ledgers loaded from Tally.")
    elif not gp_manual:
        st.info("⏸️ Get the ledger list from Tally above to continue.")
        st.stop()

    # A bill has either IGST or CGST + SGST, never both. Each rate therefore
    # needs, for local bills, a purchase/sales ledger plus CGST and SGST tax
    # ledgers, and for interstate bills a purchase/sales ledger plus an IGST
    # tax ledger — the 20 rows of the MASTER_LEDGER_NAME_LINK sheet.
    gp_map, gp_unset = {}, []
    for side in sorted({gst_portal.ledger_side(c) for c in gp_chosen.values()}):
        for igst, heading in ((False, "local bills (CGST + SGST)"),
                              (True, "interstate bills (IGST)")):
            st.markdown(f"**{side} — {heading}**")
            used = {rate: sum(gp_results[n]["group_counts"][gst_portal.group_name(rate, igst)]
                              for n, c in gp_chosen.items()
                              if gst_portal.ledger_side(c) == side)
                    for rate in gp_rates}
            for col, rate in zip(st.columns(len(gst_portal.GST_RATES)), gp_rates):
                col.caption(f"{rate}% — {used[rate]:,} rows" if used[rate]
                            else f"{rate}% — not used in this file")
            for title, cells in gst_portal.link_rows(igst, gp_rates):
                for col, (link, rate) in zip(st.columns(len(gst_portal.GST_RATES)), cells):
                    # The label is the fixed link-sheet name: a label that
                    # changes would reset the choice.
                    key = f"gp_ledger_{side}_{link}"
                    if gp_ledgers:
                        picked = col.selectbox(link, gp_ledgers, index=None, key=key,
                                               placeholder="Choose a ledger", help=title)
                    else:
                        picked = col.text_input(link, key=f"{key}_typed", help=title,
                                                placeholder="Exact Tally ledger name").strip()
                    gp_map[(side, link)] = picked or ""
                    if used[rate] and not picked:
                        gp_unset.append(f"{side} {link}")

    if gp_unset:
        st.warning("Still to pick: " + ", ".join(gp_unset) + ".")
        st.stop()
    st.success("Every GST rate in use has its Tally ledgers.")

    # ---- Step 6: the filled Purchase & Sales Bills template --------------
    ui.step("Step 6 — Download the Purchase & Sales Bills file")
    st.caption(
        "Each bill is written in the **Purchase & Sales Bills** format: its "
        "taxable value goes under its GST rate (the IGST column for an "
        "interstate bill, never both), and the ledgers you picked "
        "are filled in on the MASTER_LEDGER_NAME_LINK sheet. Bills marked "
        "“To check” keep their GST columns **blank** and are shaded **red** "
        "— fill those in by hand before converting.")
    for vch in ("Purchase", "Sales"):
        names = tuple(n for n, c in gp_chosen.items() if c == vch)
        if not names:
            continue
        for n in names:
            lacking = gst_portal.find_bill_columns(gp_results[n]["data"].columns)[1]
            if lacking:
                st.warning(f"**{n}** has no column for: {', '.join(lacking)}. "
                           "Those cells are left blank.")
        with st.spinner(f"Filling the {vch} file…"):
            book, written, to_check = gp_template(
                gp_bytes, gp_header_row, gp_rates, names, vch,
                tuple((g, name) for (side, g), name in gp_map.items()
                      if side == vch))
        st.markdown(
            f"**{vch}** — from {', '.join(names)}: **{written:,}** bills, "
            f"**{written - to_check:,}** filled and **{to_check:,}** red to check.")
        st.download_button(
            f"⬇️ Download {vch}_from_GST_Portal.xlsx", data=book,
            file_name=f"{vch}_from_GST_Portal.xlsx", mime=XLSX_MIME,
            type="primary", key=f"{gp_key}_dl_{vch}")
    gp_notes = [n for n, c in gp_chosen.items() if c not in ("Purchase", "Sales")]
    if gp_notes:
        st.info("Credit / debit note sheets (" + ", ".join(gp_notes) + ") are "
                "checked above but are not written to a file yet.")
    st.session_state["gp_ledger_map"] = gp_map

    st.stop()

tool = TOOLS[tool_name]

ui.page_header(tool_name, TOOL_META, SECTION_OF.get(tool_name, ""))
st.info(tool["help"])

# ---- Step 1: download the blank template -------------------------------
ui.step("Step 1 — Download the blank template")
try:
    template_bytes = build_template(tool_name)
    st.download_button(
        label=f"⬇️ Download {tool['template_name']}",
        data=template_bytes,
        file_name=tool["template_name"],
        mime=XLSX_MIME,
    )
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not generate template: {exc}")

# ---- Example: show & download a pre-filled template --------------------
with st.expander("👀 See an example (filled template)"):
    st.caption(
        "Sample rows showing the expected format. Use these as a guide, then "
        "replace them with your own data in the blank template above."
    )
    example_rows = EXAMPLES.get(tool_name)
    if example_rows:
        st.dataframe(pd.DataFrame(example_rows), use_container_width=True, hide_index=True)
        try:
            st.download_button(
                label=f"⬇️ Download filled example — {tool['template_name']}",
                data=build_example(tool_name),
                file_name=f"EXAMPLE_{tool['template_name']}",
                mime=XLSX_MIME,
                key=f"example_{tool['upload_key']}",
            )
        except Exception as exc:  # noqa: BLE001
            st.warning(f"Could not generate the example file: {exc}")

# ---- Optional voucher date range, before upload ------------------------
date_enabled, date_from, date_to = (False, None, None)
if tool_name == "Purchase / Sales":
    date_enabled, date_from, date_to = date_range_ui("ps")
elif tool_name == "Payment / Contra / Receipt":
    date_enabled, date_from, date_to = date_range_ui("pcr")

st.divider()

# ---- Step 2: upload the filled template and convert --------------------
ui.step("Step 2 — Upload your filled template")
uploaded = st.file_uploader(
    "Upload the completed Excel file",
    type=["xlsx", "xls"],
    key=f"uploader_{tool['upload_key']}_{uploader_nonce(tool['upload_key'])}",
)

if uploaded is not None:
    with st.expander("👀  Preview your file", expanded=False):
        ui.excel_preview(uploaded.getvalue(), uploaded.name,
                         key=f"up_{tool['upload_key']}")
    block_conversion = False

    # ---- Wrong template / required sheets: fail early with a clear message --
    required_sheets = ["TEMPLATE"]
    if tool_name == "Purchase / Sales":
        required_sheets.append("MASTER_LEDGER_NAME_LINK")
    try:
        up_wb = load_workbook(io.BytesIO(uploaded.getvalue()), read_only=True)
        sheet_names = up_wb.sheetnames
        up_headers = (sheet_header_set(up_wb["TEMPLATE"])
                      if "TEMPLATE" in sheet_names else set())
    except Exception:  # noqa: BLE001
        sheet_names = None  # let the converter report unreadable files
    if sheet_names is not None:
        match = detect_template_tool(up_headers)
        own_cov = template_coverage(tool_name, up_headers)
        if match and match[0] != tool_name and own_cov < 0.5 and match[1] > own_cov:
            st.error(
                f"🚫 **Wrong template.** This file looks like the "
                f"**{match[0]}** template, not **{tool_name}**. Open "
                f"**{match[0]}** from the sidebar and upload it there, or "
                f"download the {tool_name} template in Step 1."
            )
            st.stop()
        missing_sheets = [s for s in required_sheets if s not in sheet_names]
        if missing_sheets:
            st.error(
                "🚫 This file is missing the sheet(s) "
                + ", ".join(f"**{s}**" for s in missing_sheets)
                + ". Sheets found: " + ", ".join(f"`{s}`" for s in sheet_names)
                + ". Download the blank template in Step 1 (it includes every "
                "required sheet), copy your data into it, and upload again."
            )
            st.stop()

    # ---- Validation (Master — Ledger) ----------------------------------
    drop_rows = []  # row indexes removed before converting (duplicate resolver)
    if tool_name == "Master — Ledger":
        try:
            val_df, dup_mask, dup_values = find_duplicate_gst(uploaded.getvalue())
        except Exception as exc:  # noqa: BLE001
            val_df, dup_mask, dup_values = None, None, []
            st.warning(f"Could not run GST validation: {exc}")

        # (a) Duplicate GST numbers — warning only (does not block).
        if dup_values:
            st.warning(
                f"⚠️ Found **{len(dup_values)}** duplicate GST number(s) across "
                f"**{int(dup_mask.sum())}** row(s). Each GST number should be "
                "unique. The duplicate rows are highlighted below — review them "
                "before converting."
            )
            st.markdown("**Duplicate GST numbers:** " +
                        ", ".join(f"`{g}`" for g in dup_values))

            # Group matching GST numbers on consecutive rows (file order kept
            # within each group) and show the Excel row number for each.
            dup_rows = val_df[dup_mask].copy()
            dup_rows.insert(0, "Row", dup_rows.index + 2)  # 1-based + header
            # GST_NO right after Row, so the matching values are side by side.
            dup_rows.insert(1, "GST_NO", dup_rows.pop("GST_NO"))
            gst_key = dup_rows["GST_NO"].astype(str).str.strip().str.upper()
            dup_rows = dup_rows.assign(_gst=gst_key).sort_values(
                ["_gst", "Row"], kind="stable")
            # Alternate the highlight shade per group so neighbouring groups
            # are easy to tell apart.
            group_no = (dup_rows["_gst"] != dup_rows["_gst"].shift()).cumsum()
            shades = ["#ffd6d6" if g % 2 else "#ffb3b3" for g in group_no]
            # Resolver: keep the first row (file order) of each GST number and
            # remove the rest before converting. Tied to this exact file.
            first_rows = set(dup_rows.drop_duplicates("_gst")["Row"])
            dedupe_drop = sorted(int(r) - 2 for r in dup_rows["Row"]
                                 if r not in first_rows)
            file_sig = file_signature(uploaded.getvalue())
            resolved = st.session_state.get("ml_dedupe") == file_sig
            if resolved:
                drop_rows = dedupe_drop
                dup_rows.insert(1, "Action", ["✅ Keep" if r in first_rows
                                              else "🗑️ Remove"
                                              for r in dup_rows["Row"]])

            dup_rows = dup_rows.drop(columns="_gst").reset_index(drop=True)
            styled = dup_rows.style.apply(
                lambda col: [f"background-color: {s}" for s in shades]
                if col.name == "GST_NO" else [""] * len(col),
                axis=0,
            )
            st.dataframe(styled, use_container_width=True, hide_index=True)

            if resolved:
                st.success(
                    f"✅ Duplicates resolved — the first row of each GST number "
                    f"is kept; **{len(dedupe_drop)}** row(s) ("
                    + ", ".join(str(r + 2) for r in dedupe_drop)
                    + ") will be removed before converting.")
                if st.button("↩️ Undo resolve", key="ml_dedupe_undo"):
                    st.session_state.pop("ml_dedupe", None)
                    st.rerun()
            elif st.button(
                    f"🧹 Resolve duplicates — keep the first row of each GST "
                    f"number, remove the other {len(dedupe_drop)}",
                    key="ml_dedupe_btn", type="primary"):
                st.session_state["ml_dedupe"] = file_sig
                st.rerun()
        elif val_df is not None:
            st.success("✅ No duplicate GST numbers found.")

        # (b) Regular ledgers must have GST_NO, Country and State_Name — blocks.
        # Rows the duplicate resolver removes are not checked.
        if val_df is not None and drop_rows:
            val_df = val_df.drop(index=drop_rows)
        reg_mask, reg_detail = find_missing_regular_fields(val_df)
        if reg_detail is not None:
            block_conversion = True
            st.error(
                f"🚫 **{len(reg_detail)}** ledger(s) have **Registration_Type = "
                "Regular** but are missing required field(s). For Regular "
                "ledgers, **GST_NO**, **Country** and **State_Name** are "
                "mandatory. Fix the rows below, then re-upload."
            )
            missing_cols = ["GST_NO", "Country", "State_Name"]
            bad_rows = val_df[reg_mask]
            styled_missing = bad_rows.style.apply(
                lambda col: [
                    "background-color: #ffd6d6" if (col.name in missing_cols
                                                    and str(v).strip() == "")
                    else "" for v in col
                ],
                axis=0,
            )
            st.dataframe(styled_missing, use_container_width=True)
            st.markdown("**Rows needing attention:**")
            st.dataframe(reg_detail, use_container_width=True, hide_index=True)

    # ---- Validation: no empty Dr/Cr ledger names (Purchase/Sales & PCR) ----
    if tool_name in REQUIRED_LEDGER_COLUMNS:
        try:
            empty_df, checked_cols = find_empty_ledger_rows(
                uploaded.getvalue(), REQUIRED_LEDGER_COLUMNS[tool_name])
        except Exception as exc:  # noqa: BLE001
            empty_df, checked_cols = None, []
            st.warning(f"Could not check for empty ledger names: {exc}")

        if empty_df is not None:
            block_conversion = True
            cols_txt = " / ".join(checked_cols)
            st.error(
                f"🚫 **{len(empty_df)}** row(s) have an empty ledger name. "
                f"**{cols_txt}** must be filled in for every voucher — a blank "
                "ledger name makes the whole Tally import fail. Fix the rows "
                "below, then re-upload."
            )
            styled_empty = empty_df.style.apply(
                lambda col: [
                    "background-color: #ffd6d6"
                    if (col.name in checked_cols and str(v).strip() == "")
                    else "" for v in col
                ],
                axis=0,
            )
            st.dataframe(styled_empty, use_container_width=True, hide_index=True)
        elif checked_cols:
            st.success(f"✅ No empty {' / '.join(checked_cols)} values.")

    if block_conversion:
        st.info("Conversion is paused until the errors above are resolved.")
        st.stop()

    # ---- Step 2a: Ledger-name check against Tally --------------------------
    convert_bytes = uploaded.getvalue()
    if tool_name == "Purchase / Sales":
        convert_bytes, ledger_blocked = ps_ledger_check_ui(uploaded.getvalue())
        if ledger_blocked:
            st.stop()
    elif tool_name == "Payment / Contra / Receipt":
        convert_bytes, ledger_blocked = pcr_ledger_check_ui(uploaded.getvalue())
        if ledger_blocked:
            st.stop()
    elif tool_name == "Master — Ledger":
        convert_bytes, gst_blocked = ml_gst_check_ui(
            uploaded.getvalue(), drop_rows=drop_rows)
        if gst_blocked:
            st.stop()

    DATE_COLS = {
        "Purchase / Sales": ("Datetime", "Supplier_Invoice"),
        "Payment / Contra / Receipt": ("DATE_TIME", "PartyLedgerName"),
    }
    KEY_PREFIX = {"Purchase / Sales": "ps", "Payment / Contra / Receipt": "pcr",
                  "Master — Ledger": "ml"}

    # ---- Date-format sanity check: warn if Excel auto-converted dates -------
    if tool_name in DATE_COLS:
        date_col, id_col = DATE_COLS[tool_name]
        try:
            conv_df, n_conv = find_excel_converted_dates(
                convert_bytes, date_col, id_col)
        except Exception:  # noqa: BLE001
            conv_df, n_conv = None, 0
        if conv_df is not None:
            st.warning(
                f"⚠️ **{n_conv}** {date_col} cell(s) were stored by Excel as "
                "**real dates**, not text. If you pasted dates as DD/MM/YYYY, "
                "Excel may have swapped day↔month for days ≤ 12 (e.g. 04/01 "
                "instead of 01/04). To be safe, **re-download the template** "
                "(its date column is now Text) and paste your dates again — "
                "they'll stay exactly as typed. The conversion below reads "
                "these cells as-is."
            )
            with st.expander(f"Show {date_col} cells stored as Excel dates"):
                st.dataframe(conv_df, use_container_width=True, hide_index=True)

    # ---- Step 2b: Fix unreadable dates (blocks conversion until resolved) ---
    if tool_name in DATE_COLS:
        date_col, id_col = DATE_COLS[tool_name]
        convert_bytes, date_blocked = date_resolver_ui(
            convert_bytes, date_col, id_col, KEY_PREFIX[tool_name])
        if date_blocked:
            st.stop()

    # ---- Step 2c: Voucher date range check -----------------------------
    if date_enabled and tool_name in DATE_COLS:
        date_col, id_col = DATE_COLS[tool_name]
        ui.step("Step 2c — Voucher date range check")
        if date_from and date_to and date_from <= date_to:
            try:
                oor_df, unparsed_df = find_dates_out_of_range(
                    convert_bytes, date_from, date_to, date_col, id_col)
            except Exception as exc:  # noqa: BLE001
                oor_df, unparsed_df = None, None
                st.warning(f"Could not validate voucher dates: {exc}")

            span = f"{date_from:%d/%m/%Y} – {date_to:%d/%m/%Y}"
            if oor_df is not None:
                st.error(
                    f"🚫 **{len(oor_df)}** voucher(s) have a date outside the "
                    f"selected range (**{span}**). Fix these rows or widen the "
                    "range, then re-upload."
                )
                st.dataframe(oor_df, use_container_width=True, hide_index=True)
            if unparsed_df is not None:
                st.warning(
                    f"⚠️ **{len(unparsed_df)}** row(s) have a {date_col} that "
                    "couldn't be read — check the date format (DD/MM/YYYY). "
                    "The affected rows are listed below."
                )
                st.dataframe(unparsed_df, use_container_width=True, hide_index=True)
            if oor_df is None and unparsed_df is None:
                st.success(f"✅ All voucher dates fall within {span}.")
            if oor_df is not None:
                st.info("Conversion is paused until the dates above are fixed.")
                st.stop()
        else:
            st.warning("Set a valid From/To date range above to validate "
                       "voucher dates.")

    ui.step("Step 3 — Run")
    state_key = tool["upload_key"]
    sig = file_signature(convert_bytes)
    stored = st.session_state.get(f"{state_key}_result")
    if not (stored and stored["sig"] == sig):
        stored = None
        run_slot = st.empty()
        if run_slot.button("▶️ Run", key=f"{state_key}_run", type="primary"):
            run_slot.empty()
            try:
                # Hand the converter a fresh, seekable copy of the (corrected) bytes.
                result = run_with_progress(
                    tool["convert"], io.BytesIO(convert_bytes),
                    label="Converting to Tally XML")
                stored = {"sig": sig, "result": result}
                st.session_state[f"{state_key}_result"] = stored
            except Exception as exc:  # noqa: BLE001
                st.error(f"Conversion failed: {exc}")
                st.exception(exc)
                restart_button(state_key, KEY_PREFIX.get(tool_name))
        else:
            st.caption("Click **Run** to convert the uploaded file.")

    if stored:
        result = stored["result"]
        xml_final = result.get("xml_final") if result else None
        file_name = (result.get("file_name") if result else None) or "tally"

        if not xml_final:
            st.progress(100, text="✅ Completed — 100%")
            st.error(
                "No output was produced. Check that the uploaded file matches "
                "the template (correct sheet names and column headers)."
            )
        else:
            st.progress(100, text="✅ Completed — 100%")
            st.success("Done! Download your files below.")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    "⬇️ Tally XML",
                    data=tally_core.result_xml_bytes(result),
                    file_name=f"{file_name}.xml",
                    mime="application/xml",
                )
            with col2:
                st.download_button(
                    "⬇️ Excel (.xlsx)",
                    data=tally_core.result_excel_bytes(result),
                    file_name="data.xlsx",
                    mime=XLSX_MIME,
                )
            with col3:
                st.download_button(
                    "⬇️ Both (.zip)",
                    data=tally_core.result_zip_bytes(result),
                    file_name="files.zip",
                    mime="application/zip",
                )

            with st.expander("Preview generated XML"):
                st.code(xml_final[:5000] + ("\n…" if len(xml_final) > 5000 else ""),
                        language="xml")
        restart_button(state_key, KEY_PREFIX.get(tool_name))
