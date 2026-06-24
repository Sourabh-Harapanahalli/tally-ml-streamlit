"""
TALLY ML — Streamlit front-end.

This reuses the EXACT processing logic from the original Django app
(GSTapp/views.py) so the generated Tally XML/Excel is byte-for-byte the
same. Instead of the Django request/response + templates, we drive the
view functions with a tiny fake request and render the UI with Streamlit.

The Django project itself is left untouched in ../TALLY ML/GST — this app
just imports its view functions.

Run:  streamlit run app.py
"""

import html
import io
import os
import re
import sys

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Must be the first Streamlit command in the script.
st.set_page_config(page_title="TALLY ML", page_icon="📒", layout="centered")

# --------------------------------------------------------------------------
# Locate the original Django project, make it importable and configure it.
# No Django server is started; we only import the view functions and call
# them directly.
# --------------------------------------------------------------------------
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
# Allow overriding via env var; otherwise assume the sibling layout:
#   TALLY_ML/streamlit_tally/app.py   and   TALLY_ML/TALLY ML/GST/
DJANGO_DIR = os.environ.get(
    "TALLY_DJANGO_DIR",
    os.path.join(os.path.dirname(THIS_DIR), "TALLY ML", "GST"),
)
DJANGO_DIR = os.path.abspath(DJANGO_DIR)

if not os.path.isdir(os.path.join(DJANGO_DIR, "GSTapp")):
    st.error(
        f"Could not find the Django project at:\n\n`{DJANGO_DIR}`\n\n"
        "Set the TALLY_DJANGO_DIR environment variable to the folder that "
        "contains `manage.py` and the `GSTapp` package."
    )
    st.stop()

if DJANGO_DIR not in sys.path:
    sys.path.insert(0, DJANGO_DIR)
# The template generators write .xlsx files into the current working
# directory and the blank-template source files live in DJANGO_DIR, so pin
# CWD there.
os.chdir(DJANGO_DIR)


@st.cache_resource
def load_views():
    """Configure Django once and return the views module."""
    import django

    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "GST.settings")
    django.setup()

    from GSTapp import views

    # The view POST methods finish by calling render() to show an HTML page.
    # We don't use templates here, so neutralise render() — we read the
    # results straight out of the (fake) request session instead.
    views.render = lambda *args, **kwargs: None
    return views


views = load_views()


# --------------------------------------------------------------------------
# Minimal stand-in for a Django HttpRequest. The views only ever touch
# `.FILES` and `.session`, so that is all we need to provide.
# --------------------------------------------------------------------------
class FakeRequest:
    def __init__(self, files=None):
        self.FILES = files or {}
        self.session = {}
        self.POST = {}
        self.GET = {}
        self.method = "POST"


# --------------------------------------------------------------------------
# Tool catalogue — maps each converter to its upload key, view class and
# blank-template generator (all straight from the original urls.py/views.py).
# --------------------------------------------------------------------------
TOOLS = {
    "Purchase / Sales": {
        "upload_key": "upload_file",
        "view": lambda: views.Purchase_Sales(),
        "template_fn": lambda: views.download_excel,
        "template_name": "Template_Purchase.xlsx",
        "help": "Purchase & Sales vouchers with GST (CGST/SGST/IGST) breakup. "
                "The template has two sheets: TEMPLATE (your data) and "
                "MASTER_LEDGER_NAME_LINK (map GST columns to your Tally ledger names).",
    },
    "Payment / Contra / Receipt": {
        "upload_key": "upload_file_pay_con_rec",
        "view": lambda: views.Pay_Con_Rec(),
        "template_fn": lambda: views.download_excel_pay_con_rec,
        "template_name": "Template_Pay_Con_Rec.xlsx",
        "help": "Payment, Contra and Receipt vouchers (Dr/Cr ledger entries).",
    },
    "Master — Ledger": {
        "upload_key": "upload_file_master_ledger",
        "view": lambda: views.Master_Ledger(),
        "template_fn": lambda: views.download_excel_master_1,
        "template_name": "Template_Master_Ledger.xlsx",
        "help": "Create ledger masters (groups, GST no., opening balance, address). "
                "The REFERENCE sheet lists the valid Tally group names.",
    },
    "Master — Duties & Taxes": {
        "upload_key": "upload_file_master_duties",
        "view": lambda: views.Master_Duties(),
        "template_fn": lambda: views.download_excel_master_2,
        "template_name": "Template_Master_DPS.xlsx",
        "help": "Create Duties & Taxes ledgers (rate of tax, tax type).",
    },
    "Master — Purchase/Sales Ledgers": {
        "upload_key": "upload_file_master_ps",
        "view": lambda: views.Master_PS(),
        "template_fn": lambda: views.download_excel_master_3,
        "template_name": "Template_Master_PS.xlsx",
        "help": "Create Purchase/Sales ledgers with nature of transaction and GST rates.",
    },
}

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def build_template(tool_name):
    """Generate a blank template workbook and return its bytes."""
    fn = TOOLS[tool_name]["template_fn"]()
    response = fn(FakeRequest())
    return bytes(response.content)


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

# Column order must match the Master — Ledger TEMPLATE sheet.
LEDGER_COLUMNS = [
    "Ledger_Name", "Alias", "Group_Name", "Country", "State_Name", "Pincode",
    "Registration_Type", "GST_NO", "Opening_Balance", "Dr/Cr", "Address",
]


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
# The Ledgers query mirrors every column in the Master — Ledger Excel template
# (Ledger_Name, Alias, Group_Name, Country, State_Name, Pincode,
# Registration_Type, GST_NO, Opening_Balance, Address). Dr/Cr is not a stored
# field — it is the sign of the opening balance — so it is omitted here.
ODBC_PRESETS = {
    "Ledgers": "SELECT $Name, $Alias, $Parent, $CountryName, $LedStateName, "
               "$PinCode, $GSTRegistrationType, $PartyGSTIN, $OpeningBalance, "
               "$ClosingBalance, $Address FROM Ledger",
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
# UI
# --------------------------------------------------------------------------
st.sidebar.title("📒 TALLY ML")
st.sidebar.caption("Excel → Tally XML converter")
tool_name = st.sidebar.radio(
    "Choose a converter", list(TOOLS.keys()) + [MASTER_XML_TOOL, ODBC_TOOL])

# ==========================================================================
# Live tool: browse a running Tally instance over ODBC.
# ==========================================================================
if tool_name == ODBC_TOOL:
    st.title(ODBC_TOOL)
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

    st.subheader("Step 1 — Connect")
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
    st.subheader("Step 2 — Browse data")
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
    st.title(MASTER_XML_TOOL)
    st.info(
        "Upload a Tally **All Masters** XML export (e.g. `Master.xml`). Each "
        "ledger is extracted into the **Master — Ledger** Excel format "
        "(Ledger_Name, Alias, Group_Name, Country, State_Name, Pincode, "
        "Registration_Type, GST_NO, Opening_Balance, Dr/Cr, Address)."
    )

    st.subheader("Step 1 — Export your masters from Tally as XML")
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

    st.subheader("Step 2 — Upload your Tally Master XML")
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

tool = TOOLS[tool_name]

st.title(tool_name)
st.info(tool["help"])

# ---- Step 1: download the blank template -------------------------------
st.subheader("Step 1 — Download the blank template")
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

st.divider()

# ---- Step 2: upload the filled template and convert --------------------
st.subheader("Step 2 — Upload your filled template")
uploaded = st.file_uploader(
    "Upload the completed Excel file",
    type=["xlsx", "xls"],
    key=f"uploader_{tool['upload_key']}",
)

if uploaded is not None:
    block_conversion = False

    # ---- Validation (Master — Ledger) ----------------------------------
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

            dup_rows = val_df[dup_mask]
            styled = dup_rows.style.apply(
                lambda col: ["background-color: #ffd6d6"] * len(col)
                if col.name == "GST_NO" else [""] * len(col),
                axis=0,
            )
            st.dataframe(styled, use_container_width=True)
        elif val_df is not None:
            st.success("✅ No duplicate GST numbers found.")

        # (b) Regular ledgers must have GST_NO, Country and State_Name — blocks.
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

    if block_conversion:
        st.info("Conversion is paused until the errors above are resolved.")
        st.stop()

    with st.spinner("Converting to Tally XML…"):
        try:
            # Hand the view a fresh, seekable copy of the bytes under the
            # exact key the original Django form used.
            buffer = io.BytesIO(uploaded.getvalue())
            req = FakeRequest(files={tool["upload_key"]: buffer})

            tool["view"]().post(req)

            xml_final = req.session.get("xml_final")
            csv_string = req.session.get("data_csv")
            file_name = req.session.get("file_name") or "tally"

            if not xml_final:
                st.error(
                    "No output was produced. Check that the uploaded file matches "
                    "the template (correct sheet names and column headers)."
                )
            else:
                st.success("Done! Download your files below.")

                # Reuse the original download views for identical output.
                xml_resp = views.download_xml_file(req)
                zip_resp = views.download_xml(req)
                xlsx_resp = views.download_xml_excel(req)

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.download_button(
                        "⬇️ Tally XML",
                        data=bytes(xml_resp.content),
                        file_name=f"{file_name}.xml",
                        mime="application/xml",
                    )
                with col2:
                    st.download_button(
                        "⬇️ Excel (.xlsx)",
                        data=bytes(xlsx_resp.content),
                        file_name="data.xlsx",
                        mime=XLSX_MIME,
                    )
                with col3:
                    st.download_button(
                        "⬇️ Both (.zip)",
                        data=bytes(zip_resp.content),
                        file_name="files.zip",
                        mime="application/zip",
                    )

                with st.expander("Preview generated XML"):
                    st.code(xml_final[:5000] + ("\n…" if len(xml_final) > 5000 else ""),
                            language="xml")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Conversion failed: {exc}")
            st.exception(exc)
