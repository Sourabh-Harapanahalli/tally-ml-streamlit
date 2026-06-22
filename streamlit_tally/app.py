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
# UI
# --------------------------------------------------------------------------
st.sidebar.title("📒 TALLY ML")
st.sidebar.caption("Excel → Tally XML converter")
tool_name = st.sidebar.radio("Choose a converter", list(TOOLS.keys()) + [MASTER_XML_TOOL])

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

    st.subheader("Upload your Tally Master XML")
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
