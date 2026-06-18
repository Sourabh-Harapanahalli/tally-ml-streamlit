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

import io
import os
import sys

import streamlit as st

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
# UI
# --------------------------------------------------------------------------
st.sidebar.title("📒 TALLY ML")
st.sidebar.caption("Excel → Tally XML converter")
tool_name = st.sidebar.radio("Choose a converter", list(TOOLS.keys()))
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
