# TALLY ML — Streamlit

A self-contained Streamlit app that converts Excel sheets to Tally XML (and
browses a live Tally instance over ODBC).

All conversion logic lives in [`tally_core.py`](tally_core.py) — pure
pandas/openpyxl, no web framework. The output is byte-for-byte identical to the
original Django implementation this was ported from.

## Setup

```bash
cd streamlit_tally
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Opens at http://localhost:8501. To share on your local network:

```bash
streamlit run app.py --server.address 0.0.0.0
```

## What it does

Each converter (Purchase/Sales, Payment/Contra/Receipt, and the three Master
types):

1. **Step 1** — download the blank Excel template.
2. **Step 2** — upload the filled template; download the resulting
   `.xml`, `.xlsx`, or a `.zip` of both.

For **Purchase / Sales**, an optional step verifies the `PartyLedgerName` and
`Dr_LedgerName` values against the live ledger names in your Tally company
(over ODBC). Tally imports are case-sensitive and need exact names, so any
mismatch is flagged with the closest matching ledger and a one-click **Fix**
button that corrects the file before conversion.

There are also a few extra tools:

- **Master XML → Excel (reverse)** — turn a Tally "All Masters" XML export back
  into the Master — Ledger Excel format.
- **Consolidate Excel Files** — upload several workbooks; sheets with the same
  name are merged into one sheet. The header — single- or multi-row, detected
  as the top rows common to every file — is kept once and only the data rows
  of the other files are stacked below. Logic lives in `excel_merge.py`.
- **GSTR-1 JSON → Excel** — upload a GSTR-1 JSON from the GST portal and export
  B2B, CDNR, B2CS, NIL, HSN or ALL to Excel (detail rows plus rate-wise
  totals). Each JSON becomes its own workbook; missing sections are skipped. Logic lives in `gstr1_json.py`.
- **Find Missing Values (Excel A vs B)** — upload two workbooks, click one column
  in each preview, and see which values of each column are missing from the
  other. Logic lives in `excel_compare.py`.
- **Tally ODBC — Live Data** — query a running Tally instance over its ODBC
  server (needs the Tally ODBC driver; install `pyodbc` and run next to Tally).

## Layout

```
streamlit_tally/
├── app.py          <- Streamlit UI
├── tally_core.py   <- conversion engine (Excel <-> Tally XML)
└── requirements.txt
```
