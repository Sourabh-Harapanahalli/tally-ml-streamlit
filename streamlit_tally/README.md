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

There are also two extra tools:

- **Master XML → Excel (reverse)** — turn a Tally "All Masters" XML export back
  into the Master — Ledger Excel format.
- **Tally ODBC — Live Data** — query a running Tally instance over its ODBC
  server (needs the Tally ODBC driver; install `pyodbc` and run next to Tally).

## Layout

```
streamlit_tally/
├── app.py          <- Streamlit UI
├── tally_core.py   <- conversion engine (Excel <-> Tally XML)
└── requirements.txt
```
