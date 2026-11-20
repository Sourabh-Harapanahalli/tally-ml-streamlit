"""
Marg ERP 9+ ledger export (.dbf)  ->  Master — Ledger template rows.

Marg's "ledger export" writes its ledger master as a Visual FoxPro DBF: one
record per ledger, fixed-width text columns. This module reads that file
(no third-party DBF library needed) and reshapes it into the columns the
Master — Ledger template uses. Pure Python/pandas, no Streamlit.

Marg columns used:
    PARNAM   ledger name; party ledgers carry the city from column 31 onward
    SCODE    Marg account-group code (C6 = customers, D3x = suppliers, ...)
    GSTNO    GSTIN                 PARADD/PARADD1/PARADD2  address lines
    CITY     city                  DSM                     pincode
    OPNING   opening balance (positive = debit, negative = credit)
"""

import re
import struct

import pandas as pd

LEDGER_COLUMNS = [
    "Ledger_Name", "Alias", "Group_Name", "Country", "State_Name", "Pincode",
    "Registration_Type", "GST_NO", "Opening_Balance", "Dr/Cr", "Address",
]

# Tally groups offered in the mapping table. The first block is the Master —
# Ledger REFERENCE list; the last three have their own master tools in this
# app (they need tax rates), so ledgers mapped to them are skipped by default.
TALLY_GROUPS = [
    "Sundry Debtors", "Sundry Creditors", "Bank Accounts", "Cash-in-Hand",
    "Capital Account", "Current Assets", "Current Liabilities",
    "Direct Expenses", "Indirect Expenses", "Direct Incomes",
    "Indirect Incomes", "Fixed Assets", "Investments", "Stock-in-Hand",
    "Deposits (Asset)", "Loans & Advances (Asset)", "Loans (Liability)",
    "Secured Loans", "Unsecured Loans", "Bank OD A/c", "Provisions",
    "Reserves & Surplus", "Branch / Divisions", "Misc. Expenses (ASSET)",
    "Suspense A/c", "S.CRS.EXP.PAYABLE", "Workers Advance",
    "Duties & Taxes", "Purchase Accounts", "Sales Accounts",
]
OTHER_TOOL_GROUPS = {"Duties & Taxes", "Purchase Accounts", "Sales Accounts"}

# Marg group code -> Tally group. Inferred from the ledgers Marg keeps under
# each code, not from Marg documentation, so the UI lets the user change it.
# A code not listed falls back to its longest listed prefix (D31 -> D3).
DEFAULT_GROUPS = {
    "B": "Capital Account",
    "C": "Current Assets",
    "C1": "Bank Accounts",
    "C2": "Cash-in-Hand",
    "C5": "Stock-in-Hand",
    "C6": "Sundry Debtors",
    "Cx": "Duties & Taxes",        # GST deposit on advance
    "D": "Current Liabilities",
    "D1": "Duties & Taxes",
    "Dy": "Duties & Taxes",        # GST on advance / reverse charge
    "D2": "Current Liabilities",   # expenses payable
    "D3": "Sundry Creditors",
    "E": "Fixed Assets",
    "F": "Investments",
    "J1": "Direct Expenses",
    "J12": "Indirect Expenses",
    "J4": "Indirect Incomes",
    "J5": "Purchase Accounts",
    "J6": "Sales Accounts",
    "K": "Suspense A/c",
}
# Marg codes left out by default regardless of group: "I" holds only Marg's
# Profit & Loss A/c, which Tally already has as a reserved ledger.
DEFAULT_SKIP_CODES = {"I"}

# GST state code (first two digits of a GSTIN) -> state name as Tally spells it.
GST_STATES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
    "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman & Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh", "38": "Ladakh",
}
STATE_NAMES = sorted(set(GST_STATES.values()))

_GSTIN_RE = re.compile(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]")
_PIN_RE = re.compile(r"(?<!\d)[1-9]\d{5}(?!\d)")
_NAME_WIDTH = 30    # PARNAM: name in the first 30 characters, then the city
_ADDR_WIDTH = 45    # PARADD* column width; a full line was wrapped mid-word
PARTY_GROUPS = {"Sundry Debtors", "Sundry Creditors"}


class MargDbfError(ValueError):
    """The upload is not a readable Marg ledger DBF."""


def read_dbf(file_bytes):
    """Parse a DBF file into a list of {column: text} dicts.

    Values keep their leading spaces (PARNAM is position-sensitive) and lose
    only trailing padding. Records flagged deleted are skipped.
    """
    if len(file_bytes) < 33:
        raise MargDbfError("The file is too small to be a DBF export.")
    n_records, header_len, record_len = struct.unpack("<IHH", file_bytes[4:12])
    fields, offset = [], 32
    while offset + 32 <= len(file_bytes) and file_bytes[offset] != 0x0D:
        name = file_bytes[offset:offset + 11].split(b"\0")[0].decode("latin-1")
        fields.append((name, file_bytes[offset + 16]))
        offset += 32
    if not fields or record_len != 1 + sum(length for _, length in fields):
        raise MargDbfError("This does not look like a DBF file.")

    records = []
    for i in range(n_records):
        raw = file_bytes[header_len + i * record_len:
                         header_len + (i + 1) * record_len]
        if len(raw) < record_len:
            break  # truncated file: keep what was read
        if raw[:1] == b"*":
            continue
        pos, row = 1, {}
        for name, length in fields:
            row[name] = raw[pos:pos + length].decode("cp1252", "replace").rstrip()
            pos += length
        records.append(row)
    return records


def read_ledger_dbf(file_bytes):
    """read_dbf() plus a check that the file is Marg's ledger master."""
    records = read_dbf(file_bytes)
    if not records:
        raise MargDbfError("The DBF file has no records.")
    missing = [c for c in ("PARNAM", "SCODE") if c not in records[0]]
    if missing:
        raise MargDbfError(
            "This DBF is not a Marg ledger export (no "
            + ", ".join(missing) + " column). Columns found: "
            + ", ".join(list(records[0])[:12]) + "…")
    return records


def default_group(code):
    """(Tally group, include?) suggested for a Marg group code."""
    code = code.strip()
    group = ""
    for end in range(len(code), 0, -1):
        if code[:end] in DEFAULT_GROUPS:
            group = DEFAULT_GROUPS[code[:end]]
            break
    include = (bool(group) and group not in OTHER_TOOL_GROUPS
               and code not in DEFAULT_SKIP_CODES)
    return group, include


def split_name(record):
    """(name, city) from PARNAM, which packs the city in after column 30."""
    full = record.get("PARNAM", "")
    city = record.get("CITY", "").strip()
    if city and full[_NAME_WIDTH:].strip() == city:
        return " ".join(full[:_NAME_WIDTH].split()), city
    return " ".join(full.split()), ""


def group_summary(records):
    """One row per Marg group code, with the suggested Tally group."""
    by_code = {}
    for rec in records:
        by_code.setdefault(rec.get("SCODE", "").strip(), []).append(rec)
    rows = []
    for code in sorted(by_code):
        group, include = default_group(code)
        names = [split_name(r)[0] for r in by_code[code]]
        rows.append({
            "Include": include,
            "Marg_Code": code,
            "Ledgers": len(names),
            "Tally_Group": group or None,
            "Examples": ", ".join(names[:3]),
        })
    return pd.DataFrame(rows, columns=[
        "Include", "Marg_Code", "Ledgers", "Tally_Group", "Examples"])


def _clean_gstin(raw):
    """(gstin, problem) — problem is '' when the value is a clean GSTIN."""
    value = re.sub(r"\s+", "", raw).upper()
    if not value or _GSTIN_RE.fullmatch(value):
        return value, ""
    if _GSTIN_RE.fullmatch(value[:15]):
        return value[:15], f"had extra text after the GSTIN ({value[15:]})"
    return value, "is not a valid GSTIN format"


def _join_address(record):
    """Join PARADD* lines; a line that fills the column was cut mid-word."""
    text = ""
    for key in ("PARADD", "PARADD1", "PARADD2"):
        line = record.get(key, "")
        if not line.strip():
            continue
        if text and not text.endswith("\x00"):
            text += " "
        text = text.rstrip("\x00") + line.strip()
        if len(line) >= _ADDR_WIDTH:
            text += "\x00"  # marker: glue the next line on without a space
    return text.rstrip("\x00")


def _pincode(record, address):
    dsm = record.get("DSM", "").strip()
    if re.fullmatch(r"[1-9]\d{5}", dsm):
        return dsm
    found = _PIN_RE.findall(address)
    return found[-1] if found else ""


def _number(text):
    try:
        return float(text.strip() or 0)
    except ValueError:
        return 0.0


def convert_ledgers(records, mapping, keep_city=True, default_state=""):
    """Reshape Marg ledger records into Master — Ledger rows.

    ``mapping`` is {Marg code: (Tally group, include)}. ``default_state`` fills
    State_Name for debtors/creditors that have no GSTIN to derive it from.
    Returns (DataFrame, notes) — notes are per-ledger remarks for the user.
    """
    rows, notes = [], []
    for rec in records:
        group, include = mapping.get(rec.get("SCODE", "").strip(), ("", False))
        if not include or not group:
            continue
        name, city = split_name(rec)
        if keep_city and city:
            name = f"{name} {city}"
        if not name:
            continue

        gstin, problem = _clean_gstin(rec.get("GSTNO", ""))
        if problem:
            notes.append(f"**{name}** — GST number {problem}; "
                         f"Marg has `{rec.get('GSTNO', '').strip()}`.")
        state = GST_STATES.get(gstin[:2], "") if gstin else ""
        if not state and group in PARTY_GROUPS:
            state = default_state
        address = _join_address(rec)
        opening = _number(rec.get("OPNING", ""))

        rows.append({
            "Ledger_Name": name,
            "Alias": "",
            "Group_Name": group,
            "Country": "India",
            "State_Name": state,
            "Pincode": _pincode(rec, address),
            "Registration_Type": "Regular" if gstin else "Unregistered",
            "GST_NO": gstin,
            "Opening_Balance": abs(opening),
            "Dr/Cr": "Cr" if opening < 0 else "Dr",
            "Address": address,
        })
    return pd.DataFrame(rows, columns=LEDGER_COLUMNS), notes


def duplicate_names(df):
    """Ledger names that occur more than once (Tally matches case-insensitively)."""
    key = df["Ledger_Name"].str.upper()
    return sorted(df.loc[key.duplicated(keep=False), "Ledger_Name"].unique())
