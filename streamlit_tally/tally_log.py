"""
Tally import log (tally.imp) reader.

Every time data is imported, Tally appends a block to ``tally.imp`` in its
program folder: the file imported, the company, any ``ERROR:`` entries and an
Import Summary with counts. This module turns that text into a list of import
runs, each with its errors and a plain-language hint. No Streamlit.
"""

import re
from datetime import datetime

DEFAULT_LOG_PATH = r"C:\Program Files\Tally\Tally.ERP9\tally.imp"

_RUN_RE = re.compile(
    r"Importing Data from '(?P<file>.*?)' on (?P<date>\d{1,2}-[A-Za-z]{3}-\d{4})"
    r" at (?P<time>\d{1,2}:\d{2}:\d{2})")
_COMPANY_RE = re.compile(r"Importing Data to company: '(?P<company>.*)'")
_COUNT_RE = re.compile(
    r"^(Created|Altered|Deleted|Combined|Ignored|Errors)\s*:\s*(\d+)", re.M)
# "ERROR: <what>" followed by one or more indented message lines.
_ERROR_RE = re.compile(
    r"^ERROR:[ \t]*(?P<subject>[^\r\n]*)\r?\n(?P<detail>(?:[ \t]+[^\r\n]*\r?\n?)*)",
    re.M)
_GSTIN_RE = re.compile(r"\d{2}[A-Z]{5}\d{4}[A-Z][0-9A-Z]Z[0-9A-Z]")
COUNT_KEYS = ("Created", "Altered", "Deleted", "Combined", "Ignored", "Errors")

# (pattern in Tally's message, what it means and what to do).
_HINTS = [
    (r"Group Name not given",
     "The ledger has no group. Fill in Group_Name for it in the Master — "
     "Ledger sheet."),
    (r"Improper change of Group Name",
     "This ledger already exists in Tally under a different group, and Tally "
     "won't move it on import (common for reserved ledgers like Cash). Use "
     "the group it already has, or leave the ledger out."),
    (r"Voucher Type '.*' already exists",
     "The file tries to create a voucher type Tally already has. This is "
     "harmless — the vouchers themselves are imported separately."),
    (r"Ledger '.*' does not exist",
     "A voucher uses a ledger that is not in this Tally company. Import the "
     "masters first, or correct the ledger name (names must match exactly)."),
    (r"is Out of Range",
     "The voucher date is outside the company's financial year or period. "
     "Change the period in Tally (Alt+F2) or the books-beginning date, or "
     "correct the date in the sheet. A date of 1-Jan-1998 means the date in "
     "the sheet could not be read."),
    (r"Voucher totals do not match",
     "Debits and credits of this voucher are not equal, so Tally rejected "
     "it. Check the amounts on that row."),
    (r"already exists",
     "A ledger with this name already exists in the Tally company."),
]


def decode_log(raw_bytes):
    """Tally writes the log as UTF-16; fall back for copies saved otherwise."""
    if raw_bytes[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw_bytes.decode("utf-16", errors="replace")
    if raw_bytes[1:2] == b"\x00" or raw_bytes[:1] == b"\x00":
        return raw_bytes.decode("utf-16-le" if raw_bytes[1:2] == b"\x00"
                                else "utf-16-be", errors="replace")
    try:
        return raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw_bytes.decode("cp1252", errors="replace")


def explain(detail):
    """A plain-language hint for one of Tally's error messages ('' if none)."""
    quoted = re.search(r"Ledger '([^']*)' already exists", detail)
    if quoted and _GSTIN_RE.fullmatch(quoted.group(1).strip().upper()):
        return ("Another ledger in this Tally company already has this GST "
                "number. Tally allows a GSTIN on one ledger only — remove the "
                "duplicate or use the existing ledger.")
    for pattern, hint in _HINTS:
        if re.search(pattern, detail, re.I):
            return hint
    return ""


def _parse_subject(subject):
    """Split 'Ledger: ABC' / 'Voucher: ID:<>, Voucher Type: …' into parts."""
    kind, _, rest = subject.partition(":")
    kind, rest = kind.strip(), rest.strip()
    if kind != "Voucher":
        return kind or "Error", rest, {}
    fields = dict(re.findall(
        r"(Voucher Type|Source Voucher Number|Voucher Number|Date):\s*([^,]*)",
        rest))
    label = " · ".join(p for p in (
        fields.get("Voucher Type", "").strip(),
        f"No. {fields['Source Voucher Number'].strip()}"
        if fields.get("Source Voucher Number", "").strip() else "",
        fields.get("Date", "").strip()) if p)
    return "Voucher", label or rest, fields


def parse_import_log(text):
    """Parse the log into import runs, newest first.

    Each run is a dict: ``file``, ``file_name``, ``when`` (datetime or None),
    ``when_text``, ``company``, ``counts`` and ``errors`` — a list of
    {kind, subject, detail, hint}.
    """
    starts = [m.start() for m in _RUN_RE.finditer(text)]
    runs = []
    for i, start in enumerate(starts):
        block = text[start:starts[i + 1] if i + 1 < len(starts) else len(text)]
        head = _RUN_RE.match(block)
        when_text = f"{head.group('date')} {head.group('time')}"
        try:
            when = datetime.strptime(when_text, "%d-%b-%Y %H:%M:%S")
        except ValueError:
            when = None
        company = _COMPANY_RE.search(block)
        counts = dict.fromkeys(COUNT_KEYS, 0)
        counts.update({k: int(v) for k, v in _COUNT_RE.findall(block)})
        errors = []
        for m in _ERROR_RE.finditer(block):
            detail = " ".join(m.group("detail").split())
            kind, subject, _ = _parse_subject(m.group("subject").strip())
            errors.append({"kind": kind, "subject": subject, "detail": detail,
                           "hint": explain(detail)})
        path = head.group("file")
        runs.append({
            "file": path,
            "file_name": re.split(r"[\\/]", path)[-1],
            "when": when,
            "when_text": when_text,
            "company": company.group("company") if company else "",
            "counts": counts,
            "errors": errors,
            "order": i,
        })
    # Newest first; the log is appended in time order, so position breaks ties
    # and orders any entries whose date could not be read.
    runs.sort(key=lambda r: (r["when"] or datetime.min, r["order"]),
              reverse=True)
    return runs
