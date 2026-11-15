"""Find the values that are missing between a column of Excel A and a column
of Excel B.

The app lets the user pick one sheet + one column in each workbook; this
module reads the sheets, works out the header row, and reports which values
of each column do not appear in the other.

Pure pandas/openpyxl — no Streamlit — so it can be tested on its own.
"""

import datetime
import io
import re

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

MAX_ROWS_LISTED = 10


def _blank(value):
    if value is None:
        return True
    if isinstance(value, float) and value != value:      # NaN
        return True
    return isinstance(value, str) and not value.strip()


def load_sheets(data):
    """{sheet name: list of row tuples} for every non-empty sheet, exactly as
    laid out in Excel (row 1 first), with trailing blank rows trimmed."""
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheets = {}
    for ws in wb.worksheets:
        try:
            ws.reset_dimensions()             # ignore a stale saved range
        except AttributeError:
            pass
        rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
        while rows and all(_blank(v) for v in rows[-1]):
            rows.pop()
        if rows:
            sheets[ws.title] = rows
    wb.close()
    return sheets


def guess_header_row(rows):
    """1-based Excel row most likely to hold the column headings: the first
    row that is filled across most of the sheet's width (title lines above a
    table fill only a cell or two)."""
    counts = [sum(1 for v in r if not _blank(v)) for r in rows[:50]]
    if not counts or max(counts) == 0:
        return 1
    widest = max(sum(1 for v in r if not _blank(v)) for r in rows[:500])
    first = next((i for i, n in enumerate(counts) if n >= 0.6 * widest),
                 counts.index(max(counts)))
    # Grouped headings: a second (or third) heading row fills the gaps left
    # under a merged group heading, and is itself blank under the columns
    # whose heading spans both rows. The last such row is the header row.
    last, filled = first, {i for i, v in enumerate(rows[first]) if not _blank(v)}
    for nxt in range(first + 1, min(first + 3, len(rows))):
        cells = {i for i, v in enumerate(rows[nxt]) if not _blank(v)}
        gaps_filled = cells - filled
        if (len(gaps_filled) < 2 or not (filled - cells)
                or not all(isinstance(rows[nxt][i], str) for i in cells)):
            break
        last, filled = nxt, filled | cells
    return last + 1


def to_table(rows, header_row):
    """DataFrame of the rows below `header_row` (1-based), with unique text
    column labels. Unnamed columns are labelled by their Excel letter. The
    DataFrame index is the Excel row number of each row."""
    width = max(len(r) for r in rows)
    header = rows[header_row - 1] if 0 < header_row <= len(rows) else ()
    labels, seen = [], set()
    for i in range(width):
        letter = get_column_letter(i + 1)
        value = header[i] if i < len(header) else None
        # A heading merged down from the row(s) above (title lines, which
        # fill a single cell, do not count).
        for above in range(header_row - 2, max(header_row - 4, -1), -1):
            if not _blank(value) or above >= len(rows):
                break
            if sum(1 for v in rows[above] if not _blank(v)) >= 2:
                value = rows[above][i] if i < len(rows[above]) else None
        label = (f"Column {letter}" if _blank(value)
                 else re.sub(r"\s+", " ", str(value)).strip())
        if label in seen:
            label = f"{label} ({letter})"
        seen.add(label)
        labels.append(label)

    body = [tuple(r) + (None,) * (width - len(r)) for r in rows[header_row:]]
    df = pd.DataFrame(body, columns=labels, dtype=object)
    df.index = range(header_row + 1, header_row + 1 + len(body))
    # Drop columns that are empty in both the header and the data.
    keep = [c for i, c in enumerate(labels)
            if not (c == f"Column {get_column_letter(i + 1)}"
                    and all(_blank(v) for v in df[c]))]
    return df[keep]


def pasted_series(text):
    """Pasted text as a column: one value per line, indexed by line number.

    A line copied from several Excel columns arrives tab-separated; only its
    first cell is kept. Returns (Series, number of lines that had extra cells).
    """
    values, extra = [], 0
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        cells = line.split("\t")
        if any(c.strip() for c in cells[1:]):
            extra += 1
        values.append(cells[0])
    series = pd.Series(values, index=range(1, len(values) + 1), dtype=object)
    return series, extra


def normalise(value, ignore_case=True, ignore_spaces=True):
    """Comparable text form of a cell, or None for a blank cell. 123, 123.0
    and "123" compare equal, as do a date cell and the same date as text in
    ISO form."""
    if _blank(value):
        return None
    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, (int, float)):
        number = float(value)
        text = str(int(number)) if number.is_integer() else repr(number)
    elif isinstance(value, datetime.datetime):
        text = (value.date().isoformat() if value.time() == datetime.time(0)
                else value.isoformat(sep=" "))
    elif isinstance(value, datetime.date):
        text = value.isoformat()
    else:
        text = str(value)
        if re.fullmatch(r"-?\d+\.0+", text.strip()):
            text = text.strip().split(".")[0]
    if ignore_spaces:
        text = re.sub(r"\s+", " ", text).strip()
    if ignore_case:
        text = text.casefold()
    return text


def _index(series, ignore_case, ignore_spaces):
    """{normalised value: (value as first written, [Excel rows])}."""
    out = {}
    for excel_row, value in series.items():
        key = normalise(value, ignore_case, ignore_spaces)
        if key is None:
            continue
        if key not in out:
            out[key] = (value, [])
        out[key][1].append(excel_row)
    return out


def _missing_frame(mine, other):
    rows = []
    for key, (value, where) in mine.items():
        if key in other:
            continue
        listed = ", ".join(str(r) for r in where[:MAX_ROWS_LISTED])
        if len(where) > MAX_ROWS_LISTED:
            listed += f" … (+{len(where) - MAX_ROWS_LISTED} more)"
        rows.append({"Value": value, "Occurrences": len(where),
                     "Row(s)": listed})
    return pd.DataFrame(rows, columns=["Value", "Occurrences", "Row(s)"])


def compare(series_a, series_b, ignore_case=True, ignore_spaces=True):
    """Compare two columns (Series indexed by Excel row / pasted line number).

    Returns a dict with the two "missing" tables and the counts behind them.
    Blank cells are ignored; repeated values are counted once.
    """
    a = _index(series_a, ignore_case, ignore_spaces)
    b = _index(series_b, ignore_case, ignore_spaces)
    return {
        "only_a": _missing_frame(a, b),       # in A, missing from B
        "only_b": _missing_frame(b, a),       # in B, missing from A
        "unique_a": len(a), "unique_b": len(b),
        "common": sum(1 for k in a if k in b),
    }


def result_xlsx(result, label_a, label_b):
    """Workbook with a summary and one sheet per direction."""
    wb = Workbook()
    bold = Font(bold=True)
    ws = wb.active
    ws.title = "Summary"
    for row in [
        ("", "Excel A", "Excel B"),
        ("Column compared", label_a, label_b),
        ("Distinct values", result["unique_a"], result["unique_b"]),
        ("Found in both", result["common"], result["common"]),
        ("Missing from the other file",
         len(result["only_a"]), len(result["only_b"])),
    ]:
        ws.append(row)
    for cell in ws[1]:
        cell.font = bold
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].font = bold
    ws.column_dimensions["A"].width = 30
    for letter in ("B", "C"):
        ws.column_dimensions[letter].width = 60
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top",
                                       horizontal="left")

    for title, key, note in (
            ("In A, missing from B", "only_a",
             f"Values of {label_a} that are not in {label_b}"),
            ("In B, missing from A", "only_b",
             f"Values of {label_b} that are not in {label_a}")):
        ws = wb.create_sheet(title=title)
        ws.append([note])
        ws["A1"].font = bold
        df = result[key]
        ws.append(list(df.columns))
        for cell in ws[2]:
            cell.font = bold
        for record in df.itertuples(index=False):
            ws.append([v.item() if hasattr(v, "item") else v for v in record])
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 40
        ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
