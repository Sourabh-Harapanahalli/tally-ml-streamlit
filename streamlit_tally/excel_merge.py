"""Consolidate several Excel workbooks into one.

Sheets that share a name across the uploaded workbooks are merged into a single
sheet. The header block — which may span several rows (title lines, grouped /
merged column headings) — is written once, from the first file, and only the
data rows of the other files are stacked underneath, in upload order.

How the header is found, per sheet:

* Automatic: the header is the run of leading rows that is the same in every
  file (e.g. the 6 title + heading rows of a GSTR-2B "B2B" sheet). A row still
  counts as the same when a small share of its cells is worded differently.
* Manual: the caller passes `header_rows` and that many leading rows are
  treated as the header in every file.
* If the files share no leading rows at all, the first row is taken as a
  single-row header and columns are lined up by header name instead, so files
  with reordered or extra columns still merge.

Sheets that exist in only one file are copied over unchanged.

Pure openpyxl — no Streamlit — so it can be tested on its own.
"""

import io
from collections import Counter
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

SOURCE_COLUMN = "Source_File"
GENERAL = "General"
# Share of cells that must agree for two rows to count as the same header row.
HEADER_MATCH = 0.85


def _blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _norm(text):
    """Key used to match sheet names and headers: trimmed, case-insensitive."""
    return str(text).strip().casefold()


def _row_key(cells):
    """Comparable form of a row: its values, without trailing blanks."""
    values = [c.value.strip() if isinstance(c.value, str) else c.value
              for c in cells]
    while values and _blank(values[-1]):
        values.pop()
    return tuple(None if _blank(v) else v for v in values)


def _load_part(filename, ws):
    """Read one worksheet: its rows (trailing blank rows trimmed) and row keys."""
    rows = [list(r) for r in ws.iter_rows()]
    keys = [_row_key(r) for r in rows]
    while keys and not keys[-1]:
        keys.pop()
    return {"file": filename, "ws": ws, "rows": rows[:len(keys)], "keys": keys}


def _same_header_row(a, b):
    """True when two rows are the same header line, give or take a reworded
    cell (e.g. a heading whose spelling was corrected in later exports)."""
    if a == b:
        return True
    width = max(len(a), len(b))
    a = a + (None,) * (width - len(a))
    b = b + (None,) * (width - len(b))
    filled = [i for i in range(width) if a[i] is not None or b[i] is not None]
    same = sum(1 for i in filled if a[i] == b[i])
    return same >= HEADER_MATCH * len(filled)


def _common_prefix(parts):
    """(n, exact): the number of leading rows shared by every part, and
    whether those rows match exactly or only closely."""
    n, exact = 0, True
    for group in zip(*(p["keys"] for p in parts)):
        if not all(_same_header_row(group[0], k) for k in group[1:]):
            break
        exact = exact and all(k == group[0] for k in group[1:])
        n += 1
    return n, exact


def _copy_style(src, dst, cache):
    """Copy a cell's formatting across workbooks (cached per distinct style)."""
    try:
        if not src.has_style:
            return
        key = (id(src.parent.parent), tuple(src._style))
    except AttributeError:
        return
    cached = cache.get(key)
    if cached is not None:
        dst._style = copy(cached)
        return
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.number_format = src.number_format
    cache[key] = dst._style


def _copy_row(cells, out_ws, out_row, cache):
    for col, src in enumerate(cells, start=1):
        if src.value is None and not getattr(src, "has_style", False):
            continue
        dst = out_ws.cell(row=out_row, column=col, value=src.value)
        _copy_style(src, dst, cache)


def _copy_layout(src_ws, out_ws, last_row):
    """Column widths, plus row heights and merged cells down to `last_row`."""
    for letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            out_ws.column_dimensions[letter].width = dim.width
    for r in range(1, last_row + 1):
        height = src_ws.row_dimensions[r].height
        if height:
            out_ws.row_dimensions[r].height = height
    for rng in src_ws.merged_cells.ranges:
        if rng.max_row <= last_row:
            out_ws.merge_cells(str(rng))


def _stack(out_ws, parts, header_rows, add_source, drop_duplicates, cache):
    """Header block from the first part, then every part's data rows."""
    first = parts[0]
    for r in range(header_rows):
        _copy_row(first["rows"][r], out_ws, r + 1, cache)
    _copy_layout(first["ws"], out_ws, header_rows)

    width = max((len(k) for p in parts for k in p["keys"]), default=0)
    if add_source:
        cell = out_ws.cell(row=max(header_rows, 1), column=width + 1,
                           value=SOURCE_COLUMN)
        cell.font = Font(bold=True)

    next_row = max(header_rows, 1 if add_source else 0) + 1
    seen, removed, sources = set(), 0, []
    for part in parts:
        written = 0
        for cells, key in zip(part["rows"][header_rows:],
                              part["keys"][header_rows:]):
            if not key:
                continue                      # blank row
            if drop_duplicates:
                if key in seen:
                    removed += 1
                    continue
                seen.add(key)
            _copy_row(cells, out_ws, next_row, cache)
            if add_source:
                out_ws.cell(row=next_row, column=width + 1, value=part["file"])
            next_row += 1
            written += 1
        sources.append((part["file"], written))
    return sources, removed


def _copy_whole(out_ws, part, add_source, cache):
    """Copy a sheet exactly as it is (blank rows and merged cells included)."""
    for r, cells in enumerate(part["rows"], start=1):
        _copy_row(cells, out_ws, r, cache)
    _copy_layout(part["ws"], out_ws, len(part["rows"]))
    if add_source:
        width = max((len(k) for k in part["keys"]), default=0)
        for r, key in enumerate(part["keys"], start=1):
            if key:
                out_ws.cell(row=r, column=width + 1, value=part["file"])


# ---- Fallback: single-row header, columns matched by name -----------------

def _named_table(part):
    """(keys, labels, rows) with the first non-blank row as the header."""
    body = [(cells, key) for cells, key in zip(part["rows"], part["keys"]) if key]
    if not body:
        return [], {}, []
    width = max(len(key) for _, key in body)
    used = [i for i in range(width)
            if any(i < len(key) and key[i] is not None for _, key in body)]
    header = body[0][1]
    keys, labels, seen = [], {}, Counter()
    for i in used:
        value = header[i] if i < len(header) else None
        if value is None:
            key, label = ("", i), None        # unnamed: match by position
        else:
            label = str(value)
            seen[_norm(label)] += 1
            key = (_norm(label), seen[_norm(label)])
        keys.append(key)
        labels[key] = label
    rows = []
    for cells, key in body[1:]:
        row = {k: cells[i] for k, i in zip(keys, used)
               if i < len(key) and key[i] is not None}
        if row:
            rows.append((row, key))
    return keys, labels, rows


def _merge_by_name(out_ws, parts, add_source, drop_duplicates, cache):
    tables = [_named_table(p) for p in parts]
    keys, labels = [], {}
    for t_keys, t_labels, _ in tables:
        for key in t_keys:
            if key not in labels:
                keys.append(key)
                labels[key] = t_labels[key]

    bold = Font(bold=True)
    header = [labels[k] for k in keys] + ([SOURCE_COLUMN] if add_source else [])
    for col, label in enumerate(header, start=1):
        out_ws.cell(row=1, column=col, value=label).font = bold

    next_row, seen, removed, sources = 2, set(), 0, []
    for part, (_, _, rows) in zip(parts, tables):
        written = 0
        for row, _ in rows:
            if drop_duplicates:
                fingerprint = tuple(row[k].value if k in row else None
                                    for k in keys)
                if fingerprint in seen:
                    removed += 1
                    continue
                seen.add(fingerprint)
            for col, key in enumerate(keys, start=1):
                if key in row:
                    dst = out_ws.cell(row=next_row, column=col,
                                      value=row[key].value)
                    _copy_style(row[key], dst, cache)
            if add_source:
                out_ws.cell(row=next_row, column=len(keys) + 1,
                            value=part["file"])
            next_row += 1
            written += 1
        sources.append((part["file"], written))

    partial = [labels[k] or f"(column {k[1] + 1})" for k in keys
               if any(t[0] and k not in t[0] for t in tables)]
    return sources, removed, partial


def consolidate_workbooks(files, add_source=False, drop_duplicates=False,
                          keep_identical_once=True, header_rows=None):
    """Merge same-named sheets across workbooks.

    files: list of (filename, bytes) in the order they should be stacked.
    header_rows: None to detect the header block automatically, or the number
        of leading rows that form the header in every file.
    Returns (xlsx_bytes, summary, errors) where summary is one dict per output
    sheet and errors is a list of (filename, message) for unreadable files.
    xlsx_bytes is None when no sheet with data could be read.
    """
    groups, errors = {}, []
    for filename, data in files:
        try:
            wb = load_workbook(io.BytesIO(data), data_only=True)
        except Exception as exc:  # noqa: BLE001
            errors.append((filename, str(exc)))
            continue
        for ws in wb.worksheets:
            part = _load_part(filename, ws)
            if not part["keys"]:
                continue                      # empty sheet
            group = groups.setdefault(_norm(ws.title), {
                "title": ws.title.strip() or ws.title, "parts": []})
            group["parts"].append(part)

    if not groups:
        return None, [], errors

    out = Workbook()
    out.remove(out.active)
    cache, summary = {}, []

    for group in groups.values():
        parts = group["parts"]
        ws = out.create_sheet(title=group["title"][:31])
        notes, removed = [], 0
        identical = len(parts) > 1 and all(
            p["keys"] == parts[0]["keys"] for p in parts[1:])

        if len(parts) == 1 or (identical and keep_identical_once):
            _copy_whole(ws, parts[0], add_source, cache)
            n_header = header_rows
            data = sum(1 for k in parts[0]["keys"][(header_rows or 0):] if k)
            sources = [(parts[0]["file"], data)] + [(p["file"], 0)
                                                    for p in parts[1:]]
            notes.append("Only in one file — copied as is" if len(parts) == 1
                         else f"Identical in all {len(parts)} files — kept once")
        else:
            if header_rows is not None:
                n_header = header_rows
            elif identical:
                n_header = 1                  # nothing to tell header from data
            else:
                n_header, exact = _common_prefix(parts)
                if not exact:
                    notes.append("Header wording differs slightly between "
                                 "files — the first file's header is used")

            if header_rows is None and n_header == 0:
                sources, removed, partial = _merge_by_name(
                    ws, parts, add_source, drop_duplicates, cache)
                n_header = 1
                notes.append("Headers differ between files — columns matched "
                             "by the names in the first row")
                if partial:
                    notes.append("Columns not in every file: "
                                 + ", ".join(partial))
            else:
                sources, removed = _stack(ws, parts, n_header, add_source,
                                          drop_duplicates, cache)

        if removed:
            notes.append(f"{removed} duplicate row(s) removed")
        summary.append({"sheet": ws.title, "files": len(parts),
                        "sources": sources,
                        "rows": sum(n for _, n in sources),
                        "header_rows": n_header, "notes": notes})

    buf = io.BytesIO()
    out.save(buf)
    return buf.getvalue(), summary, errors
