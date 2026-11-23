"""
Marg ERP 9+ "Tally Export" XML  ->  this app's Excel templates.

Marg's Tally Export writes two Tally-import files:

* the **masters** file (groups, ledgers, stock items)  -> Master — Ledger,
  Master — Duties & Taxes and Master — Purchase/Sales Ledgers templates;
* the **daybook** file (vouchers) -> Purchase / Sales and
  Payment / Contra / Receipt templates.

The templates are narrower than Tally's voucher model, so every voucher is
checked before it is placed:

* A Sales/Purchase voucher goes into the Purchase / Sales template only when
  that template can reproduce it exactly — taxable value at 5/12/18/28 %,
  CGST+SGST or IGST, round-off, and a whole-rupee invoice total.
* Anything else (journals, credit notes, returns, purchases posted to other
  ledgers…) is written as plain Dr/Cr rows in the Payment / Contra / Receipt
  layout, which keeps every ledger and amount as Marg posted it.

Nothing is dropped silently: each workbook's REVIEW list says what was moved
or left out, and why. Pure Python/pandas, no Streamlit.
"""

import html
import io
import re
from collections import Counter, defaultdict

import pandas as pd
from openpyxl import load_workbook

import tally_core

LEDGER_COLUMNS = [
    "Ledger_Name", "Alias", "Group_Name", "Country", "State_Name", "Pincode",
    "Registration_Type", "GST_NO", "Opening_Balance", "Dr/Cr", "Address",
]
DUTIES_COLUMNS = ["Ledger_Name", "Group_Name", "Rate_of_Tax", "Tax_Type"]
PS_MASTER_COLUMNS = ["Ledger_Name", "Group_Name", "Nature_of_transaction",
                     "RATE_OF_CGST_SGST", "RATE_OF_IGST"]
PS_COLUMNS = [
    "Supplier_Invoice", "Datetime", "Vch_Type", "PartyLedgerName",
    "Dr_LedgerName", "Total_Amount", "GST_5", "GST_12", "GST_18", "GST_28",
    "GST_5_IGST", "GST_12_IGST", "GST_18_IGST", "GST_28_IGST", "Narration",
]
PCR_COLUMNS = [
    "DATE_TIME", "PartyLedgerName", "Dr_LedgerName", "Dr_Amount",
    "Cr_LedgerName", "Cr_Amount", "Narration", "Vch_Type",
    "NARRATION_1", "NARRATION_2",
]
REVIEW_COLUMNS = ["Vch_Type", "Voucher_No", "Date", "Party", "Amount",
                  "Placed_In", "Reason"]
MASTER_REVIEW_COLUMNS = ["Ledger_Name", "Marg_Group", "Reason"]

GST_RATES = (5, 12, 18, 28)
PCR_TYPES = ("Payment", "Receipt", "Contra")

# Standard Tally groups, looked up by their letters/digits only: Marg drops
# "&" and "/" from some group references ("DUTIES  TAXES").
_STANDARD_GROUPS = [
    "Branch / Divisions", "Capital Account", "Reserves & Surplus",
    "Current Assets", "Bank Accounts", "Cash-in-Hand", "Deposits (Asset)",
    "Loans & Advances (Asset)", "Stock-in-Hand", "Sundry Debtors",
    "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors",
    "Fixed Assets", "Investments", "Loans (Liability)", "Bank OD A/c",
    "Bank OCC A/c", "Secured Loans", "Unsecured Loans",
    "Misc. Expenses (ASSET)", "Suspense A/c", "Direct Expenses",
    "Indirect Expenses", "Direct Incomes", "Indirect Incomes",
    "Purchase Accounts", "Sales Accounts",
]
# Marg's own names for groups that Tally calls something else.
_GROUP_ALIASES = {
    "SECURITYDEPOSITASSET": "Deposits (Asset)",
    "LOANADVANCESASSET": "Loans & Advances (Asset)",
    "SUSPENSEACCOUNTTEMPORARYACS": "Suspense A/c",
    "SUSPENSEACCOUNT": "Suspense A/c",
    "PURCHASEACCOUNT": "Purchase Accounts",
    "SALESACCOUNT": "Sales Accounts",
}
_PROFIT_LOSS_KEYS = {"PROFITLOSSAC", "PROFITLOSSACCOUNT"}


class MargXmlError(ValueError):
    """The upload is not the Marg Tally-export file this page expects."""


# --------------------------------------------------------------------------
# Low-level helpers
# --------------------------------------------------------------------------
def decode_xml(raw_bytes):
    """Decode the export: Marg writes plain ASCII/ANSI, Tally writes UTF-16."""
    if raw_bytes[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw_bytes.decode("utf-16", errors="replace")
    try:
        return raw_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw_bytes.decode("cp1252", errors="replace")


def _text(value):
    """Unescape XML entities (&amp;, &#47;, …) and drop control characters."""
    value = html.unescape(value)
    return re.sub(r"[\x00-\x1f]", "", value).strip()


def _tag(block, name):
    m = re.search(r"<%s>([^<]*)</%s>" % (name, name), block)
    return _text(m.group(1)) if m else ""


def _key(name):
    """Letters and digits only, upper-cased — for tolerant name matching."""
    return re.sub(r"[^A-Z0-9]", "", name.upper())


def _number(text):
    try:
        return float(text or 0)
    except ValueError:
        return 0.0


def _rate_in(name):
    """The GST rate in a ledger name ('GST LOCAL SALES 5 %' -> 5.0), or None."""
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", name)
    return float(m.group(1)) if m else None


def _tax_kind(name):
    """'CGST' / 'SGST' / 'IGST' when the ledger name carries one, else ''."""
    m = re.search(r"\b(CGST|SGST|IGST)\b", name.upper())
    return m.group(1) if m else ""


def _fmt_rate(rate):
    return str(int(rate)) if float(rate).is_integer() else str(rate)


# --------------------------------------------------------------------------
# Masters
# --------------------------------------------------------------------------
_STANDARD_BY_KEY = {_key(g): g for g in _STANDARD_GROUPS}
_STANDARD_BY_KEY.update(_GROUP_ALIASES)


def _group_chain(group, parents):
    """The group followed by its ancestors, as normalised keys."""
    chain, seen, key = [], set(), _key(group)
    while key and key not in seen:
        seen.add(key)
        chain.append(key)
        key = parents.get(key, "")
    return chain


def _standard_group(group, parents):
    """Nearest standard Tally group at or above ``group`` ('' if none)."""
    for key in _group_chain(group, parents):
        if key in _STANDARD_BY_KEY:
            return _STANDARD_BY_KEY[key]
    return ""


def parse_masters(xml_text, standard_groups=True):
    """Split Marg's masters file into the three master templates.

    Returns a dict of DataFrames: ``ledger``, ``duties``, ``ps`` and
    ``review`` (ledgers that were left out, with the reason), plus
    ``counts`` with what else the file held.
    """
    ledger_blocks = re.findall(
        r'<LEDGER NAME="([^"]*)"[^>]*>(.*?)</LEDGER>', xml_text, re.S)
    if not ledger_blocks:
        if "<VOUCHER " in xml_text:
            raise MargXmlError(
                "This is the daybook (vouchers) file — upload it on the "
                "Marg Vouchers page. This page needs the masters file.")
        raise MargXmlError("No <LEDGER> records found. Upload the masters "
                           "file written by Marg's Tally Export.")

    parents = {}
    for name, body in re.findall(
            r'<GROUP NAME="([^"]*)"[^>]*>(.*?)</GROUP>', xml_text, re.S):
        parents[_key(_text(name))] = _key(_tag(body, "PARENT"))

    ledgers, duties, ps, review = [], [], [], []
    for raw_name, body in ledger_blocks:
        name = _text(raw_name)
        marg_group = _tag(body, "PARENT")
        chain = _group_chain(marg_group, parents)
        std = _standard_group(marg_group, parents)

        if _PROFIT_LOSS_KEYS & set(chain):
            review.append({"Ledger_Name": name, "Marg_Group": marg_group,
                           "Reason": "Profit & Loss A/c already exists in "
                                     "Tally as a reserved ledger."})
            continue

        if std == "Duties & Taxes":
            kind, rate = _tax_kind(name), _rate_in(name)
            if kind and rate:
                duties.append({
                    "Ledger_Name": name, "Group_Name": "Duties & Taxes",
                    "Rate_of_Tax": rate,
                    "Tax_Type": {"CGST": "Central Tax", "SGST": "State Tax",
                                 "IGST": "Integrated Tax"}[kind]})
            else:
                review.append({
                    "Ledger_Name": name, "Marg_Group": marg_group,
                    "Reason": "Duties & Taxes ledger without a CGST/SGST/IGST "
                              "type and rate in its name — create it in Tally "
                              "by hand if you use it."})
            continue

        if std in ("Purchase Accounts", "Sales Accounts"):
            rate = _rate_in(name)
            exempt = "EXEMPT" in name.upper() or rate == 0
            if rate is None and not exempt:
                review.append({
                    "Ledger_Name": name, "Marg_Group": marg_group,
                    "Reason": "Purchase/Sales ledger with no GST rate in its "
                              "name — the nature of transaction can't be "
                              "worked out; create it in Tally by hand."})
                continue
            side = "Purchase" if std == "Purchase Accounts" else "Sales"
            central = ("CENTRAL" in name.upper()
                       or any("CENTRAL" in k for k in chain))
            nature = (("Interstate " if central else "") + side
                      + (" Exempt" if exempt else " Taxable"))
            ps.append({
                "Ledger_Name": name, "Group_Name": std,
                "Nature_of_transaction": nature,
                "RATE_OF_CGST_SGST": 0 if exempt else rate,
                "RATE_OF_IGST": 0 if exempt else rate})
            continue

        am = re.search(r"<ADDRESS\.LIST.*?</ADDRESS\.LIST>", body, re.S)
        address = ", ".join(
            p for p in (_text(a) for a in re.findall(
                r"<ADDRESS>([^<]*)</ADDRESS>", am.group(0))) if p) if am else ""
        gstin = _tag(body, "PARTYGSTIN") or _tag(body, "GSTIN")
        opening = _number(_tag(body, "OPENINGBALANCE"))
        group = (std if standard_groups else "") or marg_group
        if standard_groups and not std:
            review.append({
                "Ledger_Name": name, "Marg_Group": marg_group,
                "Reason": "Kept, but its Marg group has no standard Tally "
                          "group above it — check Group_Name."})
        ledgers.append({
            "Ledger_Name": name,
            "Alias": "",
            "Group_Name": group,
            "Country": _tag(body, "COUNTRYNAME").title() or "India",
            "State_Name": _tag(body, "LEDSTATENAME") or _tag(body, "STATE"),
            "Pincode": _tag(body, "PINCODE"),
            # Blank would be read as "Regular" by the converter and then fail
            # its GST-number check, so a ledger without a GSTIN is Unregistered.
            "Registration_Type": (_tag(body, "GSTREGISTRATIONTYPE")
                                  or ("Regular" if gstin else "Unregistered")),
            "GST_NO": gstin,
            # Tally XML: a negative opening balance is a debit.
            "Opening_Balance": abs(opening),
            "Dr/Cr": "Cr" if opening > 0 else "Dr",
            "Address": address,
        })

    counts = Counter(re.findall(r"<TALLYMESSAGE[^>]*>\s*<([A-Z]+)", xml_text))
    return {
        "ledger": pd.DataFrame(ledgers, columns=LEDGER_COLUMNS),
        "duties": pd.DataFrame(duties, columns=DUTIES_COLUMNS),
        "ps": pd.DataFrame(ps, columns=PS_MASTER_COLUMNS),
        "review": pd.DataFrame(review, columns=MASTER_REVIEW_COLUMNS),
        "counts": dict(counts),
    }


def duplicate_names(df):
    """Ledger names that occur more than once, ignoring case."""
    key = df["Ledger_Name"].str.upper()
    return sorted(df.loc[key.duplicated(keep=False), "Ledger_Name"].unique())


# --------------------------------------------------------------------------
# Vouchers
# --------------------------------------------------------------------------
_INVENTORY_RE = re.compile(
    r"<(ALLINVENTORYENTRIES|INVENTORYALLOCATIONS)\.LIST>(.*?)</\1\.LIST>", re.S)
_LEDGER_ENTRY_RE = re.compile(
    r"<(ALLLEDGERENTRIES|LEDGERENTRIES)\.LIST>(.*?)</\1\.LIST>", re.S)
_ACC_ALLOC_RE = re.compile(
    r"<ACCOUNTINGALLOCATIONS\.LIST>(.*?)</ACCOUNTINGALLOCATIONS\.LIST>", re.S)


def _is_round_off(name):
    return _key(name) in ("ROUNDOFF", "ROUNDEDOFF", "ROUNDINGOFF")


def _parse_voucher(vch_type, body):
    """One voucher -> dict with its party and every (ledger, amount) posting.

    Amounts keep Tally's sign: negative = debit, positive = credit. Stock
    lines are folded into the ledgers they are allocated to; a stock line
    nested inside a ledger entry is already counted in that entry's amount.
    """
    inventory = _INVENTORY_RE.findall(body)
    rest = _INVENTORY_RE.sub("", body)
    party, entries = "", []
    for _, entry in _LEDGER_ENTRY_RE.findall(rest):
        name = _tag(entry, "LEDGERNAME")
        entries.append((name, _number(_tag(entry, "AMOUNT"))))
        if not party and _tag(entry, "ISPARTYLEDGER") == "Yes":
            party = name
    for _, item in inventory:
        for alloc in _ACC_ALLOC_RE.findall(item):
            entries.append((_tag(alloc, "LEDGERNAME"),
                            _number(_tag(alloc, "AMOUNT"))))
    date = _tag(body, "DATE")
    return {
        "type": vch_type,
        "number": _tag(body, "VOUCHERNUMBER") or _tag(body, "REFERENCE"),
        "date": (f"{date[6:8]}/{date[4:6]}/{date[0:4]}"
                 if re.fullmatch(r"\d{8}", date) else date),
        "narration": _tag(body, "NARRATION"),
        "party": party or _tag(body, "PARTYLEDGERNAME"),
        "entries": entries,
    }


def parse_vouchers(xml_text):
    """Every <VOUCHER> in the daybook file as a list of dicts."""
    blocks = re.findall(
        r'<VOUCHER VCHTYPE="([^"]*)"[^>]*>(.*?)</VOUCHER>', xml_text, re.S)
    if not blocks:
        if "<LEDGER NAME=" in xml_text:
            raise MargXmlError(
                "This is the masters file — upload it on the Marg Masters "
                "page. This page needs the daybook (vouchers) file.")
        raise MargXmlError("No <VOUCHER> records found. Upload the daybook "
                           "file written by Marg's Tally Export.")
    return [_parse_voucher(_text(t), b) for t, b in blocks]


def _fit_purchase_sales(voucher):
    """Can the Purchase / Sales template reproduce this voucher exactly?

    Returns (fit, reason). ``fit`` holds the taxable value per template
    column and the ledger Marg used for each column; ``reason`` explains a
    refusal.
    """
    side = voucher["type"]                      # "Sales" or "Purchase"
    sign = 1 if side == "Sales" else -1         # sales income is a credit
    word = "SALE" if side == "Sales" else "PURCHASE"
    party_total, taxable, taxes, names = 0.0, Counter(), Counter(), {}
    party_seen = False
    for name, amount in voucher["entries"]:
        if name == voucher["party"] and not party_seen:
            party_seen, party_total = True, -amount * sign
            continue
        if _is_round_off(name):
            continue
        rate, kind = _rate_in(name), _tax_kind(name)
        if kind and rate:
            taxes[(kind, rate)] += amount * sign
            names.setdefault(f"{_fmt_rate(rate)}_{kind}", name)
        elif rate in GST_RATES and word in name.upper():
            central = "CENTRAL" in name.upper()
            column = f"GST_{int(rate)}" + ("_IGST" if central else "")
            taxable[column] += amount * sign
            if names.setdefault(column, name) != name:
                return None, (f"uses two ledgers for {column} "
                              f"({names[column]} and {name})")
        else:
            return None, f"posts to '{name}', which the template has no column for"

    if not party_seen or not taxable:
        return None, "has no taxable value on a GST sales/purchase ledger"
    if party_total <= 0 or any(v <= 0 for v in taxable.values()):
        return None, "has amounts on the opposite side to a normal bill"
    if abs(party_total - round(party_total)) > 0.005:
        return None, ("invoice total is not a whole rupee; the template "
                      "rounds totals, which would change the party amount")

    expected = Counter()
    for column, value in taxable.items():
        rate = int(column.split("_")[1])
        if column.endswith("_IGST"):
            expected[("IGST", float(rate))] += value * rate / 100
        else:
            expected[("CGST", rate / 2)] += value * rate / 200
            expected[("SGST", rate / 2)] += value * rate / 200
    for key in set(expected) | set(taxes):
        if abs(expected[key] - taxes[key]) > 0.5:
            return None, (f"{key[0]} {_fmt_rate(key[1])}% in Marg is "
                          f"{taxes[key]:.2f}, but the template would compute "
                          f"{expected[key]:.2f}")
    return {"total": round(party_total), "taxable": dict(taxable),
            "names": names}, ""


def _dr_cr_rows(voucher):
    """Break a voucher into balanced one-debit / one-credit rows.

    Each debit is matched against the credits in order, so every ledger ends
    up with exactly the amount Marg posted even when a voucher has several
    lines on both sides.
    """
    debits = [[n, round(-a, 2)] for n, a in voucher["entries"] if a < 0]
    credits = [[n, round(a, 2)] for n, a in voucher["entries"] if a > 0]
    rows, d, c = [], 0, 0
    while d < len(debits) and c < len(credits):
        amount = round(min(debits[d][1], credits[c][1]), 2)
        if amount > 0:
            rows.append({
                "DATE_TIME": voucher["date"],
                "PartyLedgerName": voucher["party"],
                "Dr_LedgerName": debits[d][0], "Dr_Amount": amount,
                "Cr_LedgerName": credits[c][0], "Cr_Amount": amount,
                "Narration": voucher["narration"],
                "Vch_Type": voucher["type"],
                "NARRATION_1": "", "NARRATION_2": "",
            })
        debits[d][1] = round(debits[d][1] - amount, 2)
        credits[c][1] = round(credits[c][1] - amount, 2)
        if debits[d][1] <= 0:
            d += 1
        if credits[c][1] <= 0:
            c += 1
    return rows


# MASTER_LEDGER_NAME_LINK rows, in the order the template lists them.
LINK_ROWS = (
    [f"GST_{r}" for r in GST_RATES] + [f"GST_{r}_IGST" for r in GST_RATES]
    + [f"{_fmt_rate(r / 2)}_CGST" for r in GST_RATES]
    + [f"{_fmt_rate(r / 2)}_SGST" for r in GST_RATES]
    + [f"{r}_IGST" for r in GST_RATES])


def suggested_links(side):
    """Ledger names that follow Marg's usual naming, one per link row.

    These are guesses, not data: a name here may not exist in the company.
    They are only ever offered on screen and are written to a file solely
    for the rows the user confirms.
    """
    word = "SALES" if side == "Sales" else "PURCHASE"
    links = {}
    for rate in GST_RATES:
        half = _fmt_rate(rate / 2)
        links[f"GST_{rate}"] = f"GST LOCAL {word} {rate} %"
        links[f"GST_{rate}_IGST"] = f"GST CENTRAL {word} {rate} %"
        links[f"{half}_CGST"] = f"CGST {word} TAX {half} %"
        links[f"{half}_SGST"] = f"SGST {word} TAX {half} %"
        links[f"{rate}_IGST"] = f"IGST {word} TAX {rate} %"
    return links


def convert_vouchers(vouchers):
    """Place every voucher into a template layout.

    Returns a dict:
      ``sales`` / ``purchase``  rows for the Purchase / Sales template,
      ``sales_links`` / ``purchase_links``  MASTER_LEDGER_NAME_LINK names
          found in the vouchers (rows no bill uses are absent),
      ``sales_suggestions`` / ``purchase_suggestions``  guessed names for
          those unused rows — for display only, never written by default,
      ``pcr``    Payment / Receipt / Contra rows,
      ``other``  every other voucher as Dr/Cr rows (same layout as ``pcr``),
      ``review`` what was moved, split or left out, and why,
      ``summary`` voucher counts per type and destination.
    """
    ps_rows = {"Sales": [], "Purchase": []}
    link_votes = {"Sales": defaultdict(Counter), "Purchase": defaultdict(Counter)}
    pcr_rows, other_rows, review, summary = [], [], [], Counter()

    def note(voucher, placed, reason):
        total = sum(a for _, a in voucher["entries"] if a > 0)
        review.append({
            "Vch_Type": voucher["type"], "Voucher_No": voucher["number"],
            "Date": voucher["date"], "Party": voucher["party"],
            "Amount": round(total, 2), "Placed_In": placed, "Reason": reason})

    def as_rows(voucher, target, label, reason=None):
        if abs(sum(a for _, a in voucher["entries"])) > 0.01:
            note(voucher, "Left out", "debits and credits do not balance")
            summary[(voucher["type"], "Left out")] += 1
            return
        rows = _dr_cr_rows(voucher)
        if not rows:
            note(voucher, "Left out", "voucher has no amounts")
            summary[(voucher["type"], "Left out")] += 1
            return
        target.extend(rows)
        summary[(voucher["type"], label)] += 1
        if reason:
            note(voucher, label, reason)
        elif len(rows) > 1:
            note(voucher, label,
                 f"has more than one debit or credit line, so it is written "
                 f"as {len(rows)} rows (one voucher each in Tally)")

    for voucher in vouchers:
        kind = voucher["type"]
        if kind in ps_rows:
            fit, reason = _fit_purchase_sales(voucher)
            if fit:
                # GST columns stay empty (None), not "", so they remain numeric.
                row = dict.fromkeys(PS_COLUMNS)
                row.update({
                    "Supplier_Invoice": voucher["number"],
                    "Datetime": voucher["date"], "Vch_Type": kind,
                    "PartyLedgerName": voucher["party"],
                    # The template posts Total_Amount to this ledger.
                    "Dr_LedgerName": voucher["party"],
                    "Total_Amount": fit["total"],
                    "Narration": voucher["narration"]})
                row.update({c: round(v, 2) for c, v in fit["taxable"].items()})
                ps_rows[kind].append((row, fit["names"]))
                for column, name in fit["names"].items():
                    link_votes[kind][column][name] += 1
            else:
                as_rows(voucher, other_rows, "Other vouchers", reason)
        elif kind in PCR_TYPES:
            as_rows(voucher, pcr_rows, "Payment / Contra / Receipt")
        else:
            as_rows(voucher, other_rows, "Other vouchers",
                    f"{kind} vouchers have no template of their own, so they "
                    f"are written as Dr/Cr rows")

    result = {}
    by_number = {(v["type"], v["number"]): v for v in vouchers}
    for kind, key in (("Sales", "sales"), ("Purchase", "purchase")):
        # One link sheet per workbook: a column can name only one ledger, so
        # the commonest ledger wins and any bill using another is moved out.
        # Only ledgers that really occur in the vouchers; rows no bill uses
        # stay blank rather than being filled with a made-up name.
        links = {col: votes.most_common(1)[0][0]
                 for col, votes in link_votes[kind].items()}
        kept = []
        for row, names in ps_rows[kind]:
            clash = next((c for c, n in names.items() if links[c] != n), None)
            if clash:
                as_rows(by_number[(kind, row["Supplier_Invoice"])], other_rows,
                        "Other vouchers",
                        f"uses '{names[clash]}' where the other bills use "
                        f"'{links[clash]}' for {clash}")
            else:
                kept.append(row)
                summary[(kind, "Purchase / Sales")] += 1
        result[key] = pd.DataFrame(kept, columns=PS_COLUMNS)
        result[key + "_links"] = links
        result[key + "_suggestions"] = {
            row: name for row, name in suggested_links(kind).items()
            if row not in links}

    result["pcr"] = pd.DataFrame(pcr_rows, columns=PCR_COLUMNS)
    result["other"] = pd.DataFrame(other_rows, columns=PCR_COLUMNS)
    result["review"] = pd.DataFrame(review, columns=REVIEW_COLUMNS)
    result["summary"] = pd.DataFrame(
        [{"Vch_Type": t, "Placed_In": p, "Vouchers": n}
         for (t, p), n in sorted(summary.items())],
        columns=["Vch_Type", "Placed_In", "Vouchers"])
    return result


# --------------------------------------------------------------------------
# Excel output — always written into this app's own blank templates.
# --------------------------------------------------------------------------
def fill_template(template_bytes, df, links=None, review=None):
    """Write ``df`` into the template's TEMPLATE sheet (bytes in, bytes out).

    ``links`` fills MASTER_LEDGER_NAME_LINK; ``review`` adds a REVIEW sheet.
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb["TEMPLATE"]
    headers = [cell.value for cell in ws[1]]
    columns = [(headers.index(c) + 1, c) for c in df.columns if c in headers]
    for r, record in enumerate(df.to_dict("records"), start=2):
        for col, name in columns:
            value = record[name]
            if value != "" and not pd.isna(value):
                ws.cell(row=r, column=col, value=value)
    if links and "MASTER_LEDGER_NAME_LINK" in wb.sheetnames:
        ws2 = wb["MASTER_LEDGER_NAME_LINK"]
        for r in range(2, ws2.max_row + 1):
            key = str(ws2.cell(row=r, column=1).value)
            if key in links:
                ws2.cell(row=r, column=2, value=links[key])
    if review is not None and not review.empty:
        ws3 = wb.create_sheet("REVIEW")
        ws3.append(list(review.columns))
        for record in review.itertuples(index=False):
            ws3.append(list(record))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def master_workbooks(parsed):
    """{file name: xlsx bytes} for the masters, skipping empty templates."""
    books = {}
    if not parsed["ledger"].empty:
        books["Master_Ledger.xlsx"] = fill_template(
            tally_core.template_master_ledger(), parsed["ledger"],
            review=parsed["review"])
    if not parsed["duties"].empty:
        books["Master_Duties_Taxes.xlsx"] = fill_template(
            tally_core.template_master_duties(), parsed["duties"])
    if not parsed["ps"].empty:
        books["Master_Purchase_Sales_Ledgers.xlsx"] = fill_template(
            tally_core.template_master_ps(), parsed["ps"])
    return books


def voucher_workbooks(converted, confirmed=None):
    """{file name: xlsx bytes} for the vouchers, skipping empty templates.

    ``confirmed`` is {"sales": {link row: name}, "purchase": {...}} — names
    the user has approved for link rows the vouchers do not use. Without it
    those rows are left blank. A confirmed name never replaces one that was
    read from the vouchers.
    """
    review = converted["review"]
    confirmed = confirmed or {}
    books = {}
    for key, name in (("sales", "Sales.xlsx"), ("purchase", "Purchase.xlsx")):
        if not converted[key].empty:
            links = {row: text for row, text in confirmed.get(key, {}).items()
                     if text and row in LINK_ROWS}
            links.update(converted[key + "_links"])
            books[name] = fill_template(
                tally_core.template_purchase(), converted[key], links=links)
    if not converted["pcr"].empty:
        books["Pay_Con_Rec.xlsx"] = fill_template(
            tally_core.template_pay_con_rec(), converted["pcr"],
            review=review[review["Placed_In"] == "Payment / Contra / Receipt"])
    if not converted["other"].empty:
        books["Other_Vouchers.xlsx"] = fill_template(
            tally_core.template_pay_con_rec(), converted["other"],
            review=review[review["Placed_In"] != "Payment / Contra / Receipt"])
    return books


# --------------------------------------------------------------------------
# Tally XML — each workbook run through this app's own converter.
# --------------------------------------------------------------------------
_CONVERTERS = {
    "Master_Ledger.xlsx": tally_core.convert_master_ledger,
    "Master_Duties_Taxes.xlsx": tally_core.convert_master_duties,
    "Master_Purchase_Sales_Ledgers.xlsx": tally_core.convert_master_ps,
    "Sales.xlsx": tally_core.convert_purchase_sales,
    "Purchase.xlsx": tally_core.convert_purchase_sales,
    "Pay_Con_Rec.xlsx": tally_core.convert_pay_con_rec,
    "Other_Vouchers.xlsx": tally_core.convert_pay_con_rec,
}


_ENTRY_OPEN, _ENTRY_CLOSE = "<ALLLEDGERENTRIES.LIST>", "</ALLLEDGERENTRIES.LIST>"
_EMPTY_AMOUNT_RE = re.compile(r"<AMOUNT>\s*[+-]?\s*</AMOUNT>")


def drop_blank_link_entries(xml_text):
    """Remove the ledger lines the converter writes for blank link rows.

    The Purchase / Sales converter adds one line per MASTER_LEDGER_NAME_LINK
    row to every voucher. For a blank row that line names a ledger called
    "nan" and has no amount. Such a ledger does not exist, so the line is
    dropped; nothing with an amount is touched. Returns (xml, lines removed).
    """
    parts, pos, removed = [], 0, 0
    while True:
        start = xml_text.find(_ENTRY_OPEN, pos)
        if start < 0:
            break
        end = xml_text.find(_ENTRY_CLOSE, start)
        if end < 0:
            break
        end += len(_ENTRY_CLOSE)
        block = xml_text[start:end]
        if ("<LEDGERNAME>nan</LEDGERNAME>" in block
                and _EMPTY_AMOUNT_RE.search(block)):
            parts.append(xml_text[pos:start].rstrip(" \t"))
            removed += 1
        else:
            parts.append(xml_text[pos:end])
        pos = end
    parts.append(xml_text[pos:])
    return "".join(parts), removed


def tally_xml_files(books):
    """Convert each workbook to Tally XML with the matching converter.

    Returns ({'Sales.xml': bytes, …}, {'Sales.xml': 'why it failed', …}).
    The XML is what uploading the workbook to its tool would give, except
    that Sales/Purchase vouchers do not carry the empty "nan" ledger lines
    the converter writes for blank link rows (see drop_blank_link_entries).
    """
    xmls, failures = {}, {}
    for name, data in books.items():
        convert = _CONVERTERS.get(name)
        if convert is None:
            continue
        xml_name = name.rsplit(".", 1)[0] + ".xml"
        try:
            xml = tally_core.result_xml_bytes(convert(io.BytesIO(data)))
            if convert is tally_core.convert_purchase_sales:
                xml = drop_blank_link_entries(xml.decode("utf-8"))[0].encode("utf-8")
            xmls[xml_name] = xml
        except Exception as exc:  # noqa: BLE001 - report, keep the others
            failures[xml_name] = f"{type(exc).__name__}: {exc}"
    return xmls, failures


def bundle_zip(books, xmls):
    """One .zip with the workbooks under Excel/ and the Tally files under XML/."""
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in books.items():
            zf.writestr(f"Excel/{name}", data)
        for name, data in xmls.items():
            zf.writestr(f"XML/{name}", data)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Monthly summaries — totals per month, for checking Marg against Tally.
# --------------------------------------------------------------------------
BY_TYPE_COLUMNS = ["Month", "Vch_Type", "Vouchers", "Amount"]
GST_COLUMNS = ["Month", "Vch_Type", "Vouchers", "Value", "CGST", "SGST",
               "IGST", "Round_Off", "Total"]
LEDGER_MONTH_COLUMNS = ["Ledger", "Month", "Debit", "Credit"]
ROUND_OFF_COLUMNS = ["Month", "Vch_Type", "Vouchers", "Before_Round_Off",
                     "Bills_Rounded_Up", "Rounded_Up", "Bills_Rounded_Down",
                     "Rounded_Down", "Net_Round_Off", "Billed_Amount"]
NO_DATE = "No date"


def _month_key(date_text):
    """'17/04/2026' -> ('2026-04', 'Apr 2026'); unreadable -> ('9999-99', …)."""
    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4})", date_text or "")
    if not m or not 1 <= int(m.group(2)) <= 12:
        return "9999-99", NO_DATE
    names = ("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec").split()
    return f"{m.group(3)}-{m.group(2)}", f"{names[int(m.group(2)) - 1]} {m.group(3)}"


def monthly_summaries(vouchers):
    """Month-by-month totals of the daybook.

    Returns a dict of DataFrames, each sorted by calendar month:

    * ``by_type`` — vouchers and amount per month and voucher type. For a
      bill, return, credit/debit note or journal the amount is what was
      posted to the party (the invoice total). For cash and bank vouchers
      (Payment, Receipt, Contra) it is the voucher total.
    * ``gst`` — the same rows split so that
      Value + CGST + SGST + IGST + Round_Off = Total. Tax is whatever Marg
      posted to ledgers with CGST / SGST / IGST in their name; round off is
      signed (negative when it reduces the bill).
    * ``round_off`` — for bills and other party vouchers: the amount before
      rounding, how much Marg rounded up and down (and on how many bills),
      and the amount billed. Before + Rounded_Up - Rounded_Down = Billed.
    * ``ledger`` — debit and credit per ledger per month, as posted by Marg.
    * ``months`` — the month labels in calendar order.
    """
    by_type, gst, ledger, order = Counter(), {}, {}, {}
    for voucher in vouchers:
        sort_key, month = _month_key(voucher["date"])
        order[month] = sort_key
        entries = voucher["entries"]
        party_net = sum(a for n, a in entries if n == voucher["party"])
        # Measure a bill from the party's side: what the party owes (or is
        # owed) is the invoice total, and every other line is a part of it.
        # Summing one side of the voucher instead would count a round-off
        # posted on the party's side twice.
        from_party = voucher["type"] not in PCR_TYPES and abs(party_net) > 0.005
        if from_party:
            sign = 1 if party_net < 0 else -1
            total = abs(party_net)
        else:
            sign = 0
            total = sum(a for _, a in entries if a > 0)
        key = (month, voucher["type"])
        by_type[key + ("n",)] += 1
        by_type[key + ("amount",)] += total
        parts = gst.setdefault(key, Counter())
        parts["Vouchers"] += 1
        parts["Total"] += total
        party_left = from_party
        rounded = 0.0
        for name, amount in entries:
            slot = ledger.setdefault((name, month), [0.0, 0.0])
            slot[0 if amount < 0 else 1] += abs(amount)
            if party_left and name == voucher["party"]:
                continue
            part = amount * sign if from_party else abs(amount)
            kind = _tax_kind(name)
            if kind:
                parts[kind] += part
            elif _is_round_off(name):
                parts["Round_Off"] += part
                rounded += part
        if from_party:
            parts["Party_Vouchers"] += 1
            parts["Party_Total"] += total
            if rounded > 0.004:
                parts["Up_Bills"] += 1
                parts["Up"] += rounded
            elif rounded < -0.004:
                parts["Down_Bills"] += 1
                parts["Down"] += -rounded

    def sort_months(df):
        if df.empty:
            return df
        return (df.assign(_k=df["Month"].map(order))
                  .sort_values(["_k"] + [c for c in ("Vch_Type", "Ledger")
                                         if c in df.columns], kind="stable")
                  .drop(columns="_k").reset_index(drop=True))

    type_rows = [{"Month": m, "Vch_Type": t, "Vouchers": by_type[(m, t, "n")],
                  "Amount": round(by_type[(m, t, "amount")], 2)}
                 for (m, t) in gst]
    gst_rows = []
    for (month, kind), parts in gst.items():
        tax = parts["CGST"] + parts["SGST"] + parts["IGST"]
        gst_rows.append({
            "Month": month, "Vch_Type": kind, "Vouchers": parts["Vouchers"],
            "Value": round(parts["Total"] - tax - parts["Round_Off"], 2),
            "CGST": round(parts["CGST"], 2), "SGST": round(parts["SGST"], 2),
            "IGST": round(parts["IGST"], 2),
            "Round_Off": round(parts["Round_Off"], 2),
            "Total": round(parts["Total"], 2)})
    round_rows = []
    for (month, kind), parts in gst.items():
        if not parts["Party_Vouchers"]:
            continue
        net = parts["Up"] - parts["Down"]
        round_rows.append({
            "Month": month, "Vch_Type": kind,
            "Vouchers": parts["Party_Vouchers"],
            "Before_Round_Off": round(parts["Party_Total"] - net, 2),
            "Bills_Rounded_Up": parts["Up_Bills"],
            "Rounded_Up": round(parts["Up"], 2),
            "Bills_Rounded_Down": parts["Down_Bills"],
            "Rounded_Down": round(parts["Down"], 2),
            "Net_Round_Off": round(net, 2),
            "Billed_Amount": round(parts["Party_Total"], 2)})
    ledger_rows = [{"Ledger": name, "Month": month,
                    "Debit": round(dr, 2), "Credit": round(cr, 2)}
                   for (name, month), (dr, cr) in ledger.items()]
    return {
        "by_type": sort_months(pd.DataFrame(type_rows, columns=BY_TYPE_COLUMNS)),
        "gst": sort_months(pd.DataFrame(gst_rows, columns=GST_COLUMNS)),
        "round_off": sort_months(pd.DataFrame(round_rows,
                                              columns=ROUND_OFF_COLUMNS)),
        "ledger": sort_months(pd.DataFrame(ledger_rows,
                                           columns=LEDGER_MONTH_COLUMNS)),
        "months": sorted(order, key=order.get),
    }


def opening_balances(masters_xml_text):
    """{ledger name: opening balance} from the masters file.

    Positive = debit, negative = credit (Tally's XML stores it the other way
    round, so the sign is flipped here to read naturally).
    """
    balances = {}
    for raw_name, body in re.findall(
            r'<LEDGER NAME="([^"]*)"[^>]*>(.*?)</LEDGER>', masters_xml_text, re.S):
        balances[_text(raw_name)] = -_number(_tag(body, "OPENINGBALANCE"))
    return balances


def ledger_statement(summaries, ledger_name, opening=None):
    """One ledger month by month: debit, credit, net and running balance.

    ``Closing`` is filled only when ``opening`` is given; it is shown as an
    amount with Dr or Cr. Months with no entries are included so the running
    balance reads continuously.
    """
    rows = summaries["ledger"]
    rows = rows[rows["Ledger"] == ledger_name].set_index("Month")
    out, balance = [], opening
    for month in summaries["months"]:
        debit = float(rows["Debit"].get(month, 0.0))
        credit = float(rows["Credit"].get(month, 0.0))
        record = {"Month": month, "Debit": round(debit, 2),
                  "Credit": round(credit, 2), "Net": round(debit - credit, 2)}
        if opening is not None:
            balance = round(balance + debit - credit, 2)
            record["Closing"] = (f"{abs(balance):,.2f} "
                                 f"{'Dr' if balance >= 0 else 'Cr'}")
        out.append(record)
    return pd.DataFrame(out)


def summaries_workbook(summaries, openings=None):
    """All monthly summaries as one .xlsx (bytes), one sheet per summary."""
    month_type = summaries["by_type"].pivot_table(
        index="Month", columns="Vch_Type", values="Amount", aggfunc="sum",
        fill_value=0, sort=False).reindex(summaries["months"]).reset_index()
    ledger = summaries["ledger"].copy()
    ledger["Net"] = (ledger["Debit"] - ledger["Credit"]).round(2)
    ledger = ledger.sort_values(
        ["Ledger"], kind="stable").reset_index(drop=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        month_type.to_excel(writer, sheet_name="Month x Voucher type", index=False)
        summaries["by_type"].to_excel(writer, sheet_name="Vouchers by month", index=False)
        summaries["gst"].to_excel(writer, sheet_name="GST by month", index=False)
        summaries["round_off"].to_excel(writer, sheet_name="Round off by month", index=False)
        ledger.to_excel(writer, sheet_name="Ledgers by month", index=False)
        if openings:
            pd.DataFrame(
                [{"Ledger": n, "Opening": abs(v), "Dr/Cr": "Dr" if v >= 0 else "Cr"}
                 for n, v in sorted(openings.items())]
            ).to_excel(writer, sheet_name="Opening balances", index=False)
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                width = max(len(str(c.value)) if c.value is not None else 0
                            for c in column[:200])
                sheet.column_dimensions[column[0].column_letter].width = min(
                    max(12, width + 2), 48)
    return buf.getvalue()
