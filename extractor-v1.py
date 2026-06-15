# extractor.py
import re
from io import BytesIO
from itertools import zip_longest
import pdfplumber

# ── IGNORE PATTERNS (headers / footers) ──────────────────────────────────────
IGNORE_PATTERNS = [
    r"statement of account",
    r"account statement",
    r"for the period",
    r"period from",
    r"statement period",
    r"page\s+\d+\s+of\s+\d+",
    r"printed on",
    r"print date",
    r"account number",
    r"account no",
    r"customer id",
    r"ifsc",
    r"micr",
    r"branch",
    r"available balance",
    r"ledger balance",
    r"dear customer",
    r"computer generated",
    r"thank you",
    r"end of statement",
    r"transaction summary",
    r"opening balance",
    r"^\s*s\.?no\.?\s*$",
    r"^\s*sr\.?\s*no\.?\s*$",
]

# ── HEAD RULES ────────────────────────────────────────────────────────────────
HEAD_RULES = {
    "CASH":       ["ATM WDL", "CASH", "CASH WDL", "CSH", "SELF", "ATMWDL", "CASHWDL"],
    "SALARY":     ["SALARY", "PAYROLL"],
    "WITHDRAWAL": ["ATM ISSUER REV", "ATMISSUERREV", "UPI", "UPI REV", "UPIREV", "POS"],
}

HEADER_ALIASES = {
    "date":        ["date", "txn date", "transaction date", "value date", "tran date", "trans date"],
    "particulars": ["particulars", "description", "narration", "transaction particulars",
                    "details", "remarks", "transaction details", "narr"],
    "debit":       ["debit", "withdrawal", "dr", "withdrawal amt", "withdrawal amount",
                    "debit amount", "debits", "debit(dr)"],
    "credit":      ["credit", "deposit", "cr", "deposit amt", "deposit amount",
                    "credit amount", "credits", "credit(cr)"],
    "balance":     ["balance", "running balance", "closing balance", "bal", "avl bal",
                    "available bal", "net balance"],
}

DATE_RE   = re.compile(r'\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b')
AMOUNT_RE = re.compile(r'[-+]?\d{1,3}(?:[,\s]\d{3})*(?:\.\d+)?|\d+(?:\.\d+)?')


# ── HELPERS ───────────────────────────────────────────────────────────────────
def normalize(cell):
    return str(cell).strip().lower() if cell else ""


def map_header(h):
    h = normalize(h)
    for std, aliases in HEADER_ALIASES.items():
        for a in aliases:
            if a == h or h.startswith(a) or a in h:
                return std
    return None


def parse_amount(s):
    if s is None:
        return None
    s = str(s).strip().replace('\xa0', ' ').replace('INR', '').replace('Rs.', '').replace('Rs', '')
    s = re.sub(r'[^\d\-,.\s]', '', s).replace('', '').replace(',', '')
    if s in ('', '-'):
        return None
    try:
        return float(s)
    except ValueError:
        m = AMOUNT_RE.search(str(s))
        if m:
            try:
                return float(m.group(0).replace(',', '').replace(' ', ''))
            except ValueError:
                return None
        return None


def is_ignore_line(text):
    t = str(text or "")
    return any(re.search(p, t, re.IGNORECASE) for p in IGNORE_PATTERNS)


# ── HEAD CLASSIFICATION ───────────────────────────────────────────────────────
def classify_head(particulars):
    p = str(particulars or "").upper()

    if any(kw in p for kw in ["BAJAJ FINANCE LIMITE", "BAJAJ FINANCE LTD", "BAJAJFIN"]):
        return "BAJAJ FINANCE LTD"
    if any(kw in p for kw in ["CGST", "CHARGES", "CHGS", "CHRG", "SGST", "GST"]):
        return "CHARGES"
    if any(kw in p for kw in ["PETROL", "PETROLEUM"]):
        return "CONVEYANCE"
    if "DIVIDEND" in p:
        return "DIVIDEND"
    if any(kw in p for kw in ["ICICI SECURITIES LTD", "ICICISEC.UPI", "ICICISECURITIES"]):
        return "ICICI DIRECT"
    if any(kw in p for kw in ["IDFC FIRST BANK", "IDFCFBLIMITED"]):
        return "IDFC FIRST BANK LTD"
    if "BAJAJ ALLIANZ GEN INS COM" in p:
        return "INSURANCE"
    if any(kw in p for kw in ["INT PD", "INT CR", "INT DR", "INTPD", "INTCR", "INTDR", "INTEREST"]):
        return "INTEREST"
    if any(kw in p for kw in ["LIC OF INDIA", "LIFE INSURANCE CORPORATIO", "LIFE INSURANCE CORPORATION OF INDIA"]):
        return "LIC"
    if any(kw in p for kw in ["TAX REFUND", "TAXREFUND"]):
        return "TAX REFUND"

    for head, kws in HEAD_RULES.items():
        for kw in kws:
            if kw in p:
                return head

    return "OTHER"


# ── TABLE HEADER DETECTION ────────────────────────────────────────────────────
def find_header_row(table):
    """Return index of the row most likely to be the column header row."""
    best_idx, best_score = 0, -1
    for i, row in enumerate(table[:5]):
        score = 0
        for cell in row:
            if not cell:
                continue
            c = normalize(cell)
            for aliases in HEADER_ALIASES.values():
                for a in aliases:
                    if a in c:
                        score += 3
            if re.search(r'[a-zA-Z]', c):
                score += 1
        if score > best_score:
            best_idx, best_score = i, score
    return best_idx


# ── TABLE → TRANSACTIONS ──────────────────────────────────────────────────────
def table_to_transactions(table, meta, page_no=None):
    txns = []
    if not table or len(table) < 2:
        return txns

    header_idx  = find_header_row(table)
    raw_headers = table[header_idx]

    std_headers = [
        map_header(normalize(h)) or normalize(h) or f"col{i}"
        for i, h in enumerate(raw_headers)
    ]

    for row in table[header_idx + 1:]:
        # Normalise row length
        row = list(row or [])
        row = (row + [""] * len(raw_headers))[:len(raw_headers)]
        row_cells = [str(c or "").strip() for c in row]

        if not any(row_cells):
            continue

        joined = " ".join(row_cells)
        if is_ignore_line(joined):
            continue

        row_dict = {k: v for k, v in zip_longest(std_headers, row_cells, fillvalue="")}

        date        = row_dict.get("date", "").strip() or None
        particulars = re.sub(r'\s+', ' ', row_dict.get("particulars", "")).strip()
        debit_raw   = row_dict.get("debit", "")
        credit_raw  = row_dict.get("credit", "")
        balance_raw = row_dict.get("balance", "")

        debit_amt   = parse_amount(debit_raw)
        credit_amt  = parse_amount(credit_raw)
        balance_val = parse_amount(balance_raw)

        # Skip rows missing date/particulars or both amounts
        if not date or not particulars:
            continue
        if debit_amt is None and credit_amt is None:
            continue
        if is_ignore_line(particulars):
            continue

        # Some banks use single "Amount" column + Dr/Cr flag
        if debit_amt is None and credit_amt is None:
            amount_raw = row_dict.get("amount", "") or row_dict.get("amt", "")
            flag       = joined.upper()
            amount_val = parse_amount(amount_raw)
            if amount_val is not None:
                if "CR" in flag or "CREDIT" in flag:
                    credit_amt = amount_val
                else:
                    debit_amt  = amount_val

        txns.append({
            "Date":        date,
            "Particulars": particulars,
            "Debit":       debit_amt,
            "Credit":      credit_amt,
            "Head":        classify_head(particulars),
            "Balance":     balance_val,
            "Page":        page_no,
        })

    return txns


# ── TEXT FALLBACK ─────────────────────────────────────────────────────────────
def text_fallback_extract(page_text, meta, page_no=None):
    txns = []
    lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]

    for ln in lines:
        if is_ignore_line(ln):
            continue
        dm = DATE_RE.search(ln)
        if not dm:
            continue

        nums = [parse_amount(x) for x in AMOUNT_RE.findall(ln)]
        nums = [n for n in nums if n is not None]
        if not nums:
            continue

        date        = dm.group(0)
        debit_amt   = None
        credit_amt  = None
        balance_val = None

        # Heuristic: last number is usually balance
        if len(nums) == 1:
            debit_amt   = nums[0]
        elif len(nums) == 2:
            debit_amt   = nums[0]
            balance_val = nums[1]
        elif len(nums) >= 3:
            # Could be debit + credit + balance; one of debit/credit usually 0 or absent
            debit_amt   = nums[0] if nums[0] else None
            credit_amt  = nums[1] if nums[1] else None
            balance_val = nums[-1]

        txns.append({
            "Date":        date,
            "Particulars": ln,
            "Debit":       debit_amt,
            "Credit":      credit_amt,
            "Head":        classify_head(ln),
            "Balance":     balance_val,
            "Page":        page_no,
        })

    return txns


# ── MAIN API ──────────────────────────────────────────────────────────────────
def process_file(file_bytes, filename):
    meta         = {"filename": filename, "_logs": []}
    transactions = []

    try:
        pdf = pdfplumber.open(BytesIO(file_bytes))
    except Exception as e:
        meta["_logs"].append(f"PDF open error: {e}")
        return meta, transactions

    with pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            try:
                tables    = page.extract_tables() or []
                page_txns = []

                for table in tables:
                    page_txns.extend(table_to_transactions(table, meta, page_no=idx))

                # Fallback to raw text if table extraction gave nothing
                if not page_txns:
                    text = page.extract_text() or ""
                    if text.strip():
                        page_txns.extend(text_fallback_extract(text, meta, page_no=idx))

                transactions.extend(page_txns)

            except Exception as e:
                meta["_logs"].append(f"Page {idx} error: {e}")
                continue

    # Deduplicate
    seen, result = set(), []
    for r in transactions:
        key = (r["Date"], r["Particulars"], r["Debit"], r["Credit"], r["Page"])
        if key not in seen:
            seen.add(key)
            result.append(r)

    return meta, result
