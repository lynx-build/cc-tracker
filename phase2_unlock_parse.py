import argparse
import json
import re
import sys
import yaml
import pikepdf
import pdfplumber
import fitz
from pathlib import Path
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CONFIG_PATH    = Path("config.yaml")
STATEMENTS_DIR = Path("statements")
UNLOCKED_DIR   = Path("statements_unlocked")
PARSED_DIR     = Path("parsed")

# ---------- BDO parser (pdfplumber + regex) ----------

TX_LINE_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{2})'
    r'\s+\d{2}/\d{2}/\d{2}'
    r'\s+(.+?)'
    r'\s+(-?[\d,]+\.\d{2})'
    r'\s*$'
)
SKIP_PATTERNS = re.compile(
    r'^(PREVIOUS STATEMENT|CARD NUMBER|SUBTOTAL|TOTAL\b|Page \d|Reference:)',
    re.IGNORECASE
)


def unlock_pdf(src: Path, dest: Path, password: str) -> bool:
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        with pikepdf.open(src, password=password) as pdf:
            pdf.save(dest)
        return True
    except pikepdf.PasswordError:
        print(f"[ERROR] Wrong password for {src}")
        return False
    except Exception as e:
        print(f"[ERROR] Could not unlock {src}: {e}")
        return False


def parse_sale_date_bdo(raw: str) -> str:
    m = re.match(r'^(\d{2})/(\d{2})/(\d{2})$', raw.strip())
    if not m:
        return raw
    month, day, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{2000+yy}-{month:02d}-{day:02d}"


def parse_bdo_pdf(pdf_path: Path, debug: bool = False) -> list:
    transactions = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            for line in text.splitlines():
                line = line.strip()
                if not line or SKIP_PATTERNS.match(line):
                    continue
                m = TX_LINE_RE.match(line)
                if not m:
                    continue
                date_str = parse_sale_date_bdo(m.group(1))
                description = m.group(2).strip()
                amount = float(m.group(3).replace(",", ""))
                transactions.append({
                    "date": date_str,
                    "transaction": description,
                    "amount": amount,
                })
    seen = set()
    unique = []
    for tx in transactions:
        key = (tx["date"], tx["transaction"], tx["amount"])
        if key not in seen:
            seen.add(key)
            unique.append(tx)
    return unique


# ---------- BPI parser (fitz + custom charmap) ----------

_BPI_SKIP_DESCS = frozenset({
    "Previous Balance",
    "Past Due",
    "Ending Balance",
    "Unbilled Installment Amount",
})

BPI_CHARMAP = str.maketrans({
    # Punctuation / special
    0x40: ' ', 0x4B: '.', 0x60: '-', 0x61: '/', 0x6B: ',', 0x7A: ':',
    # Uppercase A-I: 0xC1-0xC9
    0xC1:'A',0xC2:'B',0xC3:'C',0xC4:'D',0xC5:'E',
    0xC6:'F',0xC7:'G',0xC8:'H',0xC9:'I',
    # Uppercase J-R: 0xD1-0xD9
    0xD1:'J',0xD2:'K',0xD3:'L',0xD4:'M',0xD5:'N',
    0xD6:'O',0xD7:'P',0xD8:'Q',0xD9:'R',
    # Uppercase S-Z: 0xE2-0xE9
    0xE2:'S',0xE3:'T',0xE4:'U',0xE5:'V',0xE6:'W',
    0xE7:'X',0xE8:'Y',0xE9:'Z',
    # Lowercase a-i: 0x81-0x89
    0x81:'a',0x82:'b',0x83:'c',0x84:'d',0x85:'e',
    0x86:'f',0x87:'g',0x88:'h',0x89:'i',
    # Lowercase j-r: 0x91-0x99
    0x91:'j',0x92:'k',0x93:'l',0x94:'m',0x95:'n',
    0x96:'o',0x97:'p',0x98:'q',0x99:'r',
    # Lowercase s-z: 0xA2-0xA9
    0xA2:'s',0xA3:'t',0xA4:'u',0xA5:'v',0xA6:'w',
    0xA7:'x',0xA8:'y',0xA9:'z',
    # Digits 0-9: 0xF0-0xF9
    0xF0:'0',0xF1:'1',0xF2:'2',0xF3:'3',0xF4:'4',
    0xF5:'5',0xF6:'6',0xF7:'7',0xF8:'8',0xF9:'9',
})

MONTH_NAMES = {
    'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,
    'July':7,'August':8,'September':9,'October':10,'November':11,'December':12,
}


def bpi_tr(s: str) -> str:
    return s.translate(BPI_CHARMAP)


def parse_bpi_date(text: str, stmt_date: str) -> str | None:
    parts = text.strip().split()
    if len(parts) != 2:
        return None
    month = MONTH_NAMES.get(parts[0])
    if not month:
        return None
    try:
        day = int(parts[1])
    except ValueError:
        return None
    stmt_year  = int(stmt_date[:4])
    stmt_month = int(stmt_date[5:7])
    year = stmt_year - 1 if month > stmt_month + 1 else stmt_year
    return f"{year}-{month:02d}-{day:02d}"


def parse_bpi_amount(text: str) -> float | None:
    clean = text.replace(",", "")
    try:
        return float(clean)
    except ValueError:
        return None


def extract_stmt_date_bpi(doc) -> str | None:
    """Read page 1 and find the STATEMENT DATE value."""
    page = doc[0]
    words = page.get_text("words")
    decoded_words = [(w[1], bpi_tr(w[4])) for w in words]  # (y, decoded)
    found_label = False
    for y, text in decoded_words:
        if "STATEMENT DATE" in text:
            found_label = True
            continue
        if found_label:
            # Try to parse "APRIL 26, 2026" or similar
            m = re.search(
                r'(January|February|March|April|May|June|July|August|'
                r'September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})',
                text, re.IGNORECASE
            )
            if m:
                month = MONTH_NAMES[m.group(1).capitalize()]
                day   = int(m.group(2))
                year  = int(m.group(3))
                return f"{year}-{month:02d}-{day:02d}"
    return None


def parse_bpi_pdf(pdf_path: Path, password: str, stmt_date: str, debug: bool = False) -> list:
    try:
        doc = fitz.open(str(pdf_path))
        doc.authenticate(str(password))
    except Exception as e:
        print(f"[ERROR] Could not open {pdf_path}: {e}")
        return []

    # Always refine stmt_date from PDF content (filename date can be wrong)
    pdf_date = extract_stmt_date_bpi(doc)
    if pdf_date:
        stmt_date = pdf_date

    transactions = []

    for page_num, page in enumerate(doc, 1):
        words = page.get_text("words")

        # Group by y-row: {y_key → {x_int → decoded_text}}
        rows: dict[float, dict[int, str]] = {}
        for w in words:
            x0, y0, word = w[0], w[1], w[4]
            y_key = round(y0, 1)
            decoded = bpi_tr(word)
            x_int = int(round(x0))
            if y_key not in rows:
                rows[y_key] = {}
            rows[y_key][x_int] = decoded

        last_date = None  # carry forward for dateless sub-rows (e.g. Finance Charge)

        for y_key in sorted(rows):
            row = rows[y_key]

            # Transaction-date column: x ≈ 54
            txn_date_raw = next((t for x, t in sorted(row.items()) if 48 <= x <= 70), None)
            if txn_date_raw:
                parsed = parse_bpi_date(txn_date_raw, stmt_date)
                if parsed:
                    last_date = parsed
                date = parsed
            else:
                # Dateless sub-row (Finance Charge appears directly under Payment)
                date = last_date

            if not date:
                continue

            # Description column: x ≈ 220-230
            desc_raw = next((t for x, t in sorted(row.items()) if 200 <= x <= 250), None)
            if not desc_raw:
                continue
            description = " ".join(desc_raw.split())

            # Skip balance/summary rows that are not transactions
            if description in _BPI_SKIP_DESCS:
                continue

            # Amount column: x ≈ 485-525
            amount_raw = next((t for x, t in sorted(row.items()) if 480 <= x <= 530), None)
            if not amount_raw:
                continue
            amount = parse_bpi_amount(amount_raw)
            if amount is None:
                continue

            if debug:
                print(f"  [TX] {date} | {description[:50]} | {amount}")

            transactions.append({
                "statement_date": stmt_date,
                "date": date,
                "transaction": description,
                "amount": amount,
            })

    return transactions


# ---------- Statement-date extraction from filename ----------

def stmt_date_from_filename(pdf_path: Path) -> str:
    stem = pdf_path.stem
    # New format: YYYY-MM-DD_uid
    m = re.match(r'^(\d{4}-\d{2}-\d{2})', stem)
    if m:
        return m.group(1)
    # Old format: YYYY-MM_uid
    m = re.match(r'^(\d{4}-\d{2})', stem)
    if m:
        return m.group(1) + "-01"
    # Named format: "BPI eStatement-April 26, 2026"
    m = re.search(
        r'(January|February|March|April|May|June|July|August|'
        r'September|October|November|December)\s+(\d{1,2}),\s+(\d{4})',
        stem, re.IGNORECASE
    )
    if m:
        month = MONTH_NAMES[m.group(1).capitalize()]
        day   = int(m.group(2))
        year  = int(m.group(3))
        return f"{year}-{month:02d}-{day:02d}"
    return datetime.today().strftime("%Y-%m-%d")


# ---------- Per-PDF dispatcher ----------

def process_pdf(pdf_path: Path, bank: str, password: str,
                dry_run: bool, debug: bool) -> bool:
    is_bpi = bank.startswith("BPI")
    stem = f"{bank}_{pdf_path.stem}"
    json_path  = PARSED_DIR / f"{stem}.json"
    lock_path  = PARSED_DIR / f"{stem}.lock"
    error_path = PARSED_DIR / f"{stem}.error"

    if json_path.exists() or lock_path.exists():
        print(f"[SKIP] {bank} {pdf_path.stem} — already parsed")
        return False
    if dry_run:
        print(f"[DRY]  Would parse: {pdf_path}")
        return False

    stmt_date = stmt_date_from_filename(pdf_path)

    if not is_bpi and stmt_date[:4] < "2026":
        print(f"[SKIP] {bank} {pdf_path.stem} — statement year {stmt_date[:4]} before 2026")
        return False

    if is_bpi:
        unlocked_path = UNLOCKED_DIR / bank / pdf_path.name
        if not unlocked_path.exists():
            if unlock_pdf(pdf_path, unlocked_path, password):
                print(f"[OK]   BPI unlocked -> {unlocked_path}")
        transactions = parse_bpi_pdf(pdf_path, password, stmt_date, debug)
    else:
        unlocked_path = UNLOCKED_DIR / bank / pdf_path.name
        if not unlock_pdf(pdf_path, unlocked_path, password):
            return False
        print(f"[OK]   Unlocked -> {unlocked_path}")
        transactions = parse_bdo_pdf(unlocked_path, debug)
        # Inject statement_date into BDO transactions
        for tx in transactions:
            tx["statement_date"] = stmt_date

    PARSED_DIR.mkdir(parents=True, exist_ok=True)

    if not transactions:
        error_path.write_text(
            f"0 transactions found.\nFile: {pdf_path}\n", encoding="utf-8"
        )
        print(f"[ERROR] 0 transactions for {bank} {pdf_path.stem} — wrote {error_path}")
        return True

    json_path.write_text(
        json.dumps(transactions, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    lock_path.write_text(
        f"Parsed {len(transactions)} transactions\n", encoding="utf-8"
    )
    print(f"[OK]   Parsed {len(transactions)} transactions -> {json_path}")
    return True


# ---------- run() ----------

def load_config():
    with open(CONFIG_PATH) as f:
        return yaml.safe_load(f)


def run(bank_filter=None, dry_run=False, debug=False):
    config = load_config()
    banks_cfg = config["banks"]
    results = {}

    for bank_dir in sorted(STATEMENTS_DIR.iterdir()):
        if not bank_dir.is_dir():
            continue
        bank = bank_dir.name
        if bank_filter and bank != bank_filter:
            continue
        if bank not in banks_cfg:
            print(f"[WARN] {bank} folder found but not in config.yaml — skipping")
            continue

        password = str(banks_cfg[bank]["pdf_password"])
        results[bank] = []

        for pdf_path in sorted(bank_dir.glob("*.pdf")):
            attempted = process_pdf(pdf_path, bank, password, dry_run, debug)
            if attempted:
                results[bank].append(PARSED_DIR / f"{bank}_{pdf_path.stem}.json")

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Phase 2: Unlock PDFs and parse transactions"
    )
    parser.add_argument("--bank", help="Only process this bank (e.g. BDO)")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--debug-pdf", action="store_true",
                        help="Print raw text and parse decisions")
    args = parser.parse_args()

    results = run(bank_filter=args.bank, dry_run=args.dry_run, debug=args.debug_pdf)
    total = sum(len(v) for v in results.values())
    print(f"\nDone. {total} statement(s) parsed.")
    for bank, files in results.items():
        print(f"  {bank}: {len(files)} file(s)")
