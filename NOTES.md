# CC Tracker — Dev Notes

## Completed (May 2026)

### Phase 0 — Scaffold
- Created full directory structure: `statements/`, `statements_unlocked/`, `parsed/`, `data/`, `docs/`, `logs/`
- `.gitignore` covers `.env`, `config.yaml`, `statements/`, `statements_unlocked/`, `parsed/`, `data/downloaded_uids.json`, `__pycache__/`, `venv/`
- `config.yaml` holds Yahoo IMAP credentials + bank sender emails + PDF passwords
- Git initialized, remote set to `https://github.com/lynx-build/cc-tracker`
- Git identity: name = `lynx`, email = `lincolnvillaganas@gmail.com`

### Phase 1 — Yahoo IMAP Fetcher (`phase1_fetch_emails.py`)
- Connects to `imap.mail.yahoo.com:993` using Yahoo App Password
- Searches inbox by `sender_email` per bank in `config.yaml`
- Downloads PDF attachments to `statements/<BANK>/<YYYY-MM-DD>_<uid>.pdf`
- Date derived from email received date (full date, not just YYYY-MM)
- UID tracking in `data/downloaded_uids.json` prevents re-downloads
- Default: fetches 1 most recent new email per bank. `--backfill` fetches all.
- **Per-bank IMAP override**: if a bank config has `imap_user`/`imap_password`, those are used instead of the global Yahoo creds. Banks sharing the same account reuse one IMAP connection.
- **Tested: BDO (3 PDFs) and BPI_Lincoln (1 PDF) fetched from live inbox.**

### Phase 2 — PDF Unlock + Local Parser (`phase2_unlock_parse.py`)
- **BDO path**: unlocks with `pikepdf`, parses text with `pdfplumber`
  - Format: `MM/DD/YY MM/DD/YY DESCRIPTION AMOUNT` per line, matched by `TX_LINE_RE`
  - Handles negative amounts (payments, credits)
- **BPI path**: opens directly with `PyMuPDF (fitz)` — no pikepdf step
  - BPI PDFs use custom font encoding (fonts: C0FNT000–C0FNT018) that produces garbled text with standard extractors
  - Uses a 60-entry `BPI_CHARMAP` translation table mapping encoded bytes to readable characters
  - Encoding: uppercase A–I → 0xC1–0xC9, J–R → 0xD1–0xD9, S–Z → 0xE2–0xE9 (7-byte and 8-byte gaps mirror the Latin-1 accented ranges); lowercase follows same pattern offset by 0x40; digits 0–9 → 0xF0–0xF9; space=0x40, period=0x4B, comma=0x6B, minus=0x60, slash=0x61, colon=0x7A
  - Coordinate-based extraction: transaction date x≈54, post date x≈144, description x≈220, amount x≈487; rows without both a date and an amount are discarded
  - Month names decoded with `MONTH_NAMES` dict; year inferred from statement date with rollover handling
- **`statement_date` added to all JSON output** (new field alongside `date`, `transaction`, `amount`)
  - For phase1-fetched files: extracted from `YYYY-MM-DD` in filename
  - For BPI: also extracted from the "STATEMENT DATE" field on PDF page 1 (overrides filename if filename only has YYYY-MM-01)
  - Old-format filenames (`YYYY-MM_uid`): fall back to `YYYY-MM-01`
  - Manually placed BPI files (`BPI eStatement-April 26, 2026.pdf`): parsed from filename text
- `.lock` file written on success — prevents re-parsing even if `.json` is deleted
- `.error` file written if 0 transactions found
- **Tested: 13 BPI_Lincoln transactions parsed from April 2026 statement. All merchant names, amounts, and dates correct.**

### Phase 3 — Excel Writer (`phase3_excel_writer.py`)
- Creates `data/cc_transactions.xlsx` if missing
- **New 5-column schema**: Statement Date | Date | Transaction | Amount | Hash
- Hash column is **visible** (not hidden)
- After each write, sheet is **sorted descending** by Statement Date then Date (newest statement first, newest transaction within statement first)
- Deduplication via `MD5(bank + date + transaction + amount)` — unchanged so existing hashes remain valid
- `--dedup-fix` flag removes duplicate rows from all sheets
- Bank name detection uses directory matching (longest-first) to handle underscore keys like `BPI_Lincoln` without breaking on the `_` separator
- Back-fills `statement_date` for old-format JSON (no statement_date field) from the JSON filename
- Per-bank results accumulate correctly across multiple JSON files
- **Tested: 43 BDO + 13 BPI_Lincoln rows, all sorted correctly, new schema verified.**

### Phase 4 — Dashboard Generator (`phase4_dashboard_generator.py`)
- Reads all sheets from Excel (`data/cc_transactions.xlsx`)
- **Full dark-mode redesign**: tokens `--bg:#0f172a`, `--card:#1e293b`, accent `#38bdf8`, fonts Inter + Sora via Google Fonts CDN
- **6 sections**:
  1. Header — total spend across all banks, last updated timestamp
  2. Bank Cards — per-bank total + transaction count; bank display splits on `_` (BPI/Lincoln, BPI/Xherine)
  3. Monthly Spend — stacked bar chart (Chart.js v4 CDN) per bank, calendar months on x-axis
  4. Key Merchants — spend totals for Grab, Lazada, Apple (keyword-matched)
  5. Month Comparison — previous vs current month per bank with delta
  6. Transaction Table — all transactions, paginated (50/page), Statement Date column, negative amounts green / positive amounts red
- Writes to **both** `dashboard/dashboard.html` (primary) and `docs/index.html` (GitHub Pages)
- Auto git-commits and pushes both files after every run
- Handles missing Excel gracefully — shows ₱0.00 / 0 transactions until data exists
- **Live at: https://lynx-build.github.io/cc-tracker/**

### Phase 5 — Master Pipeline (`run_pipeline.py`)
- Chains Phases 1 → 4 in sequence with per-phase error catching
- Flags: `--bank BDO`, `--dry-run`, `--backfill`, `--dedup-fix`
- Writes structured log line to `logs/run_log.txt` after every run
- `run_pipeline.bat` is the Windows Task Scheduler entry point

---

## Config Structure (`config.yaml`)

```yaml
yahoo:
  imap_user: <yahoo address>      # default account for all banks
  imap_password: <app password>

banks:
  BDO:
    sender_email: bdoesoa-noreply@bdo.com.ph
    password_type: static
    pdf_password: <password>

  BPI_Lincoln:
    sender_email: bpi_cards_estatement@bpi.com.ph
    password_type: static
    pdf_password: <password>
    # No imap_user → uses global yahoo account

  #BPI_Xherine:                   # uncomment when ready
    #imap_user: <different yahoo>  # own account → separate IMAP connection
    #imap_password: <app password>
    #sender_email: bpi_cards_estatement@bpi.com.ph
    #password_type: static
    #pdf_password: <password>
```

Banks with no `imap_user` override share the global Yahoo connection. BPI_Xherine has its own account — when uncommented, phase1 will open a second IMAP connection automatically with no further code changes.

---

## Key Decisions

| Decision | Chosen | Reason |
|----------|--------|--------|
| PDF parser (BDO) | `pdfplumber` (local) | Privacy — statements must not leave the machine |
| PDF parser (BPI) | `PyMuPDF` + custom charmap | pikepdf cannot open BPI PDFs (non-standard xref); pdfplumber gets garbled text due to custom font encoding |
| BPI charmap | Manual 60-entry translation table | No new dependencies, offline, fast, fully accurate; encoding is systematic (3 ranges each for upper/lower + digits) |
| BPI coordinate extraction | x-position column detection | BPI layouts are position-stable; filters out summary rows (no txn date at x≈54) without needing keyword lists |
| PDF filename format | `YYYY-MM-DD_uid.pdf` | Full date allows exact statement_date extraction without reading inside the file |
| statement_date field | Separate from transaction date | Enables sorting by when a statement was issued vs when a purchase was made; required for correct descending sort |
| Excel schema | 5 columns, Hash visible | Statement Date column added; hash made visible for easier debugging |
| Sort order | Descending (newest first) | Most useful for day-to-day review; last row in sheet = oldest transaction |
| Dashboard output | Write to both `dashboard/` and `docs/` | `docs/` is GitHub Pages source; `dashboard/` is the canonical local copy |
| Git push | End of Phase 4 | One phase owns publish responsibility |
| Bank name separator | `_` (underscore) | Used in config, Excel tabs, JSON filenames, directory names; display splits on `_` for card labels |
| BPI_Xherine | Configured but commented out | Feasibility confirmed (App Passwords work regardless of login state); activating requires only uncommenting config |

---

## Known Issues

1. **`PYTHONUTF8=1` may be needed in Task Scheduler** for peso sign (`₱`) output.
   `run_pipeline.bat` should have `set PYTHONUTF8=1` at the top. The scripts set `sys.stdout.reconfigure(encoding="utf-8")` internally, but the batch file controls the scheduler process encoding.

2. **BPI charmap verified only against BPI Gold Rewards Card.**
   If BPI_Xherine uses a different card type that renders differently, run `--debug-pdf` and inspect the raw bytes. The charmap structure is systematic and should hold across BPI card types, but has not been confirmed.

3. **Old BDO JSON files (pre–statement_date) get `YYYY-MM-01` as Statement Date.**
   The 3 existing BDO JSON files in `parsed/` were created before the statement_date field existed. Phase 3 back-fills from the filename (`BDO_2026-05_...` → `2026-05-01`). To get exact statement dates, delete the `.json` and `.lock` files for those BDO statements and re-run phase 2.

4. **Stale file in project root:** `.gitkeep` is still present but not committed. Safe to delete manually.

---

## Next Steps

1. **Parse new BPI_Lincoln May statement** — `statements/BPI_Lincoln/2026-05-02_271364.pdf` was fetched during testing. Run `python run_pipeline.py --bank BPI_Lincoln` (no dry-run) to parse and push it.
2. **Check dashboard** — open `https://lynx-build.github.io/cc-tracker/` and verify BPI_Lincoln data appears correctly alongside BDO.
3. **Activate BPI_Xherine** — uncomment the BPI_Xherine block in `config.yaml` when ready; no code changes needed.
4. **Set up Windows Task Scheduler** — use `run_pipeline.bat`, trigger monthly, "Run whether user is logged on or not"; add `set PYTHONUTF8=1` to the bat file.
5. **Backfill historical BPI statements** — once BPI_Lincoln has more statements emailed, run `python run_pipeline.py --bank BPI_Lincoln --backfill` to fetch and parse all history.
