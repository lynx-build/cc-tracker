import argparse
import traceback
from datetime import datetime
from pathlib import Path

import phase1_fetch_emails
import phase2_unlock_parse
import phase3_excel_writer
import phase4_dashboard_generator

LOG_PATH = Path("logs/run_log.txt")


def write_log(entry: dict):
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    line = (
        f"{entry['timestamp']} | "
        f"banks={entry['banks']} | "
        f"pdfs={entry['pdfs']} | "
        f"parsed={entry['parsed']} | "
        f"added={entry['added']} | "
        f"skipped={entry['skipped']} | "
        f"errors={entry['errors'] or 'none'}"
    )
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(f"\n[LOG]  {line}")


def run(bank_filter=None, dry_run=False, backfill=False, dedup_fix=False):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    errors = []

    pdfs_found = 0
    stmts_parsed = 0
    rows_added = 0
    rows_skipped = 0
    banks_processed = bank_filter or "all"

    print("=" * 55)
    print(f"  CC Tracker Pipeline — {timestamp}")
    print(f"  bank={banks_processed}  dry_run={dry_run}  backfill={backfill}")
    print("=" * 55)

    # Phase 1 — Fetch emails
    print("\n[ Phase 1 ] Fetching statement PDFs from Yahoo Mail...")
    try:
        p1_results = phase1_fetch_emails.run(bank_filter=bank_filter, backfill=backfill)
        pdfs_found = sum(len(v) for v in p1_results.values())
        print(f"  PDFs downloaded: {pdfs_found}")
    except Exception:
        msg = f"Phase1: {traceback.format_exc().splitlines()[-1]}"
        errors.append(msg)
        print(f"[ERROR] {msg}")
        p1_results = {}

    # Phase 2 — Unlock + parse
    print("\n[ Phase 2 ] Unlocking and parsing PDFs locally...")
    try:
        p2_results = phase2_unlock_parse.run(bank_filter=bank_filter, dry_run=dry_run)
        stmts_parsed = sum(len(v) for v in p2_results.values())
        print(f"  Statements parsed: {stmts_parsed}")
    except Exception:
        msg = f"Phase2: {traceback.format_exc().splitlines()[-1]}"
        errors.append(msg)
        print(f"[ERROR] {msg}")
        p2_results = {}

    # Phase 3 — Excel writer
    print("\n[ Phase 3 ] Writing transactions to Excel...")
    try:
        p3_results = phase3_excel_writer.run(bank_filter=bank_filter, fix_dupes=dedup_fix)
        rows_added = sum(v["appended"] for v in p3_results.values())
        rows_skipped = sum(v["skipped"] for v in p3_results.values())
        print(f"  Rows added: {rows_added}  Skipped: {rows_skipped}")
    except Exception:
        msg = f"Phase3: {traceback.format_exc().splitlines()[-1]}"
        errors.append(msg)
        print(f"[ERROR] {msg}")

    # Phase 4 — Dashboard
    print("\n[ Phase 4 ] Generating dashboard and pushing to GitHub Pages...")
    try:
        phase4_dashboard_generator.run()
    except Exception:
        msg = f"Phase4: {traceback.format_exc().splitlines()[-1]}"
        errors.append(msg)
        print(f"[ERROR] {msg}")

    # Summary
    print("\n" + "=" * 55)
    print("  Run Summary")
    print("=" * 55)
    print(f"  Banks processed : {banks_processed}")
    print(f"  PDFs downloaded : {pdfs_found}")
    print(f"  Statements parsed: {stmts_parsed}")
    print(f"  Rows added      : {rows_added}")
    print(f"  Rows skipped    : {rows_skipped}")
    print(f"  Errors          : {len(errors)}")
    for e in errors:
        print(f"    - {e}")
    print("=" * 55)

    write_log({
        "timestamp": timestamp,
        "banks": banks_processed,
        "pdfs": pdfs_found,
        "parsed": stmts_parsed,
        "added": rows_added,
        "skipped": rows_skipped,
        "errors": "; ".join(errors) if errors else None,
    })


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="CC Tracker — full pipeline runner")
    parser.add_argument("--bank", help="Only process this bank (e.g. BDO)")
    parser.add_argument("--dry-run", action="store_true", help="Skip all API calls")
    parser.add_argument("--backfill", action="store_true", help="Fetch all historical statements")
    parser.add_argument("--dedup-fix", action="store_true", help="Remove duplicate rows from Excel")
    args = parser.parse_args()

    run(
        bank_filter=args.bank,
        dry_run=args.dry_run,
        backfill=args.backfill,
        dedup_fix=args.dedup_fix,
    )
