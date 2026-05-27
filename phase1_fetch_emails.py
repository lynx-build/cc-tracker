import argparse
import imaplib
import email
import json
import yaml
from pathlib import Path
from datetime import datetime


CONFIG_PATH = Path("config.yaml")
UID_STORE_PATH = Path("data/downloaded_uids.json")
STATEMENTS_DIR = Path("statements")

FETCH_START_DATE = datetime(2026, 1, 1)


def load_config():
    with open(CONFIG_PATH) as f:
        return yaml.safe_load(f)


def load_uid_store():
    if UID_STORE_PATH.exists():
        with open(UID_STORE_PATH) as f:
            return json.load(f)
    return {}


def save_uid_store(store):
    UID_STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(UID_STORE_PATH, "w") as f:
        json.dump(store, f, indent=2)


def get_email_date(msg):
    date_str = msg.get("Date", "")
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S %Z",
        "%d %b %Y %H:%M:%S %z",
    ):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return datetime.now()


def find_pdf_attachment(msg):
    for part in msg.walk():
        if part.get_content_type() == "application/pdf":
            return part.get_payload(decode=True)
        if part.get_content_type() == "application/octet-stream":
            filename = part.get_filename() or ""
            if filename.lower().endswith(".pdf"):
                return part.get_payload(decode=True)
    return None


def imap_creds_for_bank(bank_cfg, default_yahoo):
    """Return (imap_user, imap_password) — bank-level override or default yahoo creds."""
    return (
        bank_cfg.get("imap_user", default_yahoo["imap_user"]),
        bank_cfg.get("imap_password", default_yahoo["imap_password"]),
    )


def fetch_statements(bank_name, bank_cfg, imap, uid_store, backfill):
    seen_uids = set(uid_store.get(bank_name, []))
    sender = bank_cfg["sender_email"]

    imap.select("INBOX")
    status, data = imap.uid("search", None, f'FROM "{sender}"')
    if status != "OK":
        print(f"[WARN] IMAP search failed for {bank_name}")
        return []

    all_uids = data[0].split() if data[0] else []
    new_uids = [u for u in all_uids if u.decode() not in seen_uids]

    if not new_uids:
        print(f"[WARN] No new PDFs found for {bank_name}")
        return []

    # Default: only the most recent new email. Backfill: all of them.
    selected = new_uids if backfill else [new_uids[-1]]

    bank_dir = STATEMENTS_DIR / bank_name
    bank_dir.mkdir(parents=True, exist_ok=True)

    saved = []
    for uid_bytes in selected:
        uid = uid_bytes.decode()
        status, msg_data = imap.uid("fetch", uid, "(RFC822)")
        if status != "OK" or not msg_data or not msg_data[0]:
            print(f"[WARN] Could not fetch UID {uid} for {bank_name}")
            continue

        msg = email.message_from_bytes(msg_data[0][1])
        pdf_bytes = find_pdf_attachment(msg)
        if not pdf_bytes:
            print(f"[WARN] No PDF attachment in UID {uid} for {bank_name}")
            continue

        received_date = get_email_date(msg)
        received_naive = received_date.replace(tzinfo=None) if received_date.tzinfo else received_date
        if received_naive < FETCH_START_DATE:
            date_str = received_date.strftime("%Y-%m-%d")
            print(f"[SKIP] UID {uid} dated {date_str} — before Jan 2026")
            uid_store.setdefault(bank_name, []).append(uid)
            continue

        date_str = received_date.strftime("%Y-%m-%d")
        dest = bank_dir / f"{date_str}_{uid}.pdf"

        dest.write_bytes(pdf_bytes)
        print(f"[OK]   Saved {dest}")

        uid_store.setdefault(bank_name, []).append(uid)
        saved.append(dest)

    return saved


def run(bank_filter=None, backfill=False):
    config = load_config()
    uid_store = load_uid_store()

    default_yahoo = config["yahoo"]
    banks = config["banks"]

    if bank_filter:
        if bank_filter not in banks:
            print(f"[ERROR] Bank '{bank_filter}' not found in config.yaml")
            return {}
        banks = {bank_filter: banks[bank_filter]}

    # Group banks by IMAP account to share connections
    account_groups: dict[tuple, list] = {}
    for bank_name, bank_cfg in banks.items():
        creds = imap_creds_for_bank(bank_cfg, default_yahoo)
        account_groups.setdefault(creds, []).append((bank_name, bank_cfg))

    results = {}
    for (imap_user, imap_password), bank_list in account_groups.items():
        print(f"\nConnecting to Yahoo IMAP as {imap_user}...")
        imap = imaplib.IMAP4_SSL("imap.mail.yahoo.com", 993)
        imap.login(imap_user, imap_password)
        print(f"[OK]   Connected as {imap_user}")

        try:
            for bank_name, bank_cfg in bank_list:
                print(f"\n--- {bank_name} ---")
                saved = fetch_statements(bank_name, bank_cfg, imap, uid_store, backfill)
                results[bank_name] = saved
        finally:
            imap.logout()

    save_uid_store(uid_store)
    print(f"\nUID store updated: {UID_STORE_PATH}")
    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Phase 1: Fetch credit card statement PDFs from Yahoo Mail")
    parser.add_argument("--bank", help="Only fetch for this bank (e.g. BDO)")
    parser.add_argument("--backfill", action="store_true", help="Fetch all historical statements, not just the latest")
    args = parser.parse_args()

    results = run(bank_filter=args.bank, backfill=args.backfill)

    total = sum(len(v) for v in results.values())
    print(f"\nDone. {total} PDF(s) downloaded.")
    for bank, files in results.items():
        print(f"  {bank}: {len(files)} file(s)")
