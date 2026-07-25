#!/usr/bin/env python3
"""
Seller Dashboard — iCloud IMAP sync
Reads order confirmation emails from iCloud Mail and writes to Google Sheet.
Credentials loaded from environment variables (set as GitHub Actions secrets).
"""

import imaplib, email, email.header, json, os, sys, time, hashlib, argparse, re
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import anthropic, gspread
from google.oauth2.service_account import Credentials

CONFIG = {
    "IMAP_HOST":        "imap.mail.me.com",
    "IMAP_PORT":        993,
    "ICLOUD_EMAIL":     os.environ.get("ICLOUD_EMAIL", ""),
    "ICLOUD_PASS":      os.environ.get("ICLOUD_PASS", ""),
    "IMAP_FOLDER":      "2026",
    "CLAUDE_API_KEY":   os.environ.get("CLAUDE_API_KEY", ""),
    "SHEET_ID":         os.environ.get("SHEET_ID", ""),
    "CREDENTIALS_FILE": "credentials.json",
    "ENTRIES_SHEET":    "Entries",
    "PROCESSED_SHEET":  "Processed",
    "APPS_SCRIPT_URL":  os.environ.get("APPS_SCRIPT_URL", ""),
}

ENTRY_COLS = [
    "id","item","qty","price","type","date","category","site",
    "purchasedOn","orderNumber","postage","matchedItem","status",
    "email","cardUsed","notes","platform","releaseDate","venue",
    "eventDate","row","seats","ticketType","paidBackInitial","paidStatus"
]

def get_sheets_client():
    # Support credentials from file or environment variable
    creds_json = os.environ.get("GOOGLE_CREDENTIALS", "")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    if creds_json:
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            f.write(creds_json)
            tmp = f.name
        creds = Credentials.from_service_account_file(tmp, scopes=scopes)
        os.unlink(tmp)
    else:
        creds = Credentials.from_service_account_file(CONFIG["CREDENTIALS_FILE"], scopes=scopes)
    return gspread.authorize(creds)

def get_sheet(client, tab):
    return client.open_by_key(CONFIG["SHEET_ID"]).worksheet(tab)

def get_processed_ids(client):
    try:
        records = get_sheet(client, CONFIG["PROCESSED_SHEET"]).get_all_records()
        return {str(r["messageId"]) for r in records if r.get("messageId")}
    except Exception as e:
        print(f"  Warning: {e}")
        return set()

def mark_processed(client, msg_id, subject, count):
    get_sheet(client, CONFIG["PROCESSED_SHEET"]).append_row(
        [str(msg_id), datetime.now(timezone.utc).isoformat(), subject[:200], count]
    )

def write_entries(client, entries):
    if not entries: return

    # Write to Entries sheet tab
    rows = []
    for e in entries:
        row = []
        for col in ENTRY_COLS:
            val = e.get(col, "")
            if isinstance(val, bool): val = "TRUE" if val else "FALSE"
            elif val is None: val = ""
            row.append(str(val))
        rows.append(row)
    get_sheet(client, CONFIG["ENTRIES_SHEET"]).append_rows(rows)

    # Merge into KV store via Apps Script API
    try:
        import urllib.request, urllib.parse
        get_url = CONFIG["APPS_SCRIPT_URL"] + "?action=get&key=" + urllib.parse.quote("inventory:entries")
        req = urllib.request.urlopen(get_url, timeout=30)
        data = json.loads(req.read().decode())
        existing = json.loads(data.get("value") or "[]")

        existing_ids = {e["id"] for e in existing if e.get("id")}
        existing_orders = {str(e.get("orderNumber","")).lower() for e in existing if e.get("orderNumber")}
        to_add = [e for e in entries if e.get("id") not in existing_ids
                  and str(e.get("orderNumber","")).lower() not in existing_orders]

        if not to_add:
            print("    KV store: no new entries (all duplicates)")
            return

        merged = existing + to_add
        post_data = json.dumps({"action": "set", "key": "inventory:entries", "value": json.dumps(merged)}).encode()
        post_req = urllib.request.Request(CONFIG["APPS_SCRIPT_URL"], data=post_data,
                                          headers={"Content-Type": "text/plain"}, method="POST")
        urllib.request.urlopen(post_req, timeout=30)
        print(f"    KV store updated: added {len(to_add)} entries ({len(merged)} total)")
    except Exception as e:
        print(f"    KV store sync failed (non-critical): {e}")

def connect_imap():
    print(f"  Connecting to {CONFIG['IMAP_HOST']}...")
    mail = imaplib.IMAP4_SSL(CONFIG["IMAP_HOST"], CONFIG["IMAP_PORT"])
    mail.login(CONFIG["ICLOUD_EMAIL"], CONFIG["ICLOUD_PASS"])
    print(f"  Logged in as {CONFIG['ICLOUD_EMAIL']}")
    return mail

def decode_header_value(value):
    if not value: return ""
    parts = email.header.decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(str(part))
    return "".join(decoded)

def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                try:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="replace")
                    break
                except: pass
        if not body:
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        charset = part.get_content_charset() or "utf-8"
                        html = part.get_payload(decode=True).decode(charset, errors="replace")
                        body = re.sub(r"<[^>]+>", " ", html)
                        body = re.sub(r"\s+", " ", body).strip()
                        break
                    except: pass
    else:
        try:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="replace")
        except:
            body = str(msg.get_payload())
    return body

def fetch_emails(mail, processed_ids):
    folder = CONFIG["IMAP_FOLDER"]
    opened = False
    for attempt in [folder, f'"{folder}"', f"INBOX.{folder}", f"INBOX/{folder}"]:
        result, _ = mail.select(attempt)
        if result == "OK":
            print(f"  Opened folder: {attempt}")
            opened = True
            break
    if not opened:
        _, folders = mail.list()
        available = [f.decode() for f in (folders or [])]
        print(f"  ERROR: Could not open folder '{folder}'")
        print(f"  Available: {available[:15]}")
        return []

    result, data = mail.search(None, "ALL")
    if result != "OK" or not data[0]:
        print(f"  No emails in '{folder}'")
        return []

    msg_nums = data[0].split()
    print(f"  Found {len(msg_nums)} email(s)")

    emails = []
    for num in msg_nums:
        raw = None
        for fetch_type in ["(RFC822)", "RFC822", "(BODY[])"]:
            result, data = mail.fetch(num, fetch_type)
            if result != "OK": continue
            for part in data:
                if isinstance(part, tuple):
                    for item in part:
                        if isinstance(item, bytes) and len(item) > 50:
                            raw = item; break
                elif isinstance(part, bytes) and len(part) > 50:
                    raw = part
                if raw: break
            if raw: break
        if not raw: continue

        msg = email.message_from_bytes(raw)
        subject = decode_header_value(msg.get("Subject", ""))
        date_h  = msg.get("Date", "")
        msg_id  = msg.get("Message-ID", "")
        if not msg_id:
            msg_id = hashlib.md5((subject + date_h + str(len(raw))).encode()).hexdigest()
        msg_id = "icloud:" + msg_id.strip()

        if msg_id in processed_ids: continue

        from_    = decode_header_value(msg.get("From", ""))
        try:    date = parsedate_to_datetime(date_h).strftime("%Y-%m-%d")
        except: date = datetime.now().strftime("%Y-%m-%d")

        emails.append({"msg_id": msg_id, "subject": subject, "from": from_, "date": date, "body": get_email_body(msg)})
    return emails

def looks_like_order_email(subject, body):
    keywords = ["order","confirmation","confirmed","purchase","receipt","invoice",
                "shipped","dispatched","booking","ticket","payment received",
                "thank you for your order","order number","order #","ref:","reference"]
    text = (subject + " " + body).lower()
    return any(k in text for k in keywords)

def parse_email_with_claude(client_anthropic, subject, from_, date, body):
    truncated = body[:3000] + "\n[truncated]" if len(body) > 3000 else body
    system = ("You are a JSON extraction API. Output ONLY raw JSON arrays. "
              "No explanation, no markdown, no backticks. If no data found, return []")
    prompt = "\n".join([
        "Extract every distinct purchased item from this order confirmation email.",
        "Return a JSON array where each object has:",
        '  "item" - product name', '  "qty" - quantity integer',
        '  "price" - unit price in GBP', '  "date" - YYYY-MM-DD',
        '  "site" - retailer name', '  "orderNumber" - order reference',
        '  "email" - recipient email',
        "Use null for missing fields. Return [] if not an order.",
        f"From: {from_}", f"Date: {date}", f"Subject: {subject}", "", "Body:", truncated
    ])
    response = client_anthropic.messages.create(
        model="claude-sonnet-4-6", max_tokens=1000, system=system,
        messages=[{"role": "user", "content": prompt}]
    )
    text = (response.content[0].text if response.content else "").replace("```json","").replace("```","").strip()
    try:
        parsed = json.loads(text)
        if isinstance(parsed, list): return parsed
    except: pass
    for match in re.findall(r"\[.*?\]", text, re.DOTALL):
        try:
            parsed = json.loads(match)
            if isinstance(parsed, list): return parsed
        except: pass
    return []

def build_entry(p, subject):
    return {
        "id":            hashlib.md5((str(p.get("item",""))+str(p.get("date",""))+str(p.get("price",""))).encode()).hexdigest()[:16],
        "item":          p.get("item") or "Unknown item",
        "qty":           int(p.get("qty") or 1),
        "price":         float(p.get("price") or 0),
        "type":          "purchase",
        "date":          p.get("date") or datetime.now().strftime("%Y-%m-%d"),
        "category":      "", "site": p.get("site") or "",
        "purchasedOn":   p.get("date") or datetime.now().strftime("%Y-%m-%d"),
        "orderNumber":   p.get("orderNumber") or "", "postage": 0,
        "matchedItem":   "", "status": "Pending",
        "email":         p.get("email") or "", "cardUsed": "",
        "notes":         f"Auto-imported from iCloud: {subject[:80]}",
        "platform":      "", "releaseDate": "", "venue": "", "eventDate": "",
        "row":           "", "seats": "", "ticketType": "",
        "paidBackInitial": False, "paidStatus": False,
    }

def sync_once():
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting iCloud sync...")
    missing = [k for k in ["ICLOUD_EMAIL","ICLOUD_PASS","CLAUDE_API_KEY","SHEET_ID","APPS_SCRIPT_URL"] if not CONFIG[k]]
    if missing:
        print(f"ERROR: Missing config: {', '.join(missing)}")
        sys.exit(1)

    sheets  = get_sheets_client()
    claude  = anthropic.Anthropic(api_key=CONFIG["CLAUDE_API_KEY"])
    done    = get_processed_ids(sheets)
    print(f"  Already processed: {len(done)} emails")

    mail   = connect_imap()
    emails = fetch_emails(mail, done)
    mail.logout()
    print(f"  New emails: {len(emails)}")

    total = 0
    for em in emails:
        subj = em["subject"]
        print(f"\n  Processing: {subj[:70]}")
        if not looks_like_order_email(subj, em["body"]):
            print("    Skipped")
            mark_processed(sheets, em["msg_id"], subj, 0)
            continue
        parsed = parse_email_with_claude(claude, subj, em["from"], em["date"], em["body"])
        entries = [build_entry(p, subj) for p in parsed if p.get("item") and p["item"] != "Unknown item"] if parsed else []
        if entries:
            write_entries(sheets, entries)
            total += len(entries)
            print(f"    Imported {len(entries)} item(s)")
        else:
            print("    No items found")
        mark_processed(sheets, em["msg_id"], subj, len(entries))

    print(f"\nDone. Imported {total} entries from {len(emails)} emails.")
    return total

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--watch", action="store_true")
    parser.add_argument("--interval", type=int, default=30)
    args = parser.parse_args()
    if args.watch:
        while True:
            try: sync_once()
            except Exception as e: print(f"Error: {e}")
            time.sleep(args.interval * 60)
    else:
        sync_once()

if __name__ == "__main__":
    main()
