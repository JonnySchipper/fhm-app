import imaplib
import email
import os
import re
import html as html_lib
from pathlib import Path
from datetime import datetime, timedelta
from email.header import decode_header, make_header

# =========================
# Settings / output files
# =========================

IMAP_HOST = "imap.gmail.com"

EMAIL_SNIPPETS_FILE = "etsy_sales_emails.txt"  # append extracted Item/Personalization lines only
PROCESSED_FILE = "processed_orders.txt"        # one processed key per line

CONFIG_PATH = (Path(__file__).resolve().parent.parent / "gmailconfig.txt")

DAYS_BACK = 14

# Etsy subject example:
# You made a sale on Etsy - Ship by Jan 9 - [$18.32, Order #3941425315]
ORDER_RE = re.compile(r"Order\s*#(\d+)", re.IGNORECASE)
AMOUNT_RE = re.compile(r"\[\s*\$?([0-9]+(?:\.[0-9]{2})?)\s*,\s*Order\s*#\d+\s*\]", re.IGNORECASE)
SHIPBY_RE = re.compile(r"Ship by\s+([A-Za-z]{3,9}\s+\d{1,2})", re.IGNORECASE)

# Extract Transaction ID and all content until pricing section:
TRANSACTION_ID_RE = re.compile(r"^\s*Transaction\s*ID\s*:\s*(.+)\s*$", re.IGNORECASE)
# Stop capturing when we hit pricing/totals section
STOP_CAPTURE_RE = re.compile(r"^\s*(Item total|Subtotal|Order Total|Discount|Shipping|Sales Tax|Applied discounts|Quantity)\s*:", re.IGNORECASE)
# These are the lines we want to keep
KEEP_LINE_RE = re.compile(r"^\s*(Item|Personalization|Character)\s*:\s*(.+)\s*$", re.IGNORECASE)
# Also match numbered character lines
CHARACTER_LINE_RE = re.compile(r"^\s*\d+\)\s*(.+)$", re.IGNORECASE)

# Shipping snippet (HTML spans)
SPAN_NAME_RE = re.compile(r"<span[^>]*class=['\"]name['\"][^>]*>(.*?)</span>", re.IGNORECASE | re.DOTALL)
SPAN_CITY_RE = re.compile(r"<span[^>]*class=['\"]city['\"][^>]*>(.*?)</span>", re.IGNORECASE | re.DOTALL)
SPAN_STATE_RE = re.compile(r"<span[^>]*class=['\"]state['\"][^>]*>(.*?)</span>", re.IGNORECASE | re.DOTALL)


# =========================
# Credentials
# =========================

def load_gmail_credentials():
    env_user = os.getenv("GMAIL_USER")
    env_pass = os.getenv("GMAIL_APP_PASSWORD")
    if env_user and env_pass:
        return env_user.strip(), env_pass.strip()

    if not CONFIG_PATH.exists():
        raise ValueError(
            "Missing credentials.\n"
            "Set environment variables GMAIL_USER and GMAIL_APP_PASSWORD, OR create:\n"
            f"  {CONFIG_PATH}\n"
            "with:\n"
            "  GMAIL_USER=...\n"
            "  GMAIL_APP_PASSWORD=..."
        )

    text = CONFIG_PATH.read_text(encoding="utf-8", errors="ignore")
    kv = {}
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            k, v = line.split("=", 1)
            kv[k.strip()] = v.strip().strip('"').strip("'")

    user = kv.get("GMAIL_USER")
    app_password = kv.get("GMAIL_APP_PASSWORD")

    if not user or not app_password:
        raise ValueError(
            f"Couldn't parse credentials from {CONFIG_PATH}.\n"
            "Expected:\n"
            "  GMAIL_USER=your.email@gmail.com\n"
            "  GMAIL_APP_PASSWORD=your-16-char-app-password"
        )

    return user, app_password


# =========================
# Helpers
# =========================

def decode_mime_header(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        if isinstance(value, bytes):
            return value.decode(errors="ignore")
        return str(value)


def _decode_part_payload(part) -> str:
    payload = part.get_payload(decode=True)
    if not payload:
        return ""
    charset = part.get_content_charset() or "utf-8"
    return payload.decode(charset, errors="ignore")


def extract_plain_and_html(msg):
    """Return (text_plain, text_html). Collect ALL parts to ensure nothing is missed."""
    text_plain_parts = []
    text_html_parts = []

    if msg.is_multipart():
        for part in msg.walk():
            disp = part.get_content_disposition()
            if disp == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain":
                # Collect ALL text/plain parts
                content = _decode_part_payload(part)
                if content:
                    text_plain_parts.append(content)
            elif ctype == "text/html":
                # Collect ALL text/html parts
                content = _decode_part_payload(part)
                if content:
                    text_html_parts.append(content)
    else:
        ctype = msg.get_content_type()
        content = _decode_part_payload(msg)
        if ctype == "text/html":
            text_html_parts.append(content)
        else:
            text_plain_parts.append(content)

    # Combine all parts with newlines to ensure we don't miss anything
    text_plain = "\n\n".join(text_plain_parts) if text_plain_parts else ""
    text_html = "\n\n".join(text_html_parts) if text_html_parts else ""

    return text_plain, text_html


def html_to_text_minimal(html_str: str) -> str:
    """Very light HTML -> text for parsing Item/Personalization lines."""
    if not html_str:
        return ""
    s = html_str
    # Normalize line breaks
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</p\s*>", "\n", s)
    s = re.sub(r"(?i)</div\s*>", "\n", s)
    # Strip tags
    s = re.sub(r"<[^>]+>", "", s)
    # Unescape entities
    s = html_lib.unescape(s)
    return s


def parse_subject_details(subject: str):
    order_no = None
    amount = None
    ship_by = None

    m = ORDER_RE.search(subject or "")
    if m:
        order_no = m.group(1)

    m = AMOUNT_RE.search(subject or "")
    if m:
        amount = m.group(1)

    m = SHIPBY_RE.search(subject or "")
    if m:
        ship_by = m.group(1)

    return order_no, amount, ship_by


def load_processed_set():
    p = Path(PROCESSED_FILE)
    if not p.exists():
        return set()
    return set(
        line.strip()
        for line in p.read_text(encoding="utf-8", errors="ignore").splitlines()
        if line.strip()
    )

def get_existing_orders_in_snippets():
    """Get set of order numbers that already exist in the snippets file"""
    p = Path(EMAIL_SNIPPETS_FILE)
    if not p.exists():
        return set()
    
    orders = set()
    content = p.read_text(encoding="utf-8", errors="ignore")
    
    # Find all "Order: XXXXXXX" lines
    for line in content.splitlines():
        if line.startswith("Order: "):
            order_no = line.replace("Order: ", "").strip()
            if order_no:
                orders.add(order_no)
    
    return orders

def deduplicate_snippets_file():
    """Remove duplicate orders from the snippets file (keeps first occurrence)"""
    p = Path(EMAIL_SNIPPETS_FILE)
    if not p.exists():
        print(f"No file found at {EMAIL_SNIPPETS_FILE}")
        return
    
    content = p.read_text(encoding="utf-8", errors="ignore")
    
    # Split by separator
    sections = content.split("-" * 60)
    
    seen_orders = set()
    unique_sections = []
    duplicates_removed = 0
    
    for section in sections:
        section = section.strip()
        if not section:
            continue
        
        # Find the order number in this section
        order_no = None
        for line in section.splitlines():
            if line.startswith("Order: "):
                order_no = line.replace("Order: ", "").strip()
                break
        
        if order_no:
            if order_no in seen_orders:
                duplicates_removed += 1
                continue  # Skip this duplicate
            seen_orders.add(order_no)
        
        unique_sections.append(section)
    
    # Rebuild the file
    with open(EMAIL_SNIPPETS_FILE, "w", encoding="utf-8") as f:
        for section in unique_sections:
            f.write("-" * 60 + "\n\n")
            f.write(section + "\n\n")
    
    print(f"✓ Deduplicated {EMAIL_SNIPPETS_FILE}")
    print(f"  Kept: {len(unique_sections)} unique orders")
    print(f"  Removed: {duplicates_removed} duplicates")


def append_processed_key(key: str):
    with open(PROCESSED_FILE, "a", encoding="utf-8") as f:
        f.write(key + "\n")


def extract_item_and_personalization_lines(body_text: str) -> str:
    """
    Returns lines grouped by Transaction ID.
    
    Captures everything from "Transaction ID:" until the pricing section starts
    (Item total, Subtotal, Order Total, etc.)
    
    This ensures we get ALL order details regardless of format variations.
    """
    if not body_text:
        return ""

    # Use a dict to group items by transaction ID
    # Key: transaction_id, Value: list of lines
    transactions = {}
    current_transaction_id = None
    capturing = False  # Are we currently capturing lines for a transaction?
    items_without_transaction = []
    
    # Normalize line endings and split - this ensures we get ALL lines
    normalized = body_text.replace('\r\n', '\n').replace('\r', '\n')
    
    for line in normalized.splitlines():
        # Strip leading/trailing whitespace for matching
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        # Check for Transaction ID - START capturing
        trans_match = TRANSACTION_ID_RE.match(line_stripped)
        if trans_match:
            transaction_id = trans_match.group(1).strip()
            current_transaction_id = transaction_id
            capturing = True
            # Initialize this transaction if not seen before
            if transaction_id not in transactions:
                transactions[transaction_id] = []
            continue
        
        # Check if we should STOP capturing (hit pricing section)
        if capturing and STOP_CAPTURE_RE.match(line_stripped):
            capturing = False
            current_transaction_id = None
            continue
        
        # If we're capturing, check what kind of line this is
        if capturing and current_transaction_id:
            # Check for Item/Personalization/Character lines
            item_match = KEEP_LINE_RE.match(line_stripped)
            if item_match:
                label = item_match.group(1).capitalize()
                value = item_match.group(2).strip()
                transactions[current_transaction_id].append(f"{label}: {value}")
                continue
            
            # Check for numbered lines (e.g., "1) Character: ..." or just "1) Sleeping Beauty/ Maebry")
            char_match = CHARACTER_LINE_RE.match(line_stripped)
            if char_match:
                value = char_match.group(1).strip()
                transactions[current_transaction_id].append(f"Personalization: {value}")
                continue
    
    # Build output
    result = []
    
    # Add transaction groups
    for trans_id, items in transactions.items():
        if items:  # Only add if there are items
            result.append(f"\nTransaction ID: {trans_id}")
            for item_line in items:
                result.append(f"  {item_line}")
    
    # Add items without transaction IDs (if any exist)
    for line in items_without_transaction:
        result.append(line)
    
    return "\n".join(result).strip()


def _clean_span_value(v: str) -> str:
    if v is None:
        return ""
    v = html_lib.unescape(v)
    v = re.sub(r"\s+", " ", v).strip()
    return v


def extract_name_city_state_from_html(html_body: str):
    """
    Extracts from HTML snippet like:
      <span class='name'>Miriam ...</span> ... <span class='city'>MELROSE</span>, <span class='state'>FL</span>
    Returns (name, city, state) or ("", "", "") if not found.
    """
    if not html_body:
        return "", "", ""

    name_m = SPAN_NAME_RE.search(html_body)
    city_m = SPAN_CITY_RE.search(html_body)
    state_m = SPAN_STATE_RE.search(html_body)

    name = _clean_span_value(name_m.group(1)) if name_m else ""
    city = _clean_span_value(city_m.group(1)) if city_m else ""
    state = _clean_span_value(state_m.group(1)) if state_m else ""

    return name, city, state


# =========================
# Main
# =========================

def process_recent_etsy_sales_stop_on_processed():
    # First, clean up any existing duplicates in the snippets file
    if Path(EMAIL_SNIPPETS_FILE).exists():
        print("Checking for duplicates in existing snippets file...")
        deduplicate_snippets_file()
        print()
    
    gmail_user, gmail_pass = load_gmail_credentials()
    processed = load_processed_set()
    existing_orders = get_existing_orders_in_snippets()  # Check what's already in the output file

    mail = imaplib.IMAP4_SSL(IMAP_HOST)
    mail.login(gmail_user, gmail_pass)
    mail.select("inbox")

    cutoff = (datetime.utcnow() - timedelta(days=DAYS_BACK)).strftime("%d-%b-%Y")

    status, messages = mail.search(
        None,
        "SUBJECT", '"You made a sale on Etsy"',
        "SINCE", cutoff
    )
    email_ids = messages[0].split() if messages and messages[0] else []

    now = datetime.now().isoformat(timespec="seconds")
    if not email_ids:
        print(f"{now}: No Etsy sale emails found newer than {cutoff}.")
        mail.logout()
        return

    # Newest -> oldest
    email_ids = list(reversed(email_ids))
    processed_count = 0

    for email_id in email_ids:
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_mime_header(msg.get("Subject"))
        message_id = msg.get("Message-ID", "")

        order_no, amount, ship_by = parse_subject_details(subject)

        # Stop key: prefer Order #, fallback to Message-ID
        processed_key = f"ORDER:{order_no}" if order_no else f"MSG:{message_id.strip() or str(email_id)}"

        # STOP when we reach one we've already done
        if processed_key in processed:
            print(f"{datetime.now().isoformat(timespec='seconds')}: Hit already-processed ({processed_key}). Stopping.")
            break

        plain_body, html_body = extract_plain_and_html(msg)

        # Combine plain and HTML bodies FIRST, then extract
        # This ensures Transaction ID grouping stays intact
        combined_body = ""
        if plain_body:
            combined_body = plain_body
        if html_body:
            html_as_text = html_to_text_minimal(html_body)
            # If we have both, combine them
            if combined_body:
                combined_body = combined_body + "\n\n" + html_as_text
            else:
                combined_body = html_as_text
        
        # Extract from the combined body - this keeps Transaction ID grouping intact
        snippet = extract_item_and_personalization_lines(combined_body)
        
        if not snippet:
            snippet = "[No lines starting with 'Item:' or 'Personalization:' were found in this email body.]"

        # Name / City / State from HTML portion
        name, city, state = extract_name_city_state_from_html(html_body)

        # Check if this order already exists in the snippets file (prevents duplicates)
        if order_no and order_no in existing_orders:
            print(f"⚠️ Order {order_no} already exists in {EMAIL_SNIPPETS_FILE} - skipping duplicate")
        else:
            # Write to snippets file
            with open(EMAIL_SNIPPETS_FILE, "a", encoding="utf-8") as f:
                f.write("-" * 60 + "\n\n")
                f.write(f"Order: {order_no}\n")
                f.write(f"Name: {name or '[Not found]'}\n")
                f.write(f"City: {city or '[Not found]'}\n")
                f.write(f"State: {state or '[Not found]'}\n\n")
                f.write(snippet + "\n\n")
                f.write("-" * 60 + "\n\n")
            
            # Add to our tracking set
            if order_no:
                existing_orders.add(order_no)

        # Record processed key (always mark as processed even if duplicate in snippets)
        append_processed_key(processed_key)
        processed.add(processed_key)

        processed_count += 1
        print(f"Processed: {processed_key}")

    mail.logout()
    print(f"{datetime.now().isoformat(timespec='seconds')}: Done. Processed {processed_count} email(s).")


if __name__ == "__main__":
    import sys
    
    # Allow running deduplication separately: python pullorders.py --dedupe
    if len(sys.argv) > 1 and sys.argv[1] == "--dedupe":
        deduplicate_snippets_file()
    else:
        process_recent_etsy_sales_stop_on_processed()
