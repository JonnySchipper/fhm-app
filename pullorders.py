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

# Only keep these lines from the email body:
KEEP_LINE_RE = re.compile(r"^\s*(Item|Personalization)\s*:\s*(.+)\s*$", re.IGNORECASE)

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
    """Return (text_plain, text_html). Prefer first encountered of each."""
    text_plain = ""
    text_html = ""

    if msg.is_multipart():
        for part in msg.walk():
            disp = part.get_content_disposition()
            if disp == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype == "text/plain" and not text_plain:
                text_plain = _decode_part_payload(part)
            elif ctype == "text/html" and not text_html:
                text_html = _decode_part_payload(part)
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            text_html = _decode_part_payload(msg)
        else:
            text_plain = _decode_part_payload(msg)

    return text_plain or "", text_html or ""


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


def append_processed_key(key: str):
    with open(PROCESSED_FILE, "a", encoding="utf-8") as f:
        f.write(key + "\n")


def extract_item_and_personalization_lines(body_text: str) -> str:
    """
    Returns ONLY the lines that start with:
      - Item:
      - Personalization:
    """
    if not body_text:
        return ""

    kept = []
    for line in body_text.splitlines():
        m = KEEP_LINE_RE.match(line)
        if m:
            label = m.group(1).capitalize()  # Item / Personalization
            value = m.group(2).strip()
            kept.append(f"{label}: {value}")

    return "\n".join(kept).strip()


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
    gmail_user, gmail_pass = load_gmail_credentials()
    processed = load_processed_set()

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

        # Your snippet (Item/Personalization) - try plain first, then HTML->text fallback
        snippet = extract_item_and_personalization_lines(plain_body)
        if not snippet and html_body:
            snippet = extract_item_and_personalization_lines(html_to_text_minimal(html_body))

        if not snippet:
            snippet = "[No lines starting with 'Item:' or 'Personalization:' were found in this email body.]"

        # Name / City / State from HTML portion
        name, city, state = extract_name_city_state_from_html(html_body)

        with open(EMAIL_SNIPPETS_FILE, "a", encoding="utf-8") as f:
            f.write("-" * 60 + "\n\n")
            f.write(f"Order: {order_no}\n")
            f.write(f"Name: {name or '[Not found]'}\n")
            f.write(f"City: {city or '[Not found]'}\n")
            f.write(f"State: {state or '[Not found]'}\n\n")
            f.write(snippet + "\n\n")
            f.write("-" * 60 + "\n\n")

        # Record processed key
        append_processed_key(processed_key)
        processed.add(processed_key)

        processed_count += 1
        print(f"Processed: {processed_key}")

    mail.logout()
    print(f"{datetime.now().isoformat(timespec='seconds')}: Done. Processed {processed_count} email(s).")


if __name__ == "__main__":
    process_recent_etsy_sales_stop_on_processed()
