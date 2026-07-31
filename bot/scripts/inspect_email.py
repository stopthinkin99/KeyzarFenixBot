from __future__ import annotations

import sys
from pathlib import Path

# Allows this script to import modules from the project root.
PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import (
    KEYZAR_SENDER_KEYWORDS,
    KEYZAR_SUBJECT_KEYWORDS,
    OUTLOOK_EMAIL_LIMIT,
)
from email_reader.keyzar_parser import (
    is_keyzar_email,
    parse_keyzar_email,
)
from email_reader.outlook_reader import OutlookReader


def clean_preview(text: str, max_length: int = 5000) -> str:
    text = text.replace("\r\n", "\n").strip()

    if len(text) > max_length:
        return text[:max_length] + "\n...[body shortened]"

    return text


def main() -> None:
    reader = OutlookReader()

    try:
        print("[INFO] Connecting to Classic Outlook...")
        reader.connect()
        print("[INFO] Connected successfully.\n")

        total_checked = 0
        total_matched = 0

        for email in reader.iter_recent_emails(OUTLOOK_EMAIL_LIMIT):
            total_checked += 1

            if not is_keyzar_email(
                email=email,
                sender_keywords=KEYZAR_SENDER_KEYWORDS,
                subject_keywords=KEYZAR_SUBJECT_KEYWORDS,
            ):
                continue

            total_matched += 1
            order = parse_keyzar_email(email)

            print(f"Mailbox Time : {order.mailbox_received_time}")
            print(f"Original Sent: {order.original_sent_time}")
            print(f"Order Date   : {order.order_date}")
            print(f"Sender Name  : {order.sender_name}")
            print(f"Sender Email : {order.sender_email}")
            print(f"Subject      : {order.subject}")
            print(f"Forwarded    : {order.is_forwarded}")
            print(f"Order Number : {order.order_number or 'NOT FOUND'}")
            print(f"Vendor IDs   : {order.vendor_ids or 'NOT FOUND'}")
            print(f"IGI Number   : {order.igi_number or 'NOT FOUND'}")
            print(f"Shape        : {order.shape or 'NOT FOUND'}")
            print(f"Carat        : {order.carat if order.carat is not None else 'NOT FOUND'}")
            print(f"Color        : {order.color or 'NOT FOUND'}")
            print(f"Clarity      : {order.clarity or 'NOT FOUND'}")
            print(f"Location     : {order.location or 'NOT FOUND'}")
            print(f"Size         : {order.size or 'NOT FOUND'}")
            print(f"Price        : {order.price if order.price is not None else 'NOT FOUND'}")
            print(f"Outlook ID   : {order.outlook_entry_id}")

        print("=" * 80)
        print(f"[DONE] Checked {total_checked} recent emails.")
        print(f"[DONE] Found {total_matched} possible Keyzar emails.")

        if total_matched == 0:
            print(
                "[NOTE] We may need to update the sender and subject "
                "keywords in config.py."
            )

    except Exception as exc:
        print(f"[ERROR] {exc}")
        raise SystemExit(1)

    finally:
        reader.disconnect()


if __name__ == "__main__":
    main()