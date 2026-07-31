from __future__ import annotations

import traceback
from collections.abc import Callable

from config import KEYZAR_SENDER_KEYWORDS, KEYZAR_SUBJECT_KEYWORDS, OUTLOOK_EMAIL_LIMIT
from email_reader.email_sender import OutlookEmailSender
from email_reader.keyzar_parser import is_keyzar_email, parse_keyzar_email
from email_reader.outlook_reader import OutlookReader
from fenix.browser import FenixBrowser
from fenix.keyzar_order_processor import process_keyzar_stone
from processing.database import JobDatabase

LogCallback = Callable[[str], None] | None


def _log(callback: LogCallback, message: str) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def collect_keyzar_orders(log_callback: LogCallback = None) -> list:
    reader = OutlookReader()
    orders = []
    emails_checked = 0
    subject_candidates = 0

    try:
        _log(
            log_callback,
            f"Connecting to Classic Outlook and checking the latest {OUTLOOK_EMAIL_LIMIT} emails...",
        )
        reader.connect()
        _log(log_callback, f"Outlook mailbox selected: {reader.mailbox_display_name}")

        recent_emails = list(reader.iter_recent_emails(OUTLOOK_EMAIL_LIMIT))
        emails_checked = len(recent_emails)
        _log(log_callback, f"Outlook emails actually read: {emails_checked}")

        if recent_emails:
            _log(log_callback, "Newest Outlook subjects:")
            for email in recent_emails[:5]:
                received = (
                    email.received_time.strftime("%m/%d/%Y %I:%M %p")
                    if email.received_time
                    else "Unknown time"
                )
                _log(
                    log_callback,
                    f"  {received} | {email.sender_name} | {email.subject}",
                )

        for email in recent_emails:
            normalized_subject = (email.subject or "").lower()
            if any(keyword.lower() in normalized_subject for keyword in KEYZAR_SUBJECT_KEYWORDS):
                subject_candidates += 1

            if not is_keyzar_email(
                email=email,
                sender_keywords=KEYZAR_SENDER_KEYWORDS,
                subject_keywords=KEYZAR_SUBJECT_KEYWORDS,
            ):
                continue

            order = parse_keyzar_email(email)
            if not order.vendor_ids:
                _log(
                    log_callback,
                    f"A Keyzar-looking email matched, but no Vendor ID was extracted: {order.subject}",
                )
                continue

            _log(
                log_callback,
                f"Matched Keyzar order: {order.order_number or 'No order number'} | "
                f"Vendor IDs: {', '.join(order.vendor_ids)}",
            )
            orders.append(order)

    finally:
        reader.disconnect()

    _log(
        log_callback,
        f"Keyzar orders collected: {len(orders)} "
        f"(subject candidates: {subject_candidates}, emails checked: {emails_checked})",
    )
    return orders


def run_once(log_callback: LogCallback = None) -> dict[str, int]:
    database = JobDatabase()
    email_sender = OutlookEmailSender()
    orders = collect_keyzar_orders(log_callback)

    pending_jobs = []
    skipped_count = 0
    processed_count = 0
    failed_count = 0

    for order in orders:
        for vendor_id in order.vendor_ids:
            if database.is_finished(order.outlook_entry_id, vendor_id):
                skipped_count += 1
                _log(log_callback, f"Already processed: {vendor_id} | {order.subject}")
                continue
            pending_jobs.append((order, vendor_id))

    if not pending_jobs:
        _log(log_callback, "No new Keyzar orders to process.")
        return {"processed": 0, "skipped": skipped_count, "failed": 0}

    browser = FenixBrowser()
    try:
        browser.start()
        for order, vendor_id in pending_jobs:
            order_date = order.order_date.isoformat() if order.order_date else None
            database.upsert(
                outlook_entry_id=order.outlook_entry_id,
                vendor_id=vendor_id,
                order_number=order.order_number,
                order_date=order_date,
                email_subject=order.subject,
                status="PROCESSING",
            )

            try:
                _log(log_callback, f"Processing Vendor ID {vendor_id}...")
                result = process_keyzar_stone(
                    browser=browser,
                    order=order,
                    vendor_id=vendor_id,
                    email_sender=email_sender,
                )
                database.upsert(
                    outlook_entry_id=order.outlook_entry_id,
                    vendor_id=vendor_id,
                    order_number=order.order_number,
                    order_date=order_date,
                    email_subject=order.subject,
                    status=result.result_type,
                    details=result.details,
                )
                processed_count += 1
                _log(log_callback, f"Finished {vendor_id}: {result.result_type}")

            except Exception as exc:
                failed_count += 1
                database.upsert(
                    outlook_entry_id=order.outlook_entry_id,
                    vendor_id=vendor_id,
                    order_number=order.order_number,
                    order_date=order_date,
                    email_subject=order.subject,
                    status="FAILED",
                    details=f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}",
                )
                _log(log_callback, f"Failed processing {vendor_id}: {exc}")
                if "session expired" in str(exc).lower():
                    _log(log_callback, "Fenix login is required. Click Stop Bot, then Login to Fenix.")
                    break
    finally:
        browser.stop()

    return {
        "processed": processed_count,
        "skipped": skipped_count,
        "failed": failed_count,
    }
