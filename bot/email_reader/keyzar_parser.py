from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from email.utils import parsedate_to_datetime

from email_reader.outlook_reader import OutlookEmail


@dataclass
class KeyzarOrder:
    outlook_entry_id: str
    sender_name: str
    sender_email: str
    subject: str

    # When the forwarded email reached the current Outlook mailbox.
    mailbox_received_time: datetime | None

    # When Keyzar originally sent the order.
    original_sent_time: datetime | None

    order_number: str | None
    vendor_ids: list[str]

    igi_number: str | None
    shape: str | None
    carat: float | None
    color: str | None
    clarity: str | None
    location: str | None
    size: str | None
    price: float | None

    body: str
    html_body: str
    is_forwarded: bool

    @property
    def order_date(self):
        """
        Date used for daily Excel batching.

        Prefer the original Keyzar sent date. Fall back to the Outlook
        mailbox received date when the original forwarded date is unavailable.
        """

        timestamp = self.original_sent_time or self.mailbox_received_time

        if timestamp is None:
            return None

        return timestamp.date()


def normalize_subject(subject: str) -> str:
    """
    Remove repeated FW:, FWD:, and RE: prefixes.
    """

    normalized = subject.strip()

    while True:
        updated = re.sub(
            r"^\s*(?:fw|fwd|re)\s*:\s*",
            "",
            normalized,
            flags=re.IGNORECASE,
        )

        if updated == normalized:
            break

        normalized = updated.strip()

    return normalized


def is_forwarded_email(email: OutlookEmail) -> bool:
    subject = email.subject.strip().lower()
    body = email.body.lower()

    return (
        subject.startswith("fw:")
        or subject.startswith("fwd:")
        or "\nfrom:" in body
        or "\r\nfrom:" in body
    )


def body_contains_keyzar_sender(
    body: str,
    sender_keywords: list[str],
) -> bool:
    body_lower = body.lower()

    return any(
        keyword.strip().lower() in body_lower
        for keyword in sender_keywords
        if keyword.strip()
    )


def is_keyzar_email(
    email: OutlookEmail,
    sender_keywords: list[str],
    subject_keywords: list[str],
) -> bool:
    actual_sender_text = (
        f"{email.sender_name} {email.sender_email}"
    ).lower()

    normalized_subject = normalize_subject(email.subject).lower()

    direct_sender_match = any(
        keyword.strip().lower() in actual_sender_text
        for keyword in sender_keywords
        if keyword.strip()
    )

    forwarded_sender_match = body_contains_keyzar_sender(
        body=email.body,
        sender_keywords=sender_keywords,
    )

    subject_match = any(
        keyword.strip().lower() in normalized_subject
        for keyword in subject_keywords
        if keyword.strip()
    )

    return subject_match and (
        direct_sender_match or forwarded_sender_match
    )


def clean_vendor_id(value: str) -> str:
    return value.strip().upper().rstrip(".,;:)")


def extract_vendor_ids(text: str) -> list[str]:
    patterns = [
        r"\bvendor\s*id\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})",
        r"\bparcel\s*id\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})",
        r"\bstone\s*id\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})",
    ]

    found: list[str] = []

    for pattern in patterns:
        matches = re.findall(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        for match in matches:
            vendor_id = clean_vendor_id(match)

            if vendor_id and vendor_id not in found:
                found.append(vendor_id)

    return found


def extract_order_number(subject: str, body: str) -> str | None:
    """
    Expected subject:
        New Stone Order: 811678374
    """

    searchable_text = f"{normalize_subject(subject)}\n{body}"

    match = re.search(
        r"\bnew\s+stone\s+order\s*:\s*(\d+)",
        searchable_text,
        flags=re.IGNORECASE,
    )

    return match.group(1) if match else None


def extract_original_sent_time(body: str) -> datetime | None:
    """
    Extract the Sent line from the original forwarded Keyzar header.

    Example:
        From: Keyzar <info@keyzarjewelry.com>
        Sent: Friday, July 24, 2026 11:35 AM
    """

    patterns = [
        (
            r"From:\s*Keyzar[^\r\n]*"
            r"(?:\r?\n)+"
            r"Sent:\s*([^\r\n]+)"
        ),
        r"Sent:\s*([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            body,
            flags=re.IGNORECASE,
        )

        if not match:
            continue

        raw_value = match.group(1).strip()

        formats = [
            "%A, %B %d, %Y %I:%M %p",
            "%A, %B %d, %Y at %I:%M %p",
            "%B %d, %Y %I:%M %p",
        ]

        for date_format in formats:
            try:
                return datetime.strptime(raw_value, date_format)
            except ValueError:
                continue

        try:
            return parsedate_to_datetime(raw_value)
        except (TypeError, ValueError, OverflowError):
            continue

    return None


def extract_text_value(
    text: str,
    label_pattern: str,
) -> str | None:
    match = re.search(
        rf"\b{label_pattern}\s*:\s*([^\r\n]+)",
        text,
        flags=re.IGNORECASE,
    )

    if not match:
        return None

    value = match.group(1).strip()
    value = re.sub(r"\s{2,}.*$", "", value)

    return value.rstrip(".,;")


def extract_float_value(
    text: str,
    label_pattern: str,
) -> float | None:
    raw_value = extract_text_value(text, label_pattern)

    if not raw_value:
        return None

    match = re.search(r"\d+(?:\.\d+)?", raw_value.replace(",", ""))

    if not match:
        return None

    try:
        return float(match.group(0))
    except ValueError:
        return None


def extract_price(text: str) -> float | None:
    """
    Prefer the detailed price shown in the stone section.

    The email may also contain text such as:
        $276.32 diamond was sold!
    """

    matches = re.findall(
        r"\$\s*([\d,]+(?:\.\d{2})?)",
        text,
    )

    if not matches:
        return None

    # Usually the final price occurrence is the detailed stone price.
    raw_value = matches[-1].replace(",", "")

    try:
        return float(raw_value)
    except ValueError:
        return None


def extract_igi_number(text: str) -> str | None:
    patterns = [
        r"\bIGI\s*[:#]?\s*(\d{6,})",
        r"\bcertificate\s*(?:id|number|no\.?|#)?\s*:\s*(\d{6,})",
        r"\bcert(?:ificate)?\s*(?:id|number|no\.?|#)?\s*:\s*(\d{6,})",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:
            return match.group(1)

    return None


def parse_keyzar_email(email: OutlookEmail) -> KeyzarOrder:
    searchable_text = "\n".join(
        [
            email.subject or "",
            email.body or "",
        ]
    )

    return KeyzarOrder(
        outlook_entry_id=email.entry_id,
        sender_name=email.sender_name,
        sender_email=email.sender_email,
        subject=email.subject,
        mailbox_received_time=email.received_time,
        original_sent_time=extract_original_sent_time(email.body),
        order_number=extract_order_number(
            subject=email.subject,
            body=email.body,
        ),
        vendor_ids=extract_vendor_ids(searchable_text),
        igi_number=extract_igi_number(searchable_text),
        shape=extract_text_value(searchable_text, r"shape"),
        carat=extract_float_value(searchable_text, r"carat"),
        color=extract_text_value(searchable_text, r"colou?r"),
        clarity=extract_text_value(searchable_text, r"clarity"),
        location=extract_text_value(searchable_text, r"location"),
        size=extract_text_value(searchable_text, r"size"),
        price=extract_price(searchable_text),
        body=email.body,
        html_body=email.html_body,
        is_forwarded=is_forwarded_email(email),
    )