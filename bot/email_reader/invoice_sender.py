from __future__ import annotations

import shutil
import time
from collections.abc import Callable
from datetime import date
from pathlib import Path

from config import SENT_REPORT_DIR
from email_reader.email_sender import OutlookEmailSender
from email_reader.outlook_reader import load_saved_mailbox
from excel_reports.daily_report import (
    get_current_invoice_path,
    invoice_path,
)


INVOICE_RECIPIENT = "salesinvoice@egonservices.com"
INVOICE_CC = "fenixny.bizops@fenixdiamonds.com"

SENT_INVOICE_RETENTION_SECONDS = 24 * 60 * 60


def _log(
    callback: Callable[[str], None] | None,
    message: str,
) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def cleanup_sent_invoices(
    *,
    log_callback: Callable[[str], None] | None = None,
) -> int:
    """
    Delete sent Keyzar invoice files after they have been retained
    for at least 24 hours.

    This runs whenever the app starts a processing cycle and whenever
    Send Now is used. Therefore, a file is deleted on the first app
    check after its 24-hour retention period has elapsed.
    """

    SENT_REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    now = time.time()
    deleted_count = 0

    for path in SENT_REPORT_DIR.glob(
        "Keyzar Invoice *.xlsx"
    ):
        if not path.is_file():
            continue

        try:
            age_seconds = now - path.stat().st_mtime

            if age_seconds < SENT_INVOICE_RETENTION_SECONDS:
                continue

            path.unlink()
            deleted_count += 1

            _log(
                log_callback,
                f"Deleted sent invoice after 24-hour retention: {path.name}",
            )

        except OSError as exc:
            _log(
                log_callback,
                f"Could not delete retained invoice {path.name}: {exc}",
            )

    return deleted_count


def _prepare_invoice_for_sending(
    *,
    send_date: date,
    log_callback: Callable[[str], None] | None,
) -> Path | None:
    """
    Find the newest pending Keyzar invoice and rename it to today's
    date before it is attached to the outgoing email.
    """

    current_path = get_current_invoice_path()

    if current_path is None or not current_path.exists():
        _log(
            log_callback,
            "There is no unsent Keyzar invoice to send.",
        )
        return None

    if current_path.stat().st_size == 0:
        raise RuntimeError(
            f"The invoice exists but is empty: {current_path}"
        )

    today_path = invoice_path(send_date)

    if current_path.resolve() == today_path.resolve():
        return current_path

    if today_path.exists():
        raise RuntimeError(
            "A Keyzar invoice already exists with today's date: "
            f"{today_path.name}. Please review the pending_reports "
            "folder before sending so that no workbook is overwritten."
        )

    try:
        current_path.rename(today_path)

    except PermissionError as exc:
        raise RuntimeError(
            "The Keyzar invoice appears to be open in Excel. "
            "Close the workbook and click Send Now again."
        ) from exc

    except OSError as exc:
        raise RuntimeError(
            f"Could not rename {current_path.name} to "
            f"{today_path.name}: {exc}"
        ) from exc

    _log(
        log_callback,
        (
            "Renamed invoice for today's send date: "
            f"{current_path.name} -> {today_path.name}"
        ),
    )

    return today_path


def _archive_sent_invoice(
    workbook_path: Path,
) -> Path:
    """
    Move a successfully submitted invoice to sent_reports and preserve it
    for 24 hours before cleanup.
    """

    SENT_REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    archived_path = SENT_REPORT_DIR / workbook_path.name

    if archived_path.exists():
        timestamp = int(time.time())
        archived_path = SENT_REPORT_DIR / (
            f"{workbook_path.stem} sent-{timestamp}"
            f"{workbook_path.suffix}"
        )

    try:
        shutil.move(
            str(workbook_path),
            str(archived_path),
        )
    except PermissionError as exc:
        raise RuntimeError(
            "Outlook accepted the email, but the invoice could not be "
            "moved to the 24-hour sent archive because it is open in Excel. "
            "Close Excel and move the file manually from pending_reports "
            "to sent_reports."
        ) from exc

    # Reset modified time so the 24-hour retention period begins now.
    now = time.time()
    try:
        Path(archived_path).touch(
            times=(now, now)
        )
    except TypeError:
        # pathlib.Path.touch() does not accept times on all Python versions.
        import os
        os.utime(
            archived_path,
            (now, now),
        )

    return Path(archived_path)


def send_current_keyzar_invoice(
    *,
    processing_date: date | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path | None:
    """
    Find the newest unsent invoice, rename it to today's date, submit it
    through the selected Outlook mailbox, then retain the sent workbook
    locally for 24 hours before automatic cleanup.
    """

    cleanup_sent_invoices(
        log_callback=log_callback,
    )

    send_date = processing_date or date.today()

    workbook_path = _prepare_invoice_for_sending(
        send_date=send_date,
        log_callback=log_callback,
    )

    if workbook_path is None:
        return None

    saved_mailbox = load_saved_mailbox()
    selected_sender = (
        saved_mailbox.get(
            "display_name",
            "",
        ).strip()
        or None
    )

    _log(
        log_callback,
        f"Sending {workbook_path.name}...",
    )

    if selected_sender:
        _log(
            log_callback,
            f"Using selected Outlook mailbox: {selected_sender}",
        )
    else:
        _log(
            log_callback,
            (
                "No saved Outlook mailbox was found. "
                "Outlook's default sending account will be used."
            ),
        )

    OutlookEmailSender().send(
        recipients=INVOICE_RECIPIENT,
        cc=INVOICE_CC,
        send_from=selected_sender,
        subject=(
            f"Keyzar Invoice "
            f"{send_date.strftime('%m%d%y')}"
        ),
        body=(
            "Hi,\n\n"
            "Please find attached the current Keyzar invoice.\n\n"
            "Thanks,\n"
            "Fenix Sales"
        ),
        attachments=[workbook_path],
    )

    # Outlook accepted the message. Keep the workbook for 24 hours.
    archived_path = _archive_sent_invoice(
        workbook_path
    )

    _log(
        log_callback,
        (
            f"{workbook_path.name} was submitted successfully. "
            f"It will remain in sent_reports for 24 hours: "
            f"{archived_path}"
        ),
    )

    return archived_path
