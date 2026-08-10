from __future__ import annotations

import shutil
from collections.abc import Callable
from datetime import date, datetime
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


def _log(
    callback: Callable[[str], None] | None,
    message: str,
) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def _prepare_invoice_for_sending(
    *,
    send_date: date,
    log_callback: Callable[[str], None] | None,
) -> Path | None:
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

    lock_test_path = current_path.with_name(
        f"{current_path.stem}.__send_test__{current_path.suffix}"
    )

    try:
        current_path.rename(lock_test_path)
        lock_test_path.rename(current_path)

    except PermissionError as exc:
        raise RuntimeError(
            "The Keyzar invoice is currently open in Excel. "
            "Close the workbook and click Send Now again."
        ) from exc

    except OSError as exc:
        if lock_test_path.exists() and not current_path.exists():
            try:
                lock_test_path.rename(current_path)
            except Exception:
                pass

        raise RuntimeError(
            f"Could not prepare the invoice for sending: {exc}"
        ) from exc

    if current_path.resolve() == today_path.resolve():
        return current_path

    if today_path.exists():
        raise RuntimeError(
            "A pending Keyzar invoice already exists with today's date: "
            f"{today_path.name}. Please review pending_reports before sending."
        )

    try:
        current_path.rename(today_path)

    except PermissionError as exc:
        raise RuntimeError(
            "The Keyzar invoice is currently open in Excel. "
            "Close the workbook and click Send Now again."
        ) from exc

    _log(
        log_callback,
        (
            "Renamed invoice to today's send date: "
            f"{current_path.name} -> {today_path.name}"
        ),
    )

    return today_path


def _archive_sent_invoice(
    workbook_path: Path,
) -> Path:
    SENT_REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    archived_path = SENT_REPORT_DIR / workbook_path.name

    if archived_path.exists():
        timestamp = datetime.now().strftime("%H%M%S")
        archived_path = SENT_REPORT_DIR / (
            f"{workbook_path.stem} "
            f"sent-{timestamp}"
            f"{workbook_path.suffix}"
        )

    try:
        shutil.move(
            str(workbook_path),
            str(archived_path),
        )

    except PermissionError as exc:
        raise RuntimeError(
            "Outlook accepted the email, but Windows could not move the "
            "invoice to the archive. The workbook may still be open in "
            "Excel. Close Excel and move the file from pending_reports "
            "to sent_reports manually."
        ) from exc

    return archived_path


def send_current_keyzar_invoice(
    *,
    processing_date: date | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path | None:
    send_date = processing_date or date.today()

    workbook_path = _prepare_invoice_for_sending(
        send_date=send_date,
        log_callback=log_callback,
    )

    if workbook_path is None:
        return None

    saved_mailbox = load_saved_mailbox()

    selected_store_id = saved_mailbox.get(
        "store_id",
        "",
    ).strip()

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
            (
                "Selected Outlook mailbox for sending: "
                f"{selected_sender}"
            ),
        )
    else:
        _log(
            log_callback,
            (
                "No selected Outlook mailbox was saved. "
                "Outlook will use its default account."
            ),
        )

    OutlookEmailSender().send(
        recipients=INVOICE_RECIPIENT,
        cc=INVOICE_CC,
        send_from=selected_sender,
        send_store_id=selected_store_id,
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

    archived_path = _archive_sent_invoice(
        workbook_path
    )

    _log(
        log_callback,
        (
            f"{workbook_path.name} was submitted successfully "
            f"and archived to {archived_path}."
        ),
    )

    return archived_path
