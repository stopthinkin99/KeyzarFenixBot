from __future__ import annotations

from collections.abc import Callable
from datetime import date
from pathlib import Path

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
        f"Renamed invoice for today's send date: "
        f"{current_path.name} -> {today_path.name}",
    )

    return today_path


def send_current_keyzar_invoice(
    *,
    processing_date: date | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path | None:
    """
    Find the newest unsent invoice, rename it to today's date, send it
    through the Outlook mailbox selected in Connect Outlook, and delete
    it only after Outlook accepts the message.
    """

    send_date = processing_date or date.today()

    workbook_path = _prepare_invoice_for_sending(
        send_date=send_date,
        log_callback=log_callback,
    )

    if workbook_path is None:
        return None

    saved_mailbox = load_saved_mailbox()
    selected_sender = (
        saved_mailbox.get("display_name", "").strip()
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
            "No saved Outlook mailbox was found. "
            "Outlook's default sending account will be used.",
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

    # Outlook accepted the message for delivery.
    workbook_path.unlink()

    _log(
        log_callback,
        (
            f"{workbook_path.name} was submitted successfully "
            "and deleted from the computer."
        ),
    )

    return workbook_path
