from __future__ import annotations

from collections.abc import Callable
from datetime import date
from pathlib import Path

from email_reader.email_sender import OutlookEmailSender
from excel_reports.daily_report import invoice_path


INVOICE_RECIPIENT = "salesinvoice@egonservices.com"
INVOICE_CC = "fenixny.bizops@fenixdiamonds.com"
INVOICE_SENDER = "sales@fenixdiamonds.com"


def _log(
    callback: Callable[[str], None] | None,
    message: str,
) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def send_current_keyzar_invoice(
    *,
    processing_date: date | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path | None:
    """
    Send the current unsent invoice manually.

    The workbook is deleted only after Outlook's Send call returns
    successfully. The next blocked stone then creates a new workbook.
    """

    processing_date = processing_date or date.today()
    workbook_path = invoice_path(processing_date)

    if not workbook_path.exists():
        _log(
            log_callback,
            "There is no unsent Keyzar invoice to send.",
        )
        return None

    if workbook_path.stat().st_size == 0:
        raise RuntimeError(
            f"The invoice exists but is empty: {workbook_path}"
        )

    _log(
        log_callback,
        f"Sending {workbook_path.name}...",
    )

    OutlookEmailSender().send(
        recipients=INVOICE_RECIPIENT,
        cc=INVOICE_CC,
        send_from=INVOICE_SENDER,
        subject=(
            f"Keyzar Invoice "
            f"{processing_date.strftime('%m%d%y')}"
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
