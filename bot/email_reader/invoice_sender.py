from __future__ import annotations

from collections.abc import Callable
from datetime import date
from pathlib import Path

from config import SENT_REPORT_DIR
from email_reader.email_sender import OutlookEmailSender, REQUIRED_BOT_SENDER
from email_reader.outlook_reader import load_saved_mailbox
from excel_reports.daily_report import (
    PENDING_INVOICE_FILENAME,
    get_current_invoice_path,
    invoice_filename,
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


def _next_archive_name(send_date: date) -> str:
    """Return a unique final filename for sent_reports."""
    SENT_REPORT_DIR.mkdir(parents=True, exist_ok=True)

    base_name = invoice_filename(send_date)
    base_path = SENT_REPORT_DIR / base_name

    if not base_path.exists():
        return base_name

    stem = Path(base_name).stem
    suffix = Path(base_name).suffix
    number = 2

    while True:
        candidate = f"{stem} ({number}){suffix}"
        if not (SENT_REPORT_DIR / candidate).exists():
            return candidate
        number += 1


def _prepare_invoice_for_sending(
    *,
    send_date: date,
    log_callback: Callable[[str], None] | None,
) -> tuple[Path, Path] | None:
    """Rename the one pending workbook to its final dated filename.

    Returns (pending_dated_path, final_archive_path). The source remains in
    pending_reports until Outlook accepts the message.
    """
    current_path = get_current_invoice_path()

    if current_path is None or not current_path.exists():
        _log(log_callback, "There is no unsent Keyzar invoice to send.")
        return None

    if current_path.stat().st_size == 0:
        raise RuntimeError(
            f"The invoice exists but is empty: {current_path}"
        )

    archive_name = _next_archive_name(send_date)
    pending_dated_path = current_path.parent / archive_name
    archive_path = SENT_REPORT_DIR / archive_name

    # Rename is also our Excel-lock test. If Excel still has the workbook
    # open, Windows will refuse this and nothing will be emailed.
    if current_path.resolve() != pending_dated_path.resolve():
        try:
            current_path.replace(pending_dated_path)
        except PermissionError as exc:
            raise RuntimeError(
                "The Keyzar invoice is currently open in Excel. "
                "Close the workbook and click Send Now again."
            ) from exc
        except OSError as exc:
            raise RuntimeError(
                f"Could not prepare the invoice for sending: {exc}"
            ) from exc

    return pending_dated_path, archive_path


def _restore_pending_name(workbook_path: Path) -> None:
    """Restore the active filename if Outlook fails before sending."""
    if not workbook_path.exists():
        return

    pending_path = workbook_path.parent / PENDING_INVOICE_FILENAME

    if workbook_path.resolve() == pending_path.resolve():
        return

    if pending_path.exists():
        return

    try:
        workbook_path.replace(pending_path)
    except OSError:
        pass


def _archive_sent_invoice(
    workbook_path: Path,
    archive_path: Path,
) -> Path:
    """Atomically MOVE a successfully sent workbook into sent_reports."""
    SENT_REPORT_DIR.mkdir(parents=True, exist_ok=True)

    if archive_path.exists():
        raise RuntimeError(
            f"Archive target already exists: {archive_path.name}"
        )

    try:
        # Both folders are inside the same app data tree, so Path.replace is
        # a real filesystem move/rename: it does not leave a source copy.
        workbook_path.replace(archive_path)
    except PermissionError as exc:
        raise RuntimeError(
            "Outlook accepted the email, but Windows could not MOVE the "
            "invoice to sent_reports. Close Excel and move the dated file "
            "from pending_reports to sent_reports manually."
        ) from exc

    if workbook_path.exists():
        raise RuntimeError(
            "Archive verification failed: the sent invoice still exists in "
            "pending_reports. Do not process another order until it is moved."
        )

    if not archive_path.exists():
        raise RuntimeError(
            "Archive verification failed: the sent invoice was not found in "
            "sent_reports after the move."
        )

    return archive_path


def send_current_keyzar_invoice(
    *,
    processing_date: date | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path | None:
    send_date = processing_date or date.today()

    prepared = _prepare_invoice_for_sending(
        send_date=send_date,
        log_callback=log_callback,
    )

    if prepared is None:
        return None

    workbook_path, archive_path = prepared
    saved_mailbox = load_saved_mailbox()
    selected_store_id = saved_mailbox.get("store_id", "").strip()
    selected_mailbox = saved_mailbox.get("display_name", "").strip()

    _log(log_callback, f"Sending {workbook_path.name}...")
    _log(
        log_callback,
        f"Connected/read mailbox: {selected_mailbox or 'not saved'}",
    )
    _log(
        log_callback,
        f"Required outgoing sender: {REQUIRED_BOT_SENDER}",
    )

    try:
        OutlookEmailSender().send(
            recipients=INVOICE_RECIPIENT,
            cc=INVOICE_CC,
            send_from=REQUIRED_BOT_SENDER,
            send_store_id=selected_store_id,
            subject=f"Keyzar Invoice {send_date.strftime('%m%d%y')}",
            body=(
                "Hi,\n\n"
                "Please find attached the current Keyzar invoice.\n\n"
                "Thanks,\n"
                "Fenix Sales"
            ),
            attachments=[workbook_path],
        )
    except Exception:
        _restore_pending_name(workbook_path)
        raise

    archived_path = _archive_sent_invoice(
        workbook_path,
        archive_path,
    )

    _log(
        log_callback,
        f"Sent invoice MOVED to archive: {archived_path}",
    )
    _log(
        log_callback,
        "pending_reports is now clear for the next Keyzar invoice.",
    )

    return archived_path
