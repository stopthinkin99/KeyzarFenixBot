from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from config import ALERT_RECIPIENT, BOT_DRY_RUN, PORTAL_URL
from email_reader.email_sender import OutlookEmailSender
from excel_reports.daily_report import append_stone_to_daily_report
from fenix.memo_list import determine_blocked_for, search_memo_list

# Reuse the exact portal functions that already passed your dry-run test.
from scripts.test_available_stone_memo import (
    align_headers_and_values,
    enter_stone_number,
    extract_result_headers,
    extract_result_values,
    find_memo_issue_button,
    find_note_field,
    find_save_button,
    find_search_button,
    find_status,
    find_stone_input,
    locate_result_row,
    normalize_text,
    print_result,
    select_all_stone_type,
    select_memo_row,
    select_search_result_master_checkbox,
)

MEMO_NOTE = "Keyzar"
ALERT_SENDER = "sales@fenixdiamonds.com"


@dataclass
class ProcessingResult:
    vendor_id: str
    result_type: str
    inventory_status: str = ""
    details: str = ""
    report_path: str = ""


def _search_button_count(search_button) -> int:
    try:
        text = normalize_text(search_button.inner_text())
    except Exception:
        text = normalize_text(search_button.get_attribute("value"))
    match = re.search(r"\(\s*(\d+)\s*\)", text)
    return int(match.group(1)) if match else 0

def wait_for_result_grid(page,vendor_id: str,timeout_ms: int = 60_000,) -> None:
    """
    Wait for the Vendor ID to appear after SEARCH (1+) is clicked.
    """

    print(
        f"[INFO] Waiting for {vendor_id} "
        "to appear in the Search Result grid..."
    )

    try:
        page.wait_for_function(
            r"""
            (vendorId) => {
                const target = vendorId
                    .trim()
                    .toUpperCase();

                const elements = Array.from(
                    document.querySelectorAll(
                        "a, td, span, div"
                    )
                );

                return elements.some((element) => {
                    const text = (
                        element.textContent || ""
                    )
                        .replace(/\s+/g, " ")
                        .trim()
                        .toUpperCase();

                    return text === target;
                });
            }
            """,
            arg=vendor_id,
            timeout=timeout_ms,
        )

    except Exception as exc:
        raise RuntimeError(
            f"Fenix reported a result, but {vendor_id} "
            "did not appear in the Search Result grid within "
            f"{timeout_ms // 1000} seconds."
        ) from exc

    page.wait_for_timeout(1_500)

    print(
        f"[SUCCESS] {vendor_id} is visible "
        "in the Search Result grid."
    )


def wait_for_search_count(page, search_button, timeout_ms: int = 12_000) -> int:
    print("[INFO] Waiting for Fenix to prepare the stone search...")
    elapsed = 0
    previous_count: int | None = None
    stable_reads = 0

    while elapsed < timeout_ms:
        count = _search_button_count(search_button)
        if count == previous_count:
            stable_reads += 1
        else:
            stable_reads = 0
            previous_count = count

        if count > 0 and stable_reads >= 1:
            print(f"[INFO] Search is ready: SEARCH ({count})")
            return count
        if count == 0 and elapsed >= 5_000 and stable_reads >= 3:
            print("[INFO] Search is ready: SEARCH (0)")
            return 0

        page.wait_for_timeout(500)
        elapsed += 500

    count = _search_button_count(search_button)
    print(f"[INFO] Search count after timeout: {count}")
    return count


def _order_date(order) -> date:
    return order.order_date or date.today()


def _order_context(order, vendor_id: str) -> str:
    return (
        f"Vendor ID: {vendor_id}\n"
        f"Keyzar Order: {order.order_number or 'Not found'}\n"
        f"Order Date: {_order_date(order):%m/%d/%Y}\n"
        f"Source Subject: {order.subject}\n"
    )


def send_not_found_alert(sender: OutlookEmailSender, order, vendor_id: str) -> None:
    sender.send(
        recipients=ALERT_RECIPIENT,
        subject=f"Keyzar Stone Not Found – {vendor_id}",
        body=(
            "The Keyzar Fenix Bot could not locate the requested stone.\n\n"
            f"{_order_context(order, vendor_id)}"
            "Fenix Search Count: 0\n\n"
            "No memo was created and the stone was not added to Excel.\n"
            "Please review the order manually."
        ),
        send_from=ALERT_SENDER,
    )


def send_unavailable_alert(
    sender: OutlookEmailSender,
    order,
    vendor_id: str,
    inventory_status: str,
    records,
) -> None:
    blocked_for = determine_blocked_for(records)
    if records:
        record = records[0]
        memo_details = (
            f"Memo Number: {record.memo_number or 'Not found'}\n"
            f"Memo Date: {record.memo_date or 'Not found'}\n"
            f"Customer: {record.customer or 'Not listed'}\n"
            f"Salesperson: {record.salesman or 'Not listed'}\n"
            f"Memo Type: {record.memo_type or 'Not listed'}\n"
            f"Service Location: {record.service_location or 'Not listed'}\n"
            f"Note: {record.note or 'Not listed'}\n"
            f"Memo Status: {record.status or 'Not listed'}\n"
        )
    else:
        memo_details = "Memo List: No matching Created memo was found.\n"

    sender.send(
        recipients=ALERT_RECIPIENT,
        subject=f"Keyzar Stone Unavailable – {vendor_id}",
        body=(
            "The requested Keyzar stone is not currently available.\n\n"
            f"{_order_context(order, vendor_id)}"
            f"Inventory Status: {inventory_status or 'Unknown'}\n"
            f"{memo_details}\n"
            f"The stone appears to be blocked for: {blocked_for}\n\n"
            "No new memo was created and the stone was not added to Excel."
        ),
        send_from=ALERT_SENDER,
    )



def is_blocked_for_keyzar(records) -> bool:
    """
    Return True when any matching memo record identifies Keyzar in the
    Note or Customer field.
    """

    for record in records:
        note = str(getattr(record, "note", "") or "").strip().lower()
        customer = str(
            getattr(record, "customer", "") or ""
        ).strip().lower()

        if "keyzar" in note or "keyzar" in customer:
            return True

    return False

def _click_memo_issue(page, button) -> None:
    print("[INFO] Clicking the Memo Issue button...")
    try:
        button.click(force=True, timeout=15_000)
    except Exception:
        print("[WARNING] Normal click failed; using JavaScript click.")
        button.evaluate("(element) => element.click()")
    page.wait_for_timeout(4_000)


def _save_or_dry_run(page) -> str:
    save_button = find_save_button(page)
    if BOT_DRY_RUN:
        print("=" * 80)
        print("[DRY RUN] The bot would click SAVE now.")
        print("[DRY RUN] SAVE WAS NOT CLICKED.")
        print("=" * 80)
        return "DRY_RUN_PREPARED"

    if save_button is None:
        raise RuntimeError("Memo prepared, but Save control was not found.")

    print("[LIVE] Clicking SAVE...")
    save_button.click(force=True)
    page.wait_for_timeout(5_000)
    print("[LIVE] SAVE was clicked.")
    return "COMPLETED"


def process_keyzar_stone(*, browser, order, vendor_id: str, email_sender: OutlookEmailSender) -> ProcessingResult:
    page = browser.get_active_page()
    page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
    page.wait_for_timeout(2_500)

    if "/login" in page.url.lower():
        raise RuntimeError("Fenix session expired. Run scripts\\login_once.py again.")

    print("\n" + "#" * 90)
    print(f"PROCESSING KEYZAR STONE: {vendor_id}")
    print("#" * 90)

    stone_input = find_stone_input(page)
    search_button = find_search_button(page)
    enter_stone_number(page=page, stone_input=stone_input, vendor_id=vendor_id)
    select_all_stone_type(page)

    result_count = wait_for_search_count(page, search_button)
    if result_count == 0:
        print(f"[RESULT] {vendor_id}: not found in Search Stock.")
        send_not_found_alert(email_sender, order, vendor_id)
        return ProcessingResult(vendor_id, "NOT_FOUND_ALERTED", details="Fenix Search Count: 0")

    print(f"[INFO] Clicking SEARCH ({result_count})...")
    search_button.click()
    wait_for_result_grid(page=page, vendor_id=vendor_id, timeout_ms=60_000)

    result_row = locate_result_row(page, vendor_id)
    headers = extract_result_headers(page, result_row)
    values = extract_result_values(result_row)
    headers, values = align_headers_and_values(headers, values)
    print_result(headers, values)

    inventory_status = find_status(headers, values)
    print(f"[INFO] Detected stone status: {inventory_status or 'UNKNOWN'}")

    if inventory_status not in {"A", "AVAILABLE"}:
        print(f"[RESULT] {vendor_id} unavailable; checking Memo List...")
        records = search_memo_list(
            page=page,
            vendor_id=vendor_id,
        )

        blocked_for = determine_blocked_for(records)

        if (
            inventory_status.upper() == "SM"
            and is_blocked_for_keyzar(records)
        ):
            print(
                f"[SUCCESS] {vendor_id} is already blocked for Keyzar. "
                "Appending it to the current Excel invoice."
            )

            report_path, row_added = append_stone_to_daily_report(
                order_date=_order_date(order),
                order_number=order.order_number,
                vendor_id=vendor_id,
                portal_headers=headers,
                portal_values=values,
            )

            print(
                f"[SUCCESS] Daily report: {report_path} "
                f"(row added: {row_added})"
            )

            return ProcessingResult(
                vendor_id,
                "ALREADY_BLOCKED_FOR_KEYZAR",
                inventory_status=inventory_status,
                details=blocked_for,
                report_path=str(report_path),
            )

        send_unavailable_alert(
            email_sender,
            order,
            vendor_id,
            inventory_status,
            records,
        )

        return ProcessingResult(
            vendor_id,
            "UNAVAILABLE_ALERTED",
            inventory_status=inventory_status,
            details=blocked_for,
        )

    print(f"[SUCCESS] {vendor_id} is available.")
    select_search_result_master_checkbox(page=page, result_row=result_row)
    page.wait_for_timeout(1_000)
    memo_issue_button = find_memo_issue_button(page)
    _click_memo_issue(page, memo_issue_button)

    page = browser.get_active_page()
    select_memo_row(page, vendor_id)
    note_field = find_note_field(page)
    note_field.fill(MEMO_NOTE)
    print(f"[INFO] Note entered: {MEMO_NOTE}")

    final_status = _save_or_dry_run(page)

    report_path, row_added = append_stone_to_daily_report(
        order_date=_order_date(order),
        order_number=order.order_number,
        vendor_id=vendor_id,
        portal_headers=headers,
        portal_values=values,
    )
    print(f"[SUCCESS] Daily report: {report_path} (row added: {row_added})")

    return ProcessingResult(
        vendor_id,
        final_status,
        inventory_status=inventory_status,
        details="Memo prepared" if BOT_DRY_RUN else "Memo saved",
        report_path=str(report_path),
    )
