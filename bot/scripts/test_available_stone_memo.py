from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import PENDING_REPORT_DIR, PORTAL_URL
from fenix.browser import FenixBrowser

from excel_reports.daily_report import (
    append_stone_to_daily_report,
)


TEST_VENDOR_ID = "LZ2086762"
MEMO_NOTE = "Keyzar"
DRY_RUN = True

SEARCH_RESULT_HEADERS = [
    "Select",
    "Add",
    "Information",
    "Stock ID",
    "Lab",
    "Report No",
    "Location",
    "Status",
    "Media",
    "Shape",
    "Carat",
    "Color",
    "Clarity",
    "Cut",
    "Polish",
    "Symmetry",
    "Fluorescence",
    "Rap",
    "Discount %",
    "Rate ($)",
    "Total ($)",
    "Length",
    "Width",
    "Depth",
    "Table %",
    "Depth %",
    "Ratio",
]


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", value).strip()


def find_stone_input(page):
    locator = page.locator("#txtStoneNo")
    if locator.count() == 0:
        locator = page.get_by_placeholder("Enter stone no.", exact=False)
    if locator.count() == 0:
        locator = page.locator(
            "textarea[placeholder*='stone' i], input[placeholder*='stone' i]"
        )
    if locator.count() == 0:
        raise RuntimeError("Could not locate the STONE NO field.")
    locator = locator.first
    locator.wait_for(state="visible", timeout=30_000)
    return locator


def find_search_button(page):
    locator = page.locator("#btnSearch")
    if locator.count() == 0:
        locator = page.locator(
            "button:has-text('SEARCH'), "
            "input[type='button'][value*='SEARCH' i], "
            "input[type='submit'][value*='SEARCH' i]"
        )
    if locator.count() == 0:
        raise RuntimeError("Could not locate the Search button.")
    return locator.first


def enter_stone_number(page, stone_input, vendor_id: str) -> None:
    stone_input.click()
    stone_input.fill("")
    stone_input.fill(vendor_id)

    stone_input.evaluate(
        """
        (element, value) => {
            element.value = value;
            element.dispatchEvent(new Event("input", { bubbles: true }));
            element.dispatchEvent(new Event("change", { bubbles: true }));
            element.dispatchEvent(new Event("blur", { bubbles: true }));

            if (typeof window.jQuery !== "undefined") {
                window.jQuery(element)
                    .val(value)
                    .trigger("input")
                    .trigger("change")
                    .trigger("blur");
            }
        }
        """,
        vendor_id,
    )

    page.wait_for_timeout(1_000)
    actual_value = stone_input.input_value().strip()
    print(f"[INFO] Stone field value registered as: {actual_value}")
    if actual_value.upper() != vendor_id.upper():
        raise RuntimeError("Fenix did not retain the entered Stone No.")


def select_all_stone_type(page) -> None:
    all_radio = page.locator(
        "input[type='radio'][name='rdbStoneType'][value='ALL']"
    ).first
    if all_radio.count() == 0:
        raise RuntimeError("Could not locate the ALL radio button.")

    page.evaluate(
        """
        () => {
            const radios = document.querySelectorAll(
                'input[type="radio"][name="rdbStoneType"]'
            );
            radios.forEach((radio) => {
                radio.checked = radio.value === "ALL";
            });

            const allRadio = document.querySelector(
                'input[type="radio"][name="rdbStoneType"][value="ALL"]'
            );
            if (!allRadio) return;

            allRadio.dispatchEvent(new Event("input", { bubbles: true }));
            allRadio.dispatchEvent(new Event("change", { bubbles: true }));

            if (typeof window.jQuery !== "undefined") {
                window.jQuery(allRadio)
                    .prop("checked", true)
                    .trigger("input")
                    .trigger("change");
            }
        }
        """
    )

    page.wait_for_timeout(500)
    selected_value = page.evaluate(
        """
        () => {
            const selected = document.querySelector(
                'input[type="radio"][name="rdbStoneType"]:checked'
            );
            return selected ? selected.value : null;
        }
        """
    )
    print(f"[INFO] Selected Stone No. option: {selected_value or 'NONE'}")
    if selected_value != "ALL":
        raise RuntimeError("The ALL radio option was not selected.")


def wait_until_search_ready(page, search_button, timeout_ms: int = 60_000) -> None:
    print("[INFO] Waiting for Fenix to prepare the stone search...")
    page.wait_for_function(
        """
        () => {
            const button = document.querySelector("#btnSearch");
            if (!button) return false;
            const text = (button.innerText || button.value || "")
                .replace(/\u00a0/g, " ");
            return /SEARCH\s*\(\s*[1-9]\d*\s*\)/i.test(text);
        }
        """,
        timeout=timeout_ms,
    )
    print(f"[INFO] Search is ready: {normalize_text(search_button.inner_text())}")


def wait_for_result_grid(page, vendor_id: str, timeout_ms: int = 60_000) -> None:
    print(f"[INFO] Waiting for {vendor_id} to appear in the Search Result grid...")
    page.wait_for_function(
        """
        (vendorId) => {
            const target = vendorId.trim().toUpperCase();
            return Array.from(document.querySelectorAll("a, td, span, div"))
                .some((element) => {
                    const text = (element.textContent || "")
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
    page.wait_for_timeout(2_000)
    print(f"[SUCCESS] {vendor_id} is visible in the Search Result grid.")


def locate_result_row(page, vendor_id: str):
    candidates = page.locator("a, td, span, div").filter(has_text=vendor_id)
    for index in range(candidates.count()):
        element = candidates.nth(index)
        try:
            if normalize_text(element.inner_text()).upper() != vendor_id.upper():
                continue
        except Exception:
            continue

        row = element.locator("xpath=ancestor::tr[1]")
        if row.count() > 0:
            return row.first

        row = element.locator("xpath=ancestor::*[@role='row'][1]")
        if row.count() > 0:
            return row.first

    raise RuntimeError(f"No exact Fenix result row was found for {vendor_id}.")


def extract_result_headers(page, result_row) -> list[str]:
    del page, result_row
    return SEARCH_RESULT_HEADERS.copy()


def extract_result_values(result_row) -> list[str]:
    cells = result_row.locator(":scope > td")
    if cells.count() == 0:
        cells = result_row.locator("[role='gridcell']")

    values = [normalize_text(cells.nth(i).inner_text()) for i in range(cells.count())]
    if not values:
        raise RuntimeError("The Fenix Search Result row contained no readable cells.")

    values = values[: len(SEARCH_RESULT_HEADERS)]
    while len(values) < len(SEARCH_RESULT_HEADERS):
        values.append("")
    return values


def align_headers_and_values(
    headers: list[str], values: list[str]
) -> tuple[list[str], list[str]]:
    headers = SEARCH_RESULT_HEADERS.copy()
    values = values[: len(headers)]
    while len(values) < len(headers):
        values.append("")
    return headers, values


def find_status(headers: list[str], values: list[str]) -> str:
    for index, header in enumerate(headers):
        if normalize_text(header).lower() in {"sts", "status", "stone status"}:
            return normalize_text(values[index]).upper()
    return ""


def print_result(headers: list[str], values: list[str]) -> None:
    print("\n" + "=" * 100)
    print("FENIX SEARCH RESULT")
    print("=" * 100)
    for index, value in enumerate(values):
        print(f"{index:02d} | {headers[index]:<25} | {value}")


def select_search_result_master_checkbox(page, result_row) -> None:
    table = result_row.locator("xpath=ancestor::table[1]").first
    if table.count() == 0:
        raise RuntimeError("Could not locate the Search Result table.")

    checkboxes = table.locator("input[type='checkbox']")
    if checkboxes.count() < 2:
        raise RuntimeError(
            "Expected the master checkbox and at least one result-row checkbox."
        )

    master_input = checkboxes.nth(0)
    row_checkbox = checkboxes.nth(1)

    print(f"[INFO] Search Result table checkboxes found: {checkboxes.count()}")
    print(f"[INFO] Master checkbox HTML: {master_input.evaluate('(el) => el.outerHTML')}")
    print(f"[INFO] Master initially checked: {master_input.is_checked()}")
    print(f"[INFO] Row initially checked: {row_checkbox.is_checked()}")
    print("[INFO] Clicking the top-left master checkbox...")

    # First try the real input.
    box = master_input.bounding_box()
    if box is not None:
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.wait_for_timeout(1_000)

    # The portal's visible wrapper is what has worked on this grid.
    if not row_checkbox.is_checked():
        wrapper = master_input.locator(
            "xpath=ancestor::*[self::label or contains(@class,'checkbox') "
            "or contains(@class,'check') or contains(@class,'custom') "
            "or contains(@class,'icheck') or contains(@class,'form-check')][1]"
        )
        if wrapper.count() > 0:
            wrapper_box = wrapper.first.bounding_box()
            if wrapper_box is not None:
                page.mouse.click(
                    wrapper_box["x"] + wrapper_box["width"] / 2,
                    wrapper_box["y"] + wrapper_box["height"] / 2,
                )
                page.wait_for_timeout(1_500)

    # Final fallback: click the row checkbox directly. Memo Issue only needs the row selected.
    if not row_checkbox.is_checked():
        try:
            row_checkbox.click(force=True, timeout=5_000)
        except Exception:
            row_checkbox.evaluate("(el) => el.click()")
        page.wait_for_timeout(1_000)

    master_checked = master_input.is_checked()
    row_checked = row_checkbox.is_checked()
    print(f"[INFO] Master checkbox checked: {master_checked}")
    print(f"[INFO] Result row selected: {row_checked}")

    if not row_checked:
        raise RuntimeError("Fenix did not select the result row.")

    print("[SUCCESS] Result row is selected and Memo Issue can be opened.")
    if not master_checked:
        print(
            "[INFO] The raw master input remains unchecked, but the Fenix row "
            "selection is active."
        )


def find_memo_issue_button(page):
    icon = page.locator(
        "[title*='Memo Issue' i], "
        "[aria-label*='Memo Issue' i], "
        "[data-original-title*='Memo Issue' i], "
        "[data-bs-original-title*='Memo Issue' i], "
        "i.fa-cube, i[class*='cube' i]"
    ).filter(visible=True)

    if icon.count() == 0:
        raise RuntimeError("Could not locate the Memo Issue cube icon.")

    icon = icon.first
    clickable = icon.locator(
        "xpath=ancestor::button[1] | ancestor::a[1] | "
        "ancestor::*[@role='button'][1] | ancestor::li[1]"
    ).first

    if clickable.count() == 0:
        raise RuntimeError("Memo Issue icon was found, but its clickable parent was not.")

    print(f"[INFO] Memo Issue clickable tag: {clickable.evaluate('(el) => el.tagName')}")
    return clickable


def click_memo_issue(page, memo_issue_button) -> None:
    print("[INFO] Clicking the Memo Issue button...")
    try:
        memo_issue_button.click(force=True, timeout=10_000)
    except Exception:
        print("[WARNING] Playwright click was intercepted; using native JavaScript click.")
        memo_issue_button.evaluate("(element) => element.click()")
    page.wait_for_timeout(5_000)


def select_memo_row(page, vendor_id: str) -> None:
    vendor_match = page.get_by_text(vendor_id, exact=True)
    if vendor_match.count() == 0:
        vendor_match = page.locator("a, td, span, div").filter(has_text=vendor_id)
    if vendor_match.count() == 0:
        raise RuntimeError(f"Could not find {vendor_id} on the Memo Issue page.")

    memo_row = vendor_match.first.locator("xpath=ancestor::tr[1]")
    if memo_row.count() == 0:
        memo_row = vendor_match.first.locator("xpath=ancestor::*[@role='row'][1]")
    if memo_row.count() == 0:
        raise RuntimeError("Could not identify the stone row on the Memo Issue page.")

    checkbox = memo_row.first.locator("input[type='checkbox']").first
    if checkbox.count() == 0:
        raise RuntimeError("Could not locate the Memo Issue row checkbox.")

    try:
        checkbox.click(force=True, timeout=5_000)
    except Exception:
        checkbox.evaluate("(el) => el.click()")
    page.wait_for_timeout(750)

    if not checkbox.is_checked():
        raise RuntimeError("The Memo Issue row checkbox was not selected.")
    print("[INFO] Memo Issue row selected.")


def find_note_field(page):
    selectors = [
        "#txtNote",
        "#txtNotes",
        "input[id*='note' i]",
        "textarea[id*='note' i]",
        "input[name*='note' i]",
        "textarea[name*='note' i]",
        "input[placeholder*='note' i]",
        "textarea[placeholder*='note' i]",
    ]
    for selector in selectors:
        locator = page.locator(f"{selector}:visible")
        if locator.count() > 0:
            return locator.first
    raise RuntimeError("Could not locate the Note field on the Memo Issue page.")


def find_save_button(page):
    selectors = [
        "button:has-text('SAVE')",
        "input[type='button'][value='SAVE' i]",
        "input[type='submit'][value='SAVE' i]",
        "[title='SAVE' i]",
        "[aria-label='SAVE' i]",
    ]
    for selector in selectors:
        locator = page.locator(f"{selector}:visible")
        if locator.count() > 0:
            return locator.first
    return None


def main() -> None:
    if not DRY_RUN:
        raise RuntimeError("Safety protection: DRY_RUN must remain True during this test.")

    browser = FenixBrowser()
    try:
        print("[INFO] Opening Fenix Search Stock...")
        page = browser.start()
        page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
        page.wait_for_timeout(3_000)
        page = browser.get_active_page()

        print(f"[INFO] Page title: {page.title()}")
        print(f"[INFO] Page URL  : {page.url}")

        stone_input = find_stone_input(page)
        search_button = find_search_button(page)

        print(f"[INFO] Entering stone number: {TEST_VENDOR_ID}")
        enter_stone_number(page, stone_input, TEST_VENDOR_ID)
        select_all_stone_type(page)
        wait_until_search_ready(page, search_button)

        page.wait_for_timeout(1_500)
        print("[INFO] Clicking SEARCH (1)...")
        search_button.click()
        wait_for_result_grid(page, TEST_VENDOR_ID)

        result_row = locate_result_row(page, TEST_VENDOR_ID)
        headers = extract_result_headers(page, result_row)
        values = extract_result_values(result_row)
        headers, values = align_headers_and_values(headers, values)
        print_result(headers, values)

        status = find_status(headers, values)
        print(f"\n[INFO] Detected stone status: {status or 'NOT FOUND'}")
        if status not in {"A", "AVAILABLE"}:
            raise RuntimeError(
                f"Stone {TEST_VENDOR_ID} is not available. "
                f"Detected status: {status or 'UNKNOWN'}."
            )

        print(f"[SUCCESS] {TEST_VENDOR_ID} is available.")
        excel_path, row_added = append_stone_to_daily_report(order_date=date.today(),order_number=None,vendor_id=TEST_VENDOR_ID,portal_headers=headers,portal_values=values)
        print(f"[SUCCESS] Search Result row written to:\n{excel_path}")

        select_search_result_master_checkbox(page, result_row)
        page.wait_for_timeout(1_500)

        memo_issue_button = find_memo_issue_button(page)
        click_memo_issue(page, memo_issue_button)
        page = browser.get_active_page()

        print(f"[INFO] Memo page title: {page.title()}")
        print(f"[INFO] Memo page URL  : {page.url}")

        select_memo_row(page, TEST_VENDOR_ID)
        note_field = find_note_field(page)
        note_field.fill(MEMO_NOTE)
        print(f"[INFO] Note entered: {MEMO_NOTE}")

        save_button = find_save_button(page)
        print("\n" + "=" * 100)
        print("DRY-RUN SAFETY STOP")
        print("=" * 100)
        if save_button is not None:
            print("[INFO] Save control located.")
        else:
            print("[WARNING] Save control was not identified.")
        print("[DRY RUN] The bot would click SAVE at this point.")
        print("[DRY RUN] SAVE WAS NOT CLICKED.")
        print("[DRY RUN] The stone has not been blocked by this script.")
        print("=" * 100)

        print()
        print(
            "[INFO] Background dry run completed."
        )

        print(
            "[INFO] The prepared memo was closed without saving."
        )

    except Exception as exc:
        print(f"[ERROR] Dry-run memo test failed: {exc}")
        raise
    finally:
        browser.stop()


if __name__ == "__main__":
    main()