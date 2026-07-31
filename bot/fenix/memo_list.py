from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from typing import Any

from playwright.sync_api import Page

from config import MEMO_LIST_URL


@dataclass
class MemoListRecord:
    memo_number: str
    memo_date: str
    customer: str
    salesman: str
    memo_type: str
    total_pieces: str
    total_carats: str
    service_location: str
    note: str
    status: str

    def to_dict(self) -> dict[str, str]:
        return asdict(self)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def open_memo_list(page: Page) -> None:
    print("[INFO] Opening Fenix Memo List...")

    page.goto(
        MEMO_LIST_URL,
        wait_until="domcontentloaded",
        timeout=90_000,
    )

    page.wait_for_timeout(3_000)

    print(f"[INFO] Memo List title: {page.title()}")
    print(f"[INFO] Memo List URL  : {page.url}")

    if "/login" in page.url.lower():
        raise RuntimeError(
            "The Fenix login session has expired. "
            "Run scripts\\login_once.py again."
        )

    if "/memo/list" not in page.url.lower():
        raise RuntimeError(
            "Fenix did not open the expected Memo List page."
        )


def find_stone_number_input(page: Page):
    selectors = [
        "#txtStoneNo",
        "input[placeholder*='Stone No' i]",
        "textarea[placeholder*='Stone No' i]",
        "input[placeholder*='stone' i]",
        "textarea[placeholder*='stone' i]",
    ]

    for selector in selectors:
        locator = page.locator(f"{selector}:visible")

        if locator.count() > 0:
            return locator.first

    raise RuntimeError(
        "Could not locate the STONE NO field on Memo List."
    )


def ensure_created_selected(page: Page) -> None:
    created_text = page.get_by_text(
        "Created",
        exact=True,
    )

    if created_text.count() == 0:
        print(
            "[WARNING] Created status control was not found. "
            "Using the portal default."
        )
        return

    control = created_text.first

    class_name = (
        control.get_attribute("class") or ""
    ).lower()

    already_selected = any(
        keyword in class_name
        for keyword in (
            "active",
            "selected",
            "success",
            "green",
        )
    )

    if not already_selected:
        try:
            control.click(force=True)
        except Exception:
            control.evaluate("(element) => element.click()")

        page.wait_for_timeout(500)

    print("[INFO] Created memo status selected.")


def enter_stone_number(
    page: Page,
    stone_input,
    vendor_id: str,
) -> None:
    print(
        f"[INFO] Entering Memo List Stone No.: "
        f"{vendor_id}"
    )

    stone_input.click()
    stone_input.fill("")
    stone_input.fill(vendor_id)

    stone_input.evaluate(
        """
        (element, value) => {
            element.value = value;

            element.dispatchEvent(
                new Event("input", { bubbles: true })
            );

            element.dispatchEvent(
                new Event("change", { bubbles: true })
            );

            element.dispatchEvent(
                new Event("blur", { bubbles: true })
            );

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

    page.wait_for_timeout(750)

    actual_value = normalize_text(
        stone_input.input_value()
    )

    print(
        f"[INFO] Memo List Stone No. registered as: "
        f"{actual_value}"
    )

    if actual_value.upper() != vendor_id.upper():
        raise RuntimeError(
            "Fenix Memo List did not retain the Vendor ID."
        )


def find_search_button(page: Page):
    selectors = [
        "#btnSearch",
        "button:has-text('SEARCH')",
        "input[type='button'][value*='SEARCH' i]",
        "input[type='submit'][value*='SEARCH' i]",
    ]

    for selector in selectors:
        locator = page.locator(f"{selector}:visible")

        if locator.count() > 0:
            return locator.first

    raise RuntimeError(
        "Could not locate the Memo List Search button."
    )


def wait_for_memo_search_completion(
    page: Page,
    timeout_ms: int = 15_000,
) -> None:
    """
    Wait for any numeric Memo No. to appear.

    Fenix's Memo List is not always rendered as a normal HTML table,
    so this searches all leaf elements instead of only table rows.
    """

    print("[INFO] Checking the loaded Memo List rows...")

    try:
        page.wait_for_function(
            """
            () => {
                const normalize = (value) =>
                    (value || "")
                        .replace(/\\s+/g, " ")
                        .trim();

                return Array
                    .from(document.querySelectorAll("*"))
                    .some((element) => {
                        if (element.children.length > 0) {
                            return false;
                        }

                        const text = normalize(
                            element.textContent
                        );

                        return /^\\d{8,}$/.test(text);
                    });
            }
            """,
            timeout=timeout_ms,
        )

        print("[SUCCESS] Memo List result detected.")

    except Exception:
        print(
            "[WARNING] Memo number was not detected within "
            f"{timeout_ms // 1000} seconds."
        )

    page.wait_for_timeout(500)

def scroll_result_grid_to_right(page: Page) -> None:
    """
    Scroll every horizontally scrollable grid container to the right.

    This reveals Service Location, Note, and Status when the grid
    uses lazy rendering.
    """

    print(
        "[INFO] Scrolling Memo List grid to the Note column..."
    )

    page.evaluate(
        """
        () => {
            const elements = Array.from(
                document.querySelectorAll("*")
            );

            for (const element of elements) {
                if (
                    element.scrollWidth >
                    element.clientWidth + 10
                ) {
                    element.scrollLeft =
                        element.scrollWidth;
                }
            }
        }
        """
    )

    page.wait_for_timeout(1_000)

def extract_memo_records_from_dom(
    page: Page,
) -> list[MemoListRecord]:
    """
    Extract Memo List rows without assuming Fenix uses a normal table.

    It finds numeric memo-number cells, locates their nearest row-like
    container, and reads every cell in that row.
    """

    raw_rows = page.evaluate(
        """
        () => {
            const normalize = (value) =>
                (value || "")
                    .replace(/\\s+/g, " ")
                    .trim();

            const memoElements = Array
                .from(document.querySelectorAll("*"))
                .filter((element) => {
                    if (element.children.length > 0) {
                        return false;
                    }

                    const text = normalize(
                        element.textContent
                    );

                    return /^\\d{8,}$/.test(text);
                });

            const results = [];
            const seenRows = new Set();

            for (const memoElement of memoElements) {
                let row = memoElement.closest(
                    [
                        "tr",
                        "[role='row']",
                        ".ag-row",
                        ".dx-row",
                        ".k-master-row",
                        ".k-table-row",
                        ".jqgrow"
                    ].join(",")
                );

                // Generic fallback for custom Fenix grid containers.
                if (!row) {
                    let parent =
                        memoElement.parentElement;

                    while (
                        parent &&
                        parent !== document.body
                    ) {
                        const possibleCells =
                            parent.querySelectorAll(
                                [
                                    ":scope > td",
                                    ":scope > [role='gridcell']",
                                    ":scope > .ag-cell",
                                    ":scope > .dx-cell",
                                    ":scope > .k-table-td",
                                    ":scope > div"
                                ].join(",")
                            );

                        if (
                            possibleCells.length >= 8 &&
                            normalize(
                                parent.textContent
                            ).includes(
                                normalize(
                                    memoElement.textContent
                                )
                            )
                        ) {
                            row = parent;
                            break;
                        }

                        parent =
                            parent.parentElement;
                    }
                }

                if (!row || seenRows.has(row)) {
                    continue;
                }

                seenRows.add(row);

                let cells = Array.from(
                    row.querySelectorAll(
                        [
                            ":scope > td",
                            ":scope > [role='gridcell']",
                            ":scope > .ag-cell",
                            ":scope > .dx-cell",
                            ":scope > .k-table-td"
                        ].join(",")
                    )
                );

                if (cells.length < 8) {
                    cells = Array.from(
                        row.querySelectorAll(
                            [
                                "td",
                                "[role='gridcell']",
                                ".ag-cell",
                                ".dx-cell",
                                ".k-table-td"
                            ].join(",")
                        )
                    );
                }

                const values = cells.map(
                    (cell) => normalize(
                        cell.innerText ||
                        cell.textContent
                    )
                );

                if (values.length >= 8) {
                    results.push(values);
                }
            }

            return results;
        }
        """
    )

    print(
        f"[INFO] Raw Memo List rows detected: "
        f"{len(raw_rows)}"
    )

    records: list[MemoListRecord] = []

    for row_index, values in enumerate(
        raw_rows,
        start=1,
    ):
        print()
        print(
            f"[DEBUG] Memo row #{row_index} "
            f"has {len(values)} columns"
        )

        for column_index, value in enumerate(
            values
        ):
            print(
                f"[DEBUG] Column {column_index}: "
                f"{value}"
            )

        # Expected Fenix Memo List order:
        #
        # 0 Select
        # 1 Edit
        # 2 Memo No
        # 3 Memo Date
        # 4 Customer
        # 5 Salesman
        # 6 Memo Type
        # 7 Total PCS
        # 8 Total CTS
        # 9 Service Location
        # 10 Note
        # 11 Status

        while len(values) < 12:
            values.append("")

        memo_number = normalize_text(
            values[2]
        )

        if not re.fullmatch(
            r"\d{8,}",
            memo_number,
        ):
            continue

        records.append(
            MemoListRecord(
                memo_number=memo_number,
                memo_date=normalize_text(
                    values[3]
                ),
                customer=normalize_text(
                    values[4]
                ),
                salesman=normalize_text(
                    values[5]
                ),
                memo_type=normalize_text(
                    values[6]
                ),
                total_pieces=normalize_text(
                    values[7]
                ),
                total_carats=normalize_text(
                    values[8]
                ),
                service_location=normalize_text(
                    values[9]
                ),
                note=normalize_text(
                    values[10]
                ),
                status=normalize_text(
                    values[11]
                ),
            )
        )

    return records


def extract_cell_text(cell) -> str:
    """
    Read a cell even when it is horizontally outside the visible viewport.
    """

    try:
        return normalize_text(
            cell.inner_text()
        )
    except Exception:
        return normalize_text(
            cell.text_content()
        )





def search_memo_list(
    page: Page,
    vendor_id: str,
) -> list[MemoListRecord]:
    open_memo_list(page)

    stone_input = find_stone_number_input(page)

    ensure_created_selected(page)

    enter_stone_number(
        page=page,
        stone_input=stone_input,
        vendor_id=vendor_id,
    )

    search_button = find_search_button(page)

    print("[INFO] Clicking Memo List Search...")

    search_button.click(force=True)
 
    page.wait_for_timeout(1_500)

    wait_for_memo_search_completion(page=page, timeout_ms=15_000)

    scroll_result_grid_to_right(page)

    records = extract_memo_records_from_dom(page)

    print(
        f"[INFO] Memo List records found: "
        f"{len(records)}"
    )

    for index, record in enumerate(
        records,
        start=1,
    ):
        print()
        print("=" * 80)
        print(f"MEMO RECORD #{index}")
        print("=" * 80)
        print(f"Memo No.         : {record.memo_number}")
        print(f"Memo Date        : {record.memo_date}")
        print(f"Customer         : {record.customer}")
        print(f"Salesman         : {record.salesman}")
        print(f"Memo Type        : {record.memo_type}")
        print(f"Total Pieces     : {record.total_pieces}")
        print(f"Total Carats     : {record.total_carats}")
        print(f"Service Location : {record.service_location}")
        print(f"Note             : {record.note}")
        print(f"Status           : {record.status}")

    return records


def determine_blocked_for(
    records: list[MemoListRecord],
) -> str:
    for record in records:
        note = normalize_text(record.note)

        if note:
            return note

    for record in records:
        customer = normalize_text(record.customer)

        if customer:
            return customer

    for record in records:
        salesman = normalize_text(record.salesman)

        if salesman:
            return salesman

    return "Unknown customer"