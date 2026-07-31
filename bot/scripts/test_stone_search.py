from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import PORTAL_URL
from fenix.browser import FenixBrowser


TEST_VENDOR_ID = "LZ1548978"


def main() -> None:
    browser = FenixBrowser()

    try:
        print("[INFO] Opening Fenix Search Stock...")

        page = browser.start()

        page.goto(
            PORTAL_URL,
            wait_until="domcontentloaded",
            timeout=60_000,
        )

        page.wait_for_timeout(3_000)
        page = browser.get_active_page()

        print(f"[INFO] Page title: {page.title()}")
        print(f"[INFO] Page URL  : {page.url}")

        # ----------------------------------------------------------
        # Locate the STONE NO field using its placeholder.
        # ----------------------------------------------------------
        stone_input = page.get_by_placeholder(
            "Enter stone no.",
            exact=False,
        )

        if stone_input.count() == 0:
            stone_input = page.locator(
                "textarea[placeholder*='stone' i], "
                "input[placeholder*='stone' i]"
            )

        if stone_input.count() == 0:
            raise RuntimeError(
                "Could not locate the STONE NO field."
            )

        stone_input = stone_input.first
        stone_input.wait_for(
            state="visible",
            timeout=30_000,
        )

        print("[INFO] STONE NO field located.")
        print(
            f"[INFO] Stone field tag: "
            f"{stone_input.evaluate('(el) => el.tagName')}"
        )
        print(
            f"[INFO] Stone field ID : "
            f"{stone_input.get_attribute('id')}"
        )
        print(
            f"[INFO] Placeholder    : "
            f"{stone_input.get_attribute('placeholder')}"
        )

        # ----------------------------------------------------------
        # Locate the ALL / PUBLISH / UNPUBLISH radio group.
        #
        # Based on the portal layout:
        #   Radio 0 = ALL
        #   Radio 1 = PUBLISH
        #   Radio 2 = UNPUBLISH
        # ----------------------------------------------------------
        visible_radios = page.locator(
            "input[type='radio']:visible"
        )

        radio_count = visible_radios.count()

        print(
            f"[INFO] Visible radio buttons found: "
            f"{radio_count}"
        )

        if radio_count < 3:
            raise RuntimeError(
                "Could not find the expected ALL, PUBLISH, "
                "and UNPUBLISH radio buttons."
            )

        all_radio = visible_radios.nth(0)

        print("[INFO] ALL radio button located.")
        print(
            f"[INFO] Radio name : "
            f"{all_radio.get_attribute('name')}"
        )
        print(
            f"[INFO] Radio ID   : "
            f"{all_radio.get_attribute('id')}"
        )
        print(
            f"[INFO] Radio value: "
            f"{all_radio.get_attribute('value')}"
        )

        # ----------------------------------------------------------
        # Fill Stone No. and select ALL.
        # ----------------------------------------------------------
        print(
            f"[INFO] Entering Vendor ID: "
            f"{TEST_VENDOR_ID}"
        )

        stone_input.fill(TEST_VENDOR_ID)

        all_radio.check(
            force=True,
        )

        print(
            f"[INFO] ALL selected: "
            f"{all_radio.is_checked()}"
        )

        # ----------------------------------------------------------
        # Locate and click the SEARCH button.
        # ----------------------------------------------------------
        search_button = page.locator("#btnSearch")

        if search_button.count() == 0:
            search_button = page.locator(
                "button:has-text('SEARCH'), "
                "input[type='button'][value*='SEARCH' i], "
                "input[type='submit'][value*='SEARCH' i]"
            )

        if search_button.count() == 0:
            raise RuntimeError(
                "Could not locate the SEARCH button."
            )

        search_button = search_button.first

        print("[INFO] Clicking SEARCH...")

        search_button.click()

        # Wait for the search result area to update.
        page.wait_for_timeout(5_000)

        print("[INFO] Search completed.")
        print(f"[INFO] Current URL: {page.url}")

        # ----------------------------------------------------------
        # Print visible tables after the search.
        # ----------------------------------------------------------
        visible_tables = page.locator(
            "table:visible"
        )

        table_count = visible_tables.count()

        print(
            f"[INFO] Visible tables after search: "
            f"{table_count}"
        )

        for index in range(table_count):
            table = visible_tables.nth(index)

            try:
                table_text = table.inner_text().strip()

                if not table_text:
                    continue

                print()
                print("=" * 90)
                print(f"VISIBLE TABLE #{index}")
                print("=" * 90)
                print(table_text[:8_000])

            except Exception as exc:
                print(
                    f"[WARNING] Could not inspect "
                    f"table {index}: {exc}"
                )

        # ----------------------------------------------------------
        # Search the page for the Vendor ID.
        # ----------------------------------------------------------
        vendor_matches = page.get_by_text(
            TEST_VENDOR_ID,
            exact=False,
        )

        vendor_match_count = vendor_matches.count()

        print()
        print(
            f"[INFO] Elements containing "
            f"{TEST_VENDOR_ID}: "
            f"{vendor_match_count}"
        )

        for index in range(
            min(vendor_match_count, 10)
        ):
            element = vendor_matches.nth(index)

            try:
                print()
                print("-" * 90)
                print(
                    f"VENDOR MATCH #{index}"
                )
                print("-" * 90)

                element_text = element.inner_text().strip()

                print(
                    f"Text: {element_text}"
                )

                parent_html = element.evaluate(
                    """
                    (el) => {
                        const row = el.closest("tr");

                        if (row) {
                            return row.outerHTML;
                        }

                        if (el.parentElement) {
                            return el.parentElement.outerHTML;
                        }

                        return el.outerHTML;
                    }
                    """
                )

                print("\nHTML")
                print(parent_html[:8_000])

            except Exception as exc:
                print(
                    f"[WARNING] Could not inspect "
                    f"vendor match {index}: {exc}"
                )

        # ----------------------------------------------------------
        # Print result rows containing the Vendor ID.
        # ----------------------------------------------------------
        result_rows = page.locator(
            f"tr:has-text('{TEST_VENDOR_ID}')"
        )

        result_row_count = result_rows.count()

        print()
        print(
            f"[INFO] Result rows containing "
            f"{TEST_VENDOR_ID}: "
            f"{result_row_count}"
        )

        for index in range(result_row_count):
            row = result_rows.nth(index)

            try:
                cells = row.locator(
                    "th, td"
                )

                print()
                print("=" * 90)
                print(f"RESULT ROW #{index}")
                print("=" * 90)

                for cell_index in range(
                    cells.count()
                ):
                    cell_text = (
                        cells.nth(cell_index)
                        .inner_text()
                        .strip()
                    )

                    print(
                        f"Column {cell_index}: "
                        f"{cell_text}"
                    )

            except Exception as exc:
                print(
                    f"[WARNING] Could not inspect "
                    f"result row {index}: {exc}"
                )

        print()
        print("[SAFE] No stone was selected.")
        print("[SAFE] No salesman memo was created.")
        print("[SAFE] Nothing was saved or blocked.")

        input(
            "\nPress ENTER to close the browser... "
        )

    except Exception as exc:
        print(f"[ERROR] Stone search failed: {exc}")
        raise

    finally:
        browser.stop()


if __name__ == "__main__":
    main()