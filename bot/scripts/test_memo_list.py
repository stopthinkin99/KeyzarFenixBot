from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fenix.browser import FenixBrowser
from fenix.memo_list import (
    determine_blocked_for,
    search_memo_list,
)


# Replace this with a stone currently showing SM or another
# unavailable status in Search Stock.
TEST_VENDOR_ID = "LZ2282272"


def main() -> None:
    if TEST_VENDOR_ID == "REPLACE_WITH_UNAVAILABLE_STONE":
        raise RuntimeError(
            "Set TEST_VENDOR_ID in scripts\\test_memo_list.py "
            "to a stone that is currently unavailable or on memo."
        )

    browser = FenixBrowser()

    try:
        page = browser.start()

        records = search_memo_list(
            page=page,
            vendor_id=TEST_VENDOR_ID,
        )

        if not records:
            print()
            print(
                f"[RESULT] No created memo was found for "
                f"{TEST_VENDOR_ID}."
            )
            return

        blocked_for = determine_blocked_for(
            records
        )

        print()
        print(
            f"[RESULT] {TEST_VENDOR_ID} appears to be "
            f"blocked for: {blocked_for}"
        )

    finally:
        browser.stop()


if __name__ == "__main__":
    main()