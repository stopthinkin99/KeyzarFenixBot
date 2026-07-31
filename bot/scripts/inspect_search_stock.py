from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import PORTAL_URL
from fenix.browser import FenixBrowser


TEST_VENDOR_ID = "LZ1548978"


def print_inputs(page) -> None:
    inputs = page.locator("input")

    print()
    print(f"[INFO] Found {inputs.count()} input elements.")

    for index in range(inputs.count()):
        element = inputs.nth(index)

        try:
            print("-" * 70)
            print(f"INPUT #{index}")
            print(f"Type        : {element.get_attribute('type')}")
            print(f"Name        : {element.get_attribute('name')}")
            print(f"ID          : {element.get_attribute('id')}")
            print(
                f"Placeholder : "
                f"{element.get_attribute('placeholder')}"
            )
        except Exception as exc:
            print(
                f"[WARNING] Could not inspect input {index}: {exc}"
            )


def print_buttons(page) -> None:
    buttons = page.locator(
        "button, input[type='button'], input[type='submit']"
    )

    print()
    print(f"[INFO] Found {buttons.count()} button elements.")

    for index in range(buttons.count()):
        element = buttons.nth(index)

        try:
            text = element.inner_text().strip()

            print("-" * 70)
            print(f"BUTTON #{index}")
            print(f"Text  : {text}")
            print(f"Name  : {element.get_attribute('name')}")
            print(f"ID    : {element.get_attribute('id')}")
            print(f"Value : {element.get_attribute('value')}")
        except Exception as exc:
            print(
                f"[WARNING] Could not inspect button {index}: {exc}"
            )


def print_selects(page) -> None:
    selects = page.locator("select")

    print()
    print(f"[INFO] Found {selects.count()} select elements.")

    for index in range(selects.count()):
        element = selects.nth(index)

        try:
            print("-" * 70)
            print(f"SELECT #{index}")
            print(f"Name : {element.get_attribute('name')}")
            print(f"ID   : {element.get_attribute('id')}")

            options = element.locator("option")

            values = []

            for option_index in range(options.count()):
                option = options.nth(option_index)

                values.append(
                    {
                        "text": option.inner_text().strip(),
                        "value": option.get_attribute("value"),
                    }
                )

            print(f"Options: {values}")

        except Exception as exc:
            print(
                f"[WARNING] Could not inspect select {index}: {exc}"
            )


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

        print_inputs(page)
        print_selects(page)
        print_buttons(page)

        print()
        print(f"[INFO] Test Vendor ID: {TEST_VENDOR_ID}")
        print(
            "[INFO] No fields were changed and no search was run."
        )

        input(
            "\nPress ENTER after reviewing the browser and terminal: "
        )

    finally:
        browser.stop()


if __name__ == "__main__":
    main()