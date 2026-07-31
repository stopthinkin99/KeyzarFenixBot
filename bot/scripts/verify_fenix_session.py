from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import PORTAL_URL
from fenix.browser import FenixBrowser


def is_login_page(url: str) -> bool:
    url_lower = url.lower()

    return (
        "/login" in url_lower
        or "/account/login" in url_lower
        or "signin" in url_lower
        or "sign-in" in url_lower
    )


def main() -> None:
    browser = FenixBrowser()

    try:
        print("[INFO] Starting Fenix session verification...")

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

        if is_login_page(page.url):
            print("[FAILED] The stored Fenix session is not authenticated.")
            print("[ACTION] Run scripts\\login_once.py again.")
            raise SystemExit(1)

        if "searchstock" in page.url.lower().replace("/", ""):
            print("[SUCCESS] The saved Fenix session is working.")
        else:
            print(
                "[WARNING] Login appears valid, but the browser is not "
                "on the Search Stock page."
            )

        input("\nPress ENTER to close the browser... ")

    except Exception as exc:
        print(f"[ERROR] Verification failed: {exc}")
        raise SystemExit(1)

    finally:
        browser.stop()


if __name__ == "__main__":
    main()