from __future__ import annotations

from collections.abc import Callable
from playwright.sync_api import sync_playwright

from config import BROWSER_CHANNEL, BROWSER_PROFILE_DIR, FENIX_STORAGE_STATE, PORTAL_URL
from fenix.browser import (
    PASSWORD_SELECTORS,
    SUBMIT_SELECTORS,
    USERNAME_SELECTORS,
    _first_visible_locator,
)


def _log(callback: Callable[[str], None] | None, message: str) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def save_fenix_credentials_and_session(
    *,
    username: str,
    password: str,
    log_callback: Callable[[str], None] | None = None,
) -> None:
    save_fenix_credentials(username, password)

    BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    FENIX_STORAGE_STATE.parent.mkdir(parents=True, exist_ok=True)

    _log(
        log_callback,
        "Fenix credentials saved securely in Windows Credential Manager.",
    )

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE_DIR),
            channel=BROWSER_CHANNEL,
            headless=False,
            no_viewport=True,
            args=[
                "--start-maximized",
                "--disable-notifications",
                "--disable-popup-blocking",
            ],
        )

        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
            page.wait_for_timeout(2_000)

            if "/login" in page.url.lower():
                username_field = _first_visible_locator(page, USERNAME_SELECTORS)
                password_field = _first_visible_locator(page, PASSWORD_SELECTORS)

                if username_field is None or password_field is None:
                    raise RuntimeError(
                        "Could not find the Fenix username/password fields."
                    )

                username_field.fill(username)
                password_field.fill(password)

                submit_button = _first_visible_locator(page, SUBMIT_SELECTORS)
                if submit_button is not None:
                    submit_button.click(force=True)
                else:
                    password_field.press("Enter")

                try:
                    page.wait_for_url(
                        lambda url: "/login" not in url.lower(),
                        timeout=45_000,
                    )
                except Exception:
                    page.wait_for_timeout(3_000)

            if "/login" in page.url.lower():
                raise RuntimeError(
                    "Fenix login failed. Check the username and password."
                )

            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
            context.storage_state(path=str(FENIX_STORAGE_STATE))

            _log(
                log_callback,
                "Fenix automatic login tested and session saved successfully.",
            )
        finally:
            context.close()


def save_fenix_login_session(
    *,
    confirmation_callback: Callable[[], bool],
    log_callback: Callable[[str], None] | None = None,
) -> None:
    BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    FENIX_STORAGE_STATE.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE_DIR),
            channel=BROWSER_CHANNEL,
            headless=False,
            no_viewport=True,
            args=[
                "--start-maximized",
                "--disable-notifications",
                "--disable-popup-blocking",
            ],
        )

        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)

            if not confirmation_callback():
                raise RuntimeError("Fenix login was cancelled.")

            context.storage_state(path=str(FENIX_STORAGE_STATE))
            _log(log_callback, "Fenix login session saved successfully.")
        finally:
            context.close()
