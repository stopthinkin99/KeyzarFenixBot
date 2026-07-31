from __future__ import annotations

from collections.abc import Callable

from playwright.sync_api import sync_playwright

from config import (
    BROWSER_CHANNEL,
    BROWSER_PROFILE_DIR,
    FENIX_STORAGE_STATE,
    PORTAL_URL,
)


def _log(
    callback: Callable[[str], None] | None,
    message: str,
) -> None:
    if callback:
        callback(message)
    else:
        print(message)


def save_fenix_login_session(
    *,
    confirmation_callback: Callable[[], bool],
    log_callback: Callable[[str], None] | None = None,
) -> None:
    """
    Open a visible Edge window, let the user log in manually,
    verify the authenticated session by opening Search Stock,
    and save the cookies/local storage for future background runs.
    """

    BROWSER_PROFILE_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FENIX_STORAGE_STATE.parent.mkdir(
        parents=True,
        exist_ok=True,
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

        page = (
            context.pages[0]
            if context.pages
            else context.new_page()
        )

        _log(
            log_callback,
            "Opening the Fenix login window...",
        )

        page.goto(
            PORTAL_URL,
            wait_until="domcontentloaded",
            timeout=90_000,
        )

        if not confirmation_callback():
            context.close()

            raise RuntimeError(
                "Fenix login was cancelled."
            )

        _log(
            log_callback,
            "Verifying the Fenix login by opening Search Stock...",
        )

        active_pages = [
            open_page
            for open_page in context.pages
            if not open_page.is_closed()
        ]

        verification_page = (
            active_pages[-1]
            if active_pages
            else context.new_page()
        )

        try:
            verification_page.goto(
                PORTAL_URL,
                wait_until="domcontentloaded",
                timeout=90_000,
            )

            verification_page.wait_for_timeout(
                3_000
            )

            current_url = verification_page.url

            _log(
                log_callback,
                f"Fenix verification URL: {current_url}",
            )

            if "/login" in current_url.lower():
                raise RuntimeError(
                    "Fenix redirected back to the login page. "
                    "Please complete the login inside the Edge "
                    "window opened by the app before clicking Yes."
                )

            if (
                "/search/searchstock"
                not in current_url.lower()
            ):
                raise RuntimeError(
                    "Fenix login may be complete, but Search Stock "
                    "did not open successfully. "
                    f"Current page: {current_url}"
                )

            context.storage_state(
                path=str(FENIX_STORAGE_STATE)
            )

            _log(
                log_callback,
                "Fenix login session saved successfully.",
            )

            _log(
                log_callback,
                f"Saved session file: "
                f"{FENIX_STORAGE_STATE}",
            )

        finally:
            context.close()