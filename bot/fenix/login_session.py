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
    Open a visible persistent Edge window and save its authentication
    state after the user confirms that login is complete.

    This intentionally matches the original working login_once.py flow.
    It does not rely on the browser URL because Fenix can leave the URL
    showing /login/index even after the authenticated portal is visible.
    """

    BROWSER_PROFILE_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FENIX_STORAGE_STATE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    _log(
        log_callback,
        "Starting the Fenix login browser...",
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
            page = (
                context.pages[0]
                if context.pages
                else context.new_page()
            )

            _log(
                log_callback,
                "Opening the Fenix portal...",
            )

            page.goto(
                PORTAL_URL,
                wait_until="domcontentloaded",
                timeout=90_000,
            )

            _log(
                log_callback,
                (
                    "Log in to Fenix and open Search Stock. "
                    "Then return to the app and click Yes."
                ),
            )

            if not confirmation_callback():
                raise RuntimeError(
                    "Fenix login was cancelled."
                )

            _log(
                log_callback,
                "Saving the authenticated Fenix session...",
            )

            # Save all cookies and local-storage values from the persistent
            # browser context. This is the same operation used by the
            # original working login_once.py script.
            context.storage_state(
                path=str(FENIX_STORAGE_STATE)
            )

            if not FENIX_STORAGE_STATE.exists():
                raise RuntimeError(
                    "The Fenix session file could not be created."
                )

            file_size = FENIX_STORAGE_STATE.stat().st_size

            if file_size == 0:
                raise RuntimeError(
                    "The Fenix session file was created but is empty."
                )

            _log(
                log_callback,
                "Fenix login session saved successfully.",
            )

            _log(
                log_callback,
                f"Saved session file: {FENIX_STORAGE_STATE}",
            )

        finally:
            context.close()