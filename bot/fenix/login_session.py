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
    and save the authenticated cookies/local storage for future runs.
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

        active_pages = [
            open_page
            for open_page in context.pages
            if not open_page.is_closed()
        ]

        _log(
            log_callback,
            f"Open Fenix pages detected: {len(active_pages)}",
        )

        for index, open_page in enumerate(
            active_pages,
            start=1,
        ):
            try:
                _log(
                    log_callback,
                    f"Page {index}: {open_page.url}",
                )
            except Exception:
                pass

        authenticated_page = next(
            (
                open_page
                for open_page in active_pages
                if "/search/searchstock" in open_page.url.lower()
            ),
            None,
        )

        # Otherwise accept any Fenix page that is no longer the login page.
        if authenticated_page is None:
            authenticated_page = next(
                (
                    open_page
                    for open_page in active_pages
                    if (
                        "admin.fenixdiamonds.com" in open_page.url.lower()
                        and "/login" not in open_page.url.lower()
                    )
                ),
                None,
            )

        if authenticated_page is None:
            context.close()
            raise RuntimeError(
                "No authenticated Fenix page was detected. "
                "Please complete the login and open Search Stock "
                "before clicking Yes."
            )

        _log(
            log_callback,
            f"Authenticated Fenix page detected: "
            f"{authenticated_page.url}",
        )

        context.storage_state(
            path=str(FENIX_STORAGE_STATE)
        )


        _log(
            log_callback,
            "Fenix login session saved successfully.",
        )

        context.close()
