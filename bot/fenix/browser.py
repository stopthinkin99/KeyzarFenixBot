from __future__ import annotations

from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    sync_playwright,
)

from config import (
    BROWSER_CHANNEL,
    FENIX_HEADLESS,
    FENIX_STORAGE_STATE,
)


class FenixBrowser:
    """
    Start an automated Fenix browser session.

    Normal bot runs use headless mode, so no browser window appears.
    The saved Playwright storage state provides the authenticated
    Fenix cookies and local-storage values.
    """

    def __init__(self) -> None:
        self.playwright: Playwright | None = None
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    def start(self) -> Page:
        if not FENIX_STORAGE_STATE.exists():
            raise RuntimeError(
                "The saved Fenix login state was not found. "
                "Run scripts\\login_once.py before starting the bot."
            )

        print(
            "[INFO] Starting Fenix browser in "
            f"{'background' if FENIX_HEADLESS else 'visible'} mode..."
        )

        self.playwright = sync_playwright().start()

        self.browser = self.playwright.chromium.launch(
            channel=BROWSER_CHANNEL,
            headless=FENIX_HEADLESS,
            args=[
                "--disable-notifications",
                "--disable-popup-blocking",
                "--disable-background-timer-throttling",
                "--disable-backgrounding-occluded-windows",
                "--disable-renderer-backgrounding",
            ],
        )

        self.context = self.browser.new_context(
            storage_state=str(FENIX_STORAGE_STATE),

            # A fixed viewport is important in headless mode.
            # It ensures the full desktop version of Fenix is rendered.
            viewport={
                "width": 1920,
                "height": 1080,
            },

            screen={
                "width": 1920,
                "height": 1080,
            },

            accept_downloads=True,
        )

        self.page = self.context.new_page()

        self.page.set_default_timeout(
            60_000
        )

        self.page.set_default_navigation_timeout(
            90_000
        )

        return self.page

    def get_active_page(self) -> Page:
        if self.context is None:
            raise RuntimeError(
                "The Fenix browser has not been started."
            )

        open_pages = [
            page
            for page in self.context.pages
            if not page.is_closed()
        ]

        if not open_pages:
            raise RuntimeError(
                "No active Fenix browser page was found."
            )

        self.page = open_pages[-1]

        self.page.set_default_timeout(
            60_000
        )

        return self.page

    def save_storage_state(self) -> None:
        """
        Refresh the saved cookies/local storage after a successful run.
        """

        if self.context is None:
            return

        FENIX_STORAGE_STATE.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.context.storage_state(
            path=str(FENIX_STORAGE_STATE)
        )

    def stop(self) -> None:
        if self.context is not None:
            try:
                self.save_storage_state()
            except Exception as exc:
                print(
                    "[WARNING] Could not refresh the saved "
                    f"Fenix session: {exc}"
                )

            try:
                self.context.close()
            except Exception:
                pass

        if self.browser is not None:
            try:
                self.browser.close()
            except Exception:
                pass

        if self.playwright is not None:
            try:
                self.playwright.stop()
            except Exception:
                pass

        self.page = None
        self.context = None
        self.browser = None
        self.playwright = None