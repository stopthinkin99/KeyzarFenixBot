from __future__ import annotations

from playwright.sync_api import Browser, BrowserContext, Locator, Page, Playwright, sync_playwright

from config import BROWSER_CHANNEL, FENIX_HEADLESS, FENIX_PASSWORD, FENIX_STORAGE_STATE, FENIX_USERNAME, PORTAL_URL


USERNAME_SELECTORS = [
    "#txtUserName", "#txtUsername", "#username", "#email",
    "input[name='username']", "input[name='UserName']",
    "input[name='email']", "input[type='email']",
    "input[autocomplete='username']", "input[type='text']",
]

PASSWORD_SELECTORS = [
    "#txtPassword", "#password", "input[name='password']",
    "input[name='Password']", "input[type='password']",
    "input[autocomplete='current-password']",
]

SUBMIT_SELECTORS = [
    "#btnLogin", "button[type='submit']", "input[type='submit']",
    "button:has-text('LOGIN')", "button:has-text('Login')",
    "button:has-text('Sign in')", "a:has-text('LOGIN')",
]


def _first_visible_locator(page: Page, selectors: list[str]) -> Locator | None:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible():
                return locator
        except Exception:
            continue
    return None


class FenixBrowser:
    def __init__(self) -> None:
        self.playwright: Playwright | None = None
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    def start(self) -> Page:
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

        options = {
            "viewport": {"width": 1920, "height": 1080},
            "screen": {"width": 1920, "height": 1080},
            "accept_downloads": True,
        }

        if FENIX_STORAGE_STATE.exists():
            options["storage_state"] = str(FENIX_STORAGE_STATE)

        self.context = self.browser.new_context(**options)
        self.page = self.context.new_page()
        self.page.set_default_timeout(60_000)
        self.page.set_default_navigation_timeout(90_000)

        self.ensure_logged_in()
        return self.page

    def ensure_logged_in(self) -> None:
        if self.page is None:
            raise RuntimeError("The Fenix browser has not been started.")

        print("[INFO] Checking the saved Fenix login...")
        self.page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
        self.page.wait_for_timeout(2_000)

        if "/login" not in self.page.url.lower():
            print("[INFO] Existing Fenix session is valid.")
            self.save_storage_state()
            return

        print("[INFO] Saved Fenix session expired. Attempting automatic login...")
        username = FENIX_USERNAME
        password = FENIX_PASSWORD

        if not username or not password:
            raise RuntimeError(
                "Fenix credentials are missing from local .env file "
                "'Save Fenix Login' once to securely store the credentials."
            )

        username, password = credentials
        username_field = _first_visible_locator(self.page, USERNAME_SELECTORS)
        password_field = _first_visible_locator(self.page, PASSWORD_SELECTORS)

        if username_field is None:
            raise RuntimeError("Automatic Fenix login could not find the username field.")
        if password_field is None:
            raise RuntimeError("Automatic Fenix login could not find the password field.")

        username_field.fill(username)
        password_field.fill(password)

        submit_button = _first_visible_locator(self.page, SUBMIT_SELECTORS)
        if submit_button is not None:
            submit_button.click(force=True)
        else:
            password_field.press("Enter")

        try:
            self.page.wait_for_url(
                lambda url: "/login" not in url.lower(),
                timeout=45_000,
            )
        except Exception:
            self.page.wait_for_timeout(3_000)

        if "/login" in self.page.url.lower():
            raise RuntimeError(
                "Automatic Fenix login failed. The saved username or password "
                "may be incorrect, or the portal may require another step."
            )

        print("[SUCCESS] Fenix automatic login completed.")
        self.page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=90_000)
        self.page.wait_for_timeout(1_500)

        if "/login" in self.page.url.lower():
            raise RuntimeError("Fenix redirected back to the login page.")

        self.save_storage_state()

    def get_active_page(self) -> Page:
        if self.context is None:
            raise RuntimeError("The Fenix browser has not been started.")

        open_pages = [page for page in self.context.pages if not page.is_closed()]
        if not open_pages:
            raise RuntimeError("No active Fenix browser page was found.")

        self.page = open_pages[-1]
        self.page.set_default_timeout(60_000)
        return self.page

    def save_storage_state(self) -> None:
        if self.context is None:
            return
        FENIX_STORAGE_STATE.parent.mkdir(parents=True, exist_ok=True)
        self.context.storage_state(path=str(FENIX_STORAGE_STATE))

    def stop(self) -> None:
        if self.context is not None:
            try:
                self.save_storage_state()
            except Exception as exc:
                print(f"[WARNING] Could not refresh the saved Fenix session: {exc}")
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
