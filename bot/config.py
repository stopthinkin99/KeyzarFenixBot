from pathlib import Path
import os

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


# -------------------------------------------------------------------
# Project folders
# -------------------------------------------------------------------

BROWSER_PROFILE_DIR = BASE_DIR / "playwright_profile"

DATA_DIR = BASE_DIR / "data"
LOG_DIR = DATA_DIR / "logs"
SCREENSHOT_DIR = DATA_DIR / "screenshots"
PENDING_REPORT_DIR = DATA_DIR / "pending_reports"
SENT_REPORT_DIR = DATA_DIR / "sent_reports"


# -------------------------------------------------------------------
# Outlook email settings
# -------------------------------------------------------------------

KEYZAR_SENDER_KEYWORDS = [
    "@keyzarjewelry.com",
    "keyzar",
]

KEYZAR_SUBJECT_KEYWORDS = [
    "new stone order",
]

OUTLOOK_EMAIL_LIMIT = int(os.getenv("OUTLOOK_EMAIL_LIMIT", "30"))


# -------------------------------------------------------------------
# Fenix Portal
# -------------------------------------------------------------------

PORTAL_URL = os.getenv(
    "PORTAL_URL",
    "https://admin.fenixdiamonds.com/Search/SearchStock",
)

BROWSER_CHANNEL = os.getenv(
    "BROWSER_CHANNEL",
    "msedge",
)

FENIX_HEADLESS = (
    os.getenv("FENIX_HEADLESS", "true")
    .strip()
    .lower()
    in {"1", "true", "yes", "on"}
)

FENIX_SLOW_MO_MS = 250

FENIX_STORAGE_STATE = BASE_DIR / "data" / "fenix_storage_state.json"

ALERT_RECIPIENT = os.getenv(
    "ALERT_RECIPIENT",
    "aayan.boradia@unidesignusa.com",
).strip()

BOT_DRY_RUN = (
    os.getenv("BOT_DRY_RUN", "true")
    .strip()
    .lower()
    in {"1", "true", "yes", "on"}
)

MEMO_LIST_URL = (
    "https://admin.fenixdiamonds.com/Memo/List"
)

POLL_SECONDS = int(os.getenv("POLL_SECONDS", "300"))

FENIX_USERNAME = os.getenv(
    "FENIX_USERNAME",
    "",
).strip()

FENIX_PASSWORD = os.getenv(
    "FENIX_PASSWORD",
    "",
)
