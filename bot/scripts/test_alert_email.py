from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config import ALERT_RECIPIENT
from email_reader.email_sender import OutlookEmailSender


def main() -> None:
    sender = OutlookEmailSender()

    sender.send(
        recipients=ALERT_RECIPIENT,
        subject="Keyzar Fenix Bot – Test Alert",
        body=(
            "This is a test email from the Keyzar Fenix Bot.\n\n"
            "The Outlook email-sending connection is working."
        ),
    )


if __name__ == "__main__":
    main()