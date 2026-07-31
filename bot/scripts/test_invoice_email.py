from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from email_reader.invoice_sender import (
    send_keyzar_invoice,
)


if __name__ == "__main__":
    send_keyzar_invoice(
        processing_date=date.today()
    )