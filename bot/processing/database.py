from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from config import DATA_DIR

DATABASE_PATH = DATA_DIR / "keyzar_jobs.db"


class JobDatabase:
    def __init__(self, database_path: Path = DATABASE_PATH) -> None:
        self.database_path = database_path
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self._create_tables()

    @contextmanager
    def connect(self):
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        try:
            yield connection
            connection.commit()
        finally:
            connection.close()

    def _create_tables(self) -> None:
        with self.connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    outlook_entry_id TEXT NOT NULL,
                    vendor_id TEXT NOT NULL,
                    order_number TEXT,
                    order_date TEXT,
                    email_subject TEXT,
                    status TEXT NOT NULL,
                    details TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(outlook_entry_id, vendor_id)
                )
                """
            )

    def is_finished(self, outlook_entry_id: str, vendor_id: str) -> bool:
        finished = {
            "COMPLETED",
            "DRY_RUN_PREPARED",
            "NOT_FOUND_ALERTED",
            "UNAVAILABLE_ALERTED",
        }
        with self.connect() as connection:
            row = connection.execute(
                "SELECT status FROM jobs WHERE outlook_entry_id=? AND vendor_id=?",
                (outlook_entry_id, vendor_id),
            ).fetchone()
        return bool(row and row["status"] in finished)

    def upsert(
        self,
        *,
        outlook_entry_id: str,
        vendor_id: str,
        order_number: str | None,
        order_date: str | None,
        email_subject: str,
        status: str,
        details: str = "",
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO jobs (
                    outlook_entry_id, vendor_id, order_number, order_date,
                    email_subject, status, details, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(outlook_entry_id, vendor_id) DO UPDATE SET
                    order_number=excluded.order_number,
                    order_date=excluded.order_date,
                    email_subject=excluded.email_subject,
                    status=excluded.status,
                    details=excluded.details,
                    updated_at=excluded.updated_at
                """,
                (
                    outlook_entry_id,
                    vendor_id,
                    order_number,
                    order_date,
                    email_subject,
                    status,
                    details,
                    now,
                    now,
                ),
            )
