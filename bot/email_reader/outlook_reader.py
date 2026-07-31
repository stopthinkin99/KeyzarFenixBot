from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Iterator

import pythoncom
import win32com.client


@dataclass
class OutlookEmail:
    entry_id: str
    sender_name: str
    sender_email: str
    subject: str
    received_time: datetime | None
    body: str
    html_body: str


class OutlookReader:
    """Read emails from the Classic Outlook desktop application."""

    OL_FOLDER_INBOX = 6
    OL_MAIL_ITEM = 43

    def __init__(self) -> None:
        self._outlook = None
        self._namespace = None

    def connect(self) -> None:
        """Connect to the Outlook MAPI session."""

        pythoncom.CoInitialize()

        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
        except Exception as exc:
            raise RuntimeError(
                "Could not connect to Classic Outlook. "
                "Make sure Classic Outlook is installed, open, and signed in."
            ) from exc

    def disconnect(self) -> None:
        self._namespace = None
        self._outlook = None
        pythoncom.CoUninitialize()

    def get_inbox(self):
        if self._namespace is None:
            raise RuntimeError("OutlookReader is not connected.")

        return self._namespace.GetDefaultFolder(self.OL_FOLDER_INBOX)

    def iter_recent_emails(self, limit: int = 50) -> Iterator[OutlookEmail]:
        inbox = self.get_inbox()
        items = inbox.Items

        # Newest emails first.
        items.Sort("[ReceivedTime]", True)

        count = min(items.Count, limit)

        for index in range(1, count + 1):
            item = items.Item(index)

            try:
                if getattr(item, "Class", None) != self.OL_MAIL_ITEM:
                    continue

                yield self._convert_mail_item(item)

            except Exception as exc:
                print(f"[WARNING] Could not read Outlook item {index}: {exc}")

    def _convert_mail_item(self, item) -> OutlookEmail:
        sender_email = self._get_sender_email(item)

        received_time = None
        try:
            received_time = item.ReceivedTime
        except Exception:
            pass

        return OutlookEmail(
            entry_id=str(getattr(item, "EntryID", "") or ""),
            sender_name=str(getattr(item, "SenderName", "") or ""),
            sender_email=sender_email,
            subject=str(getattr(item, "Subject", "") or ""),
            received_time=received_time,
            body=str(getattr(item, "Body", "") or ""),
            html_body=str(getattr(item, "HTMLBody", "") or ""),
        )

    @staticmethod
    def _get_sender_email(item) -> str:
        """Resolve the sender address for Exchange and regular SMTP emails."""

        try:
            sender_email_type = str(
                getattr(item, "SenderEmailType", "") or ""
            ).upper()

            if sender_email_type == "EX":
                sender = item.Sender

                if sender is not None:
                    exchange_user = sender.GetExchangeUser()

                    if exchange_user is not None:
                        smtp_address = exchange_user.PrimarySmtpAddress

                        if smtp_address:
                            return str(smtp_address)

            return str(getattr(item, "SenderEmailAddress", "") or "")

        except Exception:
            return str(getattr(item, "SenderEmailAddress", "") or "")