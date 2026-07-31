from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from typing import Iterator

import pythoncom
import win32com.client

from config import DATA_DIR

OUTLOOK_SETTINGS_PATH = DATA_DIR / "outlook_settings.json"


@dataclass
class OutlookEmail:
    entry_id: str
    sender_name: str
    sender_email: str
    subject: str
    received_time: datetime | None
    body: str
    html_body: str


@dataclass
class OutlookMailbox:
    store_id: str
    display_name: str
    inbox_name: str
    total_items: int
    unread_items: int


def load_saved_mailbox() -> dict[str, str]:
    if not OUTLOOK_SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(OUTLOOK_SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {
        "store_id": str(data.get("store_id", "") or ""),
        "display_name": str(data.get("display_name", "") or ""),
    }


def save_outlook_mailbox(*, store_id: str, display_name: str) -> None:
    OUTLOOK_SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTLOOK_SETTINGS_PATH.write_text(
        json.dumps(
            {"store_id": store_id, "display_name": display_name},
            indent=2,
        ),
        encoding="utf-8",
    )


class OutlookReader:
    """Read mail from a selected Classic Outlook mailbox."""

    OL_FOLDER_INBOX = 6
    OL_MAIL_ITEM = 43

    def __init__(self, *, store_id: str | None = None, display_name: str | None = None) -> None:
        saved = load_saved_mailbox()
        self.requested_store_id = store_id if store_id is not None else saved.get("store_id", "")
        self.requested_display_name = (
            display_name if display_name is not None else saved.get("display_name", "")
        )
        self._outlook = None
        self._namespace = None
        self._store = None
        self._inbox = None

    def connect(self) -> None:
        pythoncom.CoInitialize()
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            self._select_mailbox()
        except Exception as exc:
            self.disconnect()
            raise RuntimeError(
                "Could not connect to Classic Outlook. Make sure Classic Outlook "
                "is installed, open, and signed in for this Windows user."
            ) from exc

    def disconnect(self) -> None:
        self._inbox = None
        self._store = None
        self._namespace = None
        self._outlook = None
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _iter_stores(self):
        if self._namespace is None:
            raise RuntimeError("OutlookReader is not connected.")
        stores = self._namespace.Stores
        for index in range(1, stores.Count + 1):
            yield stores.Item(index)

    def _select_mailbox(self) -> None:
        if self._namespace is None:
            raise RuntimeError("OutlookReader is not connected.")

        selected_store = None
        if self.requested_store_id:
            for store in self._iter_stores():
                try:
                    if str(store.StoreID) == self.requested_store_id:
                        selected_store = store
                        break
                except Exception:
                    continue

        if selected_store is None and self.requested_display_name:
            requested = self.requested_display_name.strip().lower()
            for store in self._iter_stores():
                try:
                    if str(store.DisplayName).strip().lower() == requested:
                        selected_store = store
                        break
                except Exception:
                    continue

        if selected_store is not None:
            self._store = selected_store
            self._inbox = selected_store.GetDefaultFolder(self.OL_FOLDER_INBOX)
            return

        self._inbox = self._namespace.GetDefaultFolder(self.OL_FOLDER_INBOX)
        try:
            self._store = self._inbox.Store
        except Exception:
            self._store = None

    @property
    def mailbox_display_name(self) -> str:
        if self._store is not None:
            try:
                return str(self._store.DisplayName or "")
            except Exception:
                pass
        return "Default Outlook Inbox"

    @property
    def mailbox_store_id(self) -> str:
        if self._store is None:
            return ""
        try:
            return str(self._store.StoreID or "")
        except Exception:
            return ""

    def get_inbox(self):
        if self._inbox is None:
            raise RuntimeError("OutlookReader is not connected.")
        return self._inbox

    def list_mailboxes(self) -> list[OutlookMailbox]:
        mailboxes: list[OutlookMailbox] = []
        for store in self._iter_stores():
            try:
                inbox = store.GetDefaultFolder(self.OL_FOLDER_INBOX)
            except Exception:
                continue
            try:
                total_items = int(inbox.Items.Count)
            except Exception:
                total_items = 0
            try:
                unread_items = int(inbox.UnReadItemCount)
            except Exception:
                unread_items = 0
            mailboxes.append(
                OutlookMailbox(
                    store_id=str(getattr(store, "StoreID", "") or ""),
                    display_name=str(getattr(store, "DisplayName", "") or ""),
                    inbox_name=str(getattr(inbox, "Name", "Inbox") or "Inbox"),
                    total_items=total_items,
                    unread_items=unread_items,
                )
            )
        return mailboxes

    def iter_recent_emails(self, limit: int = 50) -> Iterator[OutlookEmail]:
        items = self.get_inbox().Items
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

    def get_recent_email_preview(self, limit: int = 10) -> list[OutlookEmail]:
        return list(self.iter_recent_emails(limit))

    def _convert_mail_item(self, item) -> OutlookEmail:
        try:
            received_time = item.ReceivedTime
        except Exception:
            received_time = None
        return OutlookEmail(
            entry_id=str(getattr(item, "EntryID", "") or ""),
            sender_name=str(getattr(item, "SenderName", "") or ""),
            sender_email=self._get_sender_email(item),
            subject=str(getattr(item, "Subject", "") or ""),
            received_time=received_time,
            body=str(getattr(item, "Body", "") or ""),
            html_body=str(getattr(item, "HTMLBody", "") or ""),
        )

    @staticmethod
    def _get_sender_email(item) -> str:
        try:
            if str(getattr(item, "SenderEmailType", "") or "").upper() == "EX":
                sender = item.Sender
                if sender is not None:
                    exchange_user = sender.GetExchangeUser()
                    if exchange_user is not None and exchange_user.PrimarySmtpAddress:
                        return str(exchange_user.PrimarySmtpAddress)
            return str(getattr(item, "SenderEmailAddress", "") or "")
        except Exception:
            return str(getattr(item, "SenderEmailAddress", "") or "")
