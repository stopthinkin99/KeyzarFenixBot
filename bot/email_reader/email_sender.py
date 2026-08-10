from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pythoncom
import win32com.client


class OutlookEmailSender:
    OL_MAIL_ITEM = 0

    def send(
        self,
        recipients: str | Iterable[str],
        subject: str,
        body: str,
        cc: str | Iterable[str] | None = None,
        attachments: Iterable[str | Path] | None = None,
        send_from: str | None = None,
        send_store_id: str | None = None,
    ) -> None:
        recipient_text = self._join_addresses(recipients)
        cc_text = self._join_addresses(cc)

        if not recipient_text:
            raise ValueError("At least one recipient is required.")

        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            message = outlook.CreateItem(self.OL_MAIL_ITEM)

            message.To = recipient_text
            message.CC = cc_text
            message.Subject = subject
            message.Body = body

            sender_description = self._set_sender_account(
                outlook=outlook,
                message=message,
                send_store_id=send_store_id,
                send_from=send_from,
            )

            for attachment in attachments or []:
                attachment_path = Path(attachment).resolve()

                if not attachment_path.exists():
                    raise FileNotFoundError(
                        f"Attachment not found: {attachment_path}"
                    )

                message.Attachments.Add(str(attachment_path))

            message.Send()

            print(f"[EMAIL] Sent: {subject}")
            print(f"[EMAIL] From: {sender_description}")
            print(f"[EMAIL] To: {recipient_text}")
            print(f"[EMAIL] CC: {cc_text or 'None'}")

        except Exception as exc:
            raise RuntimeError(
                f"Outlook could not send the email: {exc}"
            ) from exc

        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _set_sender_account(
        *,
        outlook,
        message,
        send_store_id: str | None,
        send_from: str | None,
    ) -> str:
        session = outlook.Session
        target_store_id = (send_store_id or "").strip()
        target_address = (send_from or "").strip().lower()

        if target_store_id:
            for account in session.Accounts:
                try:
                    delivery_store = account.DeliveryStore

                    if (
                        delivery_store is not None
                        and str(getattr(delivery_store, "StoreID", "") or "")
                        == target_store_id
                    ):
                        message.SendUsingAccount = account

                        smtp_address = str(
                            getattr(account, "SmtpAddress", "") or ""
                        ).strip()

                        display_name = str(
                            getattr(account, "DisplayName", "") or ""
                        ).strip()

                        selected_name = (
                            smtp_address
                            or display_name
                            or "selected Outlook account"
                        )

                        print(
                            "[EMAIL] Using Outlook account matched "
                            f"to selected mailbox store: {selected_name}"
                        )

                        return selected_name

                except Exception:
                    continue

        if target_address:
            for account in session.Accounts:
                try:
                    smtp_address = str(
                        getattr(account, "SmtpAddress", "") or ""
                    ).strip().lower()

                    display_name = str(
                        getattr(account, "DisplayName", "") or ""
                    ).strip().lower()

                    if target_address in {smtp_address, display_name}:
                        message.SendUsingAccount = account

                        print(
                            "[EMAIL] Using Outlook account matched "
                            f"by address/name: {send_from}"
                        )

                        return (
                            smtp_address
                            or display_name
                            or str(send_from)
                        )

                except Exception:
                    continue

            message.SentOnBehalfOfName = send_from

            print(
                "[EMAIL] No direct Outlook account match. "
                "Using send-on-behalf address: "
                f"{send_from}"
            )

            return str(send_from)

        print(
            "[EMAIL] No selected Outlook mailbox was available. "
            "Using Outlook's default sending account."
        )

        return "Default Outlook account"

    @staticmethod
    def _join_addresses(
        addresses: str | Iterable[str] | None,
    ) -> str:
        if addresses is None:
            return ""

        if isinstance(addresses, str):
            return addresses.strip()

        return "; ".join(
            str(address).strip()
            for address in addresses
            if str(address).strip()
        )
