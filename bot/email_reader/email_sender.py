from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pythoncom
import win32com.client


REQUIRED_BOT_SENDER = "sales@fenixdiamonds.com"


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

        # Keyzar bot mail must never silently fall back to Nishit's default
        # Outlook account. If a caller does not provide a sender, use the
        # Fenix Sales mailbox explicitly.
        requested_sender = (send_from or REQUIRED_BOT_SENDER).strip()

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
                send_from=requested_sender,
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
            print(f"[EMAIL] Requested From: {requested_sender}")
            print(f"[EMAIL] Sender mode: {sender_description}")
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
        send_from: str,
    ) -> str:
        """Force the intended sender instead of Outlook's default account.

        If sales@fenixdiamonds.com is a real Account in the Outlook profile,
        use SendUsingAccount. If it is a shared mailbox/store, use
        SentOnBehalfOfName. Exchange must grant the logged-in user Send As
        (preferred) or Send on Behalf permission for the shared mailbox.
        """
        session = outlook.Session
        target_address = send_from.strip().lower()
        target_store_id = (send_store_id or "").strip()

        # First and safest: exact SMTP address match. Never select another
        # account merely because it is Outlook's default.
        for account in session.Accounts:
            try:
                smtp_address = str(
                    getattr(account, "SmtpAddress", "") or ""
                ).strip().lower()

                if smtp_address == target_address:
                    message.SendUsingAccount = account
                    message.SentOnBehalfOfName = send_from
                    return f"SendUsingAccount={smtp_address}"
            except Exception:
                continue

        # Secondary check: selected mailbox store may map to a real Account,
        # but only accept it if that account is also the requested sender.
        if target_store_id:
            for account in session.Accounts:
                try:
                    delivery_store = account.DeliveryStore
                    store_id = str(
                        getattr(delivery_store, "StoreID", "") or ""
                    )
                    smtp_address = str(
                        getattr(account, "SmtpAddress", "") or ""
                    ).strip().lower()

                    if (
                        store_id == target_store_id
                        and smtp_address == target_address
                    ):
                        message.SendUsingAccount = account
                        message.SentOnBehalfOfName = send_from
                        return f"SendUsingAccount={smtp_address}"
                except Exception:
                    continue

        # Shared mailbox case. Do NOT silently use the default Nishit address
        # as the visible sender. Exchange will honor this only when mailbox
        # delegation includes Send As / Send on Behalf permission.
        message.SentOnBehalfOfName = send_from
        message.Save()

        resolved_sender = str(
            getattr(message, "SentOnBehalfOfName", "") or ""
        ).strip()

        if resolved_sender.lower() != target_address:
            raise RuntimeError(
                "Outlook did not accept sales@fenixdiamonds.com as the "
                "message sender. Confirm that the mailbox is visible in "
                "Classic Outlook and that this Windows/Outlook user has "
                "Send As permission for the sales mailbox."
            )

        return f"SentOnBehalfOfName={resolved_sender}"

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
