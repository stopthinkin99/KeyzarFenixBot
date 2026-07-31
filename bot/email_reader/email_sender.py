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
    ) -> None:
        recipient_text = self._join_addresses(
            recipients
        )

        cc_text = self._join_addresses(cc)

        if not recipient_text:
            raise ValueError(
                "At least one recipient is required."
            )

        pythoncom.CoInitialize()

        try:
            outlook = win32com.client.Dispatch(
                "Outlook.Application"
            )

            message = outlook.CreateItem(
                self.OL_MAIL_ITEM
            )

            message.To = recipient_text
            message.CC = cc_text
            message.Subject = subject
            message.Body = body

            if send_from:
                self._set_sender_account(
                    outlook=outlook,
                    message=message,
                    send_from=send_from,
                )

            for attachment in attachments or []:
                attachment_path = Path(
                    attachment
                ).resolve()

                if not attachment_path.exists():
                    raise FileNotFoundError(
                        f"Attachment not found: "
                        f"{attachment_path}"
                    )

                message.Attachments.Add(
                    str(attachment_path)
                )

            message.Send()

            print(
                f"[EMAIL] Sent: {subject}"
            )
            print(
                f"[EMAIL] From: "
                f"{send_from or 'Default Outlook account'}"
            )
            print(
                f"[EMAIL] To: {recipient_text}"
            )
            print(
                f"[EMAIL] CC: {cc_text or 'None'}"
            )

        except Exception as exc:
            raise RuntimeError(
                f"Outlook could not send the email: "
                f"{exc}"
            ) from exc

        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _set_sender_account(
        outlook,
        message,
        send_from: str,
    ) -> None:
        target_address = send_from.strip().lower()
        matched_account = None

        session = outlook.Session

        for account in session.Accounts:
            smtp_address = (
                getattr(
                    account,
                    "SmtpAddress",
                    "",
                )
                or ""
            ).strip().lower()

            display_name = (
                getattr(
                    account,
                    "DisplayName",
                    "",
                )
                or ""
            ).strip().lower()

            if target_address in {
                smtp_address,
                display_name,
            }:
                matched_account = account
                break

        if matched_account is not None:
            message.SendUsingAccount = (
                matched_account
            )

            print(
                f"[EMAIL] Outlook account found: "
                f"{send_from}"
            )
            return

        # Shared mailbox fallback.
        message.SentOnBehalfOfName = send_from

        print(
            f"[EMAIL] Using shared mailbox/"
            f"send-on-behalf address: {send_from}"
        )

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