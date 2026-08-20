from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pythoncom
import win32com.client


# The Keyzar bot must NEVER fall back to Nishit's Outlook account.
REQUIRED_BOT_SENDER = "sales@fenixdiamonds.com"

# Outlook constant: olFolderSentMail
OL_FOLDER_SENT_MAIL = 5


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

            sender = self._set_sender_account(
                outlook=outlook,
                message=message,
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

            print(
                "[EMAIL] Final sender before Send(): "
                f"{sender}"
            )

            message.Send()

            print(
                f"[EMAIL] Sent successfully: {subject}"
            )
            print(
                f"[EMAIL] From: {sender}"
            )
            print(
                f"[EMAIL] To: {recipient_text}"
            )
            print(
                f"[EMAIL] CC: {cc_text or 'None'}"
            )

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
    ) -> str:
        """
        Force the real sales@fenixdiamonds.com Outlook Account.

        There is intentionally NO fallback to:
        - Nishit's account
        - Outlook default account
        - SentOnBehalfOfName
        """

        session = outlook.Session
        target = REQUIRED_BOT_SENDER.lower()

        print(
            "[EMAIL] Required Outlook account: "
            f"{target}"
        )

        for account in session.Accounts:
            try:
                smtp_address = str(
                    getattr(
                        account,
                        "SmtpAddress",
                        "",
                    )
                    or ""
                ).strip().lower()

                display_name = str(
                    getattr(
                        account,
                        "DisplayName",
                        "",
                    )
                    or ""
                ).strip()

                print(
                    "[EMAIL] Available Outlook account: "
                    f"display={display_name}, "
                    f"smtp={smtp_address}"
                )

                if smtp_address != target:
                    continue

                #
                # Equivalent to explicitly choosing
                # sales@fenixdiamonds.com in Outlook's
                # From dropdown.
                #
                message.SendUsingAccount = account

                print(
                    "[EMAIL] SendUsingAccount explicitly "
                    f"set to {smtp_address}"
                )

                #
                # Explicitly store the sent copy under
                # the Sales account's Sent Items.
                #
                delivery_store = account.DeliveryStore

                if delivery_store is None:
                    raise RuntimeError(
                        "sales@fenixdiamonds.com was found "
                        "but Outlook returned no DeliveryStore."
                    )

                sent_folder = (
                    delivery_store.GetDefaultFolder(
                        OL_FOLDER_SENT_MAIL
                    )
                )

                message.SaveSentMessageFolder = (
                    sent_folder
                )

                try:
                    folder_path = (
                        sent_folder.FolderPath
                    )
                except Exception:
                    folder_path = (
                        "Sales Sent Items"
                    )

                print(
                    "[EMAIL] SaveSentMessageFolder set to: "
                    f"{folder_path}"
                )

                return smtp_address

            except RuntimeError:
                raise

            except Exception as exc:
                print(
                    "[EMAIL] Could not inspect "
                    f"Outlook account: {exc}"
                )

        raise RuntimeError(
            "The bot refused to send because the "
            "required Outlook account "
            "'sales@fenixdiamonds.com' could not "
            "be found. No fallback account was used."
        )

    @staticmethod
    def _join_addresses(
        addresses: str | Iterable[str] | None,
    ) -> str:
        if addresses is None:
            return ""

        if isinstance(
            addresses,
            str,
        ):
            return addresses.strip()

        return "; ".join(
            str(address).strip()
            for address in addresses
            if str(address).strip()
        )