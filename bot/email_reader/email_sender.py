REQUIRED_SENDER = "sales@fenixdiamonds.com"

OL_FOLDER_SENT_MAIL = 5

@staticmethod
def _set_sender_account(
    *,
    outlook,
    message,
    send_store_id: str | None = None,
    send_from: str | None = None,
) -> str:
    """
    Force all bot-generated emails to use the real
    sales@fenixdiamonds.com Outlook Account.

    No fallback to Nishit's account is allowed.
    """

    session = outlook.Session

    required_sender = REQUIRED_SENDER.lower()

    print(
        "[EMAIL] Required Outlook sending account: "
        f"{required_sender}"
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
                "[EMAIL] Inspecting Outlook account: "
                f"display={display_name}, "
                f"smtp={smtp_address}"
            )

            if smtp_address != required_sender:
                continue

            #
            # IMPORTANT:
            # This is the programmatic equivalent of
            # manually clicking sales@fenixdiamonds.com
            # again in Outlook's From field.
            #
            message.SendUsingAccount = account

            print(
                "[EMAIL] SendUsingAccount explicitly set to: "
                f"{smtp_address}"
            )

            #
            # ALSO explicitly tell Outlook where the
            # sent copy must be stored.
            #
            try:
                delivery_store = account.DeliveryStore

                if delivery_store is None:
                    raise RuntimeError(
                        "The Sales Outlook account does not "
                        "have a DeliveryStore."
                    )

                sent_folder = delivery_store.GetDefaultFolder(
                    OL_FOLDER_SENT_MAIL
                )

                message.SaveSentMessageFolder = sent_folder

                print(
                    "[EMAIL] Sent Items folder explicitly set to: "
                    f"{sent_folder.FolderPath}"
                )

            except Exception as exc:
                raise RuntimeError(
                    "Sales account was found, but Outlook "
                    "could not select its Sent Items folder: "
                    f"{exc}"
                ) from exc

            return smtp_address

        except RuntimeError:
            raise

        except Exception as exc:
            print(
                "[EMAIL] Could not inspect Outlook account: "
                f"{exc}"
            )

    #
    # NO FALLBACK.
    #
    raise RuntimeError(
        "Cannot send this email because Outlook does not "
        "contain the required sending account "
        "'sales@fenixdiamonds.com'. "
        "The bot refused to fall back to another account."
    )