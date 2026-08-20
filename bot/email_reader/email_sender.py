@staticmethod
def _set_sender_account(
    *,
    outlook,
    message,
    send_store_id: str | None,
    send_from: str | None,
) -> str:
    session = outlook.Session

    target_address = (
        (send_from or "")
        .strip()
        .lower()
    )

    if not target_address:
        raise RuntimeError(
            "No Outlook sending mailbox was selected."
        )

    print(
        "[EMAIL] Looking for exact Outlook sending account: "
        f"{target_address}"
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
                "[EMAIL] Outlook account found: "
                f"display={display_name}, smtp={smtp_address}"
            )

            if smtp_address == target_address:
                message.SendUsingAccount = account

                print(
                    "[EMAIL] EXACT sending account selected: "
                    f"{smtp_address}"
                )

                return smtp_address

        except Exception as exc:
            print(
                "[EMAIL] Could not inspect Outlook account: "
                f"{exc}"
            )

    raise RuntimeError(
        "The selected Outlook mailbox could not be found "
        "as a sending account. "
        f"Requested sender: {target_address}"
    )