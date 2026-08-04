from __future__ import annotations

import win32cred

CREDENTIAL_TARGET = "KeyzarFenixBot/FenixPortal"


def save_fenix_credentials(username: str, password: str) -> None:
    username = username.strip()
    if not username:
        raise ValueError("Fenix username is required.")
    if not password:
        raise ValueError("Fenix password is required.")

    win32cred.CredWrite(
        {
            "Type": win32cred.CRED_TYPE_GENERIC,
            "TargetName": CREDENTIAL_TARGET,
            "UserName": username,
            "CredentialBlob": password.encode("utf-16-le"),
            "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
            "Comment": "Saved Fenix login for Keyzar Fenix Bot",
        },
        0,
    )


def load_fenix_credentials() -> tuple[str, str] | None:
    try:
        credential = win32cred.CredRead(
            CREDENTIAL_TARGET,
            win32cred.CRED_TYPE_GENERIC,
            0,
        )
    except Exception:
        return None

    username = str(credential.get("UserName", "") or "").strip()
    blob = credential.get("CredentialBlob", b"")

    if isinstance(blob, bytes):
        try:
            password = blob.decode("utf-16-le")
        except UnicodeDecodeError:
            password = blob.decode("utf-8", errors="ignore")
    else:
        password = str(blob or "")

    if not username or not password:
        return None

    return username, password


def delete_fenix_credentials() -> bool:
    try:
        win32cred.CredDelete(
            CREDENTIAL_TARGET,
            win32cred.CRED_TYPE_GENERIC,
            0,
        )
        return True
    except Exception:
        return False
