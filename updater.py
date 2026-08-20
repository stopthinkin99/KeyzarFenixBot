from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


GITHUB_OWNER = "stopthinkin99"
GITHUB_REPO = "KeyzarFenixBot"
GITHUB_BRANCH = "main"

REMOTE_BOT_FOLDER = "bot"

APP_DIR = Path(sys.executable).resolve().parent

LOCAL_BOT_DIR = APP_DIR / "bot"

PROTECTED_NAMES = {
    ".env",
    "data",
    "playwright_profile",
    "__pycache__",
    "fenix_storage_state.json",
    "keyzar_jobs.db",
}


def _is_protected(relative_path: Path) -> bool:
    return any(
        part in PROTECTED_NAMES
        for part in relative_path.parts
    )


def _download_zip_with_powershell(
    url: str,
    destination: Path,
) -> None:
    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        (
            "$ErrorActionPreference='Stop'; "
            f"Invoke-WebRequest "
            f"-Uri '{url}' "
            f"-OutFile '{destination}' "
            "-UseBasicParsing"
        ),
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        creationflags=0x08000000,
    )

    if result.returncode != 0:
        error = (
            result.stderr.strip()
            or result.stdout.strip()
            or "Unknown PowerShell download error."
        )

        raise RuntimeError(
            f"GitHub ZIP download failed: {error}"
        )


def sync_from_github() -> tuple[bool, str]:
    """
    Download the newest repository ZIP and replace only the bot folder.

    Local data, credentials, browser profiles and databases remain untouched.
    """

    LOCAL_BOT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    zip_url = (
        f"https://github.com/"
        f"{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/archive/"
        f"refs/heads/{GITHUB_BRANCH}.zip"
    )

    try:
        with tempfile.TemporaryDirectory(
            prefix="keyzar_update_"
        ) as temp_dir:
            temp_path = Path(temp_dir)

            zip_path = (
                temp_path
                / "repo.zip"
            )

            extract_dir = (
                temp_path
                / "extracted"
            )

            _download_zip_with_powershell(
                zip_url,
                zip_path,
            )

            with zipfile.ZipFile(
                zip_path,
                "r",
            ) as archive:
                archive.extractall(
                    extract_dir
                )

            repo_root = (
                extract_dir
                / (
                    f"{GITHUB_REPO}-"
                    f"{GITHUB_BRANCH}"
                )
            )

            source_bot = (
                repo_root
                / REMOTE_BOT_FOLDER
            )

            if not source_bot.exists():
                raise RuntimeError(
                    "Downloaded repository does not "
                    "contain the bot folder."
                )

            synced_files: list[str] = []

            for source_file in (
                source_bot.rglob("*")
            ):
                if not source_file.is_file():
                    continue

                relative_path = (
                    source_file.relative_to(
                        source_bot
                    )
                )

                if _is_protected(
                    relative_path
                ):
                    continue

                destination = (
                    LOCAL_BOT_DIR
                    / relative_path
                )

                destination.parent.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                shutil.copy2(
                    source_file,
                    destination,
                )

                synced_files.append(
                    str(relative_path)
                )

            return (
                True,
                (
                    f"Synced "
                    f"{len(synced_files)} "
                    f"file(s) from GitHub ZIP."
                ),
            )

    except Exception as exc:
        return (
            False,
            (
                "GitHub update was unavailable. "
                "Using the installed files. "
                f"Reason: {exc}"
            ),
        )