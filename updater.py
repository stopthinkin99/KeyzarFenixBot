from __future__ import annotations

import shutil
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path


GITHUB_OWNER = "stopthinkin99"
GITHUB_REPO = "KeyzarFenixBot"
GITHUB_BRANCH = "main"
REMOTE_BOT_FOLDER = "bot"

APP_DIR = Path(__file__).resolve().parent
LOCAL_BOT_DIR = APP_DIR / "bot"

PROTECTED_NAMES = {
    ".env",
    "data",
    "playwright_profile",
    "__pycache__",
    "fenix_storage_state.json",
    "keyzar_jobs.db",
}

USER_AGENT = "KeyzarFenixBot-Updater/2.0"


def _is_protected(relative_path: Path) -> bool:
    return any(
        part in PROTECTED_NAMES
        for part in relative_path.parts
    )


def _download_repository_zip(
    destination: Path,
) -> None:
    url = (
        f"https://codeload.github.com/"
        f"{GITHUB_OWNER}/{GITHUB_REPO}/"
        f"zip/refs/heads/{GITHUB_BRANCH}"
    )

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        },
    )

    with urllib.request.urlopen(
        request,
        timeout=90,
    ) as response:
        with destination.open("wb") as output_file:
            shutil.copyfileobj(
                response,
                output_file,
            )


def _copy_bot_folder(
    extracted_root: Path,
    synced_files: list[str],
) -> None:
    repository_folder = (
        extracted_root
        / f"{GITHUB_REPO}-{GITHUB_BRANCH}"
    )

    remote_bot_dir = (
        repository_folder
        / REMOTE_BOT_FOLDER
    )

    if not remote_bot_dir.exists():
        raise RuntimeError(
            "The downloaded repository did not contain "
            f"the '{REMOTE_BOT_FOLDER}' folder."
        )

    LOCAL_BOT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    for source_path in remote_bot_dir.rglob("*"):
        if not source_path.is_file():
            continue

        relative_path = source_path.relative_to(
            remote_bot_dir
        )

        if _is_protected(relative_path):
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
            source_path,
            destination,
        )

        synced_files.append(
            str(relative_path)
        )


def sync_from_github() -> tuple[bool, str]:
    synced_files: list[str] = []

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            zip_path = (
                temp_path
                / "repository.zip"
            )
            extract_path = (
                temp_path
                / "extracted"
            )

            _download_repository_zip(
                zip_path
            )

            extract_path.mkdir(
                parents=True,
                exist_ok=True,
            )

            with zipfile.ZipFile(
                zip_path,
                "r",
            ) as archive:
                archive.extractall(
                    extract_path
                )

            _copy_bot_folder(
                extracted_root=extract_path,
                synced_files=synced_files,
            )

        return (
            True,
            (
                f"Synced {len(synced_files)} "
                "file(s) from GitHub."
            ),
        )

    except (
        urllib.error.URLError,
        urllib.error.HTTPError,
        TimeoutError,
        OSError,
        RuntimeError,
        zipfile.BadZipFile,
    ) as exc:
        return (
            False,
            (
                "GitHub update was unavailable. "
                "Using the installed files. "
                f"Reason: {exc}"
            ),
        )