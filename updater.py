from __future__ import annotations

import json
import shutil
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path


GITHUB_OWNER = "stopthinkin99"
GITHUB_REPO = "KeyzarFenixBot"
GITHUB_BRANCH = "main"
REMOTE_BOT_FOLDER = "bot"

# updater.py is always kept beside KeyzarFenixBot.exe.
# Using __file__ makes it update the folder containing this updater,
# whether it is called by the EXE or tested manually with Python.
APP_DIR = Path(__file__).resolve().parent

LOCAL_BOT_DIR = APP_DIR / "bot"

# These files and folders must always stay local.
PROTECTED_NAMES = {
    ".env",
    "data",
    "playwright_profile",
    "__pycache__",
    "fenix_storage_state.json",
    "keyzar_jobs.db",
}

USER_AGENT = "KeyzarFenixBot-Updater/1.0"


def _request_json(url: str):
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": USER_AGENT,
        },
    )

    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def _download_file(url: str, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)

    request = urllib.request.Request(
        url,
        headers={"User-Agent": USER_AGENT},
    )

    with urllib.request.urlopen(request, timeout=60) as response:
        with tempfile.NamedTemporaryFile(
            delete=False,
            dir=str(destination.parent),
            suffix=".download",
        ) as temporary_file:
            shutil.copyfileobj(response, temporary_file)
            temporary_path = Path(temporary_file.name)

    temporary_path.replace(destination)


def _is_protected(relative_path: Path) -> bool:
    return any(part in PROTECTED_NAMES for part in relative_path.parts)


def _sync_github_folder(
    remote_path: str,
    local_path: Path,
    synced_files: list[str],
) -> None:
    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/"
        f"{GITHUB_REPO}/contents/{remote_path}"
        f"?ref={GITHUB_BRANCH}"
    )

    entries = _request_json(api_url)

    if not isinstance(entries, list):
        raise RuntimeError(
            f"GitHub folder was not returned as a list: {remote_path}"
        )

    for entry in entries:
        entry_name = entry["name"]
        entry_type = entry["type"]
        entry_remote_path = entry["path"]

        relative_path = Path(entry_remote_path).relative_to(
            REMOTE_BOT_FOLDER
        )

        if _is_protected(relative_path):
            continue

        destination = local_path / relative_path

        if entry_type == "dir":
            _sync_github_folder(
                remote_path=entry_remote_path,
                local_path=local_path,
                synced_files=synced_files,
            )
            continue

        if entry_type != "file":
            continue

        download_url = entry.get("download_url")

        if not download_url:
            continue

        _download_file(download_url, destination)
        synced_files.append(str(relative_path))


def sync_from_github() -> tuple[bool, str]:
    """
    Download the newest files from the repository's bot folder.

    If GitHub is unavailable, existing installed files are retained and
    the bot can continue running offline.
    """

    LOCAL_BOT_DIR.mkdir(parents=True, exist_ok=True)
    synced_files: list[str] = []

    try:
        _sync_github_folder(
            remote_path=REMOTE_BOT_FOLDER,
            local_path=LOCAL_BOT_DIR,
            synced_files=synced_files,
        )

        return (
            True,
            f"Synced {len(synced_files)} file(s) from GitHub.",
        )

    except (
        urllib.error.URLError,
        urllib.error.HTTPError,
        TimeoutError,
        OSError,
        RuntimeError,
        json.JSONDecodeError,
    ) as exc:
        return (
            False,
            "GitHub update was unavailable. "
            f"Using the installed files. Reason: {exc}",
        )
