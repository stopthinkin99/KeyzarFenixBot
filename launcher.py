from __future__ import annotations

import runpy
import sys
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

from updater import sync_from_github


APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

BOT_DIR = APP_DIR / "bot"
RUNTIME_APP = BOT_DIR / "app_runtime.py"
LOG_DIR = APP_DIR / "data" / "logs"
STARTUP_LOG = LOG_DIR / "launcher.log"


def _write_log(message: str) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    with STARTUP_LOG.open(
        "a",
        encoding="utf-8",
    ) as log_file:
        log_file.write(
            f"[{datetime.now():%Y-%m-%d %H:%M:%S}] "
            f"{message}\n"
        )


def main() -> None:
    updated, message = sync_from_github()
    _write_log(f"Updater: {message}")

    if not RUNTIME_APP.exists():
        raise FileNotFoundError(
            f"The runtime application was not found:\n{RUNTIME_APP}"
        )

    # Runtime modules are downloaded outside the EXE so GitHub can
    # update them without rebuilding the installer.
    sys.path.insert(0, str(BOT_DIR))
    sys.path.insert(0, str(APP_DIR))

    runpy.run_path(
        str(RUNTIME_APP),
        run_name="__main__",
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        error_text = (
            f"{type(exc).__name__}: {exc}\n\n"
            f"{traceback.format_exc()}"
        )

        _write_log(error_text)

        try:
            messagebox.showerror(
                "Keyzar Fenix Bot",
                (
                    "The application could not start.\n\n"
                    f"{type(exc).__name__}: {exc}\n\n"
                    f"Startup log:\n{STARTUP_LOG}"
                ),
            )
        except Exception:
            pass

        raise
