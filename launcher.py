from __future__ import annotations

import runpy
import sys
import traceback
from pathlib import Path

from updater import sync_from_github


APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)

BOT_DIR = APP_DIR / "bot"
RUNTIME_APP = BOT_DIR / "app_runtime.py"


def main() -> None:
    updated, message = sync_from_github()
    print(f"[UPDATER] {message}")

    if not RUNTIME_APP.exists():
        raise FileNotFoundError(
            f"The bot runtime file was not found: {RUNTIME_APP}"
        )

    # Ensure imports such as processing.workflow resolve from the
    # downloaded bot folder before any bundled copies.
    sys.path.insert(0, str(BOT_DIR))
    sys.path.insert(0, str(APP_DIR))

    runpy.run_path(
        str(RUNTIME_APP),
        run_name="__main__",
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("Press ENTER to close...")
        raise
