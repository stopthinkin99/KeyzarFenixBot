from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fenix.login_session import save_fenix_login_session


def main() -> None:
    def confirm() -> bool:
        input(
            "Press ENTER after logging in and opening Search Stock: "
        )
        return True

    save_fenix_login_session(
        confirmation_callback=confirm,
    )


if __name__ == "__main__":
    main()
