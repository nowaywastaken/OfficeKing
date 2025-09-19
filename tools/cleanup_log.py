from __future__ import annotations

"""One-off helper to trigger logging setup and clean historical WARNING lines.

This does not change runtime behavior; `index.py` will perform the same cleanup
on startup. This script simply allows running the cleanup quickly without a
full scan.
"""

from pathlib import Path
import sys

# Ensure project root on sys.path when running from tools/
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from activity_scanner.config import CONFIG_ROOT, LOG_LEVEL  # type: ignore
from index import setup_logging  # type: ignore


def main() -> None:
    setup_logging(log_path=CONFIG_ROOT / "log.txt", level_name=LOG_LEVEL)
    print("log cleanup completed")


if __name__ == "__main__":
    main()
