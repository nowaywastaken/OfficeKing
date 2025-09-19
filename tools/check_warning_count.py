from __future__ import annotations

from pathlib import Path


def main() -> None:
    p = Path("log.txt")
    s = p.read_text(encoding="utf-8", errors="ignore")
    lines = s.splitlines()
    count = sum(1 for ln in lines if "[WARNING]" in ln)
    print("[python] lines:", len(lines), "warnings:", count)


if __name__ == "__main__":
    main()

