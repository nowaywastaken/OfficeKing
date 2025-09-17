#!/usr/bin/env python3
"""Convenience launcher that scans the ./input directory and writes results."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from activity_scanner.cli import run_cli


def build_output_filename() -> str:
    """Create a timestamped Excel filename to avoid accidental overwrites."""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"activity_hits_{timestamp}.xlsx"


def launch_scan() -> int:
    root = Path("input")
    if not root.exists():
        print(f"未找到输入目录: {root}")
        return 1
    if not root.is_dir():
        print(f"输入路径不是文件夹: {root}")
        return 1
    if not any(root.iterdir()):
        print(f"输入目录为空: {root}")
        return 1

    output = build_output_filename()
    return run_cli(["--out", output])


if __name__ == "__main__":
    raise SystemExit(launch_scan())
