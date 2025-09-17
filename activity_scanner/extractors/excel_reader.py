from __future__ import annotations

"""Spreadsheet (Excel/CSV) extraction utilities."""

import os


def read_spreadsheet_text(path: str) -> str:
    """Return a textual representation of spreadsheet-like data."""

    import pandas as pd  # type: ignore[import-not-found]

    try:
        workbook = pd.ExcelFile(path)
    except Exception:
        if os.path.splitext(path)[1].lower() == ".csv":
            dataframe = pd.read_csv(path, dtype=str).fillna("")
            return dataframe.astype(str).to_string(index=False, header=True)
        raise
    segments: list[str] = []
    for sheet_name in workbook.sheet_names:
        dataframe = workbook.parse(sheet_name, dtype=str).fillna("")
        segments.append(dataframe.astype(str).to_string(index=False, header=True))
    return "\n\n".join(segments)
