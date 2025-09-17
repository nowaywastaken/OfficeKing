from __future__ import annotations

import os


def extract_text_from_excel_like(path: str) -> str:
    """Return a textual representation of spreadsheet data (Excel/CSV)."""
    import pandas as pd  # type: ignore[import-not-found]

    try:
        xls = pd.ExcelFile(path)
    except Exception:
        if os.path.splitext(path)[1].lower() == ".csv":
            df = pd.read_csv(path, dtype=str).fillna("")
            return df.astype(str).to_string(index=False, header=True)
        raise
    parts: list[str] = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, dtype=str).fillna("")
        parts.append(df.astype(str).to_string(index=False, header=True))
    return "\n\n".join(parts)
