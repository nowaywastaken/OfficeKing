from __future__ import annotations

"""Utilities for assembling pandas DataFrames and writing Excel reports."""

from typing import Iterable, Tuple

import pandas as pd

from .schema import (
    CLASS_TAG_LABEL,
    CLASS_TAG_SHEET_NAME,
    COLUMN_ACTIVITY_FILE_COUNT,
    COLUMN_ACTIVITY_NAME,
    COLUMN_FILE_PATH,
    COLUMN_MATCH_COUNT,
    COLUMN_MATCH_TOTAL,
    COLUMN_MATCH_TYPE,
    COLUMN_MATCH_VALUE,
    COLUMN_PERSON_ACTIVITY_COUNT,
    COLUMN_PERSON_ACTIVITY_LIST,
    COLUMN_SNIPPET,
    COLUMN_STATUS,
    COLUMN_STUDENT_ID,
    COLUMN_STUDENT_NAME,
    DETAIL_SHEET_NAME,
    PER_ACTIVITY_SHEET_NAME,
    PER_PERSON_SHEET_NAME,
    STUDENT_MATCH_TYPES,
)
from .roster_store import StudentDirectory


def build_report_tables(all_rows: Iterable[dict], roster: StudentDirectory) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Transform raw scan results into the DataFrames needed for reporting."""

    detail_df = pd.DataFrame(list(all_rows))

    student_hits = detail_df[
        (detail_df[COLUMN_STATUS] == "OK")
        & (detail_df[COLUMN_MATCH_COUNT] > 0)
        & (detail_df[COLUMN_MATCH_TYPE].isin(STUDENT_MATCH_TYPES))
    ].copy()

    if not student_hits.empty:
        missing_names = student_hits[COLUMN_STUDENT_NAME].eq("") & student_hits[COLUMN_STUDENT_ID].ne("")
        if missing_names.any():
            student_hits.loc[missing_names, COLUMN_STUDENT_NAME] = student_hits.loc[missing_names, COLUMN_STUDENT_ID].map(
                lambda sid: roster.resolve_name(sid or "") or ""
            )

        grouped = student_hits.groupby([COLUMN_STUDENT_ID, COLUMN_STUDENT_NAME, COLUMN_ACTIVITY_NAME], as_index=False).agg(
            {COLUMN_FILE_PATH: "nunique", COLUMN_MATCH_COUNT: "sum"}
        )
        per_activity = grouped.rename(columns={
            COLUMN_FILE_PATH: COLUMN_ACTIVITY_FILE_COUNT,
            COLUMN_MATCH_COUNT: COLUMN_MATCH_TOTAL,
        })

        per_person = per_activity.groupby([COLUMN_STUDENT_ID, COLUMN_STUDENT_NAME], as_index=False).agg(
            {COLUMN_ACTIVITY_NAME: ["nunique", lambda values: "、".join(sorted(set(values)))]}
        )
        per_person.columns = [
            COLUMN_STUDENT_ID,
            COLUMN_STUDENT_NAME,
            COLUMN_PERSON_ACTIVITY_COUNT,
            COLUMN_PERSON_ACTIVITY_LIST,
        ]
    else:
        per_activity = pd.DataFrame(
            columns=[
                COLUMN_STUDENT_ID,
                COLUMN_STUDENT_NAME,
                COLUMN_ACTIVITY_NAME,
                COLUMN_ACTIVITY_FILE_COUNT,
                COLUMN_MATCH_TOTAL,
            ]
        )
        per_person = pd.DataFrame(
            columns=[
                COLUMN_STUDENT_ID,
                COLUMN_STUDENT_NAME,
                COLUMN_PERSON_ACTIVITY_COUNT,
                COLUMN_PERSON_ACTIVITY_LIST,
            ]
        )

    class_hits = detail_df[
        (detail_df[COLUMN_STATUS] == "OK") & (detail_df[COLUMN_MATCH_TYPE] == CLASS_TAG_LABEL)
    ].copy()

    return detail_df, per_activity, per_person, class_hits


def write_report_workbook(
    output_path: str,
    detail: pd.DataFrame,
    per_activity: pd.DataFrame,
    per_person: pd.DataFrame,
    class_hits: pd.DataFrame,
) -> None:
    """Write all report sheets to an Excel workbook."""

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name=DETAIL_SHEET_NAME)
        per_activity.to_excel(writer, index=False, sheet_name=PER_ACTIVITY_SHEET_NAME)
        per_person.to_excel(writer, index=False, sheet_name=PER_PERSON_SHEET_NAME)
        class_hits.to_excel(writer, index=False, sheet_name=CLASS_TAG_SHEET_NAME)
