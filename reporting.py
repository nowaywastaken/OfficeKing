from __future__ import annotations

from typing import Iterable, Tuple

import pandas as pd

from columns import (
    ACTIVITY_FILE_COUNT,
    ACTIVITY_NAME,
    CLASS_TAG_LABEL,
    CLASS_TAG_SHEET,
    DETAIL_SHEET,
    FILE_PATH,
    MATCH_COUNT,
    MATCH_TOTAL,
    MATCH_TYPE,
    MATCH_VALUE,
    PER_ACTIVITY_SHEET,
    PER_PERSON_SHEET,
    PERSON_ACTIVITY_COUNT,
    PERSON_ACTIVITY_LIST,
    SNIPPET,
    STATUS,
    STUDENT_ID,
    STUDENT_NAME,
    STUDENT_TYPES,
)
from roster import StudentRoster


def build_report_frames(all_rows: Iterable[dict], roster: StudentRoster) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = pd.DataFrame(list(all_rows))

    hit_df = df[(df[STATUS] == "OK") & (df[MATCH_COUNT] > 0) & (df[MATCH_TYPE].isin(STUDENT_TYPES))].copy()
    if not hit_df.empty:
        missing_names = hit_df[STUDENT_NAME].eq("") & hit_df[STUDENT_ID].ne("")
        if missing_names.any():
            hit_df.loc[missing_names, STUDENT_NAME] = hit_df.loc[missing_names, STUDENT_ID].map(lambda sid: roster.resolve_name(sid or "") or "")

        grouped = hit_df.groupby([STUDENT_ID, STUDENT_NAME, ACTIVITY_NAME], as_index=False).agg({
            FILE_PATH: "nunique",
            MATCH_COUNT: "sum",
        })
        per_activity = grouped.rename(columns={
            FILE_PATH: ACTIVITY_FILE_COUNT,
            MATCH_COUNT: MATCH_TOTAL,
        })

        per_person = per_activity.groupby([STUDENT_ID, STUDENT_NAME], as_index=False).agg({
            ACTIVITY_NAME: ["nunique", lambda s: "��".join(sorted(set(s)))],
        })
        per_person.columns = [
            STUDENT_ID,
            STUDENT_NAME,
            PERSON_ACTIVITY_COUNT,
            PERSON_ACTIVITY_LIST,
        ]
    else:
        per_activity = pd.DataFrame(columns=[STUDENT_ID, STUDENT_NAME, ACTIVITY_NAME, ACTIVITY_FILE_COUNT, MATCH_TOTAL])
        per_person = pd.DataFrame(columns=[STUDENT_ID, STUDENT_NAME, PERSON_ACTIVITY_COUNT, PERSON_ACTIVITY_LIST])

    class_hits = df[(df[STATUS] == "OK") & (df[MATCH_TYPE] == CLASS_TAG_LABEL)].copy()

    return df, per_activity, per_person, class_hits


def write_reports(output_path: str, detail: pd.DataFrame, per_activity: pd.DataFrame, per_person: pd.DataFrame, class_hits: pd.DataFrame) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name=DETAIL_SHEET)
        per_activity.to_excel(writer, index=False, sheet_name=PER_ACTIVITY_SHEET)
        per_person.to_excel(writer, index=False, sheet_name=PER_PERSON_SHEET)
        class_hits.to_excel(writer, index=False, sheet_name=CLASS_TAG_SHEET)
