#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSUST �༶�۲������ļ�ɨ��������汾
----------------------------------------------------
��;���� Word��PDF��Excel �в��ұ���ͬѧ�ġ�ѧ��/�����������������༶�ؼ��ʣ�Ĭ�ϣ�����2401����·������ɺӹ��̣���
�ص㣺����Ҫ roster.csv�������Ѿ�д���ڴ����

������pandas��python-docx��pdfplumber��openpyxl
    pip install pandas python-docx pdfplumber openpyxl

�÷�ʾ����
    python csust_activity_scan_oneoff.py --paths ./input --out activity_hits.xlsx
    python csust_activity_scan_oneoff.py --paths ./input --class-tags "����2401,��·������ɺӹ���2401��"
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from typing import List

import pandas as pd  # type: ignore[import-not-found]

from reporting import build_report_frames, write_reports
from roster import StudentRoster
from scanner import Document, gather_input_files, infer_activity_name, scan_document
from text_extractors import FAILED_PDFS, SUPPORTED_EXTS, extract_text, open_failed_pdfs


logging.basicConfig(level=logging.INFO, format="%(message)s")
logging.getLogger("pdfminer").setLevel(logging.ERROR)


def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="CSUST �༶�۲������ļ�ɨ������һ���԰棩")
    parser.add_argument("--paths", nargs="+", required=True, help="��ɨ����ļ�/�ļ���·�����ɶ����")
    parser.add_argument("--out", default="activity_hits.xlsx", help="��� Excel �ļ�����Ĭ�ϣ�activity_hits.xlsx��")
    parser.add_argument(
        "--class-tags",
        default="����2401,��·������ɺӹ���",
        help="���������İ༶�ؼ��ʣ����ŷָ���Ĭ�ϣ�����2401,��·������ɺӹ��̣�",
    )
    return parser.parse_args(argv)


def build_class_tags(raw: str | None) -> List[str]:
    if not raw:
        return []
    return [tag.strip() for tag in raw.split(",") if tag.strip()]


def main(argv: List[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        roster = StudentRoster.default()
    except Exception as exc:
        print(f"�������н���ͷ���쳣��{exc}")
        return 1

    class_tags = build_class_tags(args.class_tags)

    files = gather_input_files(args.paths, SUPPORTED_EXTS)
    if not files:
        print("δ�ҵ���ɨ����ļ���֧�֣�.docx .pdf .xlsx .xls .csv")
        return 4

    all_rows: List[dict] = []
    for path in files:
        try:
            text = extract_text(path)
        except ImportError as exc:
            missing = getattr(exc, "name", str(exc))
            print(f"ȱ��������{missing}")
            return 1
        except Exception as exc:
            all_rows.append(
                {
                    "�ļ�·��": path,
                    "�����": os.path.splitext(os.path.basename(path))[0],
                    "״̬": f"��ȡʧ�ܣ�{exc}",
                    "��������": "",
                    "����ֵ": "",
                    "ѧ��": "",
                    "����": "",
                    "֤��Ƭ��": "",
                    "���д���": 0,
                }
            )
            continue

        document = Document(
            path=path,
            text=text,
            activity=infer_activity_name(path, text),
        )
        try:
            rows = scan_document(document, roster, class_tags)
        except Exception as exc:
            all_rows.append(
                {
                    "�ļ�·��": path,
                    "�����": document.activity,
                    "״̬": f"ɨ���쳣��{exc}",
                    "��������": "",
                    "����ֵ": "",
                    "ѧ��": "",
                    "����": "",
                    "֤��Ƭ��": "",
                    "���д���": 0,
                }
            )
            continue
        all_rows.extend(rows)

    detail, per_activity, per_person, class_hits = build_report_frames(all_rows, roster)
    write_reports(args.out, detail, per_activity, per_person, class_hits)

    if FAILED_PDFS:
        logging.warning("[PDF] �����ļ�δ�ܳ�ȡ�ı�����Ϊ���Զ����Ա��飺%s", "��".join(sorted(FAILED_PDFS)))
        open_failed_pdfs(sorted(FAILED_PDFS))

    print(f"ɨ����ɣ������� {len(files)} ���ļ�������������{args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
