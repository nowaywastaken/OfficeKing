#!/usr/bin/env python3
"""一键运行：Office/PDF 转文本并检索命中的学号/姓名，导出 Excel。

用法：
  python index.py

行为：
  - 启动即从根目录 `config.yml` 读取全部参数
  - 使用 MarkItDown 转换 Word/Excel/PPT（含老旧格式）为 Markdown 文本
  - 对 PDF 同时进行：嵌入文本提取 + Tesseract OCR，再与 MarkItDown 转换结果合并
  - 只搜索配置中给定的学生映射（姓名→学号）
  - 若命中，记录：命中的内容、文件名、保存路径；导出到 Output/时间戳(含提示语) 目录
  - 同时将命中的源文件复制到该目录
  - 全程写入同一个 `log.txt` 并同步输出到终端
"""

from __future__ import annotations

import logging
import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd  # type: ignore[import-not-found]

from activity_scanner.config import (
    CONFIG_ROOT,
    LOG_LEVEL,
    STUDENT_ID_MAP,
)
from activity_scanner.extractors import SUPPORTED_EXTENSIONS, read_text_from_path


def setup_logging(log_path: Path, level_name: str) -> None:
    """Configure logging to console and a single append-only file."""

    level = getattr(logging, str(level_name).upper(), logging.INFO)
    logger = logging.getLogger()
    logger.setLevel(level)
    logger.handlers.clear()

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    ch = logging.StreamHandler()
    ch.setLevel(level)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    fh = logging.FileHandler(log_path, mode="a", encoding="utf-8")
    fh.setLevel(level)
    fh.setFormatter(fmt)
    logger.addHandler(fh)


def _coerce_paths(values: Iterable[str]) -> List[Path]:
    """Return existing files/dirs as Path list (deduplicated, absolute)."""

    out: List[Path] = []
    seen: set[str] = set()
    for raw in values:
        p = (CONFIG_ROOT / str(raw)).resolve() if not os.path.isabs(str(raw)) else Path(raw).resolve()
        if p.exists():
            key = str(p)
            if key not in seen:
                seen.add(key)
                out.append(p)
        else:
            logging.warning("[input] 路径不存在: %s", p)
    return out


def collect_supported_files(inputs: Iterable[Path], exts: Iterable[str]) -> List[Path]:
    """Recursively collect supported files under the given inputs."""

    supported = {e.lower() for e in exts}
    results: List[Path] = []
    for item in inputs:
        if item.is_file():
            if item.suffix.lower() in supported:
                results.append(item.resolve())
        elif item.is_dir():
            for root, _dirs, files in os.walk(item):
                for name in files:
                    if os.path.splitext(name)[1].lower() in supported:
                        results.append((Path(root) / name).resolve())
        else:
            logging.warning("[input] 非文件/目录: %s", item)
    # stable order + dedup
    as_str = sorted(set(str(p) for p in results))
    return [Path(s) for s in as_str]


def _load_config_values() -> Tuple[List[Path], Path, str, str]:
    """Load required values from config.yml.

    Returns: (inputs, output_root, output_folder_format, report_filename)
    """

    import yaml  # type: ignore[import-not-found]

    cfg_path = CONFIG_ROOT / "config.yml"
    data = yaml.safe_load(cfg_path.read_text(encoding="utf-8")) or {}

    # All parameters are sourced from config.yml, no in-code defaults.
    input_paths = data.get("input_paths")
    output_root = data.get("output_root_dir")
    folder_fmt = data.get("output_folder_format")
    report_name = data.get("report_filename")

    if not isinstance(input_paths, list) or not input_paths:
        raise ValueError("config.yml: input_paths 必须是非空列表")
    if not isinstance(output_root, str) or not output_root:
        raise ValueError("config.yml: output_root_dir 必须是非空字符串")
    if not isinstance(folder_fmt, str) or not folder_fmt:
        raise ValueError("config.yml: output_folder_format 必须是非空字符串")
    if not isinstance(report_name, str) or not report_name.endswith(".xlsx"):
        raise ValueError("config.yml: report_filename 必须是以 .xlsx 结尾的文件名")

    inputs = _coerce_paths(input_paths)
    if not inputs:
        raise FileNotFoundError("input_paths 指向的文件/目录均不存在")

    return inputs, (CONFIG_ROOT / output_root).resolve(), folder_fmt, report_name


def _ensure_output_folder(root: Path, folder_fmt: str) -> Path:
    """Create the timestamped output folder and return it."""

    # Allow strftime tokens in folder_fmt
    name = datetime.now().strftime(folder_fmt)
    target = (root / name).resolve()
    target.mkdir(parents=True, exist_ok=True)
    return target


def _search_hits(text: str, name_to_id: Dict[str, str]) -> List[str]:
    """Return a list of matched raw tokens (names and IDs)."""

    matches: List[str] = []
    for name, sid in name_to_id.items():
        if name and name in text:
            matches.append(name)
        if sid and sid in text:
            matches.append(sid)
        # name without middle dot variant
        if "·" in name:
            alt = name.replace("·", "")
            if alt and alt in text and alt not in matches:
                matches.append(alt)
    return matches


def main() -> int:
    """Run the full scan using YAML config only."""

    log_path = CONFIG_ROOT / "log.txt"
    setup_logging(log_path=log_path, level_name=LOG_LEVEL)

    logging.info("读取 config.yml 并初始化")
    try:
        inputs, out_root, folder_fmt, report_name = _load_config_values()
    except Exception as exc:
        logging.error("配置读取失败: %s", exc)
        return 1

    logging.info("输入路径: %s", [str(p) for p in inputs])
    files = collect_supported_files(inputs, SUPPORTED_EXTENSIONS)
    if not files:
        logging.info("未找到可处理的文件 (支持: %s)", sorted(SUPPORTED_EXTENSIONS))
        return 0

    # Resolve roster
    name_to_id = dict(STUDENT_ID_MAP)
    logging.info("学生名单条目数: %d", len(name_to_id))

    # Prepare output
    out_dir = _ensure_output_folder(out_root, folder_fmt)
    report_path = out_dir / report_name
    logging.info("输出目录: %s", out_dir)

    rows: List[Dict[str, str]] = []
    matched_files: List[Path] = []

    total = len(files)
    logging.info("开始处理文件: %d", total)
    for idx, path in enumerate(files, start=1):
        logging.info("[%d/%d] 读取: %s", idx, total, path.name)
        try:
            text = read_text_from_path(str(path))
        except ImportError as exc:
            missing = getattr(exc, "name", str(exc))
            logging.error("依赖缺失: %s", missing)
            return 2
        except Exception as exc:
            logging.warning("读取失败 %s: %s", path, exc)
            continue

        hits = _search_hits(text or "", name_to_id)
        if hits:
            matched_files.append(path)
            for token in hits:
                rows.append({
                    "命中的内容": token,
                    "文件名": path.name,
                    "保存路径": str(path),
                })

    # 写出 Excel
    try:
        df = pd.DataFrame(rows, columns=["命中的内容", "文件名", "保存路径"]) if rows else pd.DataFrame(
            columns=["命中的内容", "文件名", "保存路径"]
        )
        with pd.ExcelWriter(str(report_path), engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="命中明细")
        logging.info("写出报表: %s (命中 %d 条, 文件 %d 个)", report_path, len(rows), len(matched_files))
    except Exception as exc:
        logging.error("写出 Excel 失败: %s", exc)
        return 3

    # 复制命中文件
    for src in matched_files:
        try:
            dst = out_dir / src.name
            # 若重名，追加序号避免覆盖
            if dst.exists():
                stem, suf = src.stem, src.suffix
                k = 2
                while True:
                    cand = out_dir / f"{stem}_{k}{suf}"
                    if not cand.exists():
                        dst = cand
                        break
                    k += 1
            shutil.copy2(src, dst)
        except Exception as exc:
            logging.warning("复制失败 %s -> %s: %s", src, out_dir, exc)

    logging.info("完成：共扫描 %d 个文件，命中 %d 个文件，详情见 %s", total, len(matched_files), report_path)
    print(f"完成：共扫描 {total} 个文件，命中 {len(matched_files)} 个文件，输出：{report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
