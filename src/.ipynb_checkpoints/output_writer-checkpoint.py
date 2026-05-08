from __future__ import annotations

from pathlib import Path
from typing import Any
import importlib.util

import pandas as pd

from .config_loader import get_encoding, get_output_dir


def write_csv(df: pd.DataFrame, path: str | Path, encoding: str = "utf-8-sig", index: bool = False) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(p, index=index, encoding=encoding)
    return p


def write_output(file_name: str, df: pd.DataFrame, cfg: dict[str, Any], logger) -> Path:
    out_dir = get_output_dir(cfg)
    encoding = get_encoding(cfg)
    path = out_dir / file_name
    write_csv(df, path, encoding=encoding)
    logger.info(f"Wrote output: {path}, rows={len(df):,}, cols={len(df.columns):,}")
    return path


def write_outputs(outputs: dict[str, pd.DataFrame], cfg: dict[str, Any], logger) -> list[Path]:
    written: list[Path] = []
    for file_name, df in outputs.items():
        written.append(write_output(file_name, df, cfg, logger))
    return written


def write_split_output(name: str, df: pd.DataFrame, subdir: str, cfg: dict[str, Any], logger) -> Path:
    out_dir = get_output_dir(cfg) / subdir
    encoding = get_encoding(cfg)
    safe_name = str(name).replace("/", "_").replace("\\", "_").replace(" ", "_")
    path = out_dir / f"{safe_name}.csv"
    write_csv(df, path, encoding=encoding)
    logger.info(f"Wrote split output: {path}, rows={len(df):,}")
    return path


def write_split_outputs(split_outputs: dict[str, pd.DataFrame], subdir: str, cfg: dict[str, Any], logger) -> list[Path]:
    written: list[Path] = []
    for name, df in split_outputs.items():
        written.append(write_split_output(name, df, subdir, cfg, logger))
    return written


def build_sheet_name(file_name: str, used_names: set[str]) -> str:
    stem = Path(file_name).stem
    invalid_chars = set('[]:*?/\\')
    safe = "".join(ch if ch not in invalid_chars else "_" for ch in stem).strip()
    safe = safe or "sheet"
    base = safe[:31]
    candidate = base
    counter = 2
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1
    used_names.add(candidate)
    return candidate


def write_excel_workbook(outputs: dict[str, pd.DataFrame], workbook_name: str, cfg: dict[str, Any], logger) -> Path | None:
    if not outputs:
        logger.info("Skip Excel workbook export because there are no tabular outputs.")
        return None

    engine = None
    if importlib.util.find_spec("openpyxl") is not None:
        engine = "openpyxl"
    elif importlib.util.find_spec("xlsxwriter") is not None:
        engine = "xlsxwriter"
    else:
        logger.warning(
            "Skip Excel workbook export because no Excel engine is installed. Install openpyxl or xlsxwriter to enable .xlsx output."
        )
        return None

    out_dir = get_output_dir(cfg)
    workbook_path = out_dir / workbook_name
    workbook_path.parent.mkdir(parents=True, exist_ok=True)

    used_names: set[str] = set()
    with pd.ExcelWriter(workbook_path, engine=engine) as writer:
        for file_name, df in outputs.items():
            sheet_name = build_sheet_name(file_name, used_names)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info(f"Added worksheet: {sheet_name}, source={file_name}, rows={len(df):,}")

    logger.info(f"Wrote Excel workbook: {workbook_path}, sheets={len(outputs):,}")
    return workbook_path
