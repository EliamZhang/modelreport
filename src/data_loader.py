from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


def _normalize_key_series(series: pd.Series) -> pd.Series:
    normalized = series.astype("string").str.strip()
    normalized = normalized.str.replace(r"\.0+$", "", regex=True)
    return normalized.mask(normalized.isin(["", "nan", "None", "<NA>"]), pd.NA)


def _key_columns(cfg: dict[str, Any]) -> set[str]:
    keys = cfg["keys"]
    columns = {keys["primary_key"], keys["user_key"]}
    columns.update(join_cfg["join_key"] for join_cfg in cfg["joins"].values())
    return columns


def _required_columns(cfg: dict[str, Any]) -> dict[str, set[str] | None]:
    required: dict[str, set[str] | None] = {cfg["analysis"]["base_table"]: None}
    for table_name in cfg["input_tables"]:
        required.setdefault(table_name, set(_key_columns(cfg)))

    for score_cfg in cfg["score_binning"]["scores"]:
        table_name = score_cfg["source_table"]
        if required.get(table_name) is not None:
            required[table_name].add(score_cfg["score_field"])

    for join_cfg in cfg["joins"].values():
        table_name = join_cfg["source_table"]
        if required.get(table_name) is not None:
            required[table_name].add(join_cfg["join_key"])
            required[table_name].update(join_cfg["fields"])
    return required


def _usecols(columns: set[str] | None):
    if columns is None:
        return None
    return lambda col_name: str(col_name) in columns


def load_input_tables(cfg: dict[str, Any], logger) -> dict[str, pd.DataFrame]:
    key_columns = _key_columns(cfg)
    required = _required_columns(cfg)
    tables: dict[str, pd.DataFrame] = {}

    for table_name, table_cfg in cfg["input_tables"].items():
        path = Path(table_cfg["path"])
        columns = required.get(table_name)
        logger.info(
            f"Loading table={table_name}, path={path}, "
            f"columns={'ALL' if columns is None else len(columns)}"
        )
        df = pd.read_csv(path, low_memory=False, usecols=_usecols(columns))
        for col in key_columns:
            if col in df.columns:
                df[col] = _normalize_key_series(df[col])
        tables[table_name] = df
        logger.info(f"Loaded table={table_name}, rows={len(df):,}, cols={len(df.columns):,}")
    return tables
