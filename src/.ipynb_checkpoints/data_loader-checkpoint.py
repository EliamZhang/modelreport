from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from pandas.errors import ParserError

from .data_validator import get_primary_key
from .runtime_summary import log_table_load_result


def _normalize_key_series(series: pd.Series) -> pd.Series:
    normalized = series.astype("string").str.strip()
    # Align numeric-looking ids read as float (e.g. 2014783.0) with integer/text ids.
    normalized = normalized.str.replace(r"\.0+$", "", regex=True)
    normalized = normalized.mask(normalized.isin(["", "nan", "None", "<NA>"]), pd.NA)
    return normalized


def _collect_key_columns(cfg: dict[str, Any]) -> set[str]:
    keys = cfg.get("keys", {}) or {}
    key_columns = {
        keys.get("primary_key"),
        keys.get("user_key"),
    }

    for join_cfg in (cfg.get("joins") or {}).values():
        key_columns.add(join_cfg.get("join_key"))

    fp_cfg = cfg.get("feature_profile", {}) or {}
    key_columns.add(fp_cfg.get("join_key"))

    return {str(col) for col in key_columns if col}


def _flatten_feature_fields(cfg: dict[str, Any]) -> set[str]:
    fields: set[str] = set()
    categories = cfg.get("feature_profile", {}).get("categories", {}) or {}
    for category_cfg in categories.values():
        for field in category_cfg.get("fields", []) or []:
            if field:
                fields.add(str(field))
    return fields


def _collect_required_columns_by_table(cfg: dict[str, Any]) -> dict[str, set[str] | None]:
    required: dict[str, set[str] | None] = {}
    key_columns = _collect_key_columns(cfg)
    input_tables = cfg.get("input_tables") or {}
    base_table = cfg.get("analysis", {}).get("base_table", "base_sample")

    # Base table defaults to full load because it is the backbone output table.
    if base_table in input_tables:
        required[base_table] = None

    for table_name, table_cfg in input_tables.items():
        required.setdefault(table_name, set())
        if required[table_name] is not None:
            required[table_name].update(key_columns)

        manual_fields = table_cfg.get("fields")
        if manual_fields and required[table_name] is not None:
            required[table_name].update(str(field) for field in manual_fields if field)

    for score_cfg in cfg.get("score_binning", {}).get("scores", []) or []:
        source_table = score_cfg.get("source_table")
        score_field = score_cfg.get("score_field")
        if source_table and score_field and source_table in required and required[source_table] is not None:
            required[source_table].add(str(score_field))

    for join_name, join_cfg in (cfg.get("joins") or {}).items():
        source_table = join_cfg.get("source_table", join_name)
        if source_table not in required or required[source_table] is None:
            continue
        join_key = join_cfg.get("join_key")
        if join_key:
            required[source_table].add(str(join_key))
        for field in join_cfg.get("fields", []) or []:
            if field:
                required[source_table].add(str(field))

    fp_cfg = cfg.get("feature_profile", {}) or {}
    source_table = fp_cfg.get("source_table")
    if source_table in required and required[source_table] is not None:
        join_key = fp_cfg.get("join_key")
        if join_key:
            required[source_table].add(str(join_key))
        required[source_table].update(_flatten_feature_fields(cfg))

    return required


def _build_usecols(required_columns: set[str] | None):
    if not required_columns:
        return None
    required = {str(col) for col in required_columns if col}
    return lambda col_name: str(col_name) in required


def read_csv_table(
    path: str | Path,
    read_kwargs: dict[str, Any] | None = None,
    usecols=None,
) -> pd.DataFrame:
    """Read a CSV with safe defaults."""
    p = Path(path).expanduser()
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {p}")
    kwargs = {"low_memory": False}
    if read_kwargs:
        kwargs.update(read_kwargs)
    if usecols is not None:
        kwargs["usecols"] = usecols
    try:
        return pd.read_csv(p, **kwargs)
    except ParserError as exc:
        if "out of memory" not in str(exc).lower():
            raise
        fallback_kwargs = dict(kwargs)
        fallback_kwargs["engine"] = "python"
        return pd.read_csv(p, **fallback_kwargs)


def load_input_tables(cfg: dict[str, Any], logger) -> dict[str, pd.DataFrame]:
    """Load all configured input tables."""
    read_kwargs = cfg.get("runtime", {}).get("pandas_read_csv_kwargs", {}) or {}
    key_columns = _collect_key_columns(cfg)
    required_columns_by_table = _collect_required_columns_by_table(cfg)
    primary_key = get_primary_key(cfg)
    tables: dict[str, pd.DataFrame] = {}

    for table_name, table_cfg in (cfg.get("input_tables") or {}).items():
        if not table_cfg.get("enabled", True):
            logger.info(f"Skipped disabled table={table_name}")
            continue
        path = table_cfg.get("path")
        optional = bool(table_cfg.get("optional", False))
        required_columns = required_columns_by_table.get(table_name)
        usecols = _build_usecols(required_columns)
        if required_columns is None:
            logger.info(f"Loading table={table_name}, path={path}, columns=ALL")
        else:
            logger.info(f"Loading table={table_name}, path={path}, selected_columns={len(required_columns):,}")
        try:
            df = read_csv_table(path, read_kwargs, usecols=usecols)
        except FileNotFoundError:
            if optional:
                logger.warning(f"Optional input table missing, skipped: table={table_name}, path={path}")
                continue
            raise
        for col in key_columns:
            if col in df.columns:
                df[col] = _normalize_key_series(df[col])
        tables[table_name] = df
        log_table_load_result(table_name, df, table_cfg, required_columns, primary_key, logger)

    return tables
