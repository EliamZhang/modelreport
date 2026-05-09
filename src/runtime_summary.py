from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

import pandas as pd

from .data_validator import count_duplicate_keys, get_primary_key


def _format_preview(values: Iterable[Any], limit: int = 8) -> str:
    cleaned = [str(v) for v in values if v is not None and str(v) != ""]
    if not cleaned:
        return "-"
    if len(cleaned) <= limit:
        return ", ".join(cleaned)
    return f"{', '.join(cleaned[:limit])} ... (+{len(cleaned) - limit} more)"


def _series_numeric_summary(series: pd.Series) -> str:
    numeric = pd.to_numeric(series, errors="coerce")
    valid = numeric.dropna()
    if valid.empty:
        return "min=NA, p50=NA, max=NA"
    return (
        f"min={valid.min():.6f}, "
        f"p50={valid.median():.6f}, "
        f"max={valid.max():.6f}"
    )


def _log_value_distribution(
    df: pd.DataFrame,
    column: str,
    logger,
    title: str,
    top_n: int = 10,
) -> None:
    if column not in df.columns:
        logger.warning(f"{title}: column missing, column={column}")
        return

    display_series = df[column].astype("string").fillna("<NA>")
    counts = display_series.value_counts(dropna=False).head(top_n)
    summary = ", ".join(f"{idx}={int(val):,}" for idx, val in counts.items())
    logger.info(f"{title}: column={column}, distinct={df[column].nunique(dropna=True):,}, top_{top_n}={summary}")


def log_runtime_plan(cfg: dict[str, Any], logger) -> None:
    base_table = cfg.get("analysis", {}).get("base_table", "base_sample")
    primary_key = get_primary_key(cfg)
    keys_cfg = cfg.get("keys", {}) or {}
    datetime_candidates = [keys_cfg.get("datetime_key")] + list(keys_cfg.get("datetime_aliases", []) or [])
    score_cfgs = cfg.get("score_binning", {}).get("scores", []) or []
    input_tables = cfg.get("input_tables", {}) or {}
    joins = cfg.get("joins", {}) or {}
    group_analyses = cfg.get("analysis", {}).get("group_analyses", []) or []
    monthly_analyses = cfg.get("analysis", {}).get("monthly_analyses", []) or []

    logger.info("Run summary started")
    logger.info(
        "Core settings: "
        f"base_table={base_table}, primary_key={primary_key}, "
        f"datetime_candidates=[{_format_preview(datetime_candidates, limit=6)}]"
    )

    if input_tables:
        logger.info(f"Input table count={len(input_tables):,}")
        for table_name, table_cfg in input_tables.items():
            logger.info(
                "Input table config: "
                f"table={table_name}, enabled={bool(table_cfg.get('enabled', True))}, "
                f"optional={bool(table_cfg.get('optional', False))}, "
                f"path={Path(str(table_cfg.get('path', ''))).expanduser()}, "
                f"description={table_cfg.get('description', '-')}"
            )

    if score_cfgs:
        logger.info(f"Model score config count={len(score_cfgs):,}")
        for score_cfg in score_cfgs:
            logger.info(
                "Model score config: "
                f"name={score_cfg.get('name', '-')}, "
                f"source_table={score_cfg.get('source_table', '-')}, "
                f"score_field={score_cfg.get('score_field', '-')}, "
                f"bin_field={score_cfg.get('bin_field', '-')}, "
                f"bins={len(score_cfg.get('bins', []) or []):,}, "
                f"bin_groups={len(score_cfg.get('bin_groups', []) or []):,}"
            )

    if joins:
        logger.info(f"Join config count={len(joins):,}")
        for join_name, join_cfg in joins.items():
            fields = join_cfg.get("fields") or []
            logger.info(
                "Join config: "
                f"name={join_name}, enabled={bool(join_cfg.get('enabled', True))}, "
                f"source_table={join_cfg.get('source_table', join_name)}, "
                f"join_key={join_cfg.get('join_key', primary_key)}, "
                f"how={join_cfg.get('how', 'left')}, "
                f"field_cnt={len(fields):,}"
            )

    logger.info(
        "Analysis config: "
        f"group_analyses={len(group_analyses):,}, "
        f"monthly_analyses={len(monthly_analyses):,}, "
        f"feature_profile_enabled={bool(cfg.get('feature_profile', {}).get('enabled', True))}"
    )
    logger.info("Run summary completed")


def log_table_load_result(
    table_name: str,
    df: pd.DataFrame,
    table_cfg: dict[str, Any],
    required_columns: set[str] | None,
    primary_key: str,
    logger,
) -> None:
    missing_selected: list[str] = []
    if required_columns is not None:
        missing_selected = sorted(set(required_columns) - set(map(str, df.columns)))

    key_missing_cnt = int(df[primary_key].isna().sum()) if primary_key in df.columns else -1
    key_dup_cnt = count_duplicate_keys(df, primary_key)
    key_unique_cnt = int(df[primary_key].dropna().nunique()) if primary_key in df.columns else 0

    logger.info(
        "Table load result: "
        f"table={table_name}, rows={len(df):,}, cols={len(df.columns):,}, "
        f"primary_key_present={primary_key in df.columns}, "
        f"primary_key_missing_cnt={key_missing_cnt if key_missing_cnt >= 0 else 'KEY_NOT_FOUND'}, "
        f"primary_key_unique_cnt={key_unique_cnt:,}, "
        f"primary_key_duplicate_row_cnt={key_dup_cnt:,}"
    )
    logger.info(
        "Table column match: "
        f"table={table_name}, selected_column_mode={'ALL' if required_columns is None else 'PARTIAL'}, "
        f"configured_selected_cnt={'ALL' if required_columns is None else len(required_columns)}, "
        f"actual_loaded_cnt={len(df.columns):,}, "
        f"missing_selected_cnt={len(missing_selected):,}"
    )
    if missing_selected:
        logger.warning(
            f"Table selected columns missing: table={table_name}, first_30=[{_format_preview(missing_selected, limit=30)}]"
        )

    logger.info(
        "Table meta: "
        f"table={table_name}, optional={bool(table_cfg.get('optional', False))}, "
        f"description={table_cfg.get('description', '-')}"
    )


def log_table_overlap_with_base(
    tables: dict[str, pd.DataFrame],
    base_table_name: str,
    primary_key: str,
    logger,
) -> None:
    base_df = tables.get(base_table_name)
    if base_df is None or primary_key not in base_df.columns:
        logger.warning(
            f"Base overlap summary skipped: base_table={base_table_name}, primary_key={primary_key}, key_missing=True"
        )
        return

    base_keys = set(base_df[primary_key].dropna().unique())
    logger.info(
        f"Base sample key summary: table={base_table_name}, row_cnt={len(base_df):,}, unique_application_id_cnt={len(base_keys):,}"
    )

    for table_name, df in tables.items():
        if table_name == base_table_name:
            continue
        if primary_key not in df.columns:
            logger.warning(f"Base overlap skipped: table={table_name}, primary_key_missing={primary_key}")
            continue
        table_keys = set(df[primary_key].dropna().unique())
        overlap_cnt = len(base_keys & table_keys)
        only_in_table_cnt = len(table_keys - base_keys)
        overlap_rate = (overlap_cnt / len(base_keys)) if base_keys else 0.0
        logger.info(
            "Base overlap summary: "
            f"base_table={base_table_name}, table={table_name}, "
            f"table_unique_application_id_cnt={len(table_keys):,}, "
            f"overlap_application_id_cnt={overlap_cnt:,}, "
            f"base_match_rate={overlap_rate:.2%}, "
            f"only_in_source_cnt={only_in_table_cnt:,}"
        )


def log_enriched_sample_summary(df: pd.DataFrame, cfg: dict[str, Any], logger, stage_name: str) -> None:
    primary_key = get_primary_key(cfg)
    unique_key_cnt = int(df[primary_key].dropna().nunique()) if primary_key in df.columns else 0
    logger.info(
        f"{stage_name}: rows={len(df):,}, cols={len(df.columns):,}, unique_application_id_cnt={unique_key_cnt:,}"
    )

    for score_cfg in cfg.get("score_binning", {}).get("scores", []) or []:
        score_name = score_cfg.get("name", "-")
        source_table = score_cfg.get("source_table", "-")
        score_field = score_cfg.get("score_field")
        bin_field = score_cfg.get("bin_field")

        if score_field not in df.columns:
            logger.warning(
                f"Model score field missing: model={score_name}, source_table={source_table}, score_field={score_field}"
            )
            continue

        nonnull_cnt = int(df[score_field].notna().sum())
        missing_cnt = int(df[score_field].isna().sum())
        logger.info(
            "Model score summary: "
            f"model={score_name}, source_table={source_table}, score_field={score_field}, "
            f"nonnull_cnt={nonnull_cnt:,}, missing_cnt={missing_cnt:,}, "
            f"{_series_numeric_summary(df[score_field])}"
        )

        if bin_field in df.columns:
            bin_missing_cnt = int(df[bin_field].isna().sum())
            logger.info(
                "Model bin summary: "
                f"model={score_name}, bin_field={bin_field}, "
                f"nonnull_cnt={int(df[bin_field].notna().sum()):,}, missing_cnt={bin_missing_cnt:,}"
            )
            _log_value_distribution(df, bin_field, logger, f"Model bin distribution[{score_name}]")


def log_post_label_summary(df: pd.DataFrame, cfg: dict[str, Any], logger) -> None:
    conv = cfg.get("conversion", {}) or {}
    fields = conv.get("output_fields", {}) or {}
    stage_col = fields.get("conversion_stage", "conversion_stage")
    deal_col = fields.get("deal_flag", "is_deal_application")

    _log_value_distribution(df, stage_col, logger, "Conversion stage distribution")
    if deal_col in df.columns:
        logger.info(
            f"Deal flag summary: column={deal_col}, positive_cnt={int(pd.to_numeric(df[deal_col], errors='coerce').fillna(0).sum()):,}"
        )


def log_month_summary(df: pd.DataFrame, cfg: dict[str, Any], logger) -> None:
    month_col = cfg.get("analysis", {}).get("sample_month_field", "sample_month")
    if month_col in df.columns:
        _log_value_distribution(df, month_col, logger, "Sample month distribution")
