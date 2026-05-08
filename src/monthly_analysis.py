from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

from .data_validator import resolve_datetime_field
from .metric_calculator import calculate_group_metrics


def add_sample_month(df: pd.DataFrame, cfg: dict[str, Any], logger) -> pd.DataFrame:
    """Derive a YYYY-MM sample month column from the configured datetime field."""
    month_col = cfg.get("analysis", {}).get("sample_month_field", "sample_month")
    datetime_col = resolve_datetime_field(df, cfg)
    if not datetime_col:
        logger.warning("No datetime field resolved; sample_month will not be derived.")
        return df

    enriched = df.copy()
    parsed = pd.to_datetime(enriched[datetime_col], errors="coerce")
    invalid_cnt = int(parsed.isna().sum())
    enriched[month_col] = parsed.dt.strftime("%Y-%m")
    logger.info(
        f"Derived sample month column: source={datetime_col}, target={month_col}, invalid_datetime_cnt={invalid_cnt:,}"
    )
    return enriched


def _sort_pivot_columns(columns: list[Any]) -> list[Any]:
    normalized = [col for col in columns if pd.notna(col)]
    return sorted(normalized, key=lambda x: str(x))


def _format_pivot_values(pivot: pd.DataFrame, value_format: str | None) -> pd.DataFrame:
    if not value_format:
        return pivot

    formatted = pivot.copy()
    if value_format == "percent_1":
        return formatted.apply(lambda col: col.map(lambda x: "" if pd.isna(x) else f"{x:.1%}"))
    if value_format == "percent_2":
        return formatted.apply(lambda col: col.map(lambda x: "" if pd.isna(x) else f"{x:.2%}"))
    if value_format == "float_4":
        return formatted.apply(lambda col: col.map(lambda x: "" if pd.isna(x) else f"{x:.4f}"))
    return formatted


def _build_monthly_output_file(file_name: str, suffix: str) -> str:
    p = Path(file_name)
    return f"{p.stem}{suffix}{p.suffix or '.csv'}"


def _normalize_merged_monthly_output(
    result: pd.DataFrame,
    group_by: list[str],
    month_col: str,
    analysis_cfg: dict[str, Any],
) -> pd.DataFrame:
    merge_id_columns = analysis_cfg.get("merge_id_columns")
    if merge_id_columns is None:
        merge_id_columns = group_by[:1] if group_by else []
    merge_id_columns = list(merge_id_columns) + [month_col]

    feature_columns = [c for c in group_by if c not in merge_id_columns]
    if not feature_columns:
        raise ValueError(
            f"Merged monthly output requires at least one non-id group column; analysis={analysis_cfg.get('name')}"
        )

    normalized = result.copy()
    feature_label = analysis_cfg.get("merge_group_label")

    if len(feature_columns) == 1:
        feature_col = feature_columns[0]
        normalized = normalized.rename(columns={feature_col: "group_value"})
        normalized.insert(len(merge_id_columns) - 1, "group_feature", feature_label or feature_col)
        return normalized

    normalized.insert(len(merge_id_columns) - 1, "group_feature", feature_label or "+".join(feature_columns))
    normalized["group_value"] = (
        normalized[feature_columns]
        .fillna("NA")
        .astype(str)
        .agg(" | ".join, axis=1)
    )
    return normalized.drop(columns=feature_columns)


def _run_monthly_group_analyses(df: pd.DataFrame, cfg: dict[str, Any], logger, output_handler=None) -> dict[str, pd.DataFrame]:
    outputs: dict[str, pd.DataFrame] = {}
    monthly_group_cfg = cfg.get("analysis", {}).get("monthly_group_analyses", {}) or {}
    if not monthly_group_cfg.get("enabled", True):
        return outputs

    month_col = cfg.get("analysis", {}).get("sample_month_field", "sample_month")
    output_suffix = monthly_group_cfg.get("output_suffix", "_by_sample_month")
    merged_outputs: dict[str, list[pd.DataFrame]] = defaultdict(list)

    for analysis_cfg in cfg.get("analysis", {}).get("group_analyses", []) or []:
        if not analysis_cfg.get("enabled", True):
            continue

        name = analysis_cfg.get("name")
        group_by = analysis_cfg.get("group_by", []) or []
        include_groups = analysis_cfg.get("include_metric_groups")
        output_file = analysis_cfg.get("output_file", f"{name}.csv")
        merged_output_file = analysis_cfg.get("merged_output_file")
        monthly_output_file = _build_monthly_output_file(merged_output_file or output_file, output_suffix)
        monthly_group_by = list(group_by) + [month_col]

        missing = [c for c in monthly_group_by if c not in df.columns]
        if missing:
            logger.warning(f"Monthly group analysis skipped={name}; missing group columns={missing}")
            continue

        logger.info(f"Running monthly group analysis={name}, group_by={monthly_group_by}")
        result = calculate_group_metrics(df, monthly_group_by, cfg, include_groups)

        if merged_output_file:
            normalized = _normalize_merged_monthly_output(result, group_by, month_col, analysis_cfg)
            merged_outputs[monthly_output_file].append(normalized)
            logger.info(
                f"Monthly group analysis done={name}, rows={len(result):,}, merged_output_file={monthly_output_file}"
            )
            continue

        outputs[monthly_output_file] = result
        if output_handler is not None:
            output_handler(monthly_output_file, result)
        logger.info(f"Monthly group analysis done={name}, rows={len(result):,}, output_file={monthly_output_file}")

    for monthly_output_file, dfs in merged_outputs.items():
        outputs[monthly_output_file] = pd.concat(dfs, ignore_index=True)
        if output_handler is not None:
            output_handler(monthly_output_file, outputs[monthly_output_file])
        logger.info(f"Merged monthly group output prepared: {monthly_output_file}, parts={len(dfs):,}")

    return outputs


def run_monthly_analyses(df: pd.DataFrame, cfg: dict[str, Any], logger, output_handler=None) -> dict[str, pd.DataFrame]:
    month_col = cfg.get("analysis", {}).get("sample_month_field", "sample_month")
    if month_col not in df.columns:
        logger.warning(f"Monthly analyses skipped because month column is missing: {month_col}")
        return {}

    working_df = df[df[month_col].notna()].copy()
    if working_df.empty:
        logger.warning("Monthly analyses skipped because there are no rows with valid sample_month.")
        return {}

    outputs = _run_monthly_group_analyses(working_df, cfg, logger, output_handler=output_handler)
    analyses = cfg.get("analysis", {}).get("monthly_analyses", []) or []

    for analysis_cfg in analyses:
        if not analysis_cfg.get("enabled", True):
            continue

        name = analysis_cfg.get("name", "monthly_analysis")
        row_group_by = analysis_cfg.get("group_by", []) or []
        metric_field = analysis_cfg.get("metric_field")
        output_file = analysis_cfg.get("output_file", f"{name}.csv")
        include_groups = analysis_cfg.get("include_metric_groups")
        value_format = analysis_cfg.get("value_format")
        output_layout = analysis_cfg.get("output_layout", "table")

        if not metric_field:
            logger.warning(f"Monthly analysis skipped={name}; metric_field is missing")
            continue

        group_by = list(row_group_by) + [month_col]
        missing = [c for c in group_by if c not in working_df.columns]
        if missing:
            logger.warning(f"Monthly analysis skipped={name}; missing group columns={missing}")
            continue

        logger.info(f"Running monthly analysis={name}, group_by={group_by}, metric_field={metric_field}")
        result = calculate_group_metrics(working_df, group_by, cfg, include_groups)
        if metric_field not in result.columns:
            logger.warning(f"Monthly analysis skipped={name}; metric field not found in result={metric_field}")
            continue

        if output_layout == "pivot":
            pivot = result.pivot(index=row_group_by, columns=month_col, values=metric_field)
            pivot = pivot.reindex(columns=_sort_pivot_columns(list(pivot.columns)))
            pivot = _format_pivot_values(pivot, value_format)
            pivot.columns.name = month_col
            final_df = pivot.reset_index()
        else:
            keep_cols = group_by + [metric_field]
            final_df = result[keep_cols].copy()
            if value_format in {"percent_1", "percent_2", "float_4"}:
                final_df[metric_field] = _format_pivot_values(final_df[[metric_field]], value_format)[metric_field]

        outputs[output_file] = final_df
        if output_handler is not None:
            output_handler(output_file, final_df)
        logger.info(f"Monthly analysis done={name}, rows={len(final_df):,}, cols={len(final_df.columns):,}")

    return outputs
