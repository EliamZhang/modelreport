from __future__ import annotations

from typing import Any

import pandas as pd

from .data_validator import get_primary_key, validate_columns


def _merge_with_overwrite(base: pd.DataFrame, right: pd.DataFrame, left_key: str, right_key: str, how: str, source_table: str, logger) -> pd.DataFrame:
    overlap_cols = [c for c in right.columns if c != right_key and c in base.columns and c != left_key]
    if not overlap_cols:
        return base.merge(right, left_on=left_key, right_on=right_key, how=how)

    preview = overlap_cols[:30]
    suffix = "" if len(overlap_cols) <= 30 else f" ... (+{len(overlap_cols) - 30} more)"
    logger.warning(
        f"Join overwrite columns detected: source_table={source_table}, overlap_cnt={len(overlap_cols):,}, columns={preview}{suffix}"
    )

    renamed_overlap = {col: f"__overwrite__{col}" for col in overlap_cols}
    right_to_merge = right.rename(columns=renamed_overlap).copy()
    match_flag = "__right_matched__"
    while match_flag in base.columns or match_flag in right_to_merge.columns:
        match_flag = f"_{match_flag}"
    right_to_merge[match_flag] = True

    merged = base.merge(right_to_merge, left_on=left_key, right_on=right_key, how=how)
    matched_mask = merged[match_flag].fillna(False).astype(bool)
    for col in overlap_cols:
        incoming_col = renamed_overlap[col]
        merged[col] = merged[incoming_col].where(matched_mask, merged[col])
        merged = merged.drop(columns=[incoming_col])
    return merged.drop(columns=[match_flag])


def _prepare_right_table(
    table: pd.DataFrame,
    key: str,
    fields: list[str] | None,
    table_name: str,
    logger,
    deduplicate_key: bool = True,
) -> pd.DataFrame:
    fields = fields or [c for c in table.columns if c != key]
    available_fields = [f for f in fields if f in table.columns and f != key]
    missing_fields = [f for f in fields if f not in table.columns and f != key]
    logger.info(
        "Join field check: "
        f"table={table_name}, join_key={key}, configured_field_cnt={len(fields):,}, "
        f"matched_field_cnt={len(available_fields):,}, missing_field_cnt={len(missing_fields):,}"
    )
    if missing_fields:
        logger.warning(f"Join table={table_name} missing fields skipped: {missing_fields}")
    validate_columns(table, [key], table_name, logger, fail=False)
    cols = [key] + available_fields if key in table.columns else available_fields
    right = table[cols].copy()
    if key in right.columns and deduplicate_key:
        dup_cnt = int(right.duplicated(subset=[key], keep=False).sum())
        if dup_cnt:
            logger.warning(f"Join table={table_name} has duplicate key rows={dup_cnt:,}; keeping first by key")
            right = right.drop_duplicates(subset=[key], keep="first")
    return right


def enrich_base_sample(tables: dict[str, pd.DataFrame], cfg: dict[str, Any], logger) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """Build the enriched base sample without feature-detail columns."""
    key = get_primary_key(cfg)
    base_name = cfg.get("analysis", {}).get("base_table", "base_sample")
    if base_name not in tables:
        raise KeyError(f"Base table not loaded: {base_name}")
    base = tables[base_name].copy()
    validate_columns(base, [key], base_name, logger, fail=True)

    merge_stats: list[dict[str, Any]] = []
    logger.info(f"Base sample rows={len(base):,}, cols={len(base.columns):,}")

    for join_name, join_cfg in (cfg.get("joins") or {}).items():
        if not join_cfg.get("enabled", True):
            continue
        # Feature detail is intentionally not merged into 01_enriched_base_sample to avoid huge output.
        if join_name == "feature_detail":
            continue

        source_table = join_cfg.get("source_table", join_name)
        if source_table not in tables:
            logger.warning(f"Join source table not loaded, skipped: {source_table}")
            continue
        right_table = tables[source_table]
        join_key = join_cfg.get("join_key", key)
        how = join_cfg.get("how", "left")
        fields = join_cfg.get("fields")
        dedup = bool(join_cfg.get("deduplicate_key", True))

        if join_key != key:
            logger.warning(f"Current version expects same base key and join key. base_key={key}, join_key={join_key}")
        logger.info(
            "Preparing join: "
            f"name={join_name}, source_table={source_table}, how={how}, "
            f"base_key={key}, join_key={join_key}, configured_field_cnt={len(fields or []):,}, "
            f"source_rows={len(right_table):,}, source_cols={len(right_table.columns):,}"
        )
        right = _prepare_right_table(right_table, join_key, fields, source_table, logger, dedup)
        if join_key not in right.columns:
            logger.warning(f"Join skipped because key missing in source table={source_table}: {join_key}")
            continue

        before_rows = len(base)
        right_key_set = set(right[join_key].dropna().unique())
        match_cnt = int(base[key].isin(right_key_set).sum())
        fail_cnt = int(before_rows - match_cnt)
        match_rate = (match_cnt / before_rows) if before_rows else 0.0
        merge_stats.extend([
            {"section": "join", "table": source_table, "metric": "base_row_cnt", "value": before_rows},
            {"section": "join", "table": source_table, "metric": "join_success_cnt", "value": match_cnt},
            {"section": "join", "table": source_table, "metric": "join_fail_cnt", "value": fail_cnt},
        ])

        base = _merge_with_overwrite(base, right, key, join_key, how, source_table, logger)
        if join_key != key and join_key in base.columns:
            base = base.drop(columns=[join_key])
        logger.info(
            "Join result: "
            f"source_table={source_table}, before_rows={before_rows:,}, after_rows={len(base):,}, "
            f"match_application_id_cnt={match_cnt:,}, unmatched_application_id_cnt={fail_cnt:,}, "
            f"match_rate={match_rate:.2%}, merged_in_field_cnt={len(right.columns) - (1 if join_key in right.columns else 0):,}"
        )

    return base, merge_stats
