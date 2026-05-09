from __future__ import annotations

from typing import Any

import pandas as pd


def _merge_with_overwrite(
    base: pd.DataFrame,
    right: pd.DataFrame,
    key: str,
    source_table: str,
    logger,
) -> pd.DataFrame:
    overlap_cols = [col for col in right.columns if col != key and col in base.columns]
    if not overlap_cols:
        return base.merge(right, on=key, how="left")

    logger.warning(f"Join overwrite columns detected: source_table={source_table}, columns={overlap_cols}")
    renamed = {col: f"__overwrite__{col}" for col in overlap_cols}
    marker = "__right_matched__"
    right_to_merge = right.rename(columns=renamed).copy()
    right_to_merge[marker] = True
    merged = base.merge(right_to_merge, on=key, how="left")
    matched = merged[marker].eq(True)

    for col in overlap_cols:
        incoming_col = renamed[col]
        merged[col] = merged[incoming_col].where(matched, merged[col])
        merged = merged.drop(columns=[incoming_col])
    return merged.drop(columns=[marker])


def enrich_base_sample(tables: dict[str, pd.DataFrame], cfg: dict[str, Any], logger) -> pd.DataFrame:
    key = cfg["keys"]["primary_key"]
    base_table = cfg["analysis"]["base_table"]
    base = tables[base_table].copy()
    if key not in base.columns:
        raise ValueError(f"Base table missing key column: {key}")

    logger.info(f"Base sample rows={len(base):,}, cols={len(base.columns):,}")
    for join_cfg in cfg["joins"].values():
        source_table = join_cfg["source_table"]
        join_key = join_cfg["join_key"]
        if join_key != key:
            raise ValueError(f"Join key must match primary key. primary_key={key}, join_key={join_key}")

        right = tables[source_table][[join_key] + join_cfg["fields"]].copy()
        duplicate_rows = int(right.duplicated(subset=[join_key], keep=False).sum())
        if duplicate_rows:
            logger.warning(f"Join table has duplicate keys; keeping first: table={source_table}, rows={duplicate_rows:,}")
            right = right.drop_duplicates(subset=[join_key], keep="first")

        before_rows = len(base)
        base = _merge_with_overwrite(base, right, key, source_table, logger)
        if len(base) != before_rows:
            raise ValueError(f"Join changed row count: table={source_table}, before={before_rows}, after={len(base)}")
        logger.info(f"Joined table={source_table}, rows={len(base):,}, cols={len(base.columns):,}")
    return base
