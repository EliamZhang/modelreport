from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd

from .data_validator import get_primary_key


def _flatten_feature_config(cfg: dict[str, Any]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    categories = cfg.get("feature_profile", {}).get("categories", {}) or {}
    for category, category_cfg in categories.items():
        desc = category_cfg.get("description", "")
        for field in category_cfg.get("fields", []) or []:
            items.append({"category": category, "category_description": desc, "feature": field})
    return items


def _dedup_feature_table(feature_df: pd.DataFrame, key: str, logger) -> pd.DataFrame:
    if key not in feature_df.columns:
        raise ValueError(f"Feature table missing join key: {key}")
    dup_cnt = int(feature_df.duplicated(subset=[key], keep=False).sum())
    if dup_cnt:
        logger.warning(f"Feature table duplicate key rows={dup_cnt:,}; keeping first by key")
        feature_df = feature_df.drop_duplicates(subset=[key], keep="first")
    return feature_df


def build_feature_profile(
    enriched_df: pd.DataFrame,
    tables: dict[str, pd.DataFrame],
    cfg: dict[str, Any],
    logger,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    """Build feature profile by model bin and by category."""
    fp_cfg = cfg.get("feature_profile", {})
    if not fp_cfg.get("enabled", True):
        return pd.DataFrame(), pd.DataFrame(), {}

    key = fp_cfg.get("join_key") or get_primary_key(cfg)
    source_table = fp_cfg.get("source_table", "feature_detail")
    if source_table not in tables:
        logger.warning(f"Feature profile skipped; source table not loaded: {source_table}")
        return pd.DataFrame(), pd.DataFrame(), {}

    group_by = fp_cfg.get("group_by", ["primary_model_score_bin"])
    missing_group_cols = [c for c in group_by if c not in enriched_df.columns]
    if missing_group_cols:
        logger.warning(f"Feature profile skipped; missing group cols: {missing_group_cols}")
        return pd.DataFrame(), pd.DataFrame(), {}

    feature_items = _flatten_feature_config(cfg)
    configured_features = [x["feature"] for x in feature_items]
    feature_meta = {x["feature"]: x for x in feature_items}

    feature_df = _dedup_feature_table(tables[source_table].copy(), key, logger)
    available_features = [f for f in configured_features if f in feature_df.columns]
    missing_features = [f for f in configured_features if f not in feature_df.columns]
    if missing_features:
        logger.warning(f"Feature columns missing skipped, count={len(missing_features):,}; first_30={missing_features[:30]}")
    if not available_features:
        logger.warning("No configured feature columns found; feature profile output will be empty")
        return pd.DataFrame(), pd.DataFrame(), {}

    base_cols = [key] + group_by
    base = enriched_df[base_cols].copy()
    merged = base.merge(feature_df[[key] + available_features], on=key, how="left")
    logger.info(f"Feature profile merged rows={len(merged):,}, available_features={len(available_features):,}")

    percentiles = fp_cfg.get("percentiles", [0.25, 0.5, 0.75, 0.9])
    records: list[dict[str, Any]] = []

    grouped = merged.groupby(group_by, dropna=False, sort=True)
    for group_keys, g in grouped:
        if not isinstance(group_keys, tuple):
            group_keys = (group_keys,)
        group_dict = {col: val for col, val in zip(group_by, group_keys)}
        group_size = len(g)
        for feature in available_features:
            values = pd.to_numeric(g[feature], errors="coerce")
            meta = feature_meta.get(feature, {"category": "UNKNOWN", "category_description": ""})
            row: dict[str, Any] = {
                **group_dict,
                "category": meta.get("category"),
                "category_description": meta.get("category_description"),
                "feature": feature,
                "sample_cnt": group_size,
                "nonnull_cnt": int(values.notna().sum()),
                "missing_cnt": int(values.isna().sum()),
                "missing_rate": float(values.isna().mean()) if group_size else None,
                "mean": float(values.mean()) if values.notna().any() else None,
                "median": float(values.median()) if values.notna().any() else None,
                "std": float(values.std()) if values.notna().sum() > 1 else None,
                "min": float(values.min()) if values.notna().any() else None,
                "max": float(values.max()) if values.notna().any() else None,
            }
            for p in percentiles:
                pname = f"p{int(float(p) * 100):02d}"
                row[pname] = float(values.quantile(float(p))) if values.notna().any() else None
            records.append(row)

    by_bin = pd.DataFrame(records)
    if by_bin.empty:
        return by_bin, pd.DataFrame(), {}

    # Same long format sorted by category for direct business reading.
    by_category = by_bin.sort_values(["category"] + group_by + ["feature"]).reset_index(drop=True)

    split: dict[str, pd.DataFrame] = {
        cat: g.reset_index(drop=True)
        for cat, g in by_category.groupby("category", dropna=False, sort=True)
    }
    logger.info(f"Feature profile done: rows={len(by_bin):,}, categories={len(split):,}")
    return by_bin, by_category, split
