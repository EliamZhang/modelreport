from __future__ import annotations

from typing import Iterable, Any

import pandas as pd


def get_primary_key(cfg: dict[str, Any]) -> str:
    return cfg.get("keys", {}).get("primary_key", "application_id")


def resolve_datetime_field(df: pd.DataFrame, cfg: dict[str, Any]) -> str | None:
    keys = cfg.get("keys", {})
    candidates = [keys.get("datetime_key")] + list(keys.get("datetime_aliases", []))
    for col in candidates:
        if col and col in df.columns:
            return col
    return None


def missing_columns(df: pd.DataFrame, cols: Iterable[str]) -> list[str]:
    return [c for c in cols if c not in df.columns]


def validate_columns(
    df: pd.DataFrame,
    cols: Iterable[str],
    table_name: str,
    logger,
    fail: bool = False,
) -> list[str]:
    missing = missing_columns(df, cols)
    if missing:
        msg = f"Table={table_name} missing columns: {missing}"
        if fail:
            raise ValueError(msg)
        logger.warning(msg)
    return missing


def count_duplicate_keys(df: pd.DataFrame, key: str) -> int:
    if key not in df.columns:
        return len(df)
    return int(df.duplicated(subset=[key], keep=False).sum())


def build_table_quality_records(tables: dict[str, pd.DataFrame], cfg: dict[str, Any]) -> list[dict[str, Any]]:
    key = get_primary_key(cfg)
    records: list[dict[str, Any]] = []
    for table_name, df in tables.items():
        records.extend([
            {"section": "input_table", "table": table_name, "metric": "row_cnt", "value": len(df)},
            {"section": "input_table", "table": table_name, "metric": "column_cnt", "value": len(df.columns)},
            {"section": "input_table", "table": table_name, "metric": "primary_key", "value": key},
            {"section": "input_table", "table": table_name, "metric": "primary_key_missing_cnt", "value": int(df[key].isna().sum()) if key in df.columns else "KEY_NOT_FOUND"},
            {"section": "input_table", "table": table_name, "metric": "primary_key_duplicate_row_cnt", "value": count_duplicate_keys(df, key)},
        ])
    return records


def build_final_quality_records(enriched_df: pd.DataFrame, cfg: dict[str, Any], merge_stats: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    records.extend(merge_stats)

    for score_cfg in cfg.get("score_binning", {}).get("scores", []):
        score_field = score_cfg.get("score_field")
        bin_field = score_cfg.get("bin_field")
        if score_field in enriched_df.columns:
            records.append({"section": "score", "table": "enriched_base", "metric": f"{score_field}_missing_cnt", "value": int(enriched_df[score_field].isna().sum())})
        if bin_field in enriched_df.columns:
            records.append({"section": "score", "table": "enriched_base", "metric": f"{bin_field}_missing_or_unknown_cnt", "value": int(enriched_df[bin_field].isna().sum() + (enriched_df[bin_field].astype(str) == "UNKNOWN").sum())})

    critical_fields = []
    critical_fields.extend([m.get("field") for m in cfg.get("risk_metrics", {}).get("label_metrics", []) if m.get("field")])
    critical_fields.extend([m.get("source_field") for m in cfg.get("mean_metrics", []) if m.get("source_field")])
    for analysis_cfg in cfg.get("analysis", {}).get("group_analyses", []) or []:
        critical_fields.extend(analysis_cfg.get("group_by", []) or [])
    conv = cfg.get("conversion", {})
    critical_fields.extend([conv.get("application_status_field"), conv.get("assessment_status_field"), conv.get("deal_status_field")])

    for col in sorted(set([c for c in critical_fields if c])):
        if col in enriched_df.columns:
            records.append({
                "section": "critical_field",
                "table": "enriched_base",
                "metric": f"{col}_missing_rate",
                "value": float(enriched_df[col].isna().mean()) if len(enriched_df) else None,
            })
        else:
            records.append({"section": "critical_field", "table": "enriched_base", "metric": f"{col}_missing_rate", "value": "COLUMN_NOT_FOUND"})

    records.append({"section": "final", "table": "enriched_base", "metric": "row_cnt", "value": len(enriched_df)})
    records.append({"section": "final", "table": "enriched_base", "metric": "column_cnt", "value": len(enriched_df.columns)})
    return records
