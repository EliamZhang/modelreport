from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd


def safe_divide(numerator: float | int | None, denominator: float | int | None) -> float | None:
    if denominator is None or pd.isna(denominator) or denominator == 0:
        return None
    if numerator is None or pd.isna(numerator):
        return None
    return float(numerator) / float(denominator)


def _to_numeric_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return pd.to_numeric(df[col], errors="coerce")


def _calc_one_group(g: pd.DataFrame, total_rows: int, cfg: dict[str, Any], metric_groups: set[str] | None = None) -> dict[str, Any]:
    metric_groups = metric_groups or {"sample", "risk", "amount_risk", "mean", "conversion"}
    row: dict[str, Any] = {}

    if "sample" in metric_groups:
        row["sample_cnt"] = len(g)
        row["sample_pct"] = safe_divide(len(g), total_rows)

    if "risk" in metric_groups:
        for m in cfg.get("risk_metrics", {}).get("label_metrics", []):
            field = m.get("field")
            prefix = m.get("prefix", field)
            values = _to_numeric_series(g, field)
            bad_values = set(m.get("bad_values", [1]))
            good_values = set(m.get("good_values", [0]))
            bad_cnt = int(values.isin(bad_values).sum())
            good_cnt = int(values.isin(good_values).sum())
            valid_cnt = bad_cnt + good_cnt
            row[f"{prefix}_valid_cnt"] = valid_cnt
            row[f"{prefix}_bad_cnt"] = bad_cnt
            row[f"{prefix}_good_cnt"] = good_cnt
            row[f"{prefix}_bad_rate"] = safe_divide(bad_cnt, valid_cnt)

    if "amount_risk" in metric_groups:
        for m in cfg.get("risk_metrics", {}).get("amount_overdue_metrics", []):
            prefix = m.get("prefix", m.get("name"))
            dpd = _to_numeric_series(g, m.get("dpd_field"))
            numerator = _to_numeric_series(g, m.get("numerator_field"))
            denominator = _to_numeric_series(g, m.get("denominator_field"))
            valid_denominator = denominator.where(denominator > 0)
            overdue_mask = dpd.ge(float(m.get("overdue_threshold", 0))) & valid_denominator.notna()
            numerator_sum = float(numerator.where(overdue_mask, 0).fillna(0).sum())
            denominator_sum = float(valid_denominator.fillna(0).sum())
            row[f"{prefix}_overdue_amount"] = numerator_sum
            row[f"{prefix}_principal_amount"] = denominator_sum
            row[f"{prefix}_overdue_rate"] = safe_divide(numerator_sum, denominator_sum)

    if "mean" in metric_groups:
        for m in cfg.get("mean_metrics", []):
            src = m.get("source_field")
            out = m.get("output_field", f"avg_{src}")
            row[out] = float(_to_numeric_series(g, src).mean()) if src else None

    if "conversion" in metric_groups:
        fields = cfg.get("conversion", {}).get("output_fields", {})
        completed_col = fields.get("completed_flag", "is_completed_application")
        approved_col = fields.get("approved_flag", "is_approved_application")
        auto_col = fields.get("auto_approved_flag", "is_auto_approved_application")
        manual_col = fields.get("manual_approved_flag", "is_manual_approved_application")
        deal_col = fields.get("deal_flag", "is_deal_application")

        apply_cnt = len(g)
        completed_cnt = int(_to_numeric_series(g, completed_col).eq(1).sum())
        approved_cnt = int(_to_numeric_series(g, approved_col).eq(1).sum())
        auto_cnt = int(_to_numeric_series(g, auto_col).eq(1).sum())
        manual_cnt = int(_to_numeric_series(g, manual_col).eq(1).sum())
        deal_cnt = int(_to_numeric_series(g, deal_col).eq(1).sum())

        row.update({
            "apply_cnt": apply_cnt,
            "completed_application_cnt": completed_cnt,
            "approved_application_cnt": approved_cnt,
            "auto_approved_application_cnt": auto_cnt,
            "manual_approved_application_cnt": manual_cnt,
            "deal_sample_cnt": deal_cnt,
            "completion_rate": safe_divide(completed_cnt, apply_cnt),
            "approval_rate": safe_divide(approved_cnt, completed_cnt),
            "auto_approval_rate": safe_divide(auto_cnt, completed_cnt),
            "manual_approval_rate": safe_divide(manual_cnt, completed_cnt),
            "auto_approval_share": safe_divide(auto_cnt, approved_cnt),
            "manual_approval_share": safe_divide(manual_cnt, approved_cnt),
            "deal_rate": safe_divide(deal_cnt, approved_cnt),
        })

    return row


def calculate_group_metrics(
    df: pd.DataFrame,
    group_cols: list[str],
    cfg: dict[str, Any],
    include_metric_groups: list[str] | None = None,
) -> pd.DataFrame:
    """Calculate configurable metrics by group columns."""
    missing_group_cols = [c for c in group_cols if c not in df.columns]
    if missing_group_cols:
        raise ValueError(f"Group columns missing: {missing_group_cols}")

    metric_groups = set(include_metric_groups) if include_metric_groups else None
    total_rows = len(df)
    records: list[dict[str, Any]] = []

    if not group_cols:
        records.append(_calc_one_group(df, total_rows, cfg, metric_groups))
        return pd.DataFrame(records)

    grouped = df.groupby(group_cols, dropna=False, sort=True)
    for keys, g in grouped:
        if not isinstance(keys, tuple):
            keys = (keys,)
        row = {col: val for col, val in zip(group_cols, keys)}
        row.update(_calc_one_group(g, total_rows, cfg, metric_groups))
        records.append(row)

    return pd.DataFrame(records)
