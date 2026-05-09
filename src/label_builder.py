from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd


def build_conversion_labels(df: pd.DataFrame, cfg: dict[str, Any], logger) -> pd.DataFrame:
    """Build configurable conversion flags and conversion stage."""
    conv = cfg.get("conversion", {})
    out = df.copy()
    app_status_col = conv.get("application_status_field", "application_status")
    assess_col = conv.get("assessment_status_field", "assessment_status")
    deal_status_col = conv.get("deal_status_field", "status")
    fields = conv.get("output_fields", {})

    completed_col = fields.get("completed_flag", "is_completed_application")
    approved_col = fields.get("approved_flag", "is_approved_application")
    auto_col = fields.get("auto_approved_flag", "is_auto_approved_application")
    manual_col = fields.get("manual_approved_flag", "is_manual_approved_application")
    deal_col = fields.get("deal_flag", "is_deal_application")
    stage_col = fields.get("conversion_stage", "conversion_stage")

    for col in [app_status_col, assess_col, deal_status_col]:
        if col not in out.columns:
            logger.warning(f"Conversion source column missing: {col}; filling with NA")
            out[col] = pd.NA

    app_status = out[app_status_col].astype("string")
    assess_status = out[assess_col].astype("string")
    deal_status = out[deal_status_col].astype("string")

    incomplete_values = set(map(str, conv.get("incomplete_status_values", ["0.Incomplete", "1.In Progress"])))
    approved_prefixes = tuple(map(str, conv.get("approved_status_prefixes", ["3", "4"])))
    auto_keyword = str(conv.get("auto_approved_keyword", "Auto Approved"))
    manual_keyword = str(conv.get("manual_approved_keyword", "Manual Approved"))
    deal_values = set(map(str, conv.get("deal_status_values", [])))

    out[completed_col] = (~app_status.isin(incomplete_values) & app_status.notna()).astype(int)
    out[approved_col] = (app_status.fillna("").str.slice(0, 1).isin(approved_prefixes)).astype(int)
    out[auto_col] = ((out[approved_col] == 1) & assess_status.fillna("").str.contains(auto_keyword, case=False, regex=False)).astype(int)
    out[manual_col] = ((out[approved_col] == 1) & assess_status.fillna("").str.contains(manual_keyword, case=False, regex=False)).astype(int)
    out[deal_col] = deal_status.isin(deal_values).astype(int)

    conditions = [
        out[completed_col] == 0,
        (out[completed_col] == 1) & (out[approved_col] == 0),
        (out[approved_col] == 1) & (out[deal_col] == 0),
        (out[approved_col] == 1) & (out[deal_col] == 1),
    ]
    choices = [
        "incomplete_or_in_progress",
        "completed_not_approved",
        "approved_not_deal",
        "approved_and_deal",
    ]
    out[stage_col] = np.select(conditions, choices, default="UNKNOWN")

    logger.info(
        "Built conversion labels: "
        f"completed_cnt={int(out[completed_col].sum()):,}, "
        f"approved_cnt={int(out[approved_col].sum()):,}, "
        f"auto_approved_cnt={int(out[auto_col].sum()):,}, "
        f"manual_approved_cnt={int(out[manual_col].sum()):,}, "
        f"deal_cnt={int(out[deal_col].sum()):,}"
    )
    return out


def apply_deal_amount_filter(df: pd.DataFrame, cfg: dict[str, Any], logger) -> pd.DataFrame:
    """Create a deal-only total amount field based on configured deal statuses."""
    rules = cfg.get("amount_rules", {})
    total_col = rules.get("total_amount_field", "total_amount")
    out_col = rules.get("deal_total_amount_field", "deal_total_amount")
    status_col = rules.get("deal_status_field", "status")
    deal_values = set(map(str, rules.get("deal_status_values", [])))

    out = df.copy()
    if total_col not in out.columns:
        logger.warning(f"Amount field missing: {total_col}; {out_col} will be null")
        out[out_col] = np.nan
        return out
    if status_col not in out.columns:
        logger.warning(f"Deal status field missing: {status_col}; {out_col} will be null")
        out[out_col] = np.nan
        return out

    amount = pd.to_numeric(out[total_col], errors="coerce")
    is_deal = out[status_col].astype("string").isin(deal_values)
    out[out_col] = amount.where(is_deal, np.nan)
    logger.info(
        f"Created deal amount field: {out_col}, valid_deal_amount_cnt={int(out[out_col].notna().sum()):,}, "
        f"deal_status_match_cnt={int(is_deal.sum()):,}"
    )
    return out
