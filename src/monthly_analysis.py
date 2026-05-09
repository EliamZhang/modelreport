from __future__ import annotations

from typing import Any

import pandas as pd

from .data_validator import resolve_datetime_field


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
