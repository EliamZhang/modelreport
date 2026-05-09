from __future__ import annotations

from typing import Any

import pandas as pd


def add_sample_month(df: pd.DataFrame, cfg: dict[str, Any], logger) -> pd.DataFrame:
    month_col = cfg["analysis"]["sample_month_field"]
    candidates = [cfg["keys"]["datetime_key"], *cfg["keys"].get("datetime_aliases", [])]
    datetime_col = next((col for col in candidates if col in df.columns), None)
    if datetime_col is None:
        raise ValueError(f"None of the datetime columns exists: {candidates}")

    out = df.copy()
    parsed = pd.to_datetime(out[datetime_col], errors="coerce")
    out[month_col] = parsed.dt.strftime("%Y-%m")
    logger.info(f"Derived {month_col} from {datetime_col}, invalid_datetime_cnt={int(parsed.isna().sum()):,}")
    return out
