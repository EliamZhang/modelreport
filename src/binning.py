from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd


def _cast_bin_label(series: pd.Series, label_type: str | None) -> pd.Series:
    if label_type == "int":
        return pd.to_numeric(series, errors="coerce").astype("Int64")
    if label_type == "float":
        return pd.to_numeric(series, errors="coerce")
    return series


def _build_collapsed_label(chunk_index: int, label_type: str | None, source_labels: list[Any]) -> Any:
    if label_type == "int":
        return chunk_index + 1
    if label_type == "float":
        return float(chunk_index + 1)
    if not source_labels:
        return str(chunk_index + 1)
    if len(source_labels) == 1:
        return str(source_labels[0])
    return f"{source_labels[0]}_{source_labels[-1]}"


def _collapse_bins(
    bins: list[dict[str, Any]],
    output_bin_count: int | None,
    label_type: str | None,
) -> list[dict[str, Any]]:
    if not bins or output_bin_count in {None, len(bins)}:
        return bins

    if output_bin_count <= 0:
        raise ValueError(f"output_bin_count must be positive, got {output_bin_count}")
    if output_bin_count > len(bins):
        raise ValueError(
            f"output_bin_count={output_bin_count} cannot exceed configured bin count={len(bins)}"
        )
    if len(bins) % output_bin_count != 0:
        raise ValueError(
            f"Configured bin count={len(bins)} cannot be evenly collapsed into output_bin_count={output_bin_count}"
        )

    chunk_size = len(bins) // output_bin_count
    collapsed: list[dict[str, Any]] = []
    for chunk_index in range(output_bin_count):
        start = chunk_index * chunk_size
        end = start + chunk_size
        chunk = bins[start:end]
        first = chunk[0]
        last = chunk[-1]
        labels = [item.get("label") for item in chunk if item.get("label") is not None]
        collapsed.append(
            {
                "label": _build_collapsed_label(chunk_index, label_type, labels),
                "min_score": first.get("min_score", -np.inf),
                "max_score": last.get("max_score", np.inf),
            }
        )
    return collapsed


def _collapse_else_label(
    else_label: Any,
    original_bins: list[dict[str, Any]],
    collapsed_bins: list[dict[str, Any]],
) -> Any:
    if else_label is None or not original_bins or not collapsed_bins:
        return else_label
    if len(original_bins) == len(collapsed_bins):
        return else_label

    original_last_label = original_bins[-1].get("label")
    if else_label == original_last_label:
        return collapsed_bins[-1].get("label")
    return else_label


def apply_range_binning(
    df: pd.DataFrame,
    score_field: str,
    bin_field: str,
    bins: list[dict[str, Any]],
    label_type: str | None = None,
    unknown_label: Any = "UNKNOWN",
    null_label: Any = None,
    null_values: list[Any] | None = None,
    else_label: Any = None,
    binning_mode: str = "upper_bound",
) -> pd.DataFrame:
    """Apply configurable range-based score binning.

    Supported modes:
    - upper_bound: emulate SQL CASE WHEN score <= max_score THEN label ... ELSE else_label END
    - range: score >= min_score and score <= max_score matches the configured label

    Rows with NULL score or configured null-like values use null_label.
    Non-null scores not matched use else_label/unknown_label.
    """
    out = df.copy()
    if score_field not in out.columns:
        out[bin_field] = pd.NA
        return out

    score = pd.to_numeric(out[score_field], errors="coerce")
    labels = pd.Series(pd.NA, index=out.index, dtype="object")
    null_mask = score.isna()
    if null_values:
        null_like = pd.to_numeric(pd.Series(null_values), errors="coerce").dropna().tolist()
        if null_like:
            null_mask = null_mask | score.isin(null_like)

    for b in bins:
        label = b.get("label")
        max_score = b.get("max_score", np.inf)
        if str(binning_mode).lower() == "range":
            min_score = b.get("min_score", -np.inf)
            cond = score.ge(min_score) & score.le(max_score)
        else:
            cond = score.le(max_score)
        # Keep first matched bin to avoid boundary overlaps overwriting earlier bins.
        labels = labels.mask(cond & labels.isna(), label)

    fallback_label = else_label
    if fallback_label is None and bins:
        fallback_label = bins[-1].get("label")
    if fallback_label is None:
        fallback_label = unknown_label

    labels = labels.mask(null_mask, null_label)
    labels = labels.mask(~null_mask & labels.isna(), fallback_label)
    if label_type in {"int", "float"}:
        labels = _cast_bin_label(labels, label_type)
    out[bin_field] = labels
    return out


def apply_score_binning(df: pd.DataFrame, cfg: dict[str, Any], logger) -> pd.DataFrame:
    """Apply all score binning configs to a dataframe."""
    out = df.copy()
    bin_cfg = cfg.get("score_binning", {})
    unknown_label = bin_cfg.get("unknown_label", "UNKNOWN")
    null_label = bin_cfg.get("null_label", None)
    default_binning_mode = bin_cfg.get("binning_mode", "upper_bound")
    default_output_bin_count = bin_cfg.get("output_bin_count")

    for score_cfg in bin_cfg.get("scores", []):
        score_field = score_cfg.get("score_field")
        bin_field = score_cfg.get("bin_field")
        bins = score_cfg.get("bins", [])
        label_type = score_cfg.get("bin_label_type")
        binning_mode = score_cfg.get("binning_mode", default_binning_mode)
        output_bin_count = score_cfg.get("output_bin_count", default_output_bin_count)
        score_null_label = score_cfg.get("null_label", null_label)
        score_null_values = score_cfg.get("null_values")
        else_label = score_cfg.get("else_label")
        if not score_field or not bin_field:
            logger.warning(f"Invalid score config skipped: {score_cfg}")
            continue
        if score_field not in out.columns:
            logger.warning(f"Score field missing, bin field will be null: {score_field}")
        effective_bins = _collapse_bins(
            bins,
            int(output_bin_count) if output_bin_count is not None else None,
            label_type,
        )
        effective_else_label = _collapse_else_label(else_label, bins, effective_bins)
        out = apply_range_binning(
            out,
            score_field,
            bin_field,
            effective_bins,
            label_type,
            unknown_label,
            score_null_label,
            score_null_values,
            effective_else_label,
            binning_mode,
        )
        fail_cnt = int(out[bin_field].isna().sum()) if bin_field in out.columns else len(out)
        unknown_cnt = int((out[bin_field].astype(str) == str(unknown_label)).sum()) if bin_field in out.columns else 0
        score_nonnull_cnt = int(out[score_field].notna().sum()) if score_field in out.columns else 0
        score_missing_cnt = int(out[score_field].isna().sum()) if score_field in out.columns else len(out)
        score_name = score_cfg.get("name", score_field)
        logger.info(
            "Applied binning: "
            f"model={score_name}, score_field={score_field}, bin_field={bin_field}, "
            f"mode={binning_mode}, configured_bin_cnt={len(bins):,}, effective_bin_cnt={len(effective_bins):,}, "
            f"score_nonnull_cnt={score_nonnull_cnt:,}, "
            f"score_missing_cnt={score_missing_cnt:,}, bin_null_cnt={fail_cnt:,}, unknown_cnt={unknown_cnt:,}"
        )
    return out
