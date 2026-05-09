from __future__ import annotations

from typing import Any

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


def _normalize_bin_label(value: Any) -> str:
    if value is None or pd.isna(value):
        return "__NA__"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _resolve_group_indexes(group_cfg: Any, bins: list[dict[str, Any]]) -> tuple[Any, list[int]]:
    group_label = None
    source_values = group_cfg
    source_mode = "auto"

    if isinstance(group_cfg, dict):
        group_label = group_cfg.get("label")
        if "source_bin_indexes" in group_cfg:
            source_values = group_cfg.get("source_bin_indexes")
            source_mode = "index"
        elif "source_bin_labels" in group_cfg:
            source_values = group_cfg.get("source_bin_labels")
            source_mode = "label"
        elif "bins" in group_cfg:
            source_values = group_cfg.get("bins")

    if not isinstance(source_values, (list, tuple)) or not source_values:
        raise ValueError(f"Each bin group must contain at least one source bin, got {group_cfg}")

    label_to_index = {
        _normalize_bin_label(bin_cfg.get("label")): index
        for index, bin_cfg in enumerate(bins)
    }
    values = list(source_values)
    indexes: list[int] = []

    if source_mode in {"auto", "label"}:
        missing_labels = [
            value for value in values
            if _normalize_bin_label(value) not in label_to_index
        ]
        if not missing_labels:
            indexes = [label_to_index[_normalize_bin_label(value)] for value in values]
        elif source_mode == "label":
            raise ValueError(f"bin group references unknown source_bin_labels={missing_labels}")

    if not indexes:
        try:
            indexes = [int(value) - 1 for value in values]
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"bin group values must match source labels or 1-based indexes, got {values}"
            ) from exc

    invalid_indexes = [index + 1 for index in indexes if index < 0 or index >= len(bins)]
    if invalid_indexes:
        raise ValueError(
            f"bin group references source_bin_indexes outside 1..{len(bins)}: {invalid_indexes}"
        )

    return group_label, indexes


def _build_grouped_bins(
    bins: list[dict[str, Any]],
    bin_groups: list[Any],
    label_type: str | None,
) -> list[dict[str, Any]]:
    collapsed: list[dict[str, Any]] = []
    covered_indexes: list[int] = []
    expected_next_index = 0

    for group_index, group_cfg in enumerate(bin_groups):
        configured_label, indexes = _resolve_group_indexes(group_cfg, bins)
        if indexes != list(range(indexes[0], indexes[-1] + 1)):
            raise ValueError(f"bin group must contain adjacent source bins, got indexes={[i + 1 for i in indexes]}")
        if indexes[0] != expected_next_index:
            raise ValueError(
                "bin groups must cover source bins once and in order; "
                f"expected next source_bin_index={expected_next_index + 1}, got {indexes[0] + 1}"
            )
        expected_next_index = indexes[-1] + 1
        covered_indexes.extend(indexes)

        chunk = [bins[index] for index in indexes]
        labels = [item.get("label") for item in chunk if item.get("label") is not None]
        label = configured_label
        if label is None:
            label = _build_collapsed_label(group_index, label_type, labels)
        collapsed.append(
            {
                "label": label,
                "min_score": chunk[0].get("min_score"),
                "max_score": chunk[-1].get("max_score"),
            }
        )

    expected_indexes = list(range(len(bins)))
    if covered_indexes != expected_indexes:
        missing = sorted(set(expected_indexes) - set(covered_indexes))
        extra = sorted(index for index in covered_indexes if covered_indexes.count(index) > 1)
        raise ValueError(
            "bin groups must cover every configured source bin exactly once; "
            f"missing_source_bin_indexes={[i + 1 for i in missing]}, "
            f"duplicate_source_bin_indexes={[i + 1 for i in extra]}"
        )

    return collapsed


def _build_effective_bins(
    bins: list[dict[str, Any]],
    label_type: str | None,
    bin_groups: list[Any] | None,
) -> list[dict[str, Any]]:
    if not bins:
        return []
    if not bin_groups:
        raise ValueError("Each score config must define manual bin_groups.")
    return _build_grouped_bins(bins, bin_groups, label_type)


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
        max_score = b.get("max_score", float("inf"))
        if str(binning_mode).lower() == "range":
            min_score = b.get("min_score", float("-inf"))
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

    for score_cfg in bin_cfg.get("scores", []):
        score_field = score_cfg.get("score_field")
        bin_field = score_cfg.get("bin_field")
        bins = score_cfg.get("bins", [])
        label_type = score_cfg.get("bin_label_type")
        binning_mode = score_cfg.get("binning_mode", default_binning_mode)
        bin_groups = score_cfg.get("bin_groups")
        score_null_label = score_cfg.get("null_label", null_label)
        score_null_values = score_cfg.get("null_values")
        else_label = score_cfg.get("else_label")
        if not score_field or not bin_field:
            logger.warning(f"Invalid score config skipped: {score_cfg}")
            continue
        if score_field not in out.columns:
            logger.warning(f"Score field missing, bin field will be null: {score_field}")
        effective_bins = _build_effective_bins(
            bins,
            label_type,
            bin_groups,
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
