from __future__ import annotations

from collections import defaultdict
from typing import Any

import pandas as pd

from .metric_calculator import calculate_group_metrics


def _normalize_merged_output(
    result: pd.DataFrame,
    group_by: list[str],
    analysis_cfg: dict[str, Any],
) -> pd.DataFrame:
    merge_id_columns = analysis_cfg.get("merge_id_columns")
    if merge_id_columns is None:
        merge_id_columns = group_by[:1] if group_by else []
    merge_id_columns = list(merge_id_columns)

    missing_id_columns = [c for c in merge_id_columns if c not in result.columns]
    if missing_id_columns:
        raise ValueError(f"Merge id columns missing in result: {missing_id_columns}")

    feature_columns = [c for c in group_by if c not in merge_id_columns]
    if not feature_columns:
        raise ValueError(
            f"Merged output requires at least one non-id group column; analysis={analysis_cfg.get('name')}"
        )

    normalized = result.copy()
    feature_label = analysis_cfg.get("merge_group_label")

    if len(feature_columns) == 1:
        feature_col = feature_columns[0]
        normalized = normalized.rename(columns={feature_col: "group_value"})
        normalized.insert(len(merge_id_columns), "group_feature", feature_label or feature_col)
        return normalized

    normalized.insert(len(merge_id_columns), "group_feature", feature_label or "+".join(feature_columns))
    normalized["group_value"] = (
        normalized[feature_columns]
        .fillna("NA")
        .astype(str)
        .agg(" | ".join, axis=1)
    )
    return normalized.drop(columns=feature_columns)


def run_group_analyses(df: pd.DataFrame, cfg: dict[str, Any], logger, output_handler=None) -> dict[str, pd.DataFrame]:
    """Run configured model/bin/cross analyses."""
    outputs: dict[str, pd.DataFrame] = {}
    merged_outputs: dict[str, list[pd.DataFrame]] = defaultdict(list)
    for analysis_cfg in cfg.get("analysis", {}).get("group_analyses", []):
        name = analysis_cfg.get("name")
        group_by = analysis_cfg.get("group_by", []) or []
        output_file = analysis_cfg.get("output_file", f"{name}.csv")
        merged_output_file = analysis_cfg.get("merged_output_file")
        include_groups = analysis_cfg.get("include_metric_groups")

        missing = [c for c in group_by if c not in df.columns]
        if missing:
            logger.warning(f"Analysis skipped={name}; missing group columns={missing}")
            continue

        logger.info(f"Running analysis={name}, group_by={group_by}")
        result = calculate_group_metrics(df, group_by, cfg, include_groups)
        if merged_output_file:
            normalized = _normalize_merged_output(result, group_by, analysis_cfg)
            merged_outputs[merged_output_file].append(normalized)
            logger.info(
                f"Analysis done={name}, rows={len(result):,}, merged_output_file={merged_output_file}"
            )
            continue

        outputs[output_file] = result
        if output_handler is not None:
            output_handler(output_file, result)
        logger.info(f"Analysis done={name}, rows={len(result):,}, output_file={output_file}")

    for merged_output_file, dfs in merged_outputs.items():
        outputs[merged_output_file] = pd.concat(dfs, ignore_index=True)
        if output_handler is not None:
            output_handler(merged_output_file, outputs[merged_output_file])
        logger.info(f"Merged group analysis output prepared: {merged_output_file}, parts={len(dfs):,}")

    return outputs
