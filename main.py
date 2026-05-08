from __future__ import annotations

import argparse
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd

from src.aggregation import enrich_base_sample
from src.binning import apply_score_binning
from src.config_loader import get_output_dir, load_config
from src.cross_analysis import run_group_analyses
from src.data_loader import load_input_tables
from src.data_validator import build_final_quality_records, build_table_quality_records
from src.feature_profile import build_feature_profile
from src.label_builder import apply_deal_amount_filter, build_conversion_labels
from src.logger import setup_logger
from src.monthly_analysis import add_sample_month, run_monthly_analyses
from src.output_writer import write_excel_workbook, write_output, write_split_output
from src.pivot_exporter import generate_pivot_excel
from src.runtime_summary import (
    log_enriched_sample_summary,
    log_month_summary,
    log_post_label_summary,
    log_runtime_plan,
    log_table_overlap_with_base,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Production reusable model analysis package")
    parser.add_argument("--config", required=True, help="Path to Python config, e.g. config/analysis_config_input_sample.py")
    return parser.parse_args()


def prepare_output_dir(cfg: dict) -> Path:
    out_dir = get_output_dir(cfg)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not cfg.get("runtime", {}).get("overwrite_output", True):
        return out_dir

    project_root = Path(cfg.get("project", {}).get("root_dir", ".")).expanduser().resolve()
    try:
        out_dir.relative_to(project_root)
    except ValueError as exc:
        raise ValueError(
            f"Refuse to clear output_dir outside project root when overwrite_output=True: {out_dir}"
        ) from exc

    for child in out_dir.iterdir():
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()

    return out_dir


def main() -> None:
    args = parse_args()
    cfg = load_config(args.config)
    out_dir = prepare_output_dir(cfg)
    log_file = cfg.get("outputs", {}).get("run_log", "run_log.txt")
    logger = setup_logger(out_dir, log_file=log_file, level=cfg.get("runtime", {}).get("log_level", "INFO"))

    start_ts = datetime.now()
    logger.info("=" * 90)
    logger.info("Model analysis pipeline started")
    logger.info(f"Config path: {Path(args.config).resolve()}")
    logger.info(f"Output dir: {out_dir}")
    log_runtime_plan(cfg, logger)

    written_files: list[Path] = []
    workbook_outputs: dict[str, pd.DataFrame] = {}

    try:
        # 1. Load data
        logger.info("[1/8] Loading input tables")
        tables = load_input_tables(cfg, logger)
        logger.info(f"Loaded tables: {', '.join(sorted(tables.keys()))}")
        log_table_overlap_with_base(
            tables,
            cfg.get("analysis", {}).get("base_table", "base_sample"),
            cfg.get("keys", {}).get("primary_key", "application_id"),
            logger,
        )
        quality_records = build_table_quality_records(tables, cfg)

        # 2. Enrich base sample: score table + customer profile + risk score
        logger.info("[2/8] Enriching base sample")
        enriched, merge_stats = enrich_base_sample(tables, cfg, logger)
        base_table_name = cfg.get("analysis", {}).get("base_table", "base_sample")
        expected_base_rows = len(tables[base_table_name]) if base_table_name in tables else len(enriched)
        log_enriched_sample_summary(enriched, cfg, logger, "Base sample prepared")
        if len(enriched) != expected_base_rows:
            raise ValueError(
                f"Base sample row count changed after joins: expected={expected_base_rows}, actual={len(enriched)}"
            )

        # 3. Score binning
        logger.info("[3/8] Applying score binning")
        enriched = apply_score_binning(enriched, cfg, logger)
        enriched = add_sample_month(enriched, cfg, logger)
        log_enriched_sample_summary(enriched, cfg, logger, "Post score binning sample summary")
        log_month_summary(enriched, cfg, logger)

        # 4. Conversion labels and deal amount filter
        logger.info("[4/8] Building conversion labels and filters")
        enriched = build_conversion_labels(enriched, cfg, logger)
        enriched = apply_deal_amount_filter(enriched, cfg, logger)
        log_post_label_summary(enriched, cfg, logger)

        # 5. Data quality summary after all enrichment
        logger.info("[5/8] Building data quality summary")
        if len(enriched) != expected_base_rows:
            raise ValueError(
                f"Final enriched sample row count changed unexpectedly: expected={expected_base_rows}, actual={len(enriched)}"
            )
        log_enriched_sample_summary(enriched, cfg, logger, "Final enriched sample summary")
        quality_records.extend(build_final_quality_records(enriched, cfg, merge_stats))
        quality_df = pd.DataFrame(quality_records)

        outputs: dict[str, pd.DataFrame] = {}
        output_cfg = cfg.get("outputs", {})
        enabled = cfg.get("analysis", {}).get("enabled_outputs", {})

        if enabled.get("data_quality_summary", True):
            quality_file = output_cfg.get("data_quality_summary", "00_data_quality_summary.csv")
            outputs[quality_file] = quality_df
            written_files.append(write_output(quality_file, quality_df, cfg, logger))

        if enabled.get("enriched_base_sample", True):
            enriched_file = output_cfg.get("enriched_base_sample", "01_enriched_base_sample.csv")
            outputs[enriched_file] = enriched
            written_files.append(write_output(enriched_file, enriched, cfg, logger))

        workbook_outputs.update(outputs)

        def stream_output(file_name: str, df: pd.DataFrame) -> None:
            workbook_outputs[file_name] = df
            written_files.append(write_output(file_name, df, cfg, logger))

        # 6. Group and cross analyses
        logger.info("[6/8] Running group and cross analyses")
        group_outputs = run_group_analyses(enriched, cfg, logger, output_handler=stream_output)
        logger.info(f"Group analysis outputs prepared: {len(group_outputs):,}")

        monthly_outputs = run_monthly_analyses(enriched, cfg, logger, output_handler=stream_output)
        logger.info(f"Monthly analysis outputs prepared: {len(monthly_outputs):,}")
        workbook_outputs.update(group_outputs)
        workbook_outputs.update(monthly_outputs)

        # 7. Feature profile
        logger.info("[7/8] Building feature profile outputs")
        feature_outputs_to_write: dict[str, pd.DataFrame] = {}
        feature_split_outputs: dict[str, pd.DataFrame] = {}
        fp_cfg = cfg.get("feature_profile", {})
        if fp_cfg.get("enabled", True):
            by_bin, by_category, split = build_feature_profile(enriched, tables, cfg, logger)
            if enabled.get("feature_profile_by_group", True) and not by_bin.empty:
                feature_file_by_bin = fp_cfg.get("output_file_by_bin", "06_feature_profile_by_primary_group.csv")
                feature_outputs_to_write[feature_file_by_bin] = by_bin
                written_files.append(write_output(feature_file_by_bin, by_bin, cfg, logger))
            if enabled.get("feature_profile_by_category", True) and not by_category.empty:
                feature_file_by_category = fp_cfg.get("output_file_by_category", "07_feature_profile_by_category.csv")
                feature_outputs_to_write[feature_file_by_category] = by_category
                written_files.append(write_output(feature_file_by_category, by_category, cfg, logger))
                feature_split_outputs = split
                split_subdir = fp_cfg.get("output_split_by_category_dir", "feature_profile_by_category")
                for name, df in feature_split_outputs.items():
                    written_files.append(write_split_output(name, df, split_subdir, cfg, logger))
        logger.info(
            f"Feature profile prepared: summary_files={len(feature_outputs_to_write):,}, split_files={len(feature_split_outputs):,}"
        )
        workbook_outputs.update(feature_outputs_to_write)

        # 8. Write outputs
        logger.info("[8/8] Writing Excel workbook")
        workbook_outputs.update(
            {
                f"{fp_cfg.get('output_split_by_category_dir', 'feature_profile_by_category')}/{name}.csv": df
                for name, df in feature_split_outputs.items()
            }
        )
        workbook_name = output_cfg.get("excel_workbook", "model_analysis_outputs.xlsx")
        workbook_path = write_excel_workbook(workbook_outputs, workbook_name, cfg, logger)
        if workbook_path is not None:
            written_files.append(workbook_path)

        pivot_cfg = cfg.get("pivot_export", {})
        if pivot_cfg.get("enabled", False):
            input_file = pivot_cfg.get("input_file", "05_primary_vs_comparison_model_bin.csv")
            output_excel = pivot_cfg.get("output_excel", "model_compare_pivot.xlsx")
            input_path = Path(input_file)
            if not input_path.is_absolute():
                input_path = out_dir / input_path
            output_path = Path(output_excel)
            if not output_path.is_absolute():
                output_path = out_dir / output_path

            logger.info(f"Writing pivot Excel workbook: input={input_path}, output={output_path}")
            pivot_path = generate_pivot_excel(
                input_file=str(input_path),
                output_excel=str(output_path),
                row_dim=pivot_cfg.get("row_dim", "primary_model_score_bin"),
                col_dim=pivot_cfg.get("col_dim", "comparison_model_score_bin"),
                metrics=pivot_cfg.get("metrics"),
            )
            written_files.append(pivot_path)

        # 9. Final log
        elapsed = (datetime.now() - start_ts).total_seconds()
        logger.info("Output file list:")
        for p in written_files:
            logger.info(f" - {p}")
        logger.info(f"Pipeline finished successfully, elapsed_seconds={elapsed:.2f}")
        logger.info("=" * 90)

    except Exception as exc:
        logger.exception(f"Pipeline failed: {exc}")
        raise


if __name__ == "__main__":
    main()
