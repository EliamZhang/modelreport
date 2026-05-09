from __future__ import annotations

import shutil
from pathlib import Path

from config.analysis_config_input_sample import CONFIG
from src.aggregation import enrich_base_sample
from src.binning import apply_score_binning
from src.config_loader import resolve_config_paths
from src.data_loader import load_input_tables
from src.label_builder import apply_deal_amount_filter, build_conversion_labels
from src.logger import setup_logger
from src.monthly_analysis import add_sample_month
from src.workbook import export_cross_model_bin_user_profile_excel, write_cross_model_workbook


OUTPUT_FILE_NAME = "cross_model_bin_pivot_rate_only.xlsx"
USER_PROFILE_OUTPUT_FILE_NAME = "cross_model_bin_pivot_user_profile.xlsx"


def prepare_output_dir(cfg: dict) -> Path:
    out_dir = Path(cfg["project"]["output_dir"]).resolve()
    project_root = Path(cfg["project"]["root_dir"]).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        out_dir.relative_to(project_root)
    except ValueError as exc:
        raise ValueError(f"Refuse to clear output_dir outside project root: {out_dir}") from exc

    for child in out_dir.iterdir():
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()
    return out_dir


def main() -> None:
    cfg = resolve_config_paths(CONFIG)
    out_dir = prepare_output_dir(cfg)
    output_path = out_dir / OUTPUT_FILE_NAME
    user_profile_output_path = out_dir / USER_PROFILE_OUTPUT_FILE_NAME
    logger = setup_logger(level=cfg["project"].get("log_level", "INFO"))

    logger.info("Cross model pivot workbook generation started")
    logger.info(f"Output path: {output_path}")

    tables = load_input_tables(cfg, logger)
    enriched = enrich_base_sample(tables, cfg, logger)
    enriched = apply_score_binning(enriched, cfg, logger)
    enriched = add_sample_month(enriched, cfg, logger)
    enriched = build_conversion_labels(enriched, cfg, logger)
    enriched = apply_deal_amount_filter(enriched, cfg, logger)

    write_cross_model_workbook(
        output_path=output_path,
        df=enriched,
        cfg=cfg,
        row_bin="primary_model_score_bin",
        column_bin="comparison_model_score_bin",
        logger=logger,
    )
    export_cross_model_bin_user_profile_excel(
        df=enriched,
        cfg=cfg,
        output_path=user_profile_output_path,
        row_bin="primary_model_score_bin",
        column_bin="comparison_model_score_bin",
        logger=logger,
    )
    logger.info("Cross model pivot workbook generation finished")


if __name__ == "__main__":
    main()
