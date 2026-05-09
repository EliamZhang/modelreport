from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from scripts.generate_cross_model_pivot_excel import (
    _build_enriched_sample,
    _resolve_output_path,
    write_workbook,
)
from src.config_loader import get_output_dir, load_config
from src.logger import setup_logger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate the cross-model bin pivot workbook.")
    parser.add_argument("--config", required=True, help="Path to Python config, e.g. config/analysis_config_input_sample.py")
    parser.add_argument(
        "--output",
        default=None,
        help="Workbook path. Relative paths are written under the configured output_dir.",
    )
    parser.add_argument(
        "--row-bin",
        default="primary_model_score_bin",
        help="Row bin field for the pivot tables.",
    )
    parser.add_argument(
        "--column-bin",
        default="comparison_model_score_bin",
        help="Column bin field for the pivot tables.",
    )
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
    prepare_output_dir(cfg)
    output_path = _resolve_output_path(args.output, cfg)
    logger = setup_logger(
        output_path.parent,
        log_file=None,
        level=cfg.get("runtime", {}).get("log_level", "INFO"),
    )

    logger.info("Cross model pivot workbook generation started")
    logger.info(f"Config path: {Path(args.config).resolve()}")
    logger.info(f"Output path: {output_path}")

    enriched = _build_enriched_sample(cfg, logger)
    missing_bins = [col for col in [args.row_bin, args.column_bin] if col not in enriched.columns]
    if missing_bins:
        raise ValueError(f"Pivot bin columns missing after enrichment: {missing_bins}")

    write_workbook(output_path, enriched, cfg, args.row_bin, args.column_bin, logger)
    logger.info("Cross model pivot workbook generation finished")


if __name__ == "__main__":
    main()
