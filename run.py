from __future__ import annotations

from config import CONFIG
from model_bin_analysis import (
    build_analysis_dataset,
    export_feature_mean_profile_workbook,
    export_risk_performance_workbook,
    export_user_profile_workbook,
    prepare_output_dir,
    resolve_config_paths,
    setup_logger,
)


PRIMARY_BIN = "primary_model_score_manual_bin"
COMPARISON_BIN = "comparison_model_score_manual_bin"


def main() -> None:
    cfg = resolve_config_paths(CONFIG)
    logger = setup_logger(level=cfg["project"].get("log_level", "INFO"))
    output_dir = prepare_output_dir(cfg)
    output_files = cfg["output_files"]

    logger.info("Model bin analysis started")
    analysis_df = build_analysis_dataset(cfg, logger)

    # 输出顺序固定：风险表现 -> 特征均值画像 -> 用户画像。
    export_risk_performance_workbook(
        output_path=output_dir / output_files["risk_performance"],
        df=analysis_df,
        cfg=cfg,
        row_bin=PRIMARY_BIN,
        column_bin=COMPARISON_BIN,
        logger=logger,
    )
    export_feature_mean_profile_workbook(
        output_path=output_dir / output_files["feature_mean_profile"],
        df=analysis_df,
        cfg=cfg,
        primary_bin_col=PRIMARY_BIN,
        comparison_bin_col=COMPARISON_BIN,
        logger=logger,
    )
    export_user_profile_workbook(
        output_path=output_dir / output_files["user_profile"],
        df=analysis_df,
        cfg=cfg,
        row_bin=PRIMARY_BIN,
        column_bin=COMPARISON_BIN,
        logger=logger,
    )
    logger.info("Model bin analysis finished")


if __name__ == "__main__":
    main()
