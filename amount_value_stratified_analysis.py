"""同额度分层验证（汇报文档第六节）：价值信号在额度之外。

按固定步长切分件均总额度（total_amount），在各额度区间内部比较价值分箱 1-5 的
平均收入与 3M30 逾期率，检验价值信号是否存在于额度之外。

用法：python amount_value_stratified_analysis.py

输出（写入 output/model_analysis_20260429/）：
- amount_value_stratified_pivot.xlsx
    Sheet「样本数 / 平均收入 / 3M30逾期率 / 3M30有效样本数」：
    行=额度区间，列=价值分箱，含合计。价值分箱 -1（无有效分数）样本不计入。
- amount_value_income_by_bin.png     平均收入多线图
- amount_value_duedate30_by_bin.png  3M30逾期率多线图

说明：3M30 逾期标志仅对放款后有 3 个月表现期的申请有值（全样本中仅约 1.3 万条，
集中在件均总额度 2500 以下），高额度段逾期率无法计算属数据现实，须在报告中披露。
"""

from __future__ import annotations

import logging
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from config import CONFIG
from model_bin_analysis import (
    build_analysis_dataset,
    calculate_group_metrics,
    resolve_config_paths,
    setup_logger,
)

PRIMARY_BIN = "primary_model_score_manual_bin"
AMOUNT_FIELD = "total_amount"
AMOUNT_BUCKET = "amount_bucket"
STEP = 500
TAIL_START = 5000
BIN_LABELS = [1, 2, 3, 4, 5]

CNT_METRIC = "sample_cnt"
INCOME_METRIC = "avg_total_income"
RISK_METRIC = "duedate_3m_30_bad_rate"
RISK_VALID_METRIC = "duedate_3m_30_valid_cnt"

OUTPUT_XLSX = "amount_value_stratified_pivot.xlsx"
OUTPUT_INCOME_PNG = "amount_value_income_by_bin.png"
OUTPUT_RISK_PNG = "amount_value_duedate30_by_bin.png"


def build_amount_bucket(amount: pd.Series) -> pd.Series:
    bins = list(range(0, TAIL_START + STEP, STEP))
    labels = [f"{i}-{i + STEP}" for i in range(0, TAIL_START, STEP)]
    bins.append(np.inf)
    labels.append(f"{TAIL_START}+")
    return pd.cut(amount, bins=bins, labels=labels, right=False, include_lowest=True)


def build_pivot(
    long: pd.DataFrame,
    row_total: pd.DataFrame,
    col_total: pd.DataFrame,
    grand: pd.DataFrame,
    metric: str,
) -> pd.DataFrame:
    body = long.pivot(index=AMOUNT_BUCKET, columns=PRIMARY_BIN, values=metric)
    body["合计"] = row_total.set_index(AMOUNT_BUCKET)[metric].reindex(body.index)
    total_row = col_total.set_index(PRIMARY_BIN)[metric].reindex(BIN_LABELS)
    total_row["合计"] = grand.iloc[0][metric]
    return pd.concat([body, pd.DataFrame([total_row], index=["合计"])])


def format_workbook(path: Path, number_formats: dict[str, str]) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path)
    for sheet_name, fmt in number_formats.items():
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = fmt
    wb.save(path)


def plot_metric(df: pd.DataFrame, title: str, ylabel: str, path: Path) -> None:
    fig, ax = plt.subplots(figsize=(10, 5.5))
    x = np.arange(len(df.index))
    for label in BIN_LABELS:
        ax.plot(x, df[label].to_numpy(), marker="o", label=f"价值{label}")
    ax.set_xticks(x, df.index)
    ax.set_xlabel("件均总额度区间")
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.legend(title="价值分箱")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def main() -> None:
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
    plt.rcParams["axes.unicode_minus"] = False

    cfg = resolve_config_paths(CONFIG)
    logger = setup_logger(level=cfg["project"].get("log_level", "INFO"))
    out_dir = Path(cfg["project"]["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)

    logger.info("building analysis dataset")
    df = build_analysis_dataset(cfg, logger)
    df[AMOUNT_BUCKET] = build_amount_bucket(pd.to_numeric(df[AMOUNT_FIELD], errors="coerce"))

    special_bins = cfg["score_binning"]["schemes"][0]["scores"][0].get("null_values", [-1])
    special_cnt = int(df[PRIMARY_BIN].isin(special_bins).sum())
    df = df[~df[PRIMARY_BIN].isin(special_bins)].copy()
    if special_cnt:
        logger.info(f"excluded {special_cnt:,} rows with special value bin {special_bins}")

    metric_groups = ["sample", "risk", "mean"]
    long = calculate_group_metrics(df, [AMOUNT_BUCKET, PRIMARY_BIN], cfg, metric_groups)
    row_total = calculate_group_metrics(df, [AMOUNT_BUCKET], cfg, metric_groups)
    col_total = calculate_group_metrics(df, [PRIMARY_BIN], cfg, metric_groups)
    grand = calculate_group_metrics(df, [], cfg, metric_groups)

    bucket_counts = row_total.set_index(AMOUNT_BUCKET)[CNT_METRIC]
    for bucket in bucket_counts.index:
        logger.info(f"amount bucket {bucket}: n={int(bucket_counts[bucket]):,}")
    logger.info(f"total rows: {len(df):,}")
    risk_valid = row_total.set_index(AMOUNT_BUCKET)[RISK_VALID_METRIC]
    thin_buckets = [b for b in risk_valid.index if int(risk_valid[b]) < 100]
    if thin_buckets:
        logger.warning(
            "3M30 valid samples < 100 in buckets: " + ", ".join(str(b) for b in thin_buckets)
        )

    cnt_pivot = build_pivot(long, row_total, col_total, grand, CNT_METRIC)
    income_pivot = build_pivot(long, row_total, col_total, grand, INCOME_METRIC)
    risk_pivot = build_pivot(long, row_total, col_total, grand, RISK_METRIC)
    risk_valid_pivot = build_pivot(long, row_total, col_total, grand, RISK_VALID_METRIC)

    xlsx_path = out_dir / OUTPUT_XLSX
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        cnt_pivot.to_excel(writer, sheet_name="样本数")
        income_pivot.to_excel(writer, sheet_name="平均收入")
        risk_pivot.to_excel(writer, sheet_name="3M30逾期率")
        risk_valid_pivot.to_excel(writer, sheet_name="3M30有效样本数")
    format_workbook(
        xlsx_path,
        {
            "样本数": "#,##0",
            "平均收入": "#,##0.0",
            "3M30逾期率": "0.00%",
            "3M30有效样本数": "#,##0",
        },
    )
    logger.info(f"workbook written: {xlsx_path}")

    plot_metric(
        income_pivot.drop(index="合计", columns="合计"),
        "各额度区间内不同价值分箱的平均收入",
        "平均收入",
        out_dir / OUTPUT_INCOME_PNG,
    )
    plot_metric(
        risk_pivot.drop(index="合计", columns="合计"),
        "各额度区间内不同价值分箱的3M30逾期率",
        "3M30逾期率",
        out_dir / OUTPUT_RISK_PNG,
    )
    logger.info(f"plots written: {out_dir}")


if __name__ == "__main__":
    main()
