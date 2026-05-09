from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from .metric_calculator import calculate_group_metrics


METRIC_GROUPS: dict[str, list[str]] = {
    "risk_perf": [
        "duedate_3m_30_valid_cnt",
        "duedate_3m_30_bad_cnt",
        "duedate_3m_30_good_cnt",
        "duedate_3m_30_bad_rate",
        "duedate_1m_5_valid_cnt",
        "duedate_1m_5_bad_cnt",
        "duedate_1m_5_good_cnt",
        "duedate_1m_5_bad_rate",
    ],
    "amount_risk": [
        "duedate_3m_30_amount_overdue_amount",
        "duedate_3m_30_amount_principal_amount",
        "duedate_3m_30_amount_overdue_rate",
        "duedate_1m_5_amount_overdue_amount",
        "duedate_1m_5_amount_principal_amount",
        "duedate_1m_5_amount_overdue_rate",
    ],
    "profile": [
        "avg_PTI",
        "avg_requested_loan_amount",
        "avg_total_amount",
        "avg_gross_surplus",
        "avg_net_surplus",
        "avg_total_income",
    ],
    "conversion": [
        "apply_cnt",
        "completed_application_cnt",
        "approved_application_cnt",
        "auto_approved_application_cnt",
        "manual_approved_application_cnt",
        "deal_sample_cnt",
        "completion_rate",
        "approval_rate",
        "auto_approval_rate",
        "manual_approval_rate",
        "auto_approval_share",
        "manual_approval_share",
        "deal_rate",
    ],
}

RATE_DENOMINATORS = {
    "duedate_3m_30_bad_rate": "duedate_3m_30_valid_cnt",
    "duedate_1m_5_bad_rate": "duedate_1m_5_valid_cnt",
    "duedate_3m_30_amount_overdue_rate": "duedate_3m_30_amount_principal_amount",
    "duedate_1m_5_amount_overdue_rate": "duedate_1m_5_amount_principal_amount",
    "completion_rate": "apply_cnt",
    "approval_rate": "completed_application_cnt",
    "auto_approval_rate": "completed_application_cnt",
    "manual_approval_rate": "completed_application_cnt",
    "auto_approval_share": "approved_application_cnt",
    "manual_approval_share": "approved_application_cnt",
    "deal_rate": "approved_application_cnt",
}

COUNT_METRICS = {
    "sample_cnt",
    "duedate_3m_30_valid_cnt",
    "duedate_3m_30_bad_cnt",
    "duedate_3m_30_good_cnt",
    "duedate_1m_5_valid_cnt",
    "duedate_1m_5_bad_cnt",
    "duedate_1m_5_good_cnt",
    "apply_cnt",
    "completed_application_cnt",
    "approved_application_cnt",
    "auto_approved_application_cnt",
    "manual_approved_application_cnt",
    "deal_sample_cnt",
}

AMOUNT_METRICS = {
    "duedate_3m_30_amount_overdue_amount",
    "duedate_3m_30_amount_principal_amount",
    "duedate_1m_5_amount_overdue_amount",
    "duedate_1m_5_amount_principal_amount",
}

FOUR_DECIMAL_METRICS = {"avg_PTI"}


def _is_missing(value: Any) -> bool:
    return value is None or pd.isna(value)


def _label_key(value: Any) -> str:
    if _is_missing(value):
        return "__NA__"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _label_display(value: Any) -> str:
    if _is_missing(value):
        return "NA"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _sort_values(values: list[Any]) -> list[Any]:
    def sort_key(value: Any) -> tuple[int, float, str]:
        if _is_missing(value):
            return (2, 0.0, "")
        try:
            return (0, float(value), str(value))
        except (TypeError, ValueError):
            return (1, 0.0, str(value))

    unique: dict[str, Any] = {}
    for value in values:
        unique.setdefault(_label_key(value), value)
    return sorted(unique.values(), key=sort_key)


def _records_by_key(metrics: pd.DataFrame, group_cols: list[str]) -> dict[tuple[str, ...], pd.Series]:
    records: dict[tuple[str, ...], pd.Series] = {}
    for _, row in metrics.iterrows():
        records[tuple(_label_key(row[col]) for col in group_cols)] = row
    return records


def _format_number(value: Any, metric: str) -> str:
    if _is_missing(value):
        return ""
    value = float(value)
    if metric in COUNT_METRICS:
        return f"{int(round(value)):,}"
    if metric in AMOUNT_METRICS:
        return f"{value:,.2f}"
    if metric in FOUR_DECIMAL_METRICS:
        return f"{value:.4f}"
    if metric.startswith("avg_"):
        return f"{value:,.2f}"
    return f"{value:,.4f}"


def _format_cell(record: pd.Series | None, metric: str) -> str:
    if record is None or metric not in record.index:
        return ""
    value = record[metric]
    if metric in RATE_DENOMINATORS:
        return "" if _is_missing(value) else f"{float(value):.2%}"
    return _format_number(value, metric)


def _build_metric_pivot(
    df: pd.DataFrame,
    cfg: dict[str, Any],
    row_bin: str,
    column_bin: str,
    metric: str,
) -> pd.DataFrame:
    detail = calculate_group_metrics(df, [row_bin, column_bin], cfg)
    row_totals = calculate_group_metrics(df, [row_bin], cfg)
    column_totals = calculate_group_metrics(df, [column_bin], cfg)
    grand_total = calculate_group_metrics(df, [], cfg)

    row_values = _sort_values(list(detail[row_bin]) + list(row_totals[row_bin]))
    column_values = _sort_values(list(detail[column_bin]) + list(column_totals[column_bin]))

    detail_records = _records_by_key(detail, [row_bin, column_bin])
    row_total_records = _records_by_key(row_totals, [row_bin])
    column_total_records = _records_by_key(column_totals, [column_bin])
    grand_record = grand_total.iloc[0] if not grand_total.empty else None

    records: list[dict[str, Any]] = []
    for row_value in row_values:
        row_key = _label_key(row_value)
        record: dict[str, Any] = {row_bin: _label_display(row_value)}
        for column_value in column_values:
            column_key = _label_key(column_value)
            record[_label_display(column_value)] = _format_cell(
                detail_records.get((row_key, column_key)), metric
            )
        record["Total"] = _format_cell(row_total_records.get((row_key,)), metric)
        records.append(record)

    total_record: dict[str, Any] = {row_bin: "Total"}
    for column_value in column_values:
        column_key = _label_key(column_value)
        total_record[_label_display(column_value)] = _format_cell(
            column_total_records.get((column_key,)), metric
        )
    total_record["Total"] = _format_cell(grand_record, metric)
    records.append(total_record)
    return pd.DataFrame(records)


def _write_dataframe(ws, df: pd.DataFrame, start_row: int) -> tuple[int, int]:
    rows = list(dataframe_to_rows(df, index=False, header=True))
    for row_offset, row_values in enumerate(rows):
        for col_offset, value in enumerate(row_values):
            ws.cell(
                row=start_row + row_offset,
                column=1 + col_offset,
                value=None if _is_missing(value) else value,
            )
    return len(rows), len(df.columns)


def _style_table(ws, start_row: int, row_count: int, col_count: int) -> None:
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    total_fill = PatternFill("solid", fgColor="E2F0D9")
    border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    end_row = start_row + row_count - 1

    for col in range(1, col_count + 1):
        cell = ws.cell(row=start_row, column=col)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for row in range(start_row + 1, end_row + 1):
        is_total_row = ws.cell(row=row, column=1).value == "Total"
        for col in range(1, col_count + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if is_total_row or ws.cell(row=start_row, column=col).value == "Total":
                cell.fill = total_fill


def _auto_width(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        max_len = 8
        for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx):
            for item in cell:
                if item.value is not None:
                    max_len = max(max_len, len(str(item.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 32)


def write_cross_model_workbook(
    output_path: Path,
    df: pd.DataFrame,
    cfg: dict[str, Any],
    row_bin: str,
    column_bin: str,
    logger,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "metric_guide"

    guide_df = pd.DataFrame(
        [
            {
                "category": category,
                "metric": metric,
                "cell_format": "percent" if metric in RATE_DENOMINATORS else "value",
                "denominator_metric": RATE_DENOMINATORS.get(metric, ""),
            }
            for category, metrics in METRIC_GROUPS.items()
            for metric in metrics
        ]
    )
    row_count, col_count = _write_dataframe(ws, guide_df, 1)
    _style_table(ws, 1, row_count, col_count)
    ws.freeze_panes = "A2"
    _auto_width(ws)

    for category, metrics in METRIC_GROUPS.items():
        ws = wb.create_sheet(category)
        current_row = 1
        for metric in metrics:
            ws.cell(row=current_row, column=1, value=metric).font = Font(bold=True, size=12)
            current_row += 1
            pivot = _build_metric_pivot(df, cfg, row_bin, column_bin, metric)
            row_count, col_count = _write_dataframe(ws, pivot, current_row)
            _style_table(ws, current_row, row_count, col_count)
            current_row += row_count + 2
        _auto_width(ws)

    raw_outputs = {
        "raw_cross_metrics": calculate_group_metrics(df, [row_bin, column_bin], cfg),
        "raw_row_totals": calculate_group_metrics(df, [row_bin], cfg),
        "raw_column_totals": calculate_group_metrics(df, [column_bin], cfg),
        "raw_grand_total": calculate_group_metrics(df, [], cfg),
    }
    for sheet_name, data in raw_outputs.items():
        ws = wb.create_sheet(sheet_name)
        row_count, col_count = _write_dataframe(ws, data, 1)
        _style_table(ws, 1, row_count, col_count)
        ws.freeze_panes = "C2" if sheet_name == "raw_cross_metrics" else "A2"
        _auto_width(ws)

    wb.save(output_path)
    logger.info(f"Wrote workbook: {output_path}")
