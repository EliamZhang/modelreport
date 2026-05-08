from __future__ import annotations

from pathlib import Path
import sys

import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.config_loader import load_config


def _build_base_table(cfg: dict, app_ids: list[str]) -> pd.DataFrame:
    keys = cfg.get("keys", {})
    score_cfg = cfg.get("score_binning", {}).get("scores", [])[0]
    primary_key = keys.get("primary_key", "application_id")
    user_key = keys.get("user_key", "user_id")
    dt_key = keys.get("datetime_key", "sample_datetime")
    score_field = score_cfg["score_field"]

    scores = [0.02, 0.10, 0.13, 0.17, 0.22, 0.27, 0.31, 0.36, 0.43, 0.51, 0.80, None]
    rows = []
    for i, app_id in enumerate(app_ids, start=1):
        rows.append(
            {
                primary_key: app_id,
                user_key: f"U{i:03d}",
                dt_key: f"2026-04-{(i % 28) + 1:02d} 10:00:00",
                score_field: scores[i - 1],
                "base_note": f"sample_{i:02d}",
            }
        )
    return pd.DataFrame(rows)


def _build_customer_profile(cfg: dict, app_ids: list[str]) -> pd.DataFrame:
    join_cfg = cfg["joins"]["customer_profile"]
    fields = join_cfg.get("fields", [])
    join_key = join_cfg.get("join_key", "application_id")

    last_steps = ["apply", "apply", "review", "review", "approve", "approve", "fund", "fund", "fund", "fund", "close", "close"]
    loan_tags = ["small", "small", "medium", "medium", "large", "large", "small", "medium", "large", "small", "medium", "large"]
    requested_loan_tags = ["r_small", "r_small", "r_medium", "r_medium", "r_large", "r_large", "r_small", "r_medium", "r_large", "r_small", "r_medium", "r_large"]
    states = ["NSW", "VIC", "QLD", "NSW", "VIC", "QLD", "WA", "SA", "TAS", "NSW", "VIC", "QLD"]
    suburbs = ["Sydney", "Melbourne", "Brisbane", "Parramatta", "Geelong", "Gold Coast", "Perth", "Adelaide", "Hobart", "Newcastle", "Ballarat", "Cairns"]
    family_types = ["single", "couple", "family", "single", "couple", "family", "single", "couple", "family", "single", "couple", "family"]
    categories = ["A", "A", "B", "B", "C", "C", "A", "B", "C", "A", "B", "C"]
    app_statuses = [
        "0.Incomplete",
        "1.In Progress",
        "2.Completed",
        "3.Approved",
        "3.Approved",
        "4.Approved",
        "4.Approved",
        "2.Completed",
        "3.Approved",
        "4.Approved",
        "2.Completed",
        "3.Approved",
    ]
    assess_statuses = [
        "Pending",
        "Reviewing",
        "Manual Check",
        "Auto Approved",
        "Manual Approved",
        "Auto Approved",
        "Manual Approved",
        "Rejected",
        "Auto Approved",
        "Manual Approved",
        "Rejected",
        "Auto Approved",
    ]
    statuses = [
        "Pending",
        "Pending",
        "Rejected",
        "Active_Account",
        "Closed",
        "Blocked",
        "Active_Account",
        "Pending",
        "Closed",
        "Blocked",
        "Rejected",
        "Active_Account",
    ]

    rows = []
    for i, app_id in enumerate(app_ids, start=1):
        principal = 1000 + i * 200
        total_amount = 1500 + i * 250
        row = {
            join_key: app_id,
            "last_step": last_steps[i - 1],
            "loan_tag": loan_tags[i - 1],
            "age": 20 + i,
            "base_probability": round(0.05 + i * 0.01, 4),
            "PTI": round(0.15 + i * 0.02, 4),
            "principal": principal,
            "estimate_principal_remaining_mob3": round(principal * (0.2 + (i % 3) * 0.1), 2),
            "dpd_days_mob3": [0, 5, 10, 35, 45, 0, 0, 60, 0, 12, 0, 40][i - 1],
            "estimate_principal_remaining_mob1": round(principal * (0.1 + (i % 4) * 0.05), 2),
            "dpd_days_mob1": [0, 2, 6, 8, 0, 3, 10, 0, 7, 1, 0, 9][i - 1],
            "gross_surplus": 300 + i * 20,
            "net_surplus": 180 + i * 15,
            "total_income": 2800 + i * 120,
            "requested_loan_amount": 800 + i * 100,
            "requested_loan_tag": requested_loan_tags[i - 1],
            "state": states[i - 1],
            "suburb": suburbs[i - 1],
            "family_type": family_types[i - 1],
            "dependents": i % 4,
            "attributed_category": categories[i - 1],
            "total_amount": total_amount,
            "duedate_1m_5": int([0, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 1][i - 1]),
            "duedate_3m_30": int([0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, 1][i - 1]),
            "application_status": app_statuses[i - 1],
            "assessment_status": assess_statuses[i - 1],
            "status": statuses[i - 1],
        }
        rows.append({field: row.get(field) for field in [join_key] + fields})
    return pd.DataFrame(rows)


def _build_risk_score(cfg: dict, app_ids: list[str]) -> pd.DataFrame:
    score_cfg = cfg.get("score_binning", {}).get("scores", [])[1]
    score_field = score_cfg["score_field"]
    join_key = cfg["joins"]["risk_score"].get("join_key", "application_id")
    scores = [0.03, 0.05, 0.065, 0.08, 0.095, 0.12, 0.14, 0.17, 0.20, 0.24, 0.70, None]
    return pd.DataFrame({join_key: app_ids, score_field: scores})


def _build_feature_detail(cfg: dict, app_ids: list[str]) -> pd.DataFrame:
    fp_cfg = cfg.get("feature_profile", {})
    join_key = fp_cfg.get("join_key", "application_id")
    categories = fp_cfg.get("categories", {})
    feature_fields = [field for cat in categories.values() for field in cat.get("fields", [])]

    rows = []
    for i, app_id in enumerate(app_ids, start=1):
        row = {join_key: app_id}
        for j, field in enumerate(feature_fields, start=1):
            value = round(i * 10 + j * 0.1, 4)
            if (i + j) % 17 == 0:
                value = np.nan
            row[field] = value
        rows.append(row)
    return pd.DataFrame(rows)


def _build_txn_industry(cfg: dict, app_ids: list[str]) -> pd.DataFrame:
    join_cfg = cfg["joins"]["txn_industry_features"]
    join_key = join_cfg.get("join_key", "application_id")
    fields = join_cfg.get("fields", [])

    rows = []
    industries = ["Retail", "Tech", "Health", "Retail", "Tech", "Health", "Logistics", "Energy", "Retail", "Tech", "Health", "Logistics"]
    employers = ["EMP_A", "EMP_B", "EMP_C", "EMP_A", "EMP_D", "EMP_E", "EMP_F", "EMP_G", "EMP_H", "EMP_I", "EMP_J", "EMP_K"]
    for i, app_id in enumerate(app_ids, start=1):
        row = {
            join_key: app_id,
            "bank_txn_industry_primary_industry": industries[i - 1],
            "bank_txn_industry_primary_industry_amt": 1200 + i * 80,
            "bank_txn_industry_pay_interval_median": 14 + (i % 3),
            "bank_txn_industry_pay_interval_std": round(1.2 + i * 0.1, 4),
            "bank_txn_industry_primary_employer": employers[i - 1],
            "bank_txn_industry_primary_employer_ratio": round(0.2 + i * 0.03, 4),
            "bank_txn_industry_employer_cnt": (i % 5) + 1,
        }
        rows.append({field: row.get(field) for field in [join_key] + fields})
    return pd.DataFrame(rows)


def main() -> None:
    config_path = Path("config/analysis_config_input_sample.py")
    cfg = load_config(config_path)
    input_dir = Path(cfg["project"]["root_dir"]) / "INPUT"
    input_dir.mkdir(parents=True, exist_ok=True)

    app_ids = [f"APP{i:04d}" for i in range(1, 13)]

    outputs = {
        "main_analysis_model.csv": _build_base_table(cfg, app_ids),
        "business_analysis_variable_library.csv": _build_customer_profile(cfg, app_ids),
        "cross_model_score.csv": _build_risk_score(cfg, app_ids),
        "model_variable_library.csv": _build_feature_detail(cfg, app_ids),
        "occupation_information.csv": _build_txn_industry(cfg, app_ids),
    }

    for file_name, df in outputs.items():
        path = input_dir / file_name
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"wrote {path} rows={len(df)} cols={len(df.columns)}")


if __name__ == "__main__":
    main()
