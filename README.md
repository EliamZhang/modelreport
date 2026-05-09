# Model Analysis Project

This project generates two workbooks:

```text
output/model_analysis_20260429/cross_model_bin_pivot_rate_only.xlsx
output/model_analysis_20260429/cross_model_bin_pivot_user_profile.xlsx
```

Run:

```bash
python run.py
```

The run always clears `output/model_analysis_20260429` first, writes the business analysis workbook, then writes the user profile workbook.

## Manual Bin Groups

Primary and comparison model bin groups are configured independently in
`config/analysis_config_input_sample.py`.

Current grouping for both models:

```python
"bin_groups": [
    {"label": 1, "source_bin_indexes": [1, 2, 3, 4]},
    {"label": 2, "source_bin_indexes": [5, 6, 7, 8]},
    {"label": 3, "source_bin_indexes": [9, 10]},
]
```

Rules:

- `source_bin_indexes` are 1-based source-bin positions.
- Every source bin must be covered exactly once.
- Each group must contain adjacent source bins.
