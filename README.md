# Model Analysis Project

This project builds model-bin Excel reports for comparing the primary model score bins against the comparison model score bins.

## Run

```bash
python run.py
```

Each run clears `output/model_analysis_20260429` first, then writes the latest workbooks. The legacy root-level `OUTPUT` folder is not used for final report delivery.

## Outputs

All final workbooks are written to:

```text
output/model_analysis_20260429/
```

Generated files, in write order:

1. `model_bin_feature_mean_profile.xlsx`
   - Feature mean profile by model bins.
   - Uses variables from `INPUT/model_variable_library.csv`.
   - Converts each variable with `pd.to_numeric(errors="coerce")`.
   - Reports mean values by primary bin, comparison bin, and primary bin x comparison bin.

2. `model_bin_user_profile_pivot.xlsx`
   - User profile pivot by model bins.
   - Numeric profile metrics use primary bin x comparison bin pivots.
   - Categorical profile fields use primary-model-bin distribution tables.

3. `model_bin_risk_performance_pivot.xlsx`
   - Risk performance and business KPI pivot by model bins.
   - Includes risk performance, amount risk, existing profile means, and conversion metrics.

## Inputs

Main input paths are configured in `config/analysis_config_input_sample.py`.

Current input tables:

- `INPUT/main_analysis_model.csv`
- `INPUT/business_analysis_variable_library.csv`
- `INPUT/cross_model_score.csv`
- `INPUT/occupation_information.csv`
- `INPUT/model_variable_library.csv`

## Output Flow

`run.py` performs the shared data preparation once:

1. Load configured input tables.
2. Join configured profile, comparison-score, and occupation fields.
3. Apply primary and comparison model score binning.
4. Build month, conversion, and deal-amount helper fields.
5. Write the three final workbooks in the output order listed above.

## Configuration

Primary and comparison model bins are configured in:

```text
config/analysis_config_input_sample.py
```

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

User profile metrics are configured under:

```python
"user_profile_metrics"
```

Feature mean profile variables are read from:

```text
INPUT/model_variable_library.csv
```

If this file is in wide format, all non-key columns except `application_id`, `user_id`, and `sample_datetime` are treated as feature variables.
