# Model Analysis Project

当前项目只生成一个输出文件：

```text
output/model_analysis_20260429/cross_model_bin_pivot_rate_only.xlsx
```

运行方式：

```bash
python main.py --config config/analysis_config_input_sample.py
```

也可以直接运行生成脚本：

```bash
python scripts/generate_cross_model_pivot_excel.py --config config/analysis_config_input_sample.py
```

## 人工分箱配置

原始分箱配置在 `config/analysis_config_input_sample.py` 的 `score_binning.scores[*].bins` 中。当前样例是 10 个原始箱。

现在只支持人工指定合箱规则：每个模型都在自己的 `score_binning.scores[*].bin_groups` 中配置输出分组。主模型和对比模型可以配置相同规则，也可以配置不同规则。

当前配置：

```python
"bin_groups": [
    {"label": 1, "source_bin_indexes": [1, 2, 3, 4]},
    {"label": 2, "source_bin_indexes": [5, 6, 7, 8]},
    {"label": 3, "source_bin_indexes": [9, 10]},
]
```

含义：

- 输出第 1 箱 = 原始第 1、2、3、4 箱
- 输出第 2 箱 = 原始第 5、6、7、8 箱
- 输出第 3 箱 = 原始第 9、10 箱

规则：

- `source_bin_indexes` 是原始箱的 1-based 序号。
- 所有原始箱必须覆盖一次且只能一次。
- 每组只能合并相邻箱，不能跳箱。
- 也支持按原始箱标签配置：`{"label": 1, "source_bin_labels": [10, 20]}`。

## 输出

`cross_model_bin_pivot_rate_only.xlsx` 包含主模型分箱与对比模型分箱的交叉 pivot。默认行字段是 `primary_model_score_bin`，默认列字段是 `comparison_model_score_bin`。

以下指标已从配置、计算和 workbook 输出中移除：

- `avg_bank_txn_industry_primary_industry_amt`
- `avg_bank_txn_industry_pay_interval_median`
- `avg_bank_txn_industry_pay_interval_std`
- `avg_bank_txn_industry_primary_employer_ratio`
- `avg_bank_txn_industry_employer_cnt`
