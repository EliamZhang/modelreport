# Model Analysis Project

这个项目当前只生成一个输出文件：

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

当 `runtime.overwrite_output = True` 时，运行前会清空配置中的 `output_dir`，运行后目录中只保留
`cross_model_bin_pivot_rate_only.xlsx`。

## Output

`cross_model_bin_pivot_rate_only.xlsx` 包含主模型分箱与对比模型分箱的交叉 pivot。默认行字段是
`primary_model_score_bin`，默认列字段是 `comparison_model_score_bin`。

以下指标已从配置、计算和 workbook 输出中移除：

- `avg_bank_txn_industry_primary_industry_amt`
- `avg_bank_txn_industry_pay_interval_median`
- `avg_bank_txn_industry_pay_interval_std`
- `avg_bank_txn_industry_primary_employer_ratio`
- `avg_bank_txn_industry_employer_cnt`

旧版 CSV 输出、拆分输出、总 workbook、运行日志文件以及对应的主入口生成逻辑已删除。
