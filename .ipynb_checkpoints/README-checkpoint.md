# Model Analysis Pipeline

这是一个可配置、可复用的离线分析项目，用来把多张输入表加工成统一底表，并输出：

- 分箱表现
- 转化分析
- 主模型 / 对比模型分箱交叉分析
- 各业务维度分组分析
- 月度分组分析
- 特征画像统计
- 数据质量摘要

项目的目标是保持代码通用，把业务差异尽量放到配置里。除了你在配置中明确指定的表名、字段名、分箱字段名之外，代码和默认输出命名尽量使用通用英文语义。

## 1. 运行方式

安装依赖：

```bash
pip install -r requirements.txt
```

运行样例配置：

```bash
python main.py --config config/analysis_config_input_sample.py
```

默认输出目录：

```text
./output/model_analysis_20260429
```

如果 `runtime.overwrite_output = true`，每次运行前会清空输出目录旧文件。

## 2. 当前流程

主流程如下：

1. 读取配置
2. 加载输入表
3. 以 `analysis.base_table` 指定的底表构建基础样本
4. 按 `joins` 配置关联其他表
5. 按 `score_binning` 生成一个或多个分箱字段
6. 从 `sample_datetime` 派生 `sample_month`
7. 构建转化标签和成交金额字段
8. 输出数据质量摘要
9. 输出普通分组分析
10. 自动输出所有普通分组分析的 `by_sample_month` 版本
11. 输出额外的月度专项分析
12. 输出特征画像统计

## 3. 输入表约定

样例配置当前使用这些输入表 key：

- `base_sample`
- `customer_profile`
- `comparison_score`
- `feature_detail`
- `txn_industry_features`

其中：

- `analysis.base_table` 当前配置为 `base_sample`
- 主模型分数字段来自 `base_sample`
- 对比模型分数字段来自 `comparison_score`

你完全可以在自己的配置里改成别的表 key。代码不要求必须叫 `worthiness_score` 或 `risk_score`。

## 4. 关键配置说明

### 4.1 `analysis.base_table`

指定哪张输入表作为底表。

样例：

```python
"analysis": {
    "base_table": "base_sample",
}
```

### 4.2 `score_binning`

这里定义原始分数字段和分箱输出字段。样例里使用的是：

- `primary_model_score_bin`
- `comparison_model_score_bin`

这两个名字只是样例配置，不是代码写死的要求。

### 4.3 `group_analyses`

这里定义常规分组分析。每一项至少包括：

- `name`
- `group_by`
- `output_file`

例如：

```python
{
    "name": "primary_vs_comparison_model_bin",
    "group_by": ["primary_model_score_bin", "comparison_model_score_bin"],
    "output_file": "05_primary_vs_comparison_model_bin.csv",
}
```

### 4.4 `monthly_group_analyses`

这是自动月度扩展开关。

```python
"monthly_group_analyses": {
    "enabled": True,
    "output_suffix": "_by_sample_month",
}
```

打开后，`group_analyses` 中的每一项都会自动多产出一个按月版本。

例如：

- `05_primary_vs_comparison_model_bin.csv`
- `05_primary_vs_comparison_model_bin_by_sample_month.csv`

### 4.5 `monthly_analyses`

这是额外的月度专项分析配置。它和自动扩展的 `group_analyses by month` 不冲突。

当前默认输出为竖表，不再默认输出月份横向展开的宽表。

例如：

```python
{
    "name": "primary_model_bin_x_sample_month_duedate_1m_5_bad_rate",
    "group_by": ["primary_model_score_bin"],
    "metric_field": "duedate_1m_5_bad_rate",
    "output_file": "21_primary_model_bin_x_sample_month_duedate_1m_5_bad_rate.csv",
    "value_format": "percent_1",
}
```

输出结构类似：

```csv
primary_model_score_bin,sample_month,duedate_1m_5_bad_rate
10,2026-04,0.0%
20,2026-04,1.2%
```

如果你确实需要宽表，可以在单个配置项里加：

```python
"output_layout": "pivot"
```

### 4.6 `feature_profile`

当前样例里按 `primary_model_score_bin` 做画像统计。

主要输出：

- `06_feature_profile_by_primary_group.csv`
- `07_feature_profile_by_category.csv`
- `feature_profile_by_category/*.csv`

## 5. 当前样例输出命名规范

为了保持通用，样例配置中的输出文件名已经统一改成下面这套风格：

- `primary_model_bin_*`
- `comparison_model_*`
- `primary_vs_comparison_model_*`
- `feature_profile_by_primary_group`

例如：

- `02_primary_model_bin_performance.csv`
- `03_conversion_by_primary_model_bin.csv`
- `04_primary_model_bin_x_conversion_stage.csv`
- `05_primary_vs_comparison_model_bin.csv`
- `11_primary_model_bin_x_state.csv`
- `15_primary_model_bin_x_discrete_numeric.csv`

自动月度版本则统一追加：

- `_by_sample_month`

例如：

- `05_primary_vs_comparison_model_bin_by_sample_month.csv`
- `11_primary_model_bin_x_state_by_sample_month.csv`
- `15_primary_model_bin_x_discrete_numeric_by_sample_month.csv`

## 6. 当前主要输出文件

### 基础输出

- `00_data_quality_summary.csv`
- `01_enriched_base_sample.csv`

### 常规分组分析

- `02_primary_model_bin_performance.csv`
- `03_conversion_by_primary_model_bin.csv`
- `04_primary_model_bin_x_conversion_stage.csv`
- `05_primary_vs_comparison_model_bin.csv`
- `08_primary_model_bin_x_last_step.csv`
- `09_primary_model_bin_x_loan_tag.csv`
- `10_primary_model_bin_x_requested_loan_tag.csv`
- `11_primary_model_bin_x_state.csv`
- `12_primary_model_bin_x_suburb.csv`
- `13_primary_model_bin_x_family_type.csv`
- `14_primary_model_bin_x_attributed_category.csv`
- `15_primary_model_bin_x_discrete_numeric.csv`
- `17_primary_model_bin_x_primary_industry.csv`
- `18_primary_model_bin_x_primary_employer.csv`

### 自动月度分组分析

上面这些常规分组分析，都会自动多一份 `_by_sample_month.csv`。

### 额外月度专项分析

- `20_primary_model_bin_x_sample_month_apply_cnt.csv`
- `21_primary_model_bin_x_sample_month_duedate_1m_5_bad_rate.csv`
- `22_primary_model_bin_x_sample_month_duedate_3m_30_bad_rate.csv`
- `23_primary_model_bin_x_sample_month_duedate_1m_5_amount_overdue_rate.csv`
- `24_primary_model_bin_x_sample_month_duedate_3m_30_amount_overdue_rate.csv`

### 特征画像

- `06_feature_profile_by_primary_group.csv`
- `07_feature_profile_by_category.csv`
- `feature_profile_by_category/*.csv`

## 7. 如何迁移到你的真实模型

通常只需要改配置，不需要改代码。

重点改这些位置：

1. `input_tables`
2. `analysis.base_table`
3. `score_binning.scores[*].score_field`
4. `score_binning.scores[*].bin_field`
5. `joins`
6. `group_analyses`
7. `monthly_analyses`
8. `feature_profile.group_by`

如果你明天分析的不是 worthiness 模型，而是别的主模型：

- 保留你真实的原始字段名
- 把配置里的派生分箱字段名改成更符合场景的名字
- 或者继续复用 `primary_model_score_bin` / `comparison_model_score_bin` 这种抽象命名

推荐做法是：

- 原始输入字段名保持业务真实含义
- 配置里派生的分析字段名尽量抽象化
- 输出文件名尽量不要带具体模型代号

## 8. 常见注意点

- `feature_detail` 不会并入 `01_enriched_base_sample.csv`，这是刻意设计，避免底表过大
- 如果右表主键重复，默认保留第一条并记录 warning
- 如果 `sample_datetime` 无法解析，会导致对应行没有 `sample_month`
- 月度专项分析现在默认输出竖表
- Excel 导出依赖 `openpyxl` 或 `xlsxwriter`

## 9. 代码中已经做的通用化改动

目前这些地方已经去掉了写死的 `worthiness` 语义默认值：

- 底表默认改为 `analysis.base_table`
- `data_loader` 会对 `analysis.base_table` 做全字段加载
- `feature_profile` 默认分组字段改为 `primary_model_score_bin`
- `feature profile` 输出开关改为 `feature_profile_by_group`
- 样例配置中的输出文件名和分析名改成了通用英文命名

## 10. 快速检查建议

每次改完配置后，优先看：

1. `run_log.txt`
2. `00_data_quality_summary.csv`
3. `01_enriched_base_sample.csv`
4. 目标分析输出文件

尤其检查：

- 主键是否匹配成功
- 分箱字段是否生成
- `sample_month` 是否生成
- 月度文件是否按预期产出
