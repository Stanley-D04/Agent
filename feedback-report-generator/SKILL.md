---
name: feedback-report-generator
description: "This skill generates a complete interactive HTML feedback analysis report from an Excel file. It should be used when the user provides an Excel file containing user feedback data (fields: category, evaluation, survey stage, description, tags) and wants a visual report with category overview, issue classification cards, tag statistics, satisfaction chart, insights, recommendations, and a drill-down detail modal. Trigger phrases: 分析这个反馈excel, 给我生成分析报告, 用户反馈报告, feedback report, 从excel生成报告."
---

# 用户反馈分析报告生成器

## 用途

To generate a self-contained interactive HTML analysis report from a user feedback Excel file.

The output report includes:
- Per-category overview: bad/mid/good counts, bad rate, survey stage distribution bar chart
- Issue classification cards (A/B/C/D/E/MISC) with clickable sub-tags
- Tag statistics bar chart (Top 8)
- Satisfaction donut chart
- Core insights and optimization recommendations
- Full drill-down modal with tag/evaluation/stage filtering and real-time count update

## 使用流程

### Step 1：接收 Excel 文件

To start, confirm the Excel file path from the user. Ask if the column names differ from defaults.

To understand the expected column structure, read `references/field_mapping.md`.

### Step 2：运行生成脚本

To generate the report, run:

```bash
python3 {SKILL_DIR}/scripts/gen_report.py [excel_path] [output_path]
```

- `excel_path`: path to the user's Excel file (absolute path preferred)
- `output_path`: optional; defaults to `[excel_stem]_分析报告.html` in the same directory

The script will auto-install `pandas` and `openpyxl` if not present.

### Step 3：处理列名不匹配

If output contains empty columns or missing data, it may be a column name mismatch.

To fix, either:
- Ask the user to confirm column names, then modify `COLUMN_MAP` at the top of `gen_report.py`
- Or tell the user to rename Excel columns to match the defaults in `references/field_mapping.md`

### Step 4：自定义标签映射（可选）

If the user's tags differ from the default `TAG_CLASS` mapping, to extend it:
- Open `scripts/gen_report.py`
- Find the `TAG_CLASS` dict
- Add new entries: `"新标签名": "A"` (A/B/C/D/E/MISC)

### Step 5：启动预览

To preview the generated HTML, start a local HTTP server:

```bash
cd [output_dir] && python3 -m http.server [port] &
```

Then use `preview_url` to show the result.

## 报告功能清单

- 品类切换下拉（顶部固定导航）
- 数据总览：4格统计卡 + 环节分布横条图
- 问题分类汇总：点击分类标题 → 弹出完整明细；点击子标签 → 弹出该标签明细并自动筛选
- 标签 Top8 横条图 + 满意度甜甜圈图
- 核心洞察（P0/P1/P2 优先级）+ 优化建议
- 明细弹框：标签/评价/环节三下拉联合筛选；筛选后顶部条数实时更新

## 注意事项

- Excel 必须包含 `品类`、`评价`、`描述`、`标签` 四个核心字段（列名见 `references/field_mapping.md`）
- 一行数据有多个标签时，会同时归入多个分类的明细（这是正确行为，各分类独立计数）
- 图片字段须为 HTTP/HTTPS URL，否则图片列显示"无图"
