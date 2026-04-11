# 数据分析报告生成器 - 使用说明

## 📌 快速开始

### 一键生成完整报告

```bash
python generate_report.py <数据文件路径> [输出目录]
```

**支持的格式**：CSV、XLSX、XLS

### 示例

```bash
# 分析 CSV 文件
python generate_report.py data.csv

# 分析 Excel 文件
python generate_report.py data.xlsx ./output

# Windows 路径
python generate_report.py "C:\Users\yikon\data.xlsx"
```

## 📂 生成文件

执行后自动生成：

| 文件 | 说明 |
|------|------|
| `{文件名}_分析报告.md` | Markdown 报告 |
| `{文件名}_分析报告.docx` | Word 报告 |
| `correlation_heatmap.png` | 相关性热力图 |
| `categorical_distributions.png` | 分类分布图 |
| `distributions.png` | 数值分布图 |
| `time_series_trend.png` | 时间趋势图 |

## ️ 安装依赖

```bash
pip install -r requirements.txt
```

或：

```bash
pip install pandas openpyxl matplotlib seaborn python-docx
```

## ✨ 功能特性

- ✅ 自动识别文件格式（CSV/XLSX）
- ✅ 全面数据分析（质量、数值、分类、时间序列）
- ✅ 智能图表生成
- ✅ 专业排版 Word 报告
- ✅ 中文字体支持

## 📖 详细文档

查看 `SKILL.md` 了解完整功能和使用方法。
