---
name: data-analysis-report
description: 智能数据分析工具，支持 CSV 和 XLSX 格式，自动生成统计图表和完整的 Word 分析报告。支持销售/医疗检测数据TOP10排名分析、中文显示优化。
metadata:
  version: 4.0.0
  dependencies: python>=3.8, pandas>=2.0.0, openpyxl>=3.0.0, matplotlib>=3.7.0, seaborn>=0.12.0, python-docx>=0.8.11
---

# 数据分析报告生成器

智能数据分析 Skill，支持 **CSV** 和 **XLSX** 格式文件，自动生成：
- 📊 多维度数据分析
- 📈 可视化图表（PNG）
- 📄 Markdown 格式报告
- 📝 Word (.docx) 格式报告

## 何时使用此 Skill

当用户：
- 提供 CSV 或 Excel 文件并要求分析
- 需要数据统计、趋势分析、分布分析
- 要求生成可视化图表
- 需要完整的分析报告（Markdown 或 Word 格式）
- 询问数据质量、缺失值、异常值等问题

## ⚠️ 核心行为要求 ⚠️

**收到数据文件后立即执行**：

1. ✅ **自动识别格式** - 支持 CSV、XLSX、XLS
2. ✅ **全面分析** - 数据质量、数值分析、分类分析、时间序列
3. ✅ **生成图表** - TOP10排名、相关性热力图、分布图、分类分布图等
4. ✅ **输出报告** - Markdown + Word 双格式
5. ❌ **不询问** - 不问"想要什么分析"
6. ❌ **不等待** - 不等待用户确认
7. ❌ **不选择** - 不提供选项让用户选

## 功能特性

### 1. 多格式支持

| 格式 | 扩展名 | 支持状态 |
|------|--------|---------|
| CSV | .csv | ✅ 完全支持 |
| Excel 2007+ | .xlsx | ✅ 完全支持 |
| Excel 97-2003 | .xls | ✅ 支持 |

### 2. 销售/医疗检测数据分析（专长）

自动识别并生成 TOP 10 排名分析：

| 分析维度 | 生成图表 | 说明 |
|---------|---------|------|
| 产品排名 | chart01_top10_products.png | 按销售额统计TOP10产品 |
| 客户排名 | chart02_top10_customers.png | 按销售额统计TOP10客户 |
| 销售人员排名 | chart03_top10_salespersons.png | 按销售额统计TOP10销售人员 |
| 医院/机构排名 | chart04_top10_hospitals.png | 按销售额统计TOP10医院 |
| 月度趋势 | chart05_monthly_trend.png | 月度销售额和订单量趋势 |
| 付款方式 | chart06_payment_method.png | 对公/对私等付款分布 |
| 价格区间 | chart07_price_distribution.png | 不同价格区间订单分布 |
| 业务类别 | chart08_sales_category.png | 按业务类别统计 |
| 时段分布 | chart09_hourly_pattern.png | 按小时统计销售额 |
| 患者特征 | chart10_gender_distribution.png | 患者性别分布 |
| 年龄分布 | chart11_age_distribution.png | 患者年龄段分布 |

### 3. 自动数据分析

分析维度包括：

#### 📋 数据概览
- 记录数、字段数
- 数据类型识别
- 字段列表

#### 🔍 数据质量
- 缺失值统计
- 缺失率计算
- 高缺失字段识别

#### 📈 数值分析
- 描述性统计（均值、标准差、分位数）
- 相关性分析
- 相关性热力图

#### 📊 分类分析
- 唯一值统计
- TOP 值分布
- 分类分布图

#### 📅 时间序列（自动检测日期字段）
- 时间范围
- 月度趋势
- 趋势图

#### 📉 分布分析
- 数值分布直方图
- 分布特征

### 4. 智能图表生成

根据数据特征自动生成合适的图表：

| 数据类型 | 生成图表 |
|---------|---------|
| 销售/订单数据 | TOP10产品/客户/销售人员/医院排名图 |
| 多个数值字段 | 相关性热力图 |
| 分类字段 | 分类分布图（TOP 10） |
| 日期字段 | 时间序列趋势图、月度趋势图 |
| 数值字段 | 数值分布直方图、价格区间分布图 |
| 患者数据 | 性别分布图、年龄分布图 |

### 5. 中文显示优化

- 默认使用 Microsoft YaHei、SimHei 等中文字体
- 自动处理中文编码问题
- 支持 GBK、GB18030、UTF-8 等多种编码

### 6. 报告输出

#### Markdown 报告
- 完整的分析内容
- Markdown 表格
- 图表引用

#### Word 报告
- 专业排版
- 中文字体支持（宋体、黑体）
- 表格自动格式化
- 首行缩进、行距等排版规范

## 使用方法

### 方法 1：一键生成（推荐）

```bash
python generate_report.py <数据文件路径> [输出目录]
```

**示例**：
```bash
# 分析 CSV 文件
python generate_report.py sales_data.csv

# 分析 Excel 文件，指定输出目录
python generate_report.py orders.xlsx ./output

# Windows 路径
python generate_report.py "C:\Users\data\sales.xlsx"
```

### 方法 2：分步执行

#### 步骤 1：数据分析（生成 Markdown 和图表）
```bash
python analyze.py <数据文件路径> [输出目录]
```

#### 步骤 2：转换为 Word
```bash
python markdown_to_docx.py <Markdown文件路径> [Word文件路径]
```

### 方法 3：Python 调用

```python
from generate_report import generate_report

result = generate_report('data.xlsx', './output')

print(f"Markdown: {result['markdown_report']}")
print(f"Word: {result['docx_report']}")
print(f"Charts: {result['charts']}")
```

### 方法 4：直接使用 analyze.py

```python
from analyze import summarize_csv

results = summarize_csv('sales_data.csv', output_dir='./charts')

print(f"Charts: {results['charts']}")
print(f"Report Data: {results['report_data_path']}")
```

## 输出文件

执行后生成以下文件（保存在输出目录）：

| 文件 | 说明 |
|------|------|
| `{文件名}_分析报告.md` | Markdown 格式分析报告 |
| `{文件名}_分析报告.docx` | Word 格式分析报告 |
| `charts/` | 图表目录 |
| `chart01_top10_products.png` | TOP10产品排名图 |
| `chart02_top10_customers.png` | TOP10客户排名图 |
| `chart03_top10_salespersons.png` | TOP10销售人员排名图 |
| `chart04_top10_hospitals.png` | TOP10医院排名图 |
| `chart05_monthly_trend.png` | 月度趋势图 |
| `chart06_payment_method.png` | 付款方式分布图 |
| `chart07_price_distribution.png` | 价格区间分布图 |
| `chart08_sales_category.png` | 业务类别分布图 |
| `chart09_hourly_pattern.png` | 时段分布图 |
| `chart10_gender_distribution.png` | 性别分布图 |
| `chart11_age_distribution.png` | 年龄分布图 |
| `report_data.json` | 结构化分析数据（JSON格式） |

## 报告结构

生成的报告包含以下章节：

1. **数据概览** - 记录数、字段数、数据来源
2. **核心指标** - 总金额、订单数、客单价等关键指标
3. **TOP 10 排名分析** - 产品、客户、销售人员、医院排名
4. **月度趋势分析** - 时间变化趋势
5. **付款方式分析** - 支付渠道分布
6. **价格区间分析** - 不同价位销售分布
7. **患者特征分析** - 性别、年龄等分布（医疗数据）
8. **数据质量分析** - 缺失值、完整性
9. **关键洞察与建议** - 业务洞察、存在问题、发展建议
10. **行动计划** - 短期/中期/长期行动项
11. **风险预警** - 业务风险、运营风险
12. **总结** - 分析总结

## 配置说明

### 字体配置

Word 报告默认使用：
- **中文**：宋体（正文）、黑体（标题）
- **英文**：Times New Roman
- **字号**：正文 10.5pt（五号）、标题分级递减

图表默认使用：
- **中文**：Microsoft YaHei、SimHei、KaiTi、DengXian、FangSong
- **英文**：DejaVu Sans

### 页面设置

- **纸张**：A4 (21cm × 29.7cm)
- **页边距**：上/下 2.54cm，左/右 3.17cm
- **行距**：1.5倍
- **首行缩进**：2字符

### 图表设置

- **分辨率**：150 DPI
- **格式**：PNG
- **配色**：Seaborn 默认配色

## 示例

### 示例 1：销售数据分析

```bash
python generate_report.py sales_2025.xlsx ./reports
```

输出：
```
📊 数据分析和报告生成工具
================================================================================

📁 输入文件：sales_2025.xlsx
📂 输出目录：./reports

【步骤 1/3】正在分析数据...
✅ 数据分析完成
   - 记录数：10,582
   - 字段数：15
   - 总金额：¥5,000,000
   - 生成图表：11 张

【步骤 2/3】正在生成 Word 报告...
✅ Word 报告生成完成

【步骤 3/3】生成文件清单
📄 Markdown 报告：./reports/sales_2025_分析报告.md
📄 Word 报告：./reports/sales_2025_分析报告.docx
📊 可视化图表（11 张）：
   - chart01_top10_products.png
   - chart02_top10_customers.png
   - chart03_top10_salespersons.png
   - chart04_top10_hospitals.png
   - chart05_monthly_trend.png
   - chart06_payment_method.png
   - chart07_price_distribution.png
   - chart08_sales_category.png
   - chart09_hourly_pattern.png
   - chart10_gender_distribution.png
   - chart11_age_distribution.png
```

### 示例 2：医疗检测数据分析

```bash
python analyze.py tumor_2025.xlsx ./output
```

自动识别为医疗检测数据，生成包含患者特征（性别、年龄分布）的分析报告。

### 示例 3：Python 调用

```python
from generate_report import generate_report

result = generate_report('customers.csv')

if result:
    print("分析完成！")
    print(f"报告位置：{result['docx_report']}")
```

## 文件结构

```
data-analysis-report-skill/
├── generate_report.py        # 主脚本（一键生成）
├── analyze.py                # 数据分析核心（支持TOP10排名分析）
├── markdown_to_docx.py       # Markdown 转 Word
├── requirements.txt          # Python 依赖
├── SKILL.md                  # 本文档
└── resources/
    ├── sample.csv           # 示例数据
    └── README.md             # 附加说明
```

## 依赖安装

```bash
pip install -r requirements.txt
```

或单独安装：
```bash
pip install pandas openpyxl matplotlib seaborn python-docx
```

## 注意事项

1. **文件格式**：仅支持 .csv, .xlsx, .xls
2. **编码**：CSV 文件自动识别 UTF-8、GBK、GB18030 等编码
3. **中文支持**：需要系统安装中文字体（Windows 自带 Microsoft YaHei）
4. **大文件**：超过 100 万行建议分批处理
5. **内存**：大文件分析需要足够内存
6. **图表**：根据数据特征智能选择生成的图表类型
7. **编码列名**：Excel 文件列名乱码时，analyze.py 会尝试自动修复

## 常见问题

### Q1: 如何支持其他文件格式？

修改 `analyze.py` 中的 `read_data()` 函数，添加对应的读取逻辑。

### Q2: 如何自定义报告样式？

修改 `markdown_to_docx.py` 中的样式设置函数。

### Q3: 中文字体显示异常？

确保系统安装了中文字体（宋体、黑体），或修改代码指定字体路径。analyze.py 默认配置了多种中文字体备选。

### Q4: 图表中文乱码？

修改 `analyze.py` 开头的字体配置：
```python
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
```

### Q5: Excel 文件列名乱码？

analyze.py 内置了列名修复功能，会自动处理常见的编码问题。

## 更新日志

### v4.0.0 (2026-04-17)
- ✅ 新增中文显示优化（Microsoft YaHei字体支持）
- ✅ 新增销售/医疗数据 TOP 10 排名分析
- ✅ 新增月度趋势分析图表
- ✅ 新增付款方式分布分析
- ✅ 新增价格区间分布分析
- ✅ 新增时段分布（小时级）分析
- ✅ 新增患者特征分析（性别、年龄分布）
- ✅ 新增 Excel 文件支持（.xlsx）
- ✅ 优化中文编码处理
- ✅ 新增 report_data.json 输出

### v3.0.0 (2026-04-07)
- ✅ 新增 XLSX 格式支持
- ✅ 整合 Word 报告生成功能
- ✅ 一键生成完整报告
- ✅ 优化中文排版
- ✅ 增强数据质量分析

### v2.1.0
- ✅ 智能图表选择
- ✅ 时间序列分析
- ✅ 缺失值分析

### v2.0.0
- ✅ 多图表生成
- ✅ 相关性分析
- ✅ 分类分布分析

---

**维护者**：AI Assistant
**许可证**：MIT