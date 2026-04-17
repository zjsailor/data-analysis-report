# CSV/Excel Data Summarizer

Comprehensive data analysis skill for CSV and Excel files with specialized support for sales/order data.

## Features

- **Multi-format Support**: CSV and Excel (.xlsx, .xls) files
- **Chinese Language Support**: Full support for Chinese characters in charts and analysis
- **Sales Analytics**: Specialized TOP 10 rankings for:
  - Products
  - Customers
  - Salespersons
  - Hospitals/Institutions
- **Trend Analysis**: Monthly trends, hourly patterns
- **Distribution Analysis**: Price ranges, categories
- **Data Quality**: Missing value analysis

## Usage

### Basic Analysis
```
Analyze this sales_data.csv file
```

### Excel Files
```
Analyze this report.xlsx file
```

### Full Report with Charts
```
Summarize this data and create visualizations
```

## Output

The skill generates:
1. **Charts** saved to `charts/` subdirectory:
   - `chart01_top10_products.png`
   - `chart02_top10_customers.png`
   - `chart03_top10_salespersons.png`
   - `chart04_top10_hospitals.png`
   - `chart05_monthly_trend.png`
   - `chart06_payment_method.png`
   - `chart07_price_distribution.png`
   - `chart08_sales_category.png`
   - `chart09_hourly_pattern.png`

2. **Report Data**: `report_data.json` with structured analysis results

## Requirements

```
pip install -r requirements.txt
```

## Running Locally

```bash
uv run --with pandas,matplotlib,seaborn,openpyxl python analyze.py your_file.csv
```