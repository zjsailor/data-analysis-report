import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib

matplotlib.use("Agg")
from pathlib import Path
import os
import json
import warnings

warnings.filterwarnings("ignore")

plt.rcParams["font.sans-serif"] = [
    "Microsoft YaHei",
    "SimHei",
    "KaiTi",
    "DengXian",
    "FangSong",
    "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False


def summarize_csv(file_path, output_dir=None):
    """
    Comprehensively analyzes a CSV/Excel file and generates multiple visualizations.
    Specializes in sales/order data with TOP 10 rankings and trend analysis.

    Args:
        file_path (str): Path to the CSV or Excel file
        output_dir (str): Directory to save charts and report data

    Returns:
        dict: Contains analysis results, charts list, and data summaries
    """
    # Determine output directory
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    os.makedirs(output_dir, exist_ok=True)

    # Read file (CSV or Excel)
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, engine="openpyxl")
    else:
        # Try utf-8 first, then other encodings
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_path, encoding="gbk")
            except:
                df = pd.read_csv(file_path, encoding="latin1")

    results = {
        "total_rows": len(df),
        "total_columns": len(df.columns),
        "columns": list(df.columns),
        "charts": [],
        "data": {},
    }

    # Clean column names if they're corrupted
    if any("?" in c or "\ufffd" in c for c in df.columns):
        df.columns = [
            c.encode("utf-8", errors="ignore").decode("utf-8", errors="ignore")
            for c in df.columns
        ]

    # Identify column types
    date_cols = [
        c
        for c in df.columns
        if any(x in c.lower() for x in ["date", "time", "时间", "日期", "创建"])
    ]
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    amount_cols = [
        c
        for c in df.columns
        if any(
            x in c.lower()
            for x in ["amount", "total", "price", "金额", "总价", "单价", "总计", "sum"]
        )
    ]
    categorical_cols = df.select_dtypes(include=["object"]).columns.tolist()

    # Clean numeric columns (remove commas and spaces)
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.replace(",", "").str.strip()
            if (
                col in amount_cols
                or "price" in col.lower()
                or "amount" in col.lower()
                or "total" in col.lower()
            ):
                df[col] = pd.to_numeric(df[col], errors="coerce")

    # Convert date columns
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # Basic info
    results["data"]["overview"] = {
        "total_records": len(df),
        "total_columns": len(df.columns),
    }

    # Total amount calculation
    amount_col = amount_cols[0] if amount_cols else None
    if amount_col and amount_col in df.columns:
        total_amount = df[amount_col].sum()
        results["data"]["total_amount"] = float(total_amount)
        results["data"]["avg_order_value"] = (
            float(total_amount / len(df)) if len(df) > 0 else 0
        )

    # Missing data analysis
    missing = df.isnull().sum()
    missing_pct = (missing / len(df) * 100).round(2)
    results["data"]["missing"] = {
        col: {"count": int(missing[col]), "pct": float(missing_pct[col])}
        for col in df.columns
        if missing[col] > 0
    }

    # Numeric analysis
    if numeric_cols:
        results["data"]["numeric_summary"] = df[numeric_cols].describe().to_dict()

    # TOP 10 Analysis for key dimensions
    customer_col = None
    product_col = None
    salesperson_col = None
    hospital_col = None
    category_col = None

    # Identify columns by name patterns
    for col in categorical_cols:
        col_lower = col.lower()
        if not customer_col and any(
            x in col_lower for x in ["customer", "客户", "客人"]
        ):
            customer_col = col
        if not product_col and any(x in col_lower for x in ["product", "产品", "商品"]):
            product_col = col
        if not salesperson_col and any(
            x in col_lower for x in ["salesperson", "sales", "销售", "拥有", "owner"]
        ):
            salesperson_col = col
        if not hospital_col and any(
            x in col_lower for x in ["hospital", "医院", "机构", "送检"]
        ):
            hospital_col = col
        if not category_col and any(
            x in col_lower for x in ["category", "类别", "科室", "分类"]
        ):
            category_col = col

    # Create TOP 10 charts
    charts_dir = os.path.join(output_dir, "charts")
    os.makedirs(charts_dir, exist_ok=True)

    amount_col_for_analysis = (
        amount_cols[0] if amount_cols else numeric_cols[0] if numeric_cols else None
    )

    if amount_col_for_analysis and amount_col_for_analysis in df.columns:
        # TOP 10 Products
        if product_col and product_col in df.columns:
            top_products = (
                df.groupby(product_col)[amount_col_for_analysis]
                .sum()
                .nlargest(10)
                .sort_values(ascending=True)
            )
            if len(top_products) > 0:
                fig, ax = plt.subplots(figsize=(12, 7))
                ax.barh(
                    range(len(top_products)), top_products.values, color="steelblue"
                )
                ax.set_yticks(range(len(top_products)))
                ax.set_yticklabels(
                    [str(x)[:40] for x in top_products.index], fontsize=10
                )
                ax.set_xlabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "TOP 10 Products by Revenue", fontsize=14, fontweight="bold", pad=15
                )
                for i, v in enumerate(top_products.values):
                    ax.text(
                        v + max(top_products) * 0.01,
                        i,
                        f"{v:,.0f}",
                        va="center",
                        fontsize=9,
                    )
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart01_top10_products.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["top_products"] = {
                    str(k): float(v) for k, v in top_products.items()
                }

        # TOP 10 Customers
        if customer_col and customer_col in df.columns:
            top_customers = (
                df.groupby(customer_col)[amount_col_for_analysis]
                .sum()
                .nlargest(10)
                .sort_values(ascending=True)
            )
            if len(top_customers) > 0:
                fig, ax = plt.subplots(figsize=(12, 7))
                ax.barh(range(len(top_customers)), top_customers.values, color="coral")
                ax.set_yticks(range(len(top_customers)))
                ax.set_yticklabels(
                    [str(x)[:25] for x in top_customers.index], fontsize=10
                )
                ax.set_xlabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "TOP 10 Customers by Revenue",
                    fontsize=14,
                    fontweight="bold",
                    pad=15,
                )
                for i, v in enumerate(top_customers.values):
                    ax.text(
                        v + max(top_customers) * 0.01,
                        i,
                        f"{v:,.0f}",
                        va="center",
                        fontsize=9,
                    )
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart02_top10_customers.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["top_customers"] = {
                    str(k): float(v) for k, v in top_customers.items()
                }

        # TOP 10 Salespersons
        if salesperson_col and salesperson_col in df.columns:
            top_salespersons = (
                df.groupby(salesperson_col)[amount_col_for_analysis]
                .sum()
                .nlargest(10)
                .sort_values(ascending=True)
            )
            if len(top_salespersons) > 0:
                fig, ax = plt.subplots(figsize=(12, 7))
                ax.barh(
                    range(len(top_salespersons)),
                    top_salespersons.values,
                    color="seagreen",
                )
                ax.set_yticks(range(len(top_salespersons)))
                ax.set_yticklabels(
                    [str(x)[:20] for x in top_salespersons.index], fontsize=10
                )
                ax.set_xlabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "TOP 10 Salespersons by Revenue",
                    fontsize=14,
                    fontweight="bold",
                    pad=15,
                )
                for i, v in enumerate(top_salespersons.values):
                    ax.text(
                        v + max(top_salespersons) * 0.01,
                        i,
                        f"{v:,.0f}",
                        va="center",
                        fontsize=9,
                    )
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart03_top10_salespersons.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["top_salespersons"] = {
                    str(k): float(v) for k, v in top_salespersons.items()
                }

        # TOP 10 Hospitals
        if hospital_col and hospital_col in df.columns:
            top_hospitals = (
                df.groupby(hospital_col)[amount_col_for_analysis]
                .sum()
                .nlargest(10)
                .sort_values(ascending=True)
            )
            if len(top_hospitals) > 0:
                fig, ax = plt.subplots(figsize=(12, 7))
                ax.barh(range(len(top_hospitals)), top_hospitals.values, color="purple")
                ax.set_yticks(range(len(top_hospitals)))
                ax.set_yticklabels(
                    [str(x)[:25] for x in top_hospitals.index], fontsize=10
                )
                ax.set_xlabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "TOP 10 Hospitals/Institutions by Revenue",
                    fontsize=14,
                    fontweight="bold",
                    pad=15,
                )
                for i, v in enumerate(top_hospitals.values):
                    ax.text(
                        v + max(top_hospitals) * 0.01,
                        i,
                        f"{v:,.0f}",
                        va="center",
                        fontsize=9,
                    )
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart04_top10_hospitals.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["top_hospitals"] = {
                    str(k): float(v) for k, v in top_hospitals.items()
                }

        # Monthly Trend
        if date_cols:
            date_col = date_cols[0]
            df["month"] = df[date_col].dt.to_period("M")
            monthly_data = (
                df.groupby("month")
                .agg({amount_col_for_analysis: "sum", date_col: "count"})
                .reset_index()
            )
            monthly_data.columns = ["month", "total_amount", "order_count"]
            monthly_data["month"] = monthly_data["month"].astype(str)

            if len(monthly_data) > 1:
                fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
                ax1.plot(
                    monthly_data["month"],
                    monthly_data["total_amount"],
                    marker="o",
                    linewidth=2,
                    color="steelblue",
                    markersize=8,
                )
                ax1.set_xlabel("Month", fontsize=11)
                ax1.set_ylabel("Total Amount (Yuan)", fontsize=11)
                ax1.set_title("Monthly Revenue Trend", fontsize=12, fontweight="bold")
                ax1.grid(True, alpha=0.3)
                ax1.tick_params(axis="x", rotation=45)
                for i, v in enumerate(monthly_data["total_amount"]):
                    ax1.text(
                        i,
                        v + max(monthly_data["total_amount"]) * 0.02,
                        f"{v:,.0f}",
                        ha="center",
                        fontsize=9,
                    )

                ax2.bar(
                    monthly_data["month"],
                    monthly_data["order_count"],
                    color="coral",
                    alpha=0.8,
                )
                ax2.set_xlabel("Month", fontsize=11)
                ax2.set_ylabel("Order Count", fontsize=11)
                ax2.set_title("Monthly Order Volume", fontsize=12, fontweight="bold")
                ax2.grid(True, alpha=0.3, axis="y")
                ax2.tick_params(axis="x", rotation=45)

                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart05_monthly_trend.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["monthly_trend"] = monthly_data.to_dict("records")

        # Payment Method Distribution
        payment_col = None
        for col in categorical_cols:
            if (
                "payment" in col.lower()
                or "回款" in col
                or "付款" in col
                or "支付" in col
            ):
                payment_col = col
                break

        if payment_col and payment_col in df.columns:
            payment_data = (
                df.groupby(payment_col)[amount_col_for_analysis]
                .sum()
                .sort_values(ascending=False)
            )
            if len(payment_data) > 0 and len(payment_data) <= 10:
                fig, ax = plt.subplots(figsize=(8, 6))
                colors = ["steelblue", "coral", "seagreen", "gold", "purple"][
                    : len(payment_data)
                ]
                wedges, texts, autotexts = ax.pie(
                    payment_data.values,
                    labels=payment_data.index,
                    autopct="%1.1f%%",
                    startangle=90,
                    colors=colors,
                    explode=[0.05] * len(payment_data),
                )
                ax.set_title(
                    "Revenue by Payment Method", fontsize=14, fontweight="bold", pad=15
                )
                for text in texts:
                    text.set_fontsize(11)
                for autotext in autotexts:
                    autotext.set_fontsize(10)
                    autotext.set_color("white")
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart06_payment_method.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["payment_method"] = {
                    str(k): float(v) for k, v in payment_data.items()
                }

        # Price Range Distribution
        price_col = None
        for col in df.columns:
            if "price" in col.lower() or "单价" in col or "unit" in col.lower():
                price_col = col
                break

        if price_col and price_col in df.columns:
            price_data = df[price_col].dropna()
            if len(price_data) > 0 and price_data.max() > price_data.min():
                price_bins = [0, 500, 1000, 2000, 3000, float("inf")]
                price_labels = [
                    "<¥500",
                    "¥500-1000",
                    "¥1000-2000",
                    "¥2000-3000",
                    ">¥3000",
                ]
                df["price_range"] = pd.cut(
                    df[price_col], bins=price_bins, labels=price_labels
                )
                price_dist = (
                    df.groupby("price_range", observed=True)
                    .agg({amount_col_for_analysis: "sum", price_col: "count"})
                    .reset_index()
                )
                price_dist.columns = ["price_range", "total_amount", "order_count"]

                fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))
                ax1.bar(
                    price_dist["price_range"].astype(str),
                    price_dist["order_count"],
                    color="steelblue",
                )
                ax1.set_xlabel("Price Range", fontsize=11)
                ax1.set_ylabel("Order Count", fontsize=11)
                ax1.set_title("Orders by Price Range", fontsize=12, fontweight="bold")
                ax1.grid(True, alpha=0.3, axis="y")

                ax2.bar(
                    price_dist["price_range"].astype(str),
                    price_dist["total_amount"],
                    color="coral",
                )
                ax2.set_xlabel("Price Range", fontsize=11)
                ax2.set_ylabel("Total Amount (Yuan)", fontsize=11)
                ax2.set_title("Revenue by Price Range", fontsize=12, fontweight="bold")
                ax2.grid(True, alpha=0.3, axis="y")

                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart07_price_distribution.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["price_distribution"] = price_dist.to_dict("records")

        # Sales Category Analysis
        if category_col and category_col in df.columns:
            sales_cat = (
                df.groupby(category_col)[amount_col_for_analysis]
                .sum()
                .nlargest(10)
                .sort_values(ascending=True)
            )
            if len(sales_cat) > 0:
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.barh(range(len(sales_cat)), sales_cat.values, color="darkorange")
                ax.set_yticks(range(len(sales_cat)))
                ax.set_yticklabels([str(x)[:20] for x in sales_cat.index], fontsize=10)
                ax.set_xlabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "Revenue by Category", fontsize=14, fontweight="bold", pad=15
                )
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart08_sales_category.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)
                results["data"]["sales_category"] = {
                    str(k): float(v) for k, v in sales_cat.items()
                }

        # Hourly Pattern
        if date_cols:
            date_col = date_cols[0]
            df["hour"] = pd.to_datetime(df[date_col], errors="coerce").dt.hour
            hourly = df.groupby("hour")[amount_col_for_analysis].sum()
            if len(hourly) > 0:
                fig, ax = plt.subplots(figsize=(12, 5))
                ax.bar(hourly.index, hourly.values, color="mediumpurple")
                ax.set_xlabel("Hour of Day", fontsize=11)
                ax.set_ylabel("Total Amount (Yuan)", fontsize=11)
                ax.set_title(
                    "Revenue Distribution by Hour",
                    fontsize=14,
                    fontweight="bold",
                    pad=15,
                )
                ax.set_xticks(hourly.index)
                ax.grid(True, alpha=0.3, axis="y")
                plt.tight_layout()
                chart_path = os.path.join(charts_dir, "chart09_hourly_pattern.png")
                plt.savefig(chart_path, dpi=150, bbox_inches="tight")
                plt.close()
                results["charts"].append(chart_path)

    # Save report data as JSON
    report_data_path = os.path.join(output_dir, "report_data.json")
    with open(report_data_path, "w", encoding="utf-8") as f:
        json.dump(results["data"], f, ensure_ascii=False, indent=2, default=str)
    results["report_data_path"] = report_data_path

    return results


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = "resources/sample.csv"

    output_dir = os.path.dirname(file_path)
    results = summarize_csv(file_path, output_dir)

    print("=" * 60)
    print("ANALYSIS COMPLETE")
    print("=" * 60)
    print(f"Total Rows: {results['total_rows']}")
    print(f"Total Columns: {results['total_columns']}")
    print(f"\nCharts Created: {len(results['charts'])}")
    for chart in results["charts"]:
        print(f"  - {chart}")
    print(f"\nReport Data: {results.get('report_data_path', 'N/A')}")
