import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import re

# 配置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def read_data(file_path):
    """
    自动识别文件格式并读取数据
    
    Args:
        file_path (str): 文件路径
        
    Returns:
        pandas.DataFrame: 读取的数据
    """
    file_path = Path(file_path)
    extension = file_path.suffix.lower()
    
    if extension == '.csv':
        return pd.read_csv(file_path, encoding='utf-8')
    elif extension in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    else:
        raise ValueError(f"不支持的文件格式：{extension}。支持的格式：CSV, XLSX, XLS")


def clean_amount(val):
    """清理金额字段，处理中文格式"""
    if pd.isna(val) or str(val).strip() in ['-', '']:
        return 0
    val_str = str(val).replace(',', '').replace(' ', '').replace('-', '0').replace('¥', '')
    try:
        return float(val_str)
    except:
        return 0


def analyze_data(file_path, output_dir=None):
    """
    全面分析数据文件（支持CSV和XLSX）
    
    Args:
        file_path (str): 数据文件路径
        output_dir (str): 输出目录，默认为文件所在目录
        
    Returns:
        dict: 包含分析结果和报告路径的字典
    """
    file_path = Path(file_path)
    if output_dir is None:
        output_dir = file_path.parent
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 读取数据
    df = read_data(file_path)
    summary_lines = []
    charts_created = []
    
    # ==========================================
    # 1. 数据概览
    # ==========================================
    summary_lines.append(f"# {file_path.stem} 数据分析报告")
    summary_lines.append("")
    summary_lines.append(f"**报告生成日期：** {pd.Timestamp.now().strftime('%Y年%m月%d日')}")
    summary_lines.append(f"**数据来源：** {file_path.name}")
    summary_lines.append(f"**数据总量：** {len(df):,} 条记录 × {len(df.columns)} 个字段")
    summary_lines.append("")
    summary_lines.append("---")
    summary_lines.append("")
    
    # ==========================================
    # 2. 数据质量分析
    # ==========================================
    summary_lines.append("## 一、数据质量分析")
    summary_lines.append("")
    
    missing = df.isnull().sum().sum()
    missing_pct = (missing / (df.shape[0] * df.shape[1])) * 100
    
    summary_lines.append("### 1.1 整体数据质量")
    summary_lines.append("")
    summary_lines.append(f"- **总记录数**：{len(df):,} 条")
    summary_lines.append(f"- **总字段数**：{len(df.columns)} 个")
    summary_lines.append(f"- **总缺失值**：{missing:,} 个")
    summary_lines.append(f"- **缺失率**：{missing_pct:.2f}%")
    summary_lines.append("")
    
    # 高缺失率字段
    high_missing = df.isnull().sum()
    high_missing = high_missing[high_missing > 0].sort_values(ascending=False)
    
    if len(high_missing) > 0:
        summary_lines.append("### 1.2 缺失值分布（TOP 10）")
        summary_lines.append("")
        summary_lines.append("| 字段名 | 缺失数量 | 缺失率 |")
        summary_lines.append("|--------|---------|--------|")
        for col, count in high_missing.head(10).items():
            pct = (count / len(df)) * 100
            summary_lines.append(f"| {col} | {count:,} | {pct:.1f}% |")
        summary_lines.append("")
    
    # ==========================================
    # 3. 数值型数据分析
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 二、数值型数据分析")
    summary_lines.append("")
    
    # 尝试清理金额字段
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    
    # 检查是否有看起来像金额的文本列
    potential_amount_cols = []
    for col in df.select_dtypes(include=['object']).columns:
        sample = df[col].dropna().head(10)
        if any('元' in str(v) or '¥' in str(v) or (str(v).replace(',', '').replace('.', '').replace('-', '').isdigit()) 
               for v in sample if pd.notna(v)):
            potential_amount_cols.append(col)
    
    summary_lines.append(f"### 2.1 数值字段概览")
    summary_lines.append("")
    summary_lines.append(f"- **数值型字段**：{len(numeric_cols)} 个")
    if len(numeric_cols) > 0:
        summary_lines.append(f"- **字段列表**：{', '.join(numeric_cols[:10])}")
    summary_lines.append("")
    
    if numeric_cols:
        summary_lines.append("### 2.2 描述性统计")
        summary_lines.append("")
        summary_lines.append("| 字段 | 计数 | 均值 | 标准差 | 最小值 | 25% | 50% | 75% | 最大值 |")
        summary_lines.append("|------|------|------|--------|--------|-----|-----|-----|--------|")
        stats = df[numeric_cols].describe().T
        for col in numeric_cols[:10]:
            if col in stats.index:
                row = stats.loc[col]
                summary_lines.append(
                    f"| {col} | {row['count']:.0f} | {row['mean']:.2f} | {row['std']:.2f} | "
                    f"{row['min']:.2f} | {row['25%']:.2f} | {row['50%']:.2f} | {row['75%']:.2f} | {row['max']:.2f} |"
                )
        summary_lines.append("")
        
        # 相关性分析
        if len(numeric_cols) > 1:
            summary_lines.append("### 2.3 相关性分析")
            summary_lines.append("")
            corr_matrix = df[numeric_cols].corr()
            
            # 找出强相关性
            strong_corr = []
            for i in range(len(corr_matrix.columns)):
                for j in range(i+1, len(corr_matrix.columns)):
                    corr_val = corr_matrix.iloc[i, j]
                    if abs(corr_val) > 0.5:
                        strong_corr.append((corr_matrix.columns[i], corr_matrix.columns[j], corr_val))
            
            if strong_corr:
                summary_lines.append("**强相关字段对（|r| > 0.5）：**")
                summary_lines.append("")
                for col1, col2, corr in strong_corr[:5]:
                    summary_lines.append(f"- {col1} ↔ {col2}: r = {corr:.3f}")
                summary_lines.append("")
            
            # 生成热力图
            try:
                plt.figure(figsize=(10, 8))
                sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0,
                           square=True, linewidths=1, fmt='.2f',
                           xticklabels=numeric_cols[:10], yticklabels=numeric_cols[:10])
                plt.title('Correlation Heatmap')
                plt.tight_layout()
                chart_path = output_dir / 'correlation_heatmap.png'
                plt.savefig(chart_path, dpi=150, bbox_inches='tight')
                plt.close()
                charts_created.append('correlation_heatmap.png')
                summary_lines.append(f"*图1：相关性热力图已生成*")
                summary_lines.append("")
            except Exception as e:
                summary_lines.append(f"*相关性图生成失败：{str(e)}*")
                summary_lines.append("")
    
    # ==========================================
    # 4. 分类型数据分析
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 三、分类型数据分析")
    summary_lines.append("")
    
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    categorical_cols = [c for c in categorical_cols if 'id' not in c.lower() and 'url' not in c.lower()]
    
    if categorical_cols:
        summary_lines.append(f"### 3.1 分类字段概览")
        summary_lines.append("")
        summary_lines.append(f"- **分类字段数量**：{len(categorical_cols)} 个")
        summary_lines.append(f"- **字段列表**：{', '.join(categorical_cols[:10])}")
        summary_lines.append("")
        
        # TOP 5 分类字段的分布
        summary_lines.append("### 3.2 TOP 5 分类字段分布")
        summary_lines.append("")
        
        for col_idx, col in enumerate(categorical_cols[:5]):
            value_counts = df[col].value_counts()
            summary_lines.append(f"**{col}**（共 {len(value_counts)} 个唯一值）")
            summary_lines.append("")
            summary_lines.append(f"| 值 | 数量 | 占比 |")
            summary_lines.append(f"|----|------|------|")
            for val, count in value_counts.head(10).items():
                pct = (count / len(df)) * 100
                val_str = str(val)[:50]  # 截断过长的值
                summary_lines.append(f"| {val_str} | {count:,} | {pct:.1f}% |")
            summary_lines.append("")
        
        # 生成分布图
        try:
            n_plots = min(4, len(categorical_cols))
            fig, axes = plt.subplots(2, 2, figsize=(14, 10))
            axes = axes.flatten()
            
            for idx, col in enumerate(categorical_cols[:n_plots]):
                value_counts = df[col].value_counts().head(10)
                axes[idx].barh(range(len(value_counts)), value_counts.values, color='steelblue')
                axes[idx].set_yticks(range(len(value_counts)))
                axes[idx].set_yticklabels([str(v)[:30] for v in value_counts.index])
                axes[idx].set_title(f'TOP 10: {col}')
                axes[idx].set_xlabel('Count')
                axes[idx].grid(True, alpha=0.3, axis='x')
            
            for idx in range(n_plots, 4):
                axes[idx].set_visible(False)
            
            plt.tight_layout()
            chart_path = output_dir / 'categorical_distributions.png'
            plt.savefig(chart_path, dpi=150, bbox_inches='tight')
            plt.close()
            charts_created.append('categorical_distributions.png')
            summary_lines.append("*图2：分类字段分布图已生成*")
            summary_lines.append("")
        except Exception as e:
            summary_lines.append(f"*分类分布图生成失败：{str(e)}*")
            summary_lines.append("")
    
    # ==========================================
    # 5. 时间序列分析（如果存在日期列）
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 四、时间序列分析")
    summary_lines.append("")
    
    date_cols = [c for c in df.columns if 'date' in c.lower() or 'time' in c.lower() or '日期' in c or '时间' in c]
    
    if date_cols:
        date_col = date_cols[0]
        try:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            date_valid = df[date_col].dropna()
            
            if len(date_valid) > 0:
                summary_lines.append(f"### 4.1 时间范围")
                summary_lines.append("")
                summary_lines.append(f"- **日期字段**：{date_col}")
                summary_lines.append(f"- **时间范围**：{date_valid.min()} 至 {date_valid.max()}")
                summary_lines.append(f"- **时间跨度**：{(date_valid.max() - date_valid.min()).days} 天")
                summary_lines.append(f"- **有效记录**：{len(date_valid):,} 条")
                summary_lines.append("")
                
                # 按月统计
                if len(date_valid) > 10:
                    summary_lines.append("### 4.2 月度趋势")
                    summary_lines.append("")
                    
                    df['月份'] = df[date_col].dt.to_period('M')
                    monthly_stats = df.groupby('月份').size()
                    
                    summary_lines.append("| 月份 | 记录数 |")
                    summary_lines.append("|------|--------|")
                    for period, count in monthly_stats.items():
                        summary_lines.append(f"| {period} | {count:,} |")
                    summary_lines.append("")
                    
                    # 如果有数值列，生成时间序列图
                    if numeric_cols and len(numeric_cols) > 0:
                        try:
                            fig, ax = plt.subplots(figsize=(12, 6))
                            num_col = numeric_cols[0]
                            monthly_agg = df.groupby('月份')[num_col].agg(['mean', 'sum', 'count'])
                            monthly_agg['mean'].plot(ax=ax, marker='o', linewidth=2, markersize=6, label='均值')
                            plt.title(f'{num_col} 月度趋势')
                            plt.xlabel('月份')
                            plt.ylabel(num_col)
                            plt.legend()
                            plt.grid(True, alpha=0.3)
                            plt.xticks(rotation=45)
                            plt.tight_layout()
                            chart_path = output_dir / 'time_series_trend.png'
                            plt.savefig(chart_path, dpi=150, bbox_inches='tight')
                            plt.close()
                            charts_created.append('time_series_trend.png')
                            summary_lines.append(f"*图3：{num_col} 月度趋势图已生成*")
                            summary_lines.append("")
                        except Exception as e:
                            summary_lines.append(f"*时间序列图生成失败：{str(e)}*")
                            summary_lines.append("")
        except Exception as e:
            summary_lines.append(f"*时间序列分析失败：{str(e)}*")
            summary_lines.append("")
    else:
        summary_lines.append("*未检测到日期/时间字段，跳过时间序列分析*")
        summary_lines.append("")
    
    # ==========================================
    # 6. 数据分布分析
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 五、数据分布分析")
    summary_lines.append("")
    
    if numeric_cols:
        summary_lines.append("### 5.1 数值字段分布")
        summary_lines.append("")
        
        try:
            n_cols = min(4, len(numeric_cols))
            fig, axes = plt.subplots(2, 2, figsize=(12, 10))
            axes = axes.flatten()
            
            for idx, col in enumerate(numeric_cols[:4]):
                axes[idx].hist(df[col].dropna(), bins=30, edgecolor='black', alpha=0.7, color='steelblue')
                axes[idx].set_title(f'Distribution: {col}')
                axes[idx].set_xlabel(col)
                axes[idx].set_ylabel('Frequency')
                axes[idx].grid(True, alpha=0.3)
            
            for idx in range(n_cols, 4):
                axes[idx].set_visible(False)
            
            plt.tight_layout()
            chart_path = output_dir / 'distributions.png'
            plt.savefig(chart_path, dpi=150, bbox_inches='tight')
            plt.close()
            charts_created.append('distributions.png')
            summary_lines.append("*图4：数值字段分布图已生成*")
            summary_lines.append("")
        except Exception as e:
            summary_lines.append(f"*分布图生成失败：{str(e)}*")
            summary_lines.append("")
    
    # ==========================================
    # 7. 关键洞察
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 六、关键洞察与建议")
    summary_lines.append("")
    
    summary_lines.append("### 6.1 数据特征")
    summary_lines.append("")
    
    # 基于数据分析生成洞察
    insights = []
    
    # 记录数洞察
    if len(df) > 10000:
        insights.append(f"- ✅ 数据量充足（{len(df):,}条），适合进行深度分析")
    elif len(df) > 1000:
        insights.append(f"- ⚠️ 数据量中等（{len(df):,}条），分析结果具有一定参考价值")
    else:
        insights.append(f"- ⚠️ 数据量较小（{len(df):,}条），建议谨慎解读分析结果")
    
    # 缺失值洞察
    if missing_pct < 5:
        insights.append(f"- ✅ 数据质量良好，缺失率仅 {missing_pct:.2f}%")
    elif missing_pct < 15:
        insights.append(f"- ⚠️ 存在一定缺失（{missing_pct:.2f}%），建议补充完善")
    else:
        insights.append(f"- ❌ 缺失率较高（{missing_pct:.2f}%），可能影响分析准确性")
    
    # 字段多样性洞察
    if len(categorical_cols) > 5:
        insights.append(f"- ✅ 分类字段丰富（{len(categorical_cols)}个），支持多维度分析")
    
    if len(numeric_cols) > 3:
        insights.append(f"- ✅ 数值字段充足（{len(numeric_cols)}个），可进行统计分析")
    
    for insight in insights:
        summary_lines.append(insight)
    summary_lines.append("")
    
    summary_lines.append("### 6.2 改进建议")
    summary_lines.append("")
    summary_lines.append("1. **数据录入规范**：建立必填字段校验，降低缺失率")
    summary_lines.append("2. **字段标准化**：统一字段命名和数据格式")
    summary_lines.append("3. **定期数据审计**：每月检查数据质量和完整性")
    summary_lines.append("4. **补充维度信息**：增加业务相关的分类字段")
    summary_lines.append("")
    
    # ==========================================
    # 8. 图表清单
    # ==========================================
    if charts_created:
        summary_lines.append("---")
        summary_lines.append("")
        summary_lines.append("## 七、图表清单")
        summary_lines.append("")
        for idx, chart in enumerate(charts_created, 1):
            summary_lines.append(f"{idx}. **{chart}** - 已生成")
        summary_lines.append("")
    
    # ==========================================
    # 9. 总结
    # ==========================================
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append("## 八、总结")
    summary_lines.append("")
    summary_lines.append(f"本次分析共处理 **{len(df):,}** 条记录、**{len(df.columns)}** 个字段，")
    summary_lines.append(f"生成了 **{len(charts_created)}** 张可视化图表。")
    summary_lines.append(f"数据缺失率为 **{missing_pct:.2f}%**，")
    summary_lines.append(f"{'数据质量良好，分析结果可信。' if missing_pct < 10 else '建议进一步优化数据质量。'}")
    summary_lines.append("")
    summary_lines.append("---")
    summary_lines.append("")
    summary_lines.append(f"*报告生成时间：{pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}*")
    
    # 保存 Markdown 报告
    md_report_path = output_dir / f"{file_path.stem}_分析报告.md"
    with open(md_report_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(summary_lines))
    
    return {
        'success': True,
        'markdown_report': str(md_report_path),
        'charts': charts_created,
        'total_records': len(df),
        'total_columns': len(df.columns),
        'missing_rate': missing_pct
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    else:
        file_path = "resources/sample.csv"
        output_dir = None
    
    result = analyze_data(file_path, output_dir)
    
    if result['success']:
        print(f"\n✅ 分析完成！")
        print(f"📄 Markdown报告：{result['markdown_report']}")
        print(f"📊 生成图表：{len(result['charts'])} 张")
        for chart in result['charts']:
            print(f"   - {chart}")
    else:
        print(f"\n❌ 分析失败")
