#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据分析和报告生成主脚本

整合数据分析和 Markdown 转 DOCX 功能，一键生成完整的 Word 分析报告。

使用方法：
    python generate_report.py <数据文件路径> [输出目录]

支持的输入格式：
    - CSV (.csv)
    - Excel (.xlsx, .xls)

输出：
    - Markdown 格式分析报告
    - Word (.docx) 格式分析报告
    - 可视化图表（PNG）
"""

import sys
import os
from pathlib import Path
import traceback

# 添加当前目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from analyze import analyze_data
from markdown_to_docx import parse_markdown, setup_document
from docx import Document


def generate_report(file_path, output_dir=None):
    """
    生成完整的数据分析报告（Markdown + DOCX + 图表）
    
    Args:
        file_path (str): 数据文件路径（CSV 或 XLSX）
        output_dir (str): 输出目录，默认为数据文件所在目录
        
    Returns:
        dict: 包含所有生成文件路径的字典
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        raise FileNotFoundError(f"找不到文件：{file_path}")
    
    if output_dir is None:
        output_dir = file_path.parent
    else:
        output_dir = Path(output_dir)
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print("=" * 80)
    print("📊 数据分析和报告生成工具")
    print("=" * 80)
    print(f"\n📁 输入文件：{file_path}")
    print(f"📂 输出目录：{output_dir}")
    print("\n" + "=" * 80)
    
    # 步骤 1: 数据分析
    print("\n【步骤 1/3】正在分析数据...")
    print("-" * 80)
    
    try:
        analysis_result = analyze_data(str(file_path), str(output_dir))
        
        if not analysis_result['success']:
            print("❌ 数据分析失败")
            return None
        
        print(f"✅ 数据分析完成")
        print(f"   - 记录数：{analysis_result['total_records']:,}")
        print(f"   - 字段数：{analysis_result['total_columns']}")
        print(f"   - 缺失率：{analysis_result['missing_rate']:.2f}%")
        print(f"   - 生成图表：{len(analysis_result['charts'])} 张")
        
    except Exception as e:
        print(f"❌ 数据分析失败：{str(e)}")
        traceback.print_exc()
        return None
    
    # 步骤 2: 生成 DOCX
    print("\n【步骤 2/3】正在生成 Word 报告...")
    print("-" * 80)
    
    try:
        md_report = Path(analysis_result['markdown_report'])
        docx_report = output_dir / f"{file_path.stem}_分析报告.docx"
        
        # 创建文档
        doc = Document()
        setup_document(doc)
        
        # 解析 Markdown 并转换
        parse_markdown(str(md_report), doc)
        
        # 保存
        doc.save(str(docx_report))
        
        print(f"✅ Word 报告生成完成")
        print(f"   - {docx_report.name}")
        
    except Exception as e:
        print(f"❌ Word 报告生成失败：{str(e)}")
        traceback.print_exc()
        docx_report = None
    
    # 步骤 3: 总结
    print("\n【步骤 3/3】生成文件清单")
    print("-" * 80)
    
    generated_files = {
        'markdown_report': analysis_result['markdown_report'],
        'docx_report': str(docx_report) if docx_report else None,
        'charts': [str(output_dir / chart) for chart in analysis_result['charts']]
    }
    
    print(f"\n📄 Markdown 报告：")
    print(f"   {generated_files['markdown_report']}")
    
    if generated_files['docx_report']:
        print(f"\n📄 Word 报告：")
        print(f"   {generated_files['docx_report']}")
    
    if generated_files['charts']:
        print(f"\n📊 可视化图表（{len(generated_files['charts'])} 张）：")
        for chart in generated_files['charts']:
            print(f"   - {Path(chart).name}")
    
    print("\n" + "=" * 80)
    print("✅ 报告生成完成！")
    print("=" * 80)
    
    return generated_files


def main():
    """命令行入口"""
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n示例：")
        print("  python generate_report.py data.csv")
        print("  python generate_report.py data.xlsx ./output")
        print("  python generate_report.py 'C:\\Users\\data.xlsx'")
        sys.exit(1)
    
    file_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        result = generate_report(file_path, output_dir)
        
        if result is None:
            sys.exit(1)
        else:
            print("\n📌 生成的文件：")
            for key, value in result.items():
                if value:
                    if isinstance(value, list):
                        print(f"   {key}: {len(value)} 个文件")
                    else:
                        print(f"   {key}: {value}")
            sys.exit(0)
            
    except Exception as e:
        print(f"\n❌ 错误：{str(e)}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
