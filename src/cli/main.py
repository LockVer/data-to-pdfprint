"""
命令行主程序入口

处理命令行参数解析和主要流程控制
"""

import click
import os
from pathlib import Path


@click.command()
@click.argument('input_files', nargs=-1, type=click.Path(exists=True))
@click.option(
    "--input",
    "-i",
    "input_file",
    help="输入Excel文件路径",
    type=click.Path(exists=True),
)
@click.option("--template", "-t", default="basic", help="使用的模板名称 (默认: basic)")
@click.option("--output", "-o", "output_dir", help="输出目录路径", type=click.Path())
@click.version_option(version="0.1.0", prog_name="data-to-pdf")
def main(input_files, input_file, template, output_dir):
    """
    Excel数据到PDF标签打印工具

    支持拖拽Excel文件: data-to-pdf file1.xlsx file2.xlsx
    或使用选项: data-to-pdf --input file.xlsx
    """
    # 处理拖拽的文件
    target_file = None
    if input_files:
        target_file = input_files[0]  # 使用第一个拖拽的文件
    elif input_file:
        target_file = input_file
    
    click.echo("🚀 欢迎使用 Data to PDF Print 工具!")

    # 显示当前配置
    click.echo(f"✨ 当前配置:")
    click.echo(f"   输入文件: {target_file or '未指定'}")
    click.echo(f"   使用模板: {template}")
    click.echo(f"   输出目录: {output_dir or '默认输出目录'}")

    # 检查输入文件
    if target_file:
        file_path = Path(target_file)
        click.echo(f"📂 文件信息:")
        click.echo(f"   文件名: {file_path.name}")
        click.echo(f"   文件大小: {file_path.stat().st_size} 字节")

        # 检查文件扩展名
        if file_path.suffix.lower() in [".xlsx", ".xls"]:
            click.echo("✅ Excel文件格式正确")
        else:
            click.echo("⚠️  警告: 文件可能不是Excel格式")
    else:
        click.echo("💡 提示: 使用 --input 参数指定Excel文件")
        click.echo("   例如: data-to-pdf --input data.xlsx")

    if target_file:
        try:
            # 导入处理模块
            import sys
            import os
            sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))
            
            from src.pdf.generator import PDFGenerator
            import pandas as pd
            
            click.echo("🔄 正在处理Excel文件...")
            
            # 读取Excel数据
            df = pd.read_excel(target_file, header=None)
            total_count = df.iloc[3,5]
            
            # 提取数据
            pdf_data = {
                '客户编码': df.iloc[3,0],
                '主题': df.iloc[3,1], 
                '排列要求': df.iloc[3,2],
                '订单数量': df.iloc[3,3],
                '张/盒': df.iloc[3,4],
                '总张数': total_count
            }
            
            click.echo("📊 提取的数据:")
            for key, value in pdf_data.items():
                click.echo(f"   {key}: {value}")
            
            # 生成PDF
            generator = PDFGenerator()
            output_path = output_dir or "output"
            pdf_file = f"{output_path}/label_{pdf_data['客户编码']}.pdf"
            
            os.makedirs(output_path, exist_ok=True)
            generator.create_label_pdf(pdf_data, pdf_file)
            
            click.echo(f"✅ PDF生成成功: {pdf_file}")
            click.echo(f"📄 总张数: {total_count}")
            
        except Exception as e:
            click.echo(f"❌ 处理失败: {e}")
            return

    click.echo("👋 感谢使用!")


if __name__ == "__main__":
    main()
