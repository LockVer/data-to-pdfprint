"""
命令行主程序入口

处理命令行参数解析和主要流程控制
"""

import click
import os
from pathlib import Path

@click.command()
@click.option('--input', '-i', 'input_file', 
              help='输入Excel文件路径', 
              type=click.Path(exists=True))
@click.option('--template', '-t', 
              default='basic', 
              help='使用的模板名称 (默认: basic)')
@click.option('--output', '-o', 'output_dir', 
              help='输出目录路径', 
              type=click.Path())
@click.version_option(version='0.1.0', prog_name='data-to-pdf')
def main(input_file, template, output_dir):
    """
    Excel数据到PDF标签打印工具
    
    这是一个用于读取Excel数据并生成PDF标签的命令行工具。
    """
    click.echo("🚀 欢迎使用 Data to PDF Print 工具!")
    
    # 显示当前配置
    click.echo(f"✨ 当前配置:")
    click.echo(f"   输入文件: {input_file or '未指定'}")
    click.echo(f"   使用模板: {template}")
    click.echo(f"   输出目录: {output_dir or '默认输出目录'}")
    
    # 检查输入文件
    if input_file:
        file_path = Path(input_file)
        click.echo(f"📂 文件信息:")
        click.echo(f"   文件名: {file_path.name}")
        click.echo(f"   文件大小: {file_path.stat().st_size} 字节")
        
        # 检查文件扩展名
        if file_path.suffix.lower() in ['.xlsx', '.xls']:
            click.echo("✅ Excel文件格式正确")
        else:
            click.echo("⚠️  警告: 文件可能不是Excel格式")
    else:
        click.echo("💡 提示: 使用 --input 参数指定Excel文件")
        click.echo("   例如: data-to-pdf --input data.xlsx")
    
    if input_file:
        click.echo("🔄 处理中... (功能开发中)")
        click.echo("✅ 完成! (这是演示输出)")
    
    click.echo("👋 感谢使用!")

if __name__ == "__main__":
    main()