"""
命令行主程序入口

处理命令行参数解析和主要流程控制
"""

import click
import os
import json
from pathlib import Path
from typing import Optional

from ..automation.workflow import AutomatedWorkflow, create_default_config
from ..data.data_processor import PackagingConfig
from ..data.excel_reader import ExcelReader


@click.group()
@click.version_option(version='0.1.0', prog_name='data-to-pdf')
def cli():
    """
    Excel数据到PDF标签打印工具
    
    这是一个用于读取Excel数据并生成PDF标签的命令行工具。
    支持常规、分盒、套盒三种包装模式的自动化处理。
    """
    pass


@cli.command()
@click.option('--input', '-i', 'input_file', 
              help='输入Excel文件路径', 
              type=click.Path(exists=True))
@click.option('--box-quantity', '-b', 
              type=int, default=100,
              help='分盒张数 (默认: 100)')
@click.option('--set-quantity', '-s', 
              type=int, default=6,
              help='分套张数，仅套盒模式使用 (默认: 6)')
@click.option('--small-box-capacity', 
              type=int, default=2,
              help='小箱内的盒数 (默认: 2)')
@click.option('--large-box-capacity', 
              type=int, default=2,
              help='大箱内的小箱数 (默认: 2)')
@click.option('--output', '-o', 'output_dir', 
              help='输出目录路径', 
              type=click.Path())
def process(input_file, box_quantity, set_quantity, small_box_capacity, 
           large_box_capacity, output_dir):
    """
    处理单个Excel文件并生成PDF标签
    """
    click.echo("🚀 开始处理Excel文件...")
    
    if not input_file:
        click.echo("❌ 请指定输入文件")
        click.echo("   例如: data-to-pdf process --input 常规-LADIES_NIGHT.xlsx")
        return
    
    try:
        # 创建配置
        config = PackagingConfig(
            box_quantity=box_quantity,
            set_quantity=set_quantity,
            small_box_capacity=small_box_capacity,
            large_box_capacity=large_box_capacity
        )
        
        # 创建工作流
        workflow = AutomatedWorkflow()
        if output_dir:
            workflow.output_dir = Path(output_dir)
        
        # 处理文件
        result = workflow.process_single_file(Path(input_file), config)
        
        if result['status'] == 'success':
            click.echo("✅ 处理成功!")
            click.echo(f"📁 输出目录: {result['output_directory']}")
            click.echo(f"📄 生成文件:")
            for label_type, file_path in result['output_files'].items():
                click.echo(f"   {label_type}: {file_path}")
            
            # 显示处理信息
            variables = result['variables']
            click.echo(f"🏷️  处理信息:")
            click.echo(f"   客户编码: {variables.get('customer_code', 'N/A')}")
            click.echo(f"   主题: {variables.get('theme', 'N/A')}")
            click.echo(f"   包装模式: {result['packaging_mode']}")
            
        else:
            click.echo(f"❌ 处理失败: {result['error']}")
            
    except Exception as e:
        click.echo(f"❌ 发生错误: {e}")


@cli.command()
@click.option('--box-quantity', '-b', 
              type=int, default=100,
              help='分盒张数 (默认: 100)')
@click.option('--set-quantity', '-s', 
              type=int, default=6,
              help='分套张数，仅套盒模式使用 (默认: 6)')
@click.option('--small-box-capacity', 
              type=int, default=2,
              help='小箱内的盒数 (默认: 2)')
@click.option('--large-box-capacity', 
              type=int, default=2,
              help='大箱内的小箱数 (默认: 2)')
@click.option('--output', '-o', 'output_dir', 
              help='输出目录路径', 
              type=click.Path())
def batch(box_quantity, set_quantity, small_box_capacity, 
         large_box_capacity, output_dir):
    """
    批量处理所有模板目录中的Excel文件
    """
    click.echo("🚀 开始批量处理...")
    
    try:
        # 创建配置
        config = PackagingConfig(
            box_quantity=box_quantity,
            set_quantity=set_quantity,
            small_box_capacity=small_box_capacity,
            large_box_capacity=large_box_capacity
        )
        
        # 创建工作流
        workflow = AutomatedWorkflow()
        if output_dir:
            workflow.output_dir = Path(output_dir)
        
        # 显示配置信息
        click.echo("⚙️  处理配置:")
        click.echo(f"   分盒张数: {config.box_quantity}")
        click.echo(f"   分套张数: {config.set_quantity}")
        click.echo(f"   小箱容量: {config.small_box_capacity}")
        click.echo(f"   大箱容量: {config.large_box_capacity}")
        click.echo()
        
        # 批量处理
        results = workflow.batch_process_all_templates(config)
        
        # 生成报告
        report = workflow.generate_processing_report(results)
        
        # 显示结果
        click.echo("📊 处理报告:")
        click.echo(f"   总文件数: {report['summary']['total_files']}")
        click.echo(f"   成功: {report['summary']['successful']}")
        click.echo(f"   失败: {report['summary']['failed']}")
        click.echo(f"   成功率: {report['summary']['success_rate']}")
        
        if report['by_packaging_mode']:
            click.echo("📦 按包装模式统计:")
            for mode, count in report['by_packaging_mode'].items():
                mode_name = {'regular': '常规', 'separate_box': '分盒', 'set_box': '套盒'}.get(mode, mode)
                click.echo(f"   {mode_name}: {count}")
        
        if report['failed_files']:
            click.echo("❌ 失败文件:")
            for failed in report['failed_files']:
                click.echo(f"   {failed['excel_file']}: {failed['error']}")
        
        click.echo(f"📁 输出目录: {workflow.output_dir}")
        
    except Exception as e:
        click.echo(f"❌ 批量处理失败: {e}")


@cli.command()
@click.argument('excel_file', type=click.Path(exists=True))
def analyze(excel_file):
    """
    分析Excel文件并显示提取的信息
    """
    click.echo(f"🔍 分析文件: {excel_file}")
    
    try:
        reader = ExcelReader(excel_file)
        
        # 提取变量
        variables = reader.extract_template_variables()
        
        # 检测包装模式
        packaging_mode = reader.detect_packaging_mode()
        
        click.echo("📋 提取的信息:")
        click.echo(f"   客户编码: {variables.get('customer_code', 'N/A')}")
        click.echo(f"   主题: {variables.get('theme', 'N/A')}")
        click.echo(f"   开始号: {variables.get('start_number', 'N/A')}")
        click.echo(f"   包装模式: {packaging_mode}")
        
        # 显示包装模式说明
        mode_descriptions = {
            'regular': '常规模式 - 一盒入箱再两盒入一箱',
            'separate_box': '分盒模式 - 一盒入一小箱 再两小箱入一大箱',
            'set_box': '套盒模式 - 六盒为一套 六盒入一小箱 再两套入一大箱'
        }
        
        description = mode_descriptions.get(packaging_mode, '未知模式')
        click.echo(f"   模式说明: {description}")
        
    except Exception as e:
        click.echo(f"❌ 分析失败: {e}")


@cli.command()
def scan():
    """
    扫描输入输出模板定义目录中的所有Excel文件
    """
    click.echo("🔎 扫描模板目录...")
    
    try:
        workflow = AutomatedWorkflow()
        files_by_mode = workflow.scan_template_directories()
        
        total_files = sum(len(files) for files in files_by_mode.values())
        click.echo(f"📁 找到 {total_files} 个Excel文件:")
        
        for mode, files in files_by_mode.items():
            if files:
                mode_name = {'regular': '常规', 'separate_box': '分盒', 'set_box': '套盒'}.get(mode, mode)
                click.echo(f"\n📦 {mode_name}模式 ({len(files)}个文件):")
                for file_path in files:
                    click.echo(f"   {file_path.relative_to(workflow.project_root)}")
        
        if total_files == 0:
            click.echo("⚠️  未找到Excel文件")
            click.echo(f"   请检查目录: {workflow.template_dir}")
            
    except Exception as e:
        click.echo(f"❌ 扫描失败: {e}")


# 保持向后兼容的主命令
@cli.command()
@click.option('--input', '-i', 'input_file', 
              help='输入Excel文件路径', 
              type=click.Path(exists=True))
@click.option('--template', '-t', 
              default='basic', 
              help='使用的模板名称 (默认: basic)')
@click.option('--output', '-o', 'output_dir', 
              help='输出目录路径', 
              type=click.Path())
def legacy(input_file, template, output_dir):
    """
    旧版兼容命令 (已弃用，请使用 process 命令)
    """
    click.echo("⚠️  此命令已弃用，请使用 'process' 命令")
    click.echo("   例如: data-to-pdf process --input your_file.xlsx")


def main():
    """主入口函数"""
    cli()


if __name__ == "__main__":
    main()