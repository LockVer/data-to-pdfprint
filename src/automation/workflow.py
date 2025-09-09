"""
自动化工作流程

负责协调整个自动化处理流程
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

from ..data.excel_reader import ExcelReader, PackagingModeDetector
from ..data.data_processor import DataProcessor, PackagingConfig
from ..pdf.generator import PDFGenerator


class AutomatedWorkflow:
    """
    自动化工作流程管理器
    """
    
    def __init__(self, project_root: Optional[str] = None):
        """
        初始化自动化工作流程
        
        Args:
            project_root: 项目根目录，如果未指定则使用当前目录的父目录
        """
        if project_root:
            self.project_root = Path(project_root)
        else:
            # 假设此文件在 src/automation/ 目录下
            self.project_root = Path(__file__).parent.parent.parent
        
        self.template_dir = self.project_root / "输入输出模板定义"
        self.output_dir = self.project_root / "output"
        
        # 确保输出目录存在
        self.output_dir.mkdir(exist_ok=True)
    
    def scan_template_directories(self) -> Dict[str, List[Path]]:
        """
        扫描模板目录中的所有Excel文件
        
        Returns:
            按包装模式分组的Excel文件列表
        """
        files_by_mode = {
            'regular': [],
            'separate_box': [],
            'set_box': []
        }
        
        if not self.template_dir.exists():
            raise FileNotFoundError(f"Template directory not found: {self.template_dir}")
        
        # 扫描所有子目录
        for mode_dir in self.template_dir.iterdir():
            if mode_dir.is_dir():
                mode_name = PackagingModeDetector.detect_from_path(str(mode_dir))
                
                # 递归扫描Excel文件
                excel_files = list(mode_dir.glob("**/*.xlsx"))
                files_by_mode[mode_name].extend(excel_files)
        
        return files_by_mode
    
    def process_single_file(self, excel_file: Path, config: PackagingConfig) -> Dict[str, Any]:
        """
        处理单个Excel文件
        
        Args:
            excel_file: Excel文件路径
            config: 包装配置参数
            
        Returns:
            处理结果信息
        """
        try:
            # 读取Excel文件
            reader = ExcelReader(str(excel_file))
            
            # 提取模板变量
            variables = reader.extract_template_variables()
            
            # 检测包装模式
            packaging_mode = reader.detect_packaging_mode()
            
            # 处理数据
            processor = DataProcessor()
            processed_data = processor.process_for_packaging_mode(
                variables, config, packaging_mode
            )
            
            # 生成PDF标签
            generator = PDFGenerator()
            
            # 创建输出目录
            output_subdir = self.output_dir / f"{variables.get('customer_code', 'unknown')}"
            output_subdir.mkdir(exist_ok=True)
            
            # 生成盒标和箱标
            results = {}
            
            if packaging_mode == 'regular':
                results = self._generate_regular_labels(generator, processed_data, output_subdir)
            elif packaging_mode == 'separate_box':
                results = self._generate_separate_box_labels(generator, processed_data, output_subdir)
            elif packaging_mode == 'set_box':
                results = self._generate_set_box_labels(generator, processed_data, output_subdir)
            
            return {
                'status': 'success',
                'excel_file': str(excel_file),
                'packaging_mode': packaging_mode,
                'variables': variables,
                'processed_data': processed_data,
                'output_files': results,
                'output_directory': str(output_subdir)
            }
            
        except Exception as e:
            return {
                'status': 'error',
                'excel_file': str(excel_file),
                'error': str(e)
            }
    
    def _generate_regular_labels(self, generator: PDFGenerator, data: Dict, output_dir: Path) -> Dict[str, str]:
        """生成常规模式标签"""
        box_label_file = output_dir / f"{data['customer_code']}_box_labels.pdf"
        case_label_file = output_dir / f"{data['customer_code']}_case_labels.pdf"
        
        # 生成盒标
        generator.generate_box_labels(data, str(box_label_file), template='regular')
        
        # 生成箱标（一盒入箱再两盒入一箱）
        generator.generate_case_labels(data, str(case_label_file), template='regular')
        
        return {
            'box_labels': str(box_label_file),
            'case_labels': str(case_label_file)
        }
    
    def _generate_separate_box_labels(self, generator: PDFGenerator, data: Dict, output_dir: Path) -> Dict[str, str]:
        """生成分盒模式标签"""
        box_label_file = output_dir / f"{data['customer_code']}_separate_box_labels.pdf"
        case_label_file = output_dir / f"{data['customer_code']}_separate_case_labels.pdf"
        
        # 生成盒标
        generator.generate_box_labels(data, str(box_label_file), template='separate_box')
        
        # 生成箱标（一盒入一小箱 再两小箱入一大箱）
        generator.generate_case_labels(data, str(case_label_file), template='separate_box')
        
        return {
            'box_labels': str(box_label_file),
            'case_labels': str(case_label_file)
        }
    
    def _generate_set_box_labels(self, generator: PDFGenerator, data: Dict, output_dir: Path) -> Dict[str, str]:
        """生成套盒模式标签"""
        box_label_file = output_dir / f"{data['customer_code']}_set_box_labels.pdf"
        case_label_file = output_dir / f"{data['customer_code']}_set_case_labels.pdf"
        
        # 生成盒标
        generator.generate_box_labels(data, str(box_label_file), template='set_box')
        
        # 生成箱标（六盒为一套 六盒入一小箱 再两套入一大箱）
        generator.generate_case_labels(data, str(case_label_file), template='set_box')
        
        return {
            'box_labels': str(box_label_file),
            'case_labels': str(case_label_file)
        }
    
    def batch_process_all_templates(self, config: PackagingConfig) -> List[Dict[str, Any]]:
        """
        批量处理所有模板文件
        
        Args:
            config: 包装配置参数
            
        Returns:
            所有文件的处理结果列表
        """
        files_by_mode = self.scan_template_directories()
        results = []
        
        # 处理所有文件
        for mode, files in files_by_mode.items():
            for excel_file in files:
                print(f"Processing: {excel_file}")
                result = self.process_single_file(excel_file, config)
                results.append(result)
                
                if result['status'] == 'success':
                    print(f"  ✅ Success: Generated labels for {result['variables'].get('customer_code', 'unknown')}")
                else:
                    print(f"  ❌ Error: {result['error']}")
        
        return results
    
    def generate_processing_report(self, results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        生成处理报告
        
        Args:
            results: 处理结果列表
            
        Returns:
            汇总报告
        """
        total = len(results)
        successful = len([r for r in results if r['status'] == 'success'])
        failed = total - successful
        
        # 按包装模式统计
        by_mode = {}
        for result in results:
            if result['status'] == 'success':
                mode = result['packaging_mode']
                if mode not in by_mode:
                    by_mode[mode] = 0
                by_mode[mode] += 1
        
        return {
            'summary': {
                'total_files': total,
                'successful': successful,
                'failed': failed,
                'success_rate': f"{(successful/total*100):.1f}%" if total > 0 else "0%"
            },
            'by_packaging_mode': by_mode,
            'failed_files': [r for r in results if r['status'] == 'error']
        }


def create_default_config() -> PackagingConfig:
    """
    创建默认的包装配置
    
    Returns:
        默认包装配置
    """
    return PackagingConfig(
        box_quantity=100,  # 默认分盒张数
        set_quantity=6,    # 默认分套张数
        small_box_capacity=1,  # 默认每小箱盒数（分盒模式可由用户修改）
        large_box_capacity=2   # 默认每大箱小箱数
    )