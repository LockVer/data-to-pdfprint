#!/usr/bin/env python3
"""
更新后Serial逻辑的完整测试用例
覆盖所有场景和边界情况
"""

import math
import sys
import os
import json
from datetime import datetime
from typing import Dict, Any, List

# 添加项目路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


class SerialLogicTester:
    """Serial逻辑完整测试器"""
    
    def __init__(self):
        self.processor = SplitBoxDataProcessor()
        self.passed_tests = 0
        self.failed_tests = 0
        self.test_results = []
        self.start_time = datetime.now()
    
    def run_all_tests(self):
        """运行所有测试用例"""
        print("🧪 Serial逻辑完整测试套件")
        print("=" * 80)
        
        # 1. 基础功能测试
        print("\n📋 第1组：基础功能测试")
        self._test_basic_functionality()
        
        # 2. 模式判断测试
        print("\n📋 第2组：模式自动判断测试")
        self._test_mode_detection()
        
        # 3. 一套分多箱测试
        print("\n📋 第3组：一套分多箱场景测试")
        self._test_one_set_multiple_boxes()
        
        # 4. 多套分一箱测试  
        print("\n📋 第4组：多套分一箱场景测试")
        self._test_multiple_sets_one_box()
        
        # 5. 一套分一箱测试
        print("\n📋 第5组：一套分一箱场景测试")
        self._test_one_set_one_box()
        
        # 6. 边界情况测试
        print("\n📋 第6组：边界情况测试")
        self._test_edge_cases()
        
        # 7. 格式化测试
        print("\n📋 第7组：格式化测试")
        self._test_format_requirements()
        
        # 8. 兼容性测试
        print("\n📋 第8组：向后兼容性测试")
        self._test_backward_compatibility()
        
        # 输出结果
        self._print_summary()
        self._export_results()
    
    def _test_basic_functionality(self):
        """测试基础功能"""
        
        # 1.1 单盒Serial生成
        self._run_test({
            "name": "基础_单盒Serial生成",
            "type": "box_serial",
            "params": {
                "box_num": 15,
                "base_number": "DSK01001-01", 
                "boxes_per_set": 10
            },
            "expected": "DSK01002-05"
        })
        
        # 1.2 小箱Serial范围生成
        self._run_test({
            "name": "基础_小箱Serial范围",
            "type": "small_box_range",
            "params": {
                "small_box_num": 2,
                "base_number": "DSK01001-01",
                "boxes_per_set": 6,
                "boxes_per_small_box": 3,
                "total_boxes": 12
            },
            "expected": "DSK01001-04-DSK01001-06"
        })
        
        # 1.3 大箱Serial范围生成
        self._run_test({
            "name": "基础_大箱Serial范围",
            "type": "large_box_range", 
            "params": {
                "large_box_num": 1,
                "base_number": "DSK01001-01",
                "boxes_per_set": 8,
                "boxes_per_small_box": 4,
                "small_boxes_per_large_box": 2,
                "total_boxes": 16
            },
            "expected": "DSK01001-01-DSK01001-08"
        })
    
    def _test_mode_detection(self):
        """测试模式自动判断"""
        
        # 2.1 检测多套分一箱模式
        self._run_test({
            "name": "模式_多套分一箱检测",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "JAW01001-01",
                "boxes_per_set": 3,           # 每套3盒
                "boxes_per_small_box": 8,     # 大箱容量8盒 >= 每套3盒
                "small_boxes_per_large_box": 1,
                "total_boxes": 12
            },
            "expected": "JAW01001-01-JAW01003-02"  # 跨套显示
        })
        
        # 2.2 检测一套分多箱模式
        self._run_test({
            "name": "模式_一套分多箱检测", 
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "JAW01001-01",
                "boxes_per_set": 15,          # 每套15盒
                "boxes_per_small_box": 8,     # 大箱容量8盒 < 每套15盒
                "small_boxes_per_large_box": 1,
                "total_boxes": 30
            },
            "expected": "JAW01001-01-JAW01001-08"  # 套内显示
        })
        
        # 2.3 检测一套分一箱模式
        self._run_test({
            "name": "模式_一套分一箱检测",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "BAT01001-01",
                "boxes_per_set": 6,           # 每套6盒
                "boxes_per_small_box": 6,     # 大箱容量6盒 = 每套6盒
                "small_boxes_per_large_box": 1,
                "total_boxes": 12
            },
            "expected": "BAT01001-01-BAT01001-06"  # 套内全部显示
        })
    
    def _test_one_set_multiple_boxes(self):
        """测试一套分多箱场景"""
        
        # 3.1 一套分2个大箱
        self._run_test({
            "name": "一套分多箱_2个大箱",
            "type": "large_box_range",
            "params": {
                "large_box_num": 2,
                "base_number": "DSK01001-01",
                "boxes_per_set": 15,
                "boxes_per_small_box": 8,
                "small_boxes_per_large_box": 1,
                "total_boxes": 30
            },
            "expected": "DSK01001-09-DSK01001-15"  # 套1的第2个大箱
        })
        
        # 3.2 一套分3个小箱
        self._run_test({
            "name": "一套分多箱_3个小箱",
            "type": "small_box_range", 
            "params": {
                "small_box_num": 3,
                "base_number": "DSK01001-01",
                "boxes_per_set": 9,
                "boxes_per_small_box": 3,
                "total_boxes": 18
            },
            "expected": "DSK01001-07-DSK01001-09"  # 套1的第3个小箱
        })
        
        # 3.3 跨套验证（第3个大箱应该属于套2）
        self._run_test({
            "name": "一套分多箱_跨套验证",
            "type": "large_box_range",
            "params": {
                "large_box_num": 3,
                "base_number": "DSK01001-01", 
                "boxes_per_set": 15,
                "boxes_per_small_box": 8,
                "small_boxes_per_large_box": 1,
                "total_boxes": 30
            },
            "expected": "DSK01002-01-DSK01002-08"  # 套2的第1个大箱
        })
    
    def _test_multiple_sets_one_box(self):
        """测试多套分一箱场景"""
        
        # 4.1 3套分1个大箱
        self._run_test({
            "name": "多套分一箱_3套分1大箱",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "JAW01001-01",
                "boxes_per_set": 3,
                "boxes_per_small_box": 10,
                "small_boxes_per_large_box": 1,
                "total_boxes": 15
            },
            "expected": "JAW01001-01-JAW01004-01"  # 包含套1-3全部+套4部分
        })
        
        # 4.2 2套分1个小箱
        self._run_test({
            "name": "多套分一箱_2套分1小箱",
            "type": "small_box_range",
            "params": {
                "small_box_num": 1,
                "base_number": "JAW01001-01",
                "boxes_per_set": 4,
                "boxes_per_small_box": 8,
                "total_boxes": 16
            },
            "expected": "JAW01001-01-JAW01002-04"  # 包含套1全部+套2全部
        })
        
        # 4.3 大容量箱子跨多套
        self._run_test({
            "name": "多套分一箱_大容量跨套",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "TEST01001-01",
                "boxes_per_set": 2,
                "boxes_per_small_box": 15,
                "small_boxes_per_large_box": 1,
                "total_boxes": 20
            },
            "expected": "TEST01001-01-TEST01008-01"  # 跨8套（15盒=7.5套，实际到第8套第1盒）
        })
    
    def _test_one_set_one_box(self):
        """测试一套分一箱场景"""
        
        # 5.1 标准一套一箱
        self._run_test({
            "name": "一套一箱_标准情况",
            "type": "large_box_range",
            "params": {
                "large_box_num": 2,
                "base_number": "BAT01001-01",
                "boxes_per_set": 6,
                "boxes_per_small_box": 6,
                "small_boxes_per_large_box": 1,
                "total_boxes": 18
            },
            "expected": "BAT01002-01-BAT01002-06"  # 第2套的完整范围
        })
        
        # 5.2 小箱一套一箱
        self._run_test({
            "name": "一套一箱_小箱情况",
            "type": "small_box_range",
            "params": {
                "small_box_num": 3,
                "base_number": "BAT01001-01",
                "boxes_per_set": 4,
                "boxes_per_small_box": 4,
                "total_boxes": 12
            },
            "expected": "BAT01003-01-BAT01003-04"  # 第3套的完整范围
        })
    
    def _test_edge_cases(self):
        """测试边界情况"""
        
        # 6.1 最后一箱盒数不足
        self._run_test({
            "name": "边界_最后箱盒数不足",
            "type": "large_box_range",
            "params": {
                "large_box_num": 2,
                "base_number": "DSK01001-01",
                "boxes_per_set": 15,
                "boxes_per_small_box": 8,
                "small_boxes_per_large_box": 1,
                "total_boxes": 22  # 第2套只有7盒
            },
            "expected": "DSK01001-09-DSK01001-15"  # 第1套第2箱，正常范围
        })
        
        # 6.2 总盒数边界检查
        self._run_test({
            "name": "边界_总盒数限制",
            "type": "large_box_range",
            "params": {
                "large_box_num": 3,
                "base_number": "DSK01001-01",
                "boxes_per_set": 15,
                "boxes_per_small_box": 8,
                "small_boxes_per_large_box": 1,
                "total_boxes": 22  # 第2套只有7盒，第3箱应该是第2套第1-7盒
            },
            "expected": "DSK01002-01-DSK01002-07"  # 受总盒数限制
        })
        
        # 6.3 单盒情况
        self._run_test({
            "name": "边界_单盒场景",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "DSK01001-01",
                "boxes_per_set": 1,
                "boxes_per_small_box": 1,
                "small_boxes_per_large_box": 1,
                "total_boxes": 3
            },
            "expected": "DSK01001-01-DSK01001-01"  # 起始结束相同但仍显示为范围
        })
        
        # 6.4 大数量测试
        self._run_test({
            "name": "边界_大数量",
            "type": "large_box_range",
            "params": {
                "large_box_num": 10,
                "base_number": "DSK01001-01",
                "boxes_per_set": 100,
                "boxes_per_small_box": 50,
                "small_boxes_per_large_box": 1,
                "total_boxes": 2000
            },
            "expected": "DSK01005-51-DSK01005-100"  # 大箱#10：套5的第2个大箱（盒451-500 → 套5的51-100）
        })
    
    def _test_format_requirements(self):
        """测试格式化要求"""
        
        # 7.1 范围分隔符测试
        self._run_test({
            "name": "格式_范围分隔符",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "TEST01001-01",
                "boxes_per_set": 10,
                "boxes_per_small_box": 5,
                "small_boxes_per_large_box": 1,
                "total_boxes": 20
            },
            "expected": "TEST01001-01-TEST01001-05",
            "check_format": True,
            "format_rules": ["must_contain_dash", "no_tilde"]
        })
        
        # 7.2 始终显示范围格式
        self._run_test({
            "name": "格式_始终范围格式",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "BAT01001-01",
                "boxes_per_set": 5,
                "boxes_per_small_box": 5,
                "small_boxes_per_large_box": 1,
                "total_boxes": 10
            },
            "expected": "BAT01001-01-BAT01001-05",
            "check_format": True,
            "format_rules": ["always_range_format"]
        })
        
        # 7.3 序列号主号递增
        self._run_test({
            "name": "格式_主号递增验证",
            "type": "box_serial",
            "params": {
                "box_num": 25,  # 第3套的第5盒
                "base_number": "DSK01001-01",
                "boxes_per_set": 10
            },
            "expected": "DSK01003-05",
            "check_format": True,
            "format_rules": ["correct_main_number"]
        })
    
    def _test_backward_compatibility(self):
        """测试向后兼容性"""
        
        # 8.1 传统分盒模式（盒/套=1）
        # 注意：这里应该调用传统逻辑，但我们测试新逻辑在盒/套=1时的表现
        self._run_test({
            "name": "兼容_传统分盒模式",
            "type": "large_box_range",
            "params": {
                "large_box_num": 1,
                "base_number": "DSK01001-01",
                "boxes_per_set": 1,  # 传统模式标识
                "boxes_per_small_box": 8,
                "small_boxes_per_large_box": 1,
                "total_boxes": 24
            },
            "expected": "DSK01001-01-DSK01008-01",  # 多套分一箱模式
            "note": "传统模式下应调用原有逻辑，这里测试新逻辑的表现"
        })
        
        # 8.2 参数兼容性
        self._run_test({
            "name": "兼容_参数兼容性",
            "type": "box_serial",
            "params": {
                "box_num": 15,
                "base_number": "OLD01001-01",
                "boxes_per_set": 1  # 应该仍能正常工作
            },
            "expected": "OLD01015-01"  # 每盒都是独立套
        })
    
    def _run_test(self, test_case: Dict[str, Any]):
        """运行单个测试用例"""
        test_start_time = datetime.now()
        test_result = {
            "name": test_case['name'],
            "type": test_case['type'],
            "start_time": test_start_time.strftime("%H:%M:%S"),
            "params": test_case['params'],
            "expected": test_case['expected'],
            "status": "unknown",
            "errors": [],
            "actual_result": None,
            "duration_ms": 0
        }
        
        try:
            print(f"\n🧪 {test_case['name']}")
            
            # 根据测试类型调用不同方法
            if test_case['type'] == 'box_serial':
                actual = self.processor.generate_set_based_box_serial(
                    test_case['params']['box_num'],
                    test_case['params']['base_number'],
                    test_case['params']['boxes_per_set']
                )
            elif test_case['type'] == 'small_box_range':
                actual = self.processor.generate_set_based_small_box_serial_range(
                    test_case['params']['small_box_num'],
                    test_case['params']['base_number'],
                    test_case['params']['boxes_per_set'],
                    test_case['params']['boxes_per_small_box'],
                    test_case['params'].get('total_boxes')
                )
            elif test_case['type'] == 'large_box_range':
                actual = self.processor.generate_set_based_large_box_serial_range(
                    test_case['params']['large_box_num'],
                    test_case['params']['base_number'],
                    test_case['params']['boxes_per_set'],
                    test_case['params']['boxes_per_small_box'],
                    test_case['params']['small_boxes_per_large_box'],
                    test_case['params'].get('total_boxes')
                )
            else:
                raise ValueError(f"未知的测试类型: {test_case['type']}")
            
            test_result["actual_result"] = actual
            expected = test_case['expected']
            
            # 验证结果
            errors = []
            if actual != expected:
                errors.append(f"结果不匹配: 期望'{expected}', 实际'{actual}'")
            
            # 格式验证
            if test_case.get('check_format'):
                format_errors = self._check_format_rules(actual, test_case.get('format_rules', []))
                errors.extend(format_errors)
            
            # 记录结果
            test_result["errors"] = errors
            
            if errors:
                print(f"❌ 失败: {'; '.join(errors)}")
                self.failed_tests += 1
                test_result["status"] = "failed"
            else:
                print(f"✅ 通过: {actual}")
                self.passed_tests += 1
                test_result["status"] = "passed"
                
        except Exception as e:
            print(f"❌ 异常: {str(e)}")
            self.failed_tests += 1
            test_result["status"] = "error"
            test_result["errors"] = [f"异常: {str(e)}"]
        
        # 记录耗时
        test_end_time = datetime.now()
        test_result["duration_ms"] = int((test_end_time - test_start_time).total_seconds() * 1000)
        
        self.test_results.append(test_result)
    
    def _check_format_rules(self, result: str, rules: List[str]) -> List[str]:
        """检查格式规则"""
        errors = []
        
        for rule in rules:
            if rule == "must_contain_dash" and "-" not in result:
                errors.append("格式错误: 结果必须包含横线'-'")
            elif rule == "no_tilde" and "~" in result:
                errors.append("格式错误: 结果不应包含波浪线'~'")
            elif rule == "always_range_format":
                if "-" not in result or result.count("-") < 3:  # 至少要有起始Serial-结束Serial的形式
                    errors.append("格式错误: 必须始终显示为范围格式")
            elif rule == "correct_main_number":
                # 验证主号是否正确递增
                pass  # 这里可以添加更复杂的主号验证逻辑
        
        return errors
    
    def _print_summary(self):
        """输出测试总结"""
        total_tests = self.passed_tests + self.failed_tests
        print("\n" + "=" * 80)
        print(f"🏁 Serial逻辑测试完成!")
        print(f"   总测试数: {total_tests}")
        print(f"   通过: {self.passed_tests}")
        print(f"   失败: {self.failed_tests}")
        print(f"   成功率: {round(self.passed_tests/total_tests*100, 1) if total_tests > 0 else 0}%")
        
        if self.failed_tests == 0:
            print(f"🎉 所有测试通过! Serial逻辑工作正常!")
        else:
            print(f"⚠️  有{self.failed_tests}个测试失败，请检查逻辑!")
    
    def _export_results(self):
        """导出测试结果"""
        end_time = datetime.now()
        total_duration = (end_time - self.start_time).total_seconds()
        
        timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
        json_filename = f"total-result/serial_test_results_{timestamp}.json"
        
        export_data = {
            "test_session": {
                "start_time": self.start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "end_time": end_time.strftime("%Y-%m-%d %H:%M:%S"),
                "duration_seconds": round(total_duration, 2),
                "total_tests": len(self.test_results),
                "passed_tests": self.passed_tests,
                "failed_tests": self.failed_tests,
                "success_rate": round(self.passed_tests / len(self.test_results) * 100, 1) if self.test_results else 0
            },
            "test_results": self.test_results
        }
        
        try:
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            print(f"\n📄 测试结果已导出: {json_filename}")
        except Exception as e:
            print(f"\n❌ 结果导出失败: {str(e)}")


def main():
    """主函数"""
    tester = SerialLogicTester()
    tester.run_all_tests()
    return 0


if __name__ == "__main__":
    exit(main())