#!/usr/bin/env python3
"""
Carton Number Logic Test Suite
测试分盒/套盒模板的Carton No计算逻辑

使用方法:
python test_carton_logic.py
"""

import math
import sys
import os
import json
from datetime import datetime
from typing import Dict, Any

# 添加项目路径，方便导入模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.pdf.split_box.data_processor import SplitBoxDataProcessor


class CartonLogicTester:
    """Carton逻辑测试器"""
    
    def __init__(self):
        self.processor = SplitBoxDataProcessor()
        self.passed_tests = 0
        self.failed_tests = 0
        self.test_results = []  # 存储所有测试结果
        self.start_time = datetime.now()
        
    def run_all_tests(self):
        """运行所有测试用例"""
        print("🧪 开始运行Carton Number逻辑测试")
        print("=" * 60)
        
        # 二级模式测试
        print("\n📦 二级模式测试 (无小箱)")
        self._test_two_level_mode()
        
        # 三级模式测试
        print("\n📦 三级模式测试 (有小箱)")
        self._test_three_level_mode()
        
        # 边界情况测试
        print("\n🔍 边界情况测试")
        self._test_edge_cases()
        
        # 性能测试
        print("\n⚡ 性能和压力测试")
        self._test_performance_cases()
        
        # 输出测试结果
        self._print_summary()
        
        # 导出测试结果到文档
        self._export_results()
    
    def _test_two_level_mode(self):
        """测试二级模式的所有情况"""
        
        # 1.1 一套分多个大箱（每套大箱数 > 1）
        
        # 测试用例1.1：基础场景
        self._run_test({
            "name": "二级_一套分2个大箱",
            "params": {
                "张/盒": 730, "盒/套": 15, "盒/小箱": 8, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "109500"},  # 150盒 -> 10套
            "expected": {
                "total_boxes": 150,           # ceil(109500/730)
                "total_sets": 10,             # ceil(150/15)  
                "large_boxes_per_set": 2,     # ceil(15/8) = 2
                "total_large_boxes": 20,      # 10套 × 2箱/套
                "carton_no_sample": ["1-1", "1-2", "2-1", "2-2", "10-1", "10-2"]
            }
        })
        
        # 测试用例1.2：不同比例
        self._run_test({
            "name": "二级_一套分3个大箱",
            "params": {
                "张/盒": 1000, "盒/套": 10, "盒/小箱": 4, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "20000"},    # 20盒 -> 2套
            "expected": {
                "total_boxes": 20,            # ceil(20000/1000)
                "total_sets": 2,              # ceil(20/10)
                "large_boxes_per_set": 3,     # ceil(10/4) = 3
                "total_large_boxes": 6,       # 2套 × 3箱/套
                "carton_no_sample": ["1-1", "1-2", "1-3", "2-1", "2-2", "2-3"]
            }
        })
        
        # 1.2 一套分一个大箱（每套大箱数 = 1）
        self._run_test({
            "name": "二级_一套分1个大箱",
            "params": {
                "张/盒": 1000, "盒/套": 5, "盒/小箱": 5, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "25000"},    # 25盒 -> 5套
            "expected": {
                "total_boxes": 25,
                "total_sets": 5,
                "large_boxes_per_set": 1,     # ceil(5/5) = 1
                "total_large_boxes": 5,       # 5套 × 1箱/套
                "carton_no_sample": ["1", "2", "3", "4", "5"]
            }
        })
        
        # 1.3 多套分一个大箱（每套大箱数 < 1）
        
        # 测试用例1.3a：2套分1个大箱
        self._run_test({
            "name": "二级_2套分1个大箱",
            "params": {
                "张/盒": 500, "盒/套": 3, "盒/小箱": 8, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "12000"},    # 24盒 -> 8套
            "expected": {
                "total_boxes": 24,            # ceil(12000/500)
                "total_sets": 8,              # ceil(24/3)
                "large_boxes_per_set": 0.375, # 3/8 = 0.375，即8/3 ≈ 2.67套/箱
                "sets_per_large_box": 2,      # floor(8/3) = 2套/箱  
                "total_large_boxes": 4,       # ceil(8套 ÷ 2套/箱) = 4箱
                "carton_no_sample": ["1-2", "3-4", "5-6", "7-8"]
            }
        })
        
        # 测试用例1.3b：5套分1个大箱
        self._run_test({
            "name": "二级_5套分1个大箱",
            "params": {
                "张/盒": 200, "盒/套": 1, "盒/小箱": 6, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "2000"},     # 10盒 -> 10套
            "expected": {
                "total_boxes": 10,            # ceil(2000/200)
                "total_sets": 10,             # ceil(10/1)
                "large_boxes_per_set": 0.167, # 1/6 = 0.167，即6套/箱
                "sets_per_large_box": 6,      # 6套/箱（但最后一箱只有4套）
                "total_large_boxes": 2,       # ceil(10套 ÷ 6套/箱) = 2箱
                "carton_no_sample": ["1-6", "7-10"]
            }
        })
    
    def _test_three_level_mode(self):
        """测试三级模式的所有情况"""
        
        # 2.1 小箱标的三种情况
        
        # 测试用例2.1a：一套分多个小箱
        self._run_test({
            "name": "三级_一套分多个小箱",
            "params": {
                "张/盒": 1000, "盒/套": 6, "盒/小箱": 2, "小箱/大箱": 2, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "12000"},    # 12盒 -> 2套
            "expected": {
                "total_boxes": 12,
                "total_sets": 2,              # ceil(12/6)
                "small_boxes_per_set": 3,     # ceil(6/2) = 3
                "large_boxes_per_set": 2,     # ceil(3/2) = 2  
                "total_small_boxes": 6,       # 2套 × 3小箱/套
                "total_large_boxes": 4,       # 2套 × 2大箱/套
                "small_carton_no_sample": ["1-1", "1-2", "1-3", "2-1", "2-2", "2-3"],
                "large_carton_no_sample": ["1-1", "1-2", "2-1", "2-2"]
            }
        })
        
        # 测试用例2.1b：一套分一个小箱
        self._run_test({
            "name": "三级_一套分1个小箱",
            "params": {
                "张/盒": 1000, "盒/套": 4, "盒/小箱": 4, "小箱/大箱": 2, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "8000"},     # 8盒 -> 2套
            "expected": {
                "total_boxes": 8,
                "total_sets": 2,              # ceil(8/4)
                "small_boxes_per_set": 1,     # ceil(4/4) = 1
                "large_boxes_per_set": 1,     # ceil(1/2) = 1
                "total_small_boxes": 2,       # 2套 × 1小箱/套
                "total_large_boxes": 2,       # 2套 × 1大箱/套
                "small_carton_no_sample": ["01", "02"],  # 单级编号
                "large_carton_no_sample": ["1", "2"]
            }
        })
        
        # 测试用例2.1c：没有小箱标（每套小箱数 < 1）
        self._run_test({
            "name": "三级_没有小箱标",
            "params": {
                "张/盒": 500, "盒/套": 2, "盒/小箱": 8, "小箱/大箱": 1, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "4000"},     # 8盒 -> 4套
            "expected": {
                "total_boxes": 8,
                "total_sets": 4,              # ceil(8/2)
                "small_boxes_per_set": 0.25,  # 2/8 = 0.25 < 1
                "should_generate_small_box": False,  # 不生成小箱标
                "large_boxes_per_set": 1,     # ceil(0.25/1) = 1
                "total_large_boxes": 4,       # 4套 × 1大箱/套
                "large_carton_no_sample": ["1", "2", "3", "4"]
            }
        })
        
        # 2.2 大箱标的三种情况
        
        # 测试用例2.2a：一套分多个大箱
        self._run_test({
            "name": "三级_一套分多个大箱",
            "params": {
                "张/盒": 1000, "盒/套": 8, "盒/小箱": 2, "小箱/大箱": 2, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "16000"},    # 16盒 -> 2套
            "expected": {
                "total_boxes": 16,
                "total_sets": 2,              # ceil(16/8)
                "small_boxes_per_set": 4,     # ceil(8/2) = 4
                "large_boxes_per_set": 2,     # ceil(4/2) = 2
                "total_small_boxes": 8,       # 2套 × 4小箱/套
                "total_large_boxes": 4,       # 2套 × 2大箱/套
                "large_carton_no_sample": ["1-1", "1-2", "2-1", "2-2"]
            }
        })
        
        # 测试用例2.2b：一套分一个大箱
        self._run_test({
            "name": "三级_一套分1个大箱",
            "params": {
                "张/盒": 1000, "盒/套": 6, "盒/小箱": 2, "小箱/大箱": 3, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "18000"},    # 18盒 -> 3套
            "expected": {
                "total_boxes": 18,
                "total_sets": 3,              # ceil(18/6)
                "small_boxes_per_set": 3,     # ceil(6/2) = 3
                "large_boxes_per_set": 1,     # ceil(3/3) = 1
                "total_small_boxes": 9,       # 3套 × 3小箱/套
                "total_large_boxes": 3,       # 3套 × 1大箱/套
                "large_carton_no_sample": ["1", "2", "3"]
            }
        })
        
        # 测试用例2.2c：多套分一个大箱
        self._run_test({
            "name": "三级_多套分1个大箱",
            "params": {
                "张/盒": 200, "盒/套": 2, "盒/小箱": 1, "小箱/大箱": 6, "是否有小箱": True,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "2400"},     # 12盒 -> 6套
            "expected": {
                "total_boxes": 12,
                "total_sets": 6,              # ceil(12/2)
                "small_boxes_per_set": 2,     # ceil(2/1) = 2
                "large_boxes_per_set": 0.33,  # 2/6 = 0.33，即3套/大箱
                "sets_per_large_box": 3,      # 3套/大箱
                "total_small_boxes": 12,      # 6套 × 2小箱/套
                "total_large_boxes": 2,       # ceil(6套 ÷ 3套/箱) = 2箱
                "large_carton_no_sample": ["1-3", "4-6"]
            }
        })
    
    def _test_edge_cases(self):
        """测试边界情况"""
        
        # 测试用例3.1：最小值
        self._run_test({
            "name": "边界_最小值",
            "params": {
                "张/盒": 1, "盒/套": 1, "盒/小箱": 1, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "1"},
            "expected": {
                "total_boxes": 1,
                "total_sets": 1,
                "large_boxes_per_set": 1,
                "total_large_boxes": 1,
                "carton_no_sample": ["1"]
            }
        })
        
        # 测试用例3.2：刚好整除
        self._run_test({
            "name": "边界_刚好整除",
            "params": {
                "张/盒": 1000, "盒/套": 10, "盒/小箱": 5, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "50000"},    # 50盒 -> 5套，每套2箱，总共10箱
            "expected": {
                "total_boxes": 50,
                "total_sets": 5,
                "large_boxes_per_set": 2,     # ceil(10/5)
                "total_large_boxes": 10,      # 5套 × 2箱/套
                "carton_no_sample": ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2", "4-1", "4-2", "5-1", "5-2"]
            }
        })
        
        # 测试用例3.3：有余数
        self._run_test({
            "name": "边界_有余数",
            "params": {
                "张/盒": 1000, "盒/套": 7, "盒/小箱": 3, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "22000"},    # 22盒 -> 4套（余数为1盒）
            "expected": {
                "total_boxes": 22,            # ceil(22000/1000)
                "total_sets": 4,              # ceil(22/7) = 4套（余数为1盒）
                "large_boxes_per_set": 3,     # ceil(7/3) = 3
                "total_large_boxes": 12,      # 4套 × 3箱/套
                "carton_no_sample": ["1-1", "1-2", "1-3", "2-1", "2-2", "2-3", "3-1", "3-2", "3-3", "4-1", "4-2", "4-3"]
            }
        })
    
    def _test_performance_cases(self):
        """测试性能和压力情况"""
        
        # 测试用例4.1：大数量
        self._run_test({
            "name": "性能_大数量",
            "params": {
                "张/盒": 10000, "盒/套": 100, "盒/小箱": 50, "是否有小箱": False,
                "标签模版": "无纸卡备注", "中文名称": "测试", "序列号字体大小": 10, "是否有盒标": False
            },
            "data": {"总张数": "10000000"},  # 1000盒 -> 10套
            "expected": {
                "total_boxes": 1000,          # ceil(10000000/10000)
                "total_sets": 10,             # ceil(1000/100)
                "large_boxes_per_set": 2,     # ceil(100/50) = 2
                "total_large_boxes": 20,      # 10套 × 2箱/套
                "carton_no_sample": ["1-1", "1-2", "2-1", "2-2", "10-1", "10-2"]
            }
        })
    
    def _run_test(self, test_case: Dict[str, Any]):
        """运行单个测试用例"""
        test_start_time = datetime.now()
        test_result = {
            "name": test_case['name'],
            "start_time": test_start_time.strftime("%H:%M:%S"),
            "params": test_case['params'],
            "data": test_case['data'],
            "expected": test_case['expected'],
            "status": "unknown",
            "errors": [],
            "actual_result": {},
            "duration_ms": 0
        }
        
        try:
            print(f"\n🧪 测试: {test_case['name']}")
            
            # 计算实际结果
            actual = self._calculate_carton_logic(test_case['params'], test_case['data'])
            test_result["actual_result"] = actual
            expected = test_case['expected']
            
            # 验证计算结果
            errors = []
            
            # 验证基础数量计算
            if 'total_boxes' in expected:
                if actual['total_boxes'] != expected['total_boxes']:
                    errors.append(f"总盒数不匹配: 期望{expected['total_boxes']}, 实际{actual['total_boxes']}")
            
            if 'total_sets' in expected:
                if actual['total_sets'] != expected['total_sets']:
                    errors.append(f"总套数不匹配: 期望{expected['total_sets']}, 实际{actual['total_sets']}")
            
            if 'total_large_boxes' in expected:
                if actual['total_large_boxes'] != expected['total_large_boxes']:
                    errors.append(f"总大箱数不匹配: 期望{expected['total_large_boxes']}, 实际{actual['total_large_boxes']}")
            
            # 验证Carton No样本
            if 'carton_no_sample' in expected:
                for i, expected_carton in enumerate(expected['carton_no_sample'], 1):
                    if i <= len(actual['large_carton_nos']):
                        actual_carton = actual['large_carton_nos'][i-1]
                        if actual_carton != expected_carton:
                            errors.append(f"大箱Carton No[{i}]不匹配: 期望'{expected_carton}', 实际'{actual_carton}'")
            
            # 验证小箱Carton No (如果有)
            if 'small_carton_no_sample' in expected:
                for i, expected_carton in enumerate(expected['small_carton_no_sample'], 1):
                    if i <= len(actual.get('small_carton_nos', [])):
                        actual_carton = actual['small_carton_nos'][i-1]
                        if actual_carton != expected_carton:
                            errors.append(f"小箱Carton No[{i}]不匹配: 期望'{expected_carton}', 实际'{actual_carton}'")
            
            # 验证大箱Carton No样本 (如果有)
            if 'large_carton_no_sample' in expected:
                for i, expected_carton in enumerate(expected['large_carton_no_sample'], 1):
                    if i <= len(actual['large_carton_nos']):
                        actual_carton = actual['large_carton_nos'][i-1]
                        if actual_carton != expected_carton:
                            errors.append(f"大箱Carton No[{i}]不匹配: 期望'{expected_carton}', 实际'{actual_carton}'")
            
            # 验证小箱数量相关字段
            if 'small_boxes_per_set' in expected:
                if abs(actual.get('small_boxes_per_set', 0) - expected['small_boxes_per_set']) > 0.01:
                    errors.append(f"每套小箱数不匹配: 期望{expected['small_boxes_per_set']}, 实际{actual.get('small_boxes_per_set', 0)}")
            
            if 'large_boxes_per_set' in expected:
                if abs(actual.get('large_boxes_per_set', 0) - expected['large_boxes_per_set']) > 0.01:
                    errors.append(f"每套大箱数不匹配: 期望{expected['large_boxes_per_set']}, 实际{actual.get('large_boxes_per_set', 0)}")
            
            # 验证特殊情况标识
            if 'should_generate_small_box' in expected:
                # 检查是否应该生成小箱标
                has_small_carton = len(actual.get('small_carton_nos', [])) > 0
                if expected['should_generate_small_box'] != has_small_carton:
                    errors.append(f"小箱标生成状态不匹配: 期望{expected['should_generate_small_box']}, 实际{has_small_carton}")
            
            # 记录错误和状态
            test_result["errors"] = errors
            
            # 输出结果
            if errors:
                print(f"❌ 测试失败:")
                for error in errors:
                    print(f"   {error}")
                self.failed_tests += 1
                test_result["status"] = "failed"
            else:
                print(f"✅ 测试通过")
                self.passed_tests += 1
                test_result["status"] = "passed"
            
            # 输出实际计算详情
            print(f"   📊 计算详情:")
            print(f"      总盒数: {actual['total_boxes']}, 总套数: {actual['total_sets']}")
            print(f"      总大箱数: {actual['total_large_boxes']}")
            if actual.get('total_small_boxes'):
                print(f"      总小箱数: {actual['total_small_boxes']}")
            print(f"      大箱Carton No前5个: {actual['large_carton_nos'][:5]}")
            if actual.get('small_carton_nos'):
                print(f"      小箱Carton No前5个: {actual['small_carton_nos'][:5]}")
            
        except Exception as e:
            print(f"❌ 测试异常: {str(e)}")
            self.failed_tests += 1
            test_result["status"] = "error"
            test_result["errors"] = [f"异常: {str(e)}"]
        
        # 记录测试时长
        test_end_time = datetime.now()
        test_result["duration_ms"] = int((test_end_time - test_start_time).total_seconds() * 1000)
        
        # 保存测试结果
        self.test_results.append(test_result)
    
    def _calculate_carton_logic(self, params: Dict[str, Any], data: Dict[str, Any]) -> Dict[str, Any]:
        """计算Carton逻辑，模拟实际的计算过程"""
        
        # 基础计算
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        boxes_per_set = int(params.get("盒/套", 1))
        total_sets = math.ceil(total_boxes / boxes_per_set)
        
        has_small_box = params.get("是否有小箱", True)
        
        result = {
            "total_pieces": total_pieces,
            "total_boxes": total_boxes,
            "total_sets": total_sets
        }
        
        if has_small_box:
            # 三级模式
            boxes_per_small_box = int(params["盒/小箱"])
            small_boxes_per_large_box = int(params["小箱/大箱"])
            
            # 基于套数的正确计算
            small_boxes_per_set = math.ceil(boxes_per_set / boxes_per_small_box)
            large_boxes_per_set = math.ceil(small_boxes_per_set / small_boxes_per_large_box)
            
            total_small_boxes = total_sets * small_boxes_per_set
            total_large_boxes = total_sets * large_boxes_per_set
            
            result.update({
                "total_small_boxes": total_small_boxes,
                "total_large_boxes": total_large_boxes,
                "small_boxes_per_set": small_boxes_per_set,
                "large_boxes_per_set": large_boxes_per_set
            })
            
            # 计算小箱Carton No (只有当small_boxes_per_set >= 1时才生成)
            small_carton_nos = []
            if small_boxes_per_set >= 1:
                for i in range(1, total_small_boxes + 1):
                    carton_no = self.processor.calculate_carton_number_for_small_box(i, boxes_per_set, boxes_per_small_box)
                    small_carton_nos.append(carton_no)
            result["small_carton_nos"] = small_carton_nos
            
            # 计算大箱Carton No
            boxes_per_large_box = boxes_per_small_box * small_boxes_per_large_box
            large_carton_nos = []
            for i in range(1, total_large_boxes + 1):
                carton_no = self.processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
                large_carton_nos.append(carton_no)
            result["large_carton_nos"] = large_carton_nos
            
        else:
            # 二级模式
            boxes_per_large_box = int(params["盒/小箱"])  # 在二级模式下这是盒/大箱
            
            # 基于套数的正确计算
            large_boxes_per_set = math.ceil(boxes_per_set / boxes_per_large_box)
            total_large_boxes = total_sets * large_boxes_per_set
            
            result.update({
                "total_large_boxes": total_large_boxes,
                "large_boxes_per_set": large_boxes_per_set
            })
            
            # 计算大箱Carton No
            large_carton_nos = []
            for i in range(1, total_large_boxes + 1):
                carton_no = self.processor.calculate_carton_range_for_large_box(i, boxes_per_set, boxes_per_large_box, total_sets)
                large_carton_nos.append(carton_no)
            result["large_carton_nos"] = large_carton_nos
        
        return result
    
    def _print_summary(self):
        """输出测试总结"""
        total_tests = self.passed_tests + self.failed_tests
        print("\n" + "=" * 60)
        print(f"🏁 测试完成!")
        print(f"   总测试数: {total_tests}")
        print(f"   通过: {self.passed_tests}")
        print(f"   失败: {self.failed_tests}")
        
        if self.failed_tests == 0:
            print(f"🎉 所有测试通过! Carton Number逻辑正确!")
        else:
            print(f"⚠️  有{self.failed_tests}个测试失败，请检查逻辑!")
    
    def _export_results(self):
        """导出测试结果到文档"""
        end_time = datetime.now()
        total_duration = (end_time - self.start_time).total_seconds()
        
        # 生成文件名
        timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
        json_filename = f"test_results_{timestamp}.json"
        md_filename = f"test_report_{timestamp}.md"
        
        # 准备导出数据
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
        
        # 导出JSON格式 (详细数据)
        try:
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            print(f"\n📄 详细测试数据已导出: {json_filename}")
        except Exception as e:
            print(f"\n❌ JSON导出失败: {str(e)}")
        
        # 导出Markdown格式 (可读性报告)
        try:
            self._export_markdown_report(md_filename, export_data)
            print(f"📄 测试报告已导出: {md_filename}")
        except Exception as e:
            print(f"❌ Markdown导出失败: {str(e)}")
    
    def _export_markdown_report(self, filename: str, data: Dict[str, Any]):
        """导出Markdown格式的测试报告"""
        session = data["test_session"]
        
        with open(filename, 'w', encoding='utf-8') as f:
            # 报告标题
            f.write(f"# Carton Number Logic 测试报告\n\n")
            f.write(f"**生成时间**: {session['end_time']}\n\n")
            
            # 测试概览
            f.write(f"## 📊 测试概览\n\n")
            f.write(f"| 项目 | 值 |\n")
            f.write(f"|------|----|\n")
            f.write(f"| 测试开始时间 | {session['start_time']} |\n")
            f.write(f"| 测试结束时间 | {session['end_time']} |\n")
            f.write(f"| 总耗时 | {session['duration_seconds']}秒 |\n")
            f.write(f"| 总测试数 | {session['total_tests']} |\n")
            f.write(f"| 通过测试 | {session['passed_tests']} |\n")
            f.write(f"| 失败测试 | {session['failed_tests']} |\n")
            f.write(f"| 成功率 | {session['success_rate']}% |\n\n")
            
            # 测试结果状态
            if session['failed_tests'] == 0:
                f.write(f"🎉 **所有测试通过！Carton Number逻辑正确！**\n\n")
            else:
                f.write(f"⚠️ **有{session['failed_tests']}个测试失败，需要检查逻辑**\n\n")
            
            # 详细测试结果
            f.write(f"## 📋 详细测试结果\n\n")
            
            for i, result in enumerate(data["test_results"], 1):
                status_icon = "✅" if result["status"] == "passed" else "❌"
                f.write(f"### {i}. {status_icon} {result['name']}\n\n")
                
                # 基本信息
                f.write(f"**状态**: {result['status']}\n")
                f.write(f"**耗时**: {result['duration_ms']}ms\n")
                f.write(f"**开始时间**: {result['start_time']}\n\n")
                
                # 测试参数
                f.write(f"**测试参数**:\n")
                params = result['params']
                for key, value in params.items():
                    f.write(f"- {key}: {value}\n")
                f.write(f"- 总张数: {result['data']['总张数']}\n\n")
                
                # 计算结果
                actual = result.get('actual_result', {})
                if actual:
                    f.write(f"**计算结果**:\n")
                    f.write(f"- 总盒数: {actual.get('total_boxes', 'N/A')}\n")
                    f.write(f"- 总套数: {actual.get('total_sets', 'N/A')}\n")
                    if 'total_small_boxes' in actual:
                        f.write(f"- 总小箱数: {actual['total_small_boxes']}\n")
                    f.write(f"- 总大箱数: {actual.get('total_large_boxes', 'N/A')}\n")
                    
                    # Carton No样本
                    if 'large_carton_nos' in actual:
                        large_sample = actual['large_carton_nos'][:10]
                        f.write(f"- 大箱Carton No前10个: {large_sample}\n")
                    
                    if 'small_carton_nos' in actual:
                        small_sample = actual['small_carton_nos'][:10]
                        f.write(f"- 小箱Carton No前10个: {small_sample}\n")
                    
                    f.write(f"\n")
                
                # 错误信息
                if result['errors']:
                    f.write(f"**错误信息**:\n")
                    for error in result['errors']:
                        f.write(f"- {error}\n")
                    f.write(f"\n")
                
                f.write(f"---\n\n")
            
            # 总结
            f.write(f"## 🏁 测试总结\n\n")
            f.write(f"本次测试覆盖了Carton Number计算逻辑的所有主要场景：\n\n")
            f.write(f"- **二级模式** (无小箱): 一套分多个/一个/多套分一个大箱\n")
            f.write(f"- **三级模式** (有小箱): 小箱标和大箱标的所有Carton No模式\n")
            f.write(f"- **边界情况**: 最小值、整除、余数等特殊情况\n\n")
            
            if session['failed_tests'] == 0:
                f.write(f"所有测试场景均通过验证，Carton Number计算逻辑正确无误。\n")
            else:
                f.write(f"发现{session['failed_tests']}个测试失败，建议检查相关计算逻辑。\n")
            
            f.write(f"\n**测试工具**: Carton Logic Test Suite v1.0\n")
            f.write(f"**生成时间**: {session['end_time']}\n")


def main():
    """主函数"""
    # 临时禁用data_processor中的调试输出，保持测试输出清晰
    import sys
    from io import StringIO
    
    # 可选：如果想看详细的计算过程，删除下面三行
    old_stdout = sys.stdout
    sys.stdout = StringIO()  # 重定向输出
    
    try:
        tester = CartonLogicTester()
        
        # 恢复输出
        sys.stdout = old_stdout
        
        tester.run_all_tests()
    except Exception as e:
        # 恢复输出
        sys.stdout = old_stdout
        print(f"测试运行异常: {str(e)}")
        return 1
    
    return 0


if __name__ == "__main__":
    exit(main())