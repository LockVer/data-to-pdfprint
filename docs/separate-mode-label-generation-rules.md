# 分盒模式箱标生成规律

## 小箱标
- **quantity**: 每盒张数 (如：3500PCS)
- **serial**: 
  - 格式: `{serial_prefix}{box_num:05d}-{sub_box:02d}-{serial_prefix}{box_num:05d}-{sub_box:02d}`
  - 示例: `LGM01001-01-LGM01001-01`, `LGM01001-02-LGM01001-02`
  - 循环逻辑: 
    - 外层循环: 小箱数量 (total_small_cases)
    - 内层循环: 每小箱盒数 (boxes_per_small_case)
    - box_num = start_num + case_idx
    - sub_box = sub_box + 1
- **carton_no**:
  - 格式: `{case_idx + 1}-{sub_box + 1}`
  - 示例: `1-1`, `1-2`, `1-3`, `1-4`, `2-1`, `2-2`...
  - 第一个数字: 当前小箱编号 (从1开始)
  - 第二个数字: 小箱内盒子编号 (从1开始)

## 大箱标
- **quantity**:
  - 计算方式: (该大箱包含的小箱数) × (每小箱盒数) × (每盒张数)
  - 如: 2个小箱 × 4盒/小箱 × 3500张/盒 = 28000PCS
- **serial**:
  - 格式: `{start_serial}-{end_serial}`
  - start_serial: `{serial_prefix}{start_box_num:05d}-01`
  - end_serial: `{serial_prefix}{end_box_num:05d}-{boxes_per_small_case:02d}`
  - 示例: `LGM01029-01-LGM01030-04` (跨越多个小箱的范围)
- **carton_no**:
  - 格式: `{start_case + 1}-{end_case + 1}`
  - 示例: `29-30`, `31-32`
  - 表示该大箱包含的小箱编号范围

## 关键规律
- **小箱标**: 双层循环生成 (小箱 → 盒子)
  - 每个小箱内的每个盒子都有一页小箱标
  - Serial开始号和结束号相同 (单个盒子)
  - Carton_no使用双数字格式表示层级关系

- **大箱标**: 按大箱生成
  - 每个大箱一页大箱标
  - Serial跨越该大箱内所有小箱的盒子范围
  - Carton_no表示小箱编号范围

- **页数计算**:
  - 小箱标页数: total_small_cases × boxes_per_small_case
  - 大箱标页数: total_large_cases
  - 总页数: 小箱标页数 + 大箱标页数

## 语义化描述

### 小箱标
- **quantity (双层循环)**
  - 每盒张数
- **serial**
  - 父级编号
    - 开始号和结束号相同
      - 小箱数量-循环递增
  - 子级编号（开始号和结束号相同）
    - 开始号
      - 每小箱盒数-循环递增
    - 结束号  
      - 每小箱盒数-循环递增
- **carton_no (双层循环)**
  - 父级编码
    - 小箱数量-循环递增
  - 子级编码
    - 每小箱盒数-循环递增

### 大箱标
- **quantity**
  - 每盒张数 × 小箱中盒数 × 大箱中小箱数
- **serial**
  - 父级编号
    - 开始号和结束号不同（跨范围）
      - 大箱数量-循环递增
  - 子级编号
    - 开始号
      - 01（固定）
    - 结束号
      - 每小箱盒数（最大值）
- **carton_no**
  - 小箱编号范围表示
    - 起始小箱编号-结束小箱编号