    # ===================== 套盒模板方法 =====================

    def create_taohebox_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None
    ) -> Dict[str, str]:
        """创建套盒模板的多级标签PDF"""
        # 计算数量 - 三级结构：张→盒→小箱→大箱
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])  # 这个参数用于确定结束号
        small_boxes_per_large_box = int(params["小箱/大箱"])

        # 计算各级数量
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        total_small_boxes = math.ceil(total_boxes / boxes_per_small_box)
        total_large_boxes = math.ceil(total_small_boxes / small_boxes_per_large_box)

        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成套盒模板的盒标 - 第二个参数用于结束号
        selected_appearance = params["选择外观"]
        box_label_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒盒标+{selected_appearance}.pdf"

        self._create_taohebox_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        # 生成套盒模板小箱标
        small_box_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒小箱标.pdf"
        self._create_taohebox_small_box_label(
            data, params, str(small_box_path), total_small_boxes, excel_file_path
        )
        generated_files["小箱标"] = str(small_box_path)

        return generated_files

    def _create_taohebox_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建套盒模板的盒标 - 第二个参数用于结束号逻辑"""
        # 使用提供的Excel文件路径分析数据结构
        excel_path = excel_file_path
        print(f"🔍 正在分析套盒模板Excel文件: {excel_path}")
        
        # 先用pandas读取文件了解结构
        try:
            import pandas as pd
            df = pd.read_excel(excel_path, header=None)
            print(f"✅ Excel文件已加载: {df.shape[0]}行 x {df.shape[1]}列")
            
            # 打印前几行数据以了解结构
            print("📊 Excel文件内容预览:")
            for i in range(min(10, df.shape[0])):
                row_data = []
                for j in range(min(8, df.shape[1])):
                    cell_value = df.iloc[i, j]
                    if pd.isna(cell_value):
                        row_data.append("(空)")
                    else:
                        row_data.append(str(cell_value)[:20])  # 限制长度
                print(f"  行{i+1}: {' | '.join(row_data)}")
                
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {e}")
            # 使用默认值
            df = None
        
        # 使用关键字提取数据（如果有的话）
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        end_number = excel_data.get('结束号') or base_number  # 尝试提取结束号
        
        print(f"✅ 套盒盒标使用关键字提取:")
        print(f"   标签名称: '{theme_text}'")
        print(f"   开始号: '{base_number}'") 
        print(f"   结束号: '{end_number}'")
        
        # 第二个参数(盒/小箱)在套盒模板中用于确定结束号的逻辑
        boxes_per_ending_unit = int(params["盒/小箱"])  # 套盒模板的特殊用途
        group_size = int(params["小箱/大箱"])
        
        print(f"✅ 套盒模板参数:")
        print(f"   张/盒: {params['张/盒']}")
        print(f"   盒/小箱(用于结束号): {boxes_per_ending_unit}")
        print(f"   小箱/大箱: {group_size}")
        
        # 计算需要生成的盒标总数
        total_pieces = int(float(data["总张数"]))
        pieces_per_box = int(params["张/盒"])
        total_boxes = math.ceil(total_pieces / pieces_per_box)
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        c.setTitle("套盒模板盒标")
        
        width, height = self.page_size
        blank_height = height / 5
        top_text_y = height - 1.5 * blank_height
        serial_number_y = height - 3.5 * blank_height
        
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)
        
        # 生成套盒盒标 - 使用第二个参数的特殊逻辑
        for box_num in range(1, total_boxes + 1):
            if box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 套盒模板序列号生成逻辑 - 需要根据Excel文件进一步分析
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                main_digit_start = match.start()
                prefix_part = base_number[:main_digit_start]
                base_main_num = int(match.group(1))
                
                # 套盒模板的特殊计算逻辑（基于第二个参数）
                box_index = box_num - 1
                main_number_offset = box_index // group_size
                suffix_number = (box_index % group_size) + 1
                
                new_main_number = base_main_num + main_number_offset
                current_number = f"{prefix_part}{new_main_number:05d}-{suffix_number:02d}"
            else:
                current_number = f"DSK{box_num:05d}-01"
            
            print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
            
            # 渲染外观
            if style == "外观一":
                self._render_taohebox_appearance_one(c, width, theme_text, current_number, top_text_y, serial_number_y)
            else:
                self._render_taohebox_appearance_two(c, width, theme_text, current_number, top_text_y, serial_number_y)

        c.save()
        print(f"✅ 套盒模板盒标PDF已生成: {output_path}")

    def _create_taohebox_small_box_label(
        self,
        data: Dict[str, Any],
        params: Dict[str, Any],
        output_path: str,
        total_small_boxes: int,
        excel_file_path: str = None,
    ):
        """创建套盒模板的小箱标 - 借鉴分盒模板的计算逻辑"""
        # 获取Excel数据 - 使用关键字提取
        excel_path = excel_file_path or '/Users/trq/Desktop/project/Python-project/data-to-pdfprint/test.xlsx'
        
        excel_data = self._extract_excel_data_by_keywords(excel_path)
        theme_text = excel_data.get('标签名称') or 'Unknown Title'
        base_number = excel_data.get('开始号') or 'DEFAULT01001'
        remark_text = excel_data.get('客户编码') or 'Unknown Client'
        print(f"✅ 套盒小箱标使用关键字提取: 标签名称='{theme_text}', 开始号='{base_number}', 客户编码='{remark_text}'")
        
        # 获取用户输入的分组大小（从"小箱/大箱"参数获取） - 借鉴分盒模板逻辑
        try:
            group_size = int(params["小箱/大箱"])  # 用户的第三个参数，控制副号满几进一
            if group_size <= 0:
                group_size = 2
            print(f"✅ 套盒小箱标使用用户输入分组大小: {group_size} (小箱/大箱)")
        except (ValueError, KeyError) as e:
            print(f"⚠️ 获取小箱/大箱参数失败: {e}")
            group_size = 2  # 默认分组大小
        
        # 计算参数 - 借鉴分盒模板逻辑
        pieces_per_box = int(params["张/盒"])
        boxes_per_small_box = int(params["盒/小箱"])
        pieces_per_small_box = pieces_per_box * boxes_per_small_box
        
        # 创建PDF
        c = canvas.Canvas(output_path, pagesize=self.page_size)
        width, height = self.page_size

        # 设置PDF/X兼容模式和CMYK颜色
        c.setPageCompression(1)
        c.setTitle(f"套盒小箱标-1到{total_small_boxes}")
        c.setSubject("Taohebox Small Box Label")
        c.setCreator("Data-to-PDF Print")

        # 使用CMYK黑色
        cmyk_black = CMYKColor(0, 0, 0, 1)
        c.setFillColor(cmyk_black)

        # 生成指定范围的套盒小箱标
        for small_box_num in range(1, total_small_boxes + 1):
            if small_box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 计算套盒模板的序列号范围 - 借鉴分盒模板逻辑
            import re
            match = re.search(r'(\d+)', base_number)
            if match:
                # 获取第一个数字（主号）的起始位置
                digit_start = match.start()
                # 截取主号前面的所有字符作为前缀
                prefix_part = base_number[:digit_start]
                base_main_num = int(match.group(1))  # 主号
                
                # 套盒模板小箱标的逻辑：借鉴分盒模板计算方式
                # 每个小箱标包含boxes_per_small_box个盒标的序列号范围
                # 计算当前小箱标包含的盒标范围
                start_box_index = (small_box_num - 1) * boxes_per_small_box  # 起始盒标索引(0基数)
                end_box_index = start_box_index + boxes_per_small_box - 1    # 结束盒标索引(0基数)
                
                # 计算起始盒标的序列号
                start_main_offset = start_box_index // group_size
                start_suffix = (start_box_index % group_size) + 1
                start_main_number = base_main_num + start_main_offset
                start_serial = f"{prefix_part}{start_main_number:05d}-{start_suffix:02d}"
                
                # 计算结束盒标的序列号
                end_main_offset = end_box_index // group_size
                end_suffix = (end_box_index % group_size) + 1
                end_main_number = base_main_num + end_main_offset
                end_serial = f"{prefix_part}{end_main_number:05d}-{end_suffix:02d}"
                
                # 套盒小箱标显示序列号范围
                serial_range = f"{start_serial}-{end_serial}"
            else:
                serial_range = f"DSK{small_box_num:05d}-DSK{small_box_num:05d}"

            # 计算套盒小箱标的Carton No（主箱号-副箱号格式） - 借鉴分盒模板逻辑
            main_box_num = ((small_box_num - 1) // group_size) + 1  # 主箱号
            sub_box_num = ((small_box_num - 1) % group_size) + 1    # 副箱号
            carton_no = f"{main_box_num}-{sub_box_num}"

            # 绘制套盒小箱标表格
            self._draw_taohebox_small_box_table(c, width, height, theme_text, pieces_per_small_box, 
                                               serial_range, carton_no, remark_text)

        c.save()
        print(f"✅ 套盒模板小箱标PDF已生成: {output_path}")

    def _draw_taohebox_small_box_table(self, c, width, height, theme_text, pieces_per_small_box, 
                                      serial_range, carton_no, remark_text):
        """绘制套盒小箱标表格 - 借鉴分盒模板的表格绘制逻辑"""
        # 表格尺寸和位置
        table_width = width - 4 * mm
        table_height = height - 4 * mm
        table_x = 2 * mm
        table_y = 2 * mm
        
        # 高度分配：Quantity行占2/6，其他4行各占1/6
        base_row_height = table_height / 6
        quantity_row_height = base_row_height * 2  # Quantity行双倍高度
        
        # 列宽 (标签列:数据列 = 1:2)
        label_col_width = table_width / 3
        data_col_width = table_width * 2 / 3
        
        # 绘制表格边框
        c.setStrokeColor(CMYKColor(0, 0, 0, 1))
        c.setLineWidth(1)
        c.rect(table_x, table_y, table_width, table_height)
        
        # 计算各行的Y坐标
        row_positions = []
        current_y = table_y
        # 从底部开始：Remark, Carton No, Quantity(双倍), Theme, Item
        for height_val in [base_row_height, base_row_height, quantity_row_height, base_row_height, base_row_height]:
            row_positions.append(current_y)
            current_y += height_val
        
        # 绘制行分割线
        for y_pos in row_positions[1:]:  # 跳过底部边框
            c.line(table_x, y_pos, table_x + table_width, y_pos)
        
        # 绘制列分割线
        c.line(table_x + label_col_width, table_y, table_x + label_col_width, table_y + table_height)
        
        # 设置字体和颜色
        c.setFillColor(CMYKColor(0, 0, 0, 1))
        
        # 定义行数据 (从底部到顶部)
        rows_data = [
            ("Remark", remark_text),
            ("Carton No", carton_no),
            ("Quantity", f"{pieces_per_small_box}PCS\n{serial_range}"),
            ("Theme", theme_text),
            ("Item", "")  # Item通常为空
        ]
        
        # 绘制每行内容
        for i, (label, content) in enumerate(rows_data):
            row_y = row_positions[i]
            row_height = quantity_row_height if label == "Quantity" else base_row_height
            
            # 绘制标签列（左侧）
            label_font_size = 11 if label != "Quantity" else 10
            c.setFont("Microsoft-YaHei-Bold", label_font_size)
            label_text_y = row_y + row_height / 2 - label_font_size / 3
            c.drawCentredText(table_x + label_col_width / 2, label_text_y, label)
            
            # 绘制内容列（右侧）
            if content:
                if label == "Quantity":
                    # Quantity内容分两行显示
                    lines = content.split('\n')
                    if len(lines) == 2:
                        # 第一行：数量
                        c.setFont("Microsoft-YaHei-Bold", 11)
                        first_line_y = row_y + row_height * 0.7 - 11/3
                        c.drawCentredText(table_x + label_col_width + data_col_width / 2, first_line_y, lines[0])
                        
                        # 第二行：序列号范围
                        c.setFont("Microsoft-YaHei-Bold", 9)
                        second_line_y = row_y + row_height * 0.3 - 9/3
                        c.drawCentredText(table_x + label_col_width + data_col_width / 2, second_line_y, lines[1])
                else:
                    # 其他行的内容
                    content_font_size = 11
                    c.setFont("Microsoft-YaHei-Bold", content_font_size)
                    
                    # 处理文本换行
                    wrapped_text = self._wrap_text_to_fit(content, data_col_width - 4*mm, content_font_size)
                    
                    if len(wrapped_text) == 1:
                        # 单行文本居中
                        content_text_y = row_y + row_height / 2 - content_font_size / 3
                        c.drawCentredText(table_x + label_col_width + data_col_width / 2, content_text_y, wrapped_text[0])
                    else:
                        # 多行文本
                        line_spacing = content_font_size * 1.2
                        total_text_height = len(wrapped_text) * line_spacing
                        start_y = row_y + (row_height + total_text_height) / 2 - line_spacing
                        
                        for j, line in enumerate(wrapped_text):
                            line_y = start_y - j * line_spacing
                            if line_y >= row_y + 2:  # 确保文本在行内
                                c.drawCentredText(table_x + label_col_width + data_col_width / 2, line_y, line)