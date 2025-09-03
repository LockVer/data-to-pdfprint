    # ===================== 套盒模板方法 =====================

    def create_taohebox_multi_level_pdfs(
        self, data: Dict[str, Any], params: Dict[str, Any], output_dir: str, excel_file_path: str = None
    ) -> Dict[str, str]:
        """创建套盒模板的多级标签PDF"""
        # 创建输出目录
        clean_theme = data['主题'].replace('\n', ' ').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('*', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_').replace('!', '_')
        folder_name = f"{data['客户编码']}+{clean_theme}+标签"
        full_output_dir = Path(output_dir) / folder_name
        full_output_dir.mkdir(parents=True, exist_ok=True)

        generated_files = {}

        # 生成套盒模板的盒标 - 第二个参数用于结束号逻辑
        selected_appearance = params["选择外观"]
        box_label_path = full_output_dir / f"{data['客户编码']}+{clean_theme}+套盒盒标+{selected_appearance}.pdf"

        self._create_taohebox_box_label(data, params, str(box_label_path), selected_appearance, excel_file_path)
        generated_files["盒标"] = str(box_label_path)

        return generated_files

    def _create_taohebox_box_label(
        self, data: Dict[str, Any], params: Dict[str, Any], output_path: str, style: str, excel_file_path: str = None
    ):
        """创建套盒模板的盒标 - 基于Excel文件的开始号和结束号"""
        # 分析Excel文件获取套盒特有的数据
        excel_path = excel_file_path
        print(f"🔍 正在分析套盒模板Excel文件: {excel_path}")
        
        try:
            import pandas as pd
            df = pd.read_excel(excel_path, header=None)
            print(f"✅ Excel文件已加载: {df.shape[0]}行 x {df.shape[1]}列")
            
            # 根据分析的结果提取数据
            # 标签名称：第10行第9列 (索引9,8)
            theme_text = df.iloc[9, 8] if pd.notna(df.iloc[9, 8]) else 'Unknown Title'
            
            # 开始号：第10行第2列 (索引9,1) 
            base_number = df.iloc[9, 1] if pd.notna(df.iloc[9, 1]) else 'DEFAULT01001'
            
            # 结束号：第10行第3列 (索引9,2)
            end_number = df.iloc[9, 2] if pd.notna(df.iloc[9, 2]) else base_number
            
            # 主题：第4行第2列 (索引3,1)
            full_theme = df.iloc[3, 1] if pd.notna(df.iloc[3, 1]) else 'Unknown Theme'
            
            print(f"✅ 套盒模板数据提取:")
            print(f"   标签名称: '{theme_text}'")
            print(f"   开始号: '{base_number}'")
            print(f"   结束号: '{end_number}'")
            print(f"   完整主题: '{full_theme}'")
            
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {e}")
            # 回退到关键字提取
            excel_data = self._extract_excel_data_by_keywords(excel_path)
            theme_text = excel_data.get('标签名称') or 'Unknown Title'
            base_number = excel_data.get('开始号') or 'DEFAULT01001'
            end_number = excel_data.get('结束号') or base_number
        
        # 套盒模板参数分析
        pieces_per_box = int(params["张/盒"])
        boxes_per_ending_unit = int(params["盒/小箱"])  # 在套盒模板中，这个参数用于结束号的范围计算
        group_size = int(params["小箱/大箱"])
        
        print(f"✅ 套盒模板参数:")
        print(f"   张/盒: {pieces_per_box}")
        print(f"   盒/小箱(结束号范围): {boxes_per_ending_unit}")
        print(f"   小箱/大箱(分组大小): {group_size}")
        
        # 解析开始号和结束号的格式
        import re
        start_match = re.search(r'(.+?)(\d+)-(\d+)', base_number)
        end_match = re.search(r'(.+?)(\d+)-(\d+)', end_number)
        
        if start_match and end_match:
            start_prefix = start_match.group(1)
            start_main = int(start_match.group(2))
            start_suffix = int(start_match.group(3))
            
            end_prefix = end_match.group(1)
            end_main = int(end_match.group(2))
            end_suffix = int(end_match.group(3))
            
            print(f"✅ 解析序列号格式:")
            print(f"   开始: {start_prefix}{start_main:05d}-{start_suffix:02d}")
            print(f"   结束: {end_prefix}{end_main:05d}-{end_suffix:02d}")
            
        else:
            print("⚠️ 无法解析序列号格式，使用默认逻辑")
            start_prefix = "JAW"
            start_main = 1001
            start_suffix = 1
            end_suffix = boxes_per_ending_unit
        
        # 计算需要生成的盒标数量
        total_pieces = int(float(data["总张数"]))
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
        
        # 生成套盒盒标 - 基于开始号到结束号的范围
        print(f"📝 开始生成套盒盒标，预计生成 {total_boxes} 个标签")
        
        for box_num in range(1, total_boxes + 1):
            if box_num > 1:
                c.showPage()
                c.setFillColor(cmyk_black)

            # 套盒模板序列号生成逻辑 - 基于开始号和结束号范围
            box_index = box_num - 1
            
            # 计算当前盒的序列号在范围内的位置
            main_offset = box_index // boxes_per_ending_unit
            suffix_in_range = (box_index % boxes_per_ending_unit) + start_suffix
            
            current_main = start_main + main_offset
            current_number = f"{start_prefix}{current_main:05d}-{suffix_in_range:02d}"
            
            print(f"📝 生成套盒盒标 #{box_num}: {current_number}")
            
            # 渲染外观
            if style == "外观一":
                self._render_taohebox_appearance_one(c, width, theme_text, current_number, top_text_y, serial_number_y)
            else:
                self._render_taohebox_appearance_two(c, width, theme_text, current_number, top_text_y, serial_number_y)

        c.save()
        print(f"✅ 套盒模板盒标PDF已生成: {output_path}")

    def _render_taohebox_appearance_one(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观一渲染"""
        clean_top_text = self._clean_text_for_font(top_text)
        self._set_best_font(c, 14, bold=True)
        
        # 绘制Game title和序列号 - 加粗效果
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], top_text_y + offset[1], clean_top_text)
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)

    def _render_taohebox_appearance_two(self, c, width, top_text, current_number, top_text_y, serial_number_y):
        """套盒模板盒标外观二渲染"""
        clean_top_text = self._clean_text_for_font(top_text)
        self._set_best_font(c, 14, bold=True)
        
        # 外观二：Game title左对齐，但溢出文本居中
        max_width = width * 0.8
        title_lines = self._wrap_text_to_fit(c, clean_top_text, max_width, 14)
        
        if len(title_lines) > 1:
            # 首行左对齐，其他行居中
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + 15 + offset[1], title_lines[0])
            for i, line in enumerate(title_lines[1:], 1):
                for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                    c.drawCentredString(width / 2 + offset[0], top_text_y + 15 - i * 16 + offset[1], line)
        else:
            for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
                c.drawString(width * 0.1 + offset[0], top_text_y + offset[1], title_lines[0])
        
        # 绘制序列号
        for offset in [(-0.3, 0), (0.3, 0), (0, -0.3), (0, 0.3), (0, 0)]:
            c.drawCentredString(width / 2 + offset[0], serial_number_y + offset[1], current_number)