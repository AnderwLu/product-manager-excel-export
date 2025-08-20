# -*- coding: utf-8 -*-
"""
导出服务层 - 原图插入方案
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.comments import Comment
import os
from PIL import Image
import io
import tempfile
import uuid

class ExportService:
    def __init__(self):
        # 模板文件路径
        self.template_path = "templates/product_template.xlsm"
    
    def _insert_original_image(self, image_path, ws, row_num, col_num):
        """插入原图到Excel，不做缩放"""
        temp_path = None
        try:
            if not os.path.exists(image_path):
                return False
                
            # 获取原图尺寸
            with Image.open(image_path) as img:
                original_width = img.width
                original_height = img.height
                
                # 转换为RGB模式（如果需要）
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                
                # 使用临时目录和唯一文件名
                temp_dir = tempfile.gettempdir()
                temp_filename = f"excel_image_{uuid.uuid4().hex}.png"
                temp_path = os.path.join(temp_dir, temp_filename)
                
                # 保存为临时文件
                img.save(temp_path, format='PNG', quality=100)
                
                # 插入原图到Excel
                img_excel = XLImage(temp_path)
                img_excel.anchor = f"{get_column_letter(col_num)}{row_num}"
                
                # 设置图片属性
                img_excel.width = original_width
                img_excel.height = original_height
                
                # 添加到工作表
                ws.add_image(img_excel)
                
                # 设置行高和列宽以适应原图
                # 行高：图片高度 * 0.75（像素转点），最小100，最大200
                row_height = max(min(original_height * 0.75, 200), 100)
                ws.row_dimensions[row_num].height = row_height
                
                # 列宽：使用更精确的像素到列宽转换
                # Excel列宽单位：1列宽 ≈ 7-8像素（根据字体和分辨率调整）
                # 使用7.5作为转换比例，这是比较标准的比例
                pixel_width = original_width + 20  # 图片宽度 + 20像素边距
                col_width = max(min(pixel_width / 7.5, 50), 8)  # 最小8，最大50
                ws.column_dimensions[get_column_letter(col_num)].width = col_width
                
                return True, temp_path  # 返回临时文件路径，稍后清理
                
        except Exception as e:
            print(f"插入原图失败: {e}")
            return False, None

    def export_to_excel(self, products, selected_columns):
        """导出商品数据到Excel，图片以原图方式插入"""
        
        # 尝试使用模板文件
        if os.path.exists(self.template_path):
            try:
                # 加载模板文件，保留所有内容（包括JS宏）
                wb = openpyxl.load_workbook(self.template_path, keep_vba=True)
                
                # 检查模板文件结构
                print(f"✓ 成功加载模板文件: {self.template_path}")
                print(f"✓ 工作表: {wb.sheetnames}")
                print(f"✓ JS宏模板已加载")
                
                # 获取或创建活动工作表
                if '商品信息模板' in wb.sheetnames:
                    ws = wb['商品信息模板']
                    # 重命名工作表
                    ws.title = "商品信息"
                else:
                    ws = wb.active
                    ws.title = "商品信息"
                    
            except Exception as e:
                print(f"加载模板失败: {e}，使用新工作簿")
                wb = openpyxl.Workbook()
                ws = wb.active
        else:
            print(f"模板文件不存在: {self.template_path}，使用新工作簿")
            wb = openpyxl.Workbook()
            ws = wb.active
        
        # 设置工作表标题
        ws.title = "商品信息"
        
        # 存储需要清理的临时文件
        temp_files = []
        
        try:
            # 完全清空模板内容，只保留宏代码
            self._clear_template_completely(ws)
            
            # 根据用户选择的列动态创建表头
            col_num = 1
            for col in selected_columns:
                cell = ws.cell(row=1, column=col_num)
                cell.value = self._get_column_title(col)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                col_num += 1
            
            # 填充数据
            for row_num, product in enumerate(products, start=2):
                col_num = 1
                for col in selected_columns:
                    cell = ws.cell(row=row_num, column=col_num)
                    
                    if col == 'image_path' and product.get('image_path'):
                        # 图片列：插入原图
                        image_path = f"uploads/{product['image_path']}"
                        
                        if os.path.exists(image_path):
                            # 插入原图
                            success, temp_file = self._insert_original_image(image_path, ws, row_num, col_num)
                            
                            if success:
                                # 清空单元格文本内容
                                cell.value = ""
                                # 添加提示信息到批注
                                cell.comment = Comment("双击图片查看原图，或拖拽调整大小", "系统")
                                # 记录临时文件路径
                                if temp_file:
                                    temp_files.append(temp_file)
                            else:
                                cell.value = "图片插入失败"
                        else:
                            cell.value = "图片文件不存在"
                    else:
                        # 其他列：正常显示数据
                        value = product.get(col, '')
                        if col == 'price':
                            cell.value = float(value) if value else 0
                            cell.number_format = '¥#,##0.00'
                        elif col == 'quantity':
                            cell.value = int(value) if value else 0
                        elif col == 'create_time':
                            cell.value = str(value) if value else ''
                        else:
                            cell.value = str(value) if value else ''
                    
                    # 设置单元格样式
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    col_num += 1
            
            # 智能调整列宽 - 根据列类型和内容设置最佳宽度
            self._adjust_all_column_widths(ws, selected_columns, products)
            
            # 保存到内存
            output = io.BytesIO()
            
            # 如果使用模板文件，确保保存为.xlsm格式以保留宏
            if os.path.exists(self.template_path):
                # 保存为.xlsm格式，保留JS宏代码
                wb.save(output)
                print("✓ 文件已保存为.xlsm格式，JS宏代码已保留")
            else:
                # 新工作簿保存为.xlsx格式
                wb.save(output)
                print("✓ 文件已保存为.xlsx格式")
            
            output.seek(0)
            return output
            
        finally:
            # 清理临时文件
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                except:
                    pass  # 忽略删除失败的错误
    
    def _clear_template_completely(self, ws):
        """完全清空模板内容，只保留JS宏代码"""
        try:
            # 获取当前工作表的名称
            original_name = ws.title
            
            # 删除所有行，但保留工作表本身和JS宏
            max_row = ws.max_row
            if max_row > 0:
                # 从最后一行开始删除，避免索引问题
                for row in range(max_row, 0, -1):
                    ws.delete_rows(row)
                print("✓ 模板行内容已清空")
            
            # 重新设置工作表标题
            ws.title = original_name
            
            print("✓ 模板内容已完全清空，JS宏代码已保留")
                
        except Exception as e:
            print(f"清空模板内容失败: {e}")
            # 如果删除行失败，尝试清空单元格内容
            try:
                max_row = ws.max_row
                max_col = ws.max_column
                
                if max_row > 0 and max_col > 0:
                    for row in range(1, max_row + 1):
                        for col in range(1, max_col + 1):
                            cell = ws.cell(row=row, column=col)
                            cell.value = None
                            cell.comment = None
                    print("✓ 模板内容已清空（备用方案）")
            except Exception as e2:
                print(f"备用清空方案也失败: {e2}")
    
    def _clear_template_data(self, ws):
        """清空模板中的现有数据，保留宏代码"""
        try:
            # 保留第1行（标题行），清空其他数据行
            max_row = ws.max_row
            if max_row > 1:
                # 删除第2行到最后一行
                ws.delete_rows(2, max_row - 1)
            print("✓ 模板数据已清空")
        except Exception as e:
            print(f"清空模板数据失败: {e}")
    
    def _adjust_all_column_widths(self, ws, selected_columns, products_data):
        """智能调整所有列的宽度"""
        try:
            for col_idx, col_name in enumerate(selected_columns, start=1):
                col_letter = get_column_letter(col_idx)
                
                if col_name == 'image_path':
                    # 图片列：根据图片尺寸设置宽度
                    self._adjust_image_column_width(ws, col_letter, col_idx)
                else:
                    # 文本列：根据内容设置宽度
                    self._adjust_text_column_width(ws, col_letter, col_name, products_data)
                    
        except Exception as e:
            print(f"调整列宽失败: {e}")
    
    def _adjust_image_column_width(self, ws, col_letter, col_idx):
        """调整图片列的宽度"""
        try:
            # 查找该列中的图片
            for row in range(2, ws.max_row + 1):  # 从第2行开始（跳过标题）
                cell = ws.cell(row=row, column=col_idx)
                if cell.comment and "双击图片查看原图" in str(cell.comment):
                    # 这是图片行，根据图片尺寸设置列宽
                    # 默认图片列宽度：15-20之间
                    ws.column_dimensions[col_letter].width = 18
                    break
            else:
                # 没有找到图片，设置默认宽度
                ws.column_dimensions[col_letter].width = 15
                
        except Exception as e:
            print(f"调整图片列宽失败: {e}")
            ws.column_dimensions[col_letter].width = 15  # 设置默认宽度
    
    def _adjust_text_column_width(self, ws, col_letter, col_name, products_data):
        """调整文本列的宽度"""
        try:
            # 获取列标题
            title = self._get_column_title(col_name)
            title_length = len(title)
            
            # 获取该列所有数据的最大长度
            max_length = title_length
            
            for product in products_data:
                value = str(product.get(col_name, ''))
                if len(value) > max_length:
                    max_length = len(value)
            
            # 根据列类型设置不同的宽度策略
            if col_name == 'id':
                # ID列：固定宽度，不需要太宽
                target_width = max(8, min(max_length + 2, 12))
            elif col_name == 'name':
                # 商品名称列：需要较宽，但有限制
                target_width = max(15, min(max_length + 3, 40))
            elif col_name == 'price':
                # 价格列：固定宽度，包含货币符号
                target_width = max(10, min(max_length + 2, 15))
            elif col_name == 'quantity':
                # 数量列：固定宽度
                target_width = max(8, min(max_length + 2, 12))
            elif col_name == 'spec':
                # 规格列：根据内容调整
                target_width = max(10, min(max_length + 2, 25))
            elif col_name == 'create_time':
                # 时间列：固定宽度，时间格式固定
                target_width = max(15, min(max_length + 2, 20))
            else:
                # 其他列：通用策略
                target_width = max(8, min(max_length + 2, 30))
            
            # 设置列宽
            ws.column_dimensions[col_letter].width = target_width
            
        except Exception as e:
            print(f"调整文本列宽失败 {col_name}: {e}")
            # 设置默认宽度
            ws.column_dimensions[col_letter].width = 15
    
    def _optimize_image_column_widths(self, ws, selected_columns):
        """优化图片列的列宽，确保图片能完整显示"""
        try:
            for col_idx, col_name in enumerate(selected_columns, start=1):
                if col_name == 'image_path':
                    # 图片列：确保列宽足够显示图片
                    col_letter = get_column_letter(col_idx)
                    current_width = ws.column_dimensions[col_letter].width
                    
                    # 如果当前列宽小于15，设置为15（确保基本显示）
                    if current_width < 15:
                        ws.column_dimensions[col_letter].width = 15
                    
                    # 如果当前列宽大于50，限制为50（避免过宽）
                    elif current_width > 50:
                        ws.column_dimensions[col_letter].width = 50
                        
        except Exception as e:
            print(f"优化图片列宽失败: {e}")
    
    def _get_column_title(self, column):
        """获取列标题"""
        titles = {
            'id': 'ID',
            'name': '商品名称',
            'price': '价格',
            'quantity': '数量',
            'spec': '规格',
            'image_path': '图片',
            'create_time': '创建时间'
        }
        return titles.get(column, column)
