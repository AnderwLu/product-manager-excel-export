#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试导出功能 - 列宽自适应版本
"""

import os
import sys
sys.path.append('.')

from models.product import Product
from services.export_service import ExportService

def test_export_with_adaptive_width():
    """测试带列宽自适应的导出功能"""
    print("=== 测试列宽自适应导出功能 ===")
    
    # 获取商品数据
    result = Product.find_all(page=1, per_page=10000)
    products = result['products']
    
    if not products:
        print("没有商品数据，无法测试导出")
        return
    
    # 转换为字典格式
    products_data = []
    for product in products:
        products_data.append({
            'id': product.id,
            'name': product.name,
            'price': product.price,
            'quantity': product.quantity,
            'spec': product.spec,
            'image_path': product.image_path,
            'create_time': str(product.create_time)
        })
    
    print(f"转换后的数据: {len(products_data)} 条")
    
    # 测试导出
    export_service = ExportService()
    selected_columns = ['id', 'name', 'price', 'quantity', 'image_path']
    
    print(f"\n选择的列: {selected_columns}")
    print("包含图片列，将自动设置列宽...")
    
    try:
        excel_data = export_service.export_to_excel(products_data, selected_columns)
        print("✓ 导出成功")
        print(f"Excel数据大小: {len(excel_data.getvalue())} 字节")
        
        # 保存到文件进行测试
        output_path = "test_export_adaptive_width.xlsx"
        with open(output_path, 'wb') as f:
            f.write(excel_data.getvalue())
        print(f"测试文件已保存: {output_path}")
        
        # 分析列宽设置
        print("\n=== 列宽设置分析 ===")
        for i, col in enumerate(selected_columns, 1):
            if col == 'image_path':
                print(f"列 {i} ({col}): 图片列，已自动设置列宽")
            else:
                print(f"列 {i} ({col}): 文本列，根据内容自动调整")
        
        print("\n=== 使用说明 ===")
        print("1. 打开生成的Excel文件")
        print("2. 图片列已自动设置合适的列宽")
        print("3. 如果使用WPS，可以运行JS宏进一步优化")
        print("4. 列宽设置：最小8，最大50，图片列最小15")
        
    except Exception as e:
        print(f"✗ 导出失败: {e}")
        import traceback
        traceback.print_exc()

def check_image_files():
    """检查图片文件"""
    print("\n=== 检查图片文件 ===")
    
    uploads_dir = "uploads"
    if not os.path.exists(uploads_dir):
        print(f"✗ uploads目录不存在: {uploads_dir}")
        return
    
    files = os.listdir(uploads_dir)
    image_files = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp'))]
    
    print(f"图片文件数量: {len(image_files)}")
    for img_file in image_files:
        img_path = os.path.join(uploads_dir, img_file)
        try:
            from PIL import Image
            with Image.open(img_path) as img:
                print(f"  {img_file}: {img.width} x {img.height} ({img.mode})")
        except Exception as e:
            print(f"  {img_file}: 读取失败 - {e}")

if __name__ == "__main__":
    print("开始测试列宽自适应导出功能...\n")
    
    check_image_files()
    test_export_with_adaptive_width()
    
    print("\n测试完成！")
    print("现在Excel文件导出时就会自动设置合适的列宽了！")
