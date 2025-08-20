#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试导出功能调试脚本
"""

import os
import sys
sys.path.append('.')

from models.product import Product
from services.export_service import ExportService

def test_data_retrieval():
    """测试数据获取"""
    print("=== 测试数据获取 ===")
    
    # 获取所有商品
    result = Product.find_all(page=1, per_page=10000)
    products = result['products']
    
    print(f"总商品数量: {result['total']}")
    print(f"当前页商品数量: {len(products)}")
    
    for i, product in enumerate(products):
        print(f"\n商品 {i+1}:")
        print(f"  ID: {product.id}")
        print(f"  名称: {product.name}")
        print(f"  价格: {product.price}")
        print(f"  数量: {product.quantity}")
        print(f"  规格: {product.spec}")
        print(f"  图片路径: {product.image_path}")
        print(f"  创建时间: {product.create_time}")
        
        # 检查图片文件是否存在
        if product.image_path:
            image_path = f"uploads/{product.image_path}"
            if os.path.exists(image_path):
                print(f"  ✓ 图片文件存在: {image_path}")
                # 获取图片尺寸
                try:
                    from PIL import Image
                    with Image.open(image_path) as img:
                        print(f"    图片尺寸: {img.width} x {img.height}")
                        print(f"    图片模式: {img.mode}")
                except Exception as e:
                    print(f"    ✗ 读取图片失败: {e}")
            else:
                print(f"  ✗ 图片文件不存在: {image_path}")
        else:
            print("  - 无图片路径")

def test_export_service():
    """测试导出服务"""
    print("\n=== 测试导出服务 ===")
    
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
    for i, data in enumerate(products_data):
        print(f"  {i+1}: {data}")
    
    # 测试导出
    export_service = ExportService()
    selected_columns = ['id', 'name', 'price', 'quantity', 'image_path']
    
    print(f"\n选择的列: {selected_columns}")
    
    try:
        excel_data = export_service.export_to_excel(products_data, selected_columns)
        print("✓ 导出成功")
        print(f"Excel数据大小: {len(excel_data.getvalue())} 字节")
        
        # 保存到文件进行测试
        output_path = "test_export_debug.xlsx"
        with open(output_path, 'wb') as f:
            f.write(excel_data.getvalue())
        print(f"测试文件已保存: {output_path}")
        
    except Exception as e:
        print(f"✗ 导出失败: {e}")
        import traceback
        traceback.print_exc()

def check_uploads_directory():
    """检查uploads目录"""
    print("\n=== 检查uploads目录 ===")
    
    uploads_dir = "uploads"
    if not os.path.exists(uploads_dir):
        print(f"✗ uploads目录不存在: {uploads_dir}")
        return
    
    print(f"✓ uploads目录存在: {uploads_dir}")
    
    # 列出所有文件
    files = os.listdir(uploads_dir)
    print(f"文件数量: {len(files)}")
    
    for file in files:
        file_path = os.path.join(uploads_dir, file)
        if os.path.isfile(file_path):
            size = os.path.getsize(file_path)
            print(f"  {file}: {size} 字节")
        else:
            print(f"  {file}: 目录")

if __name__ == "__main__":
    print("开始调试导出功能...\n")
    
    check_uploads_directory()
    test_data_retrieval()
    test_export_service()
    
    print("\n调试完成！")
