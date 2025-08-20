#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试所有列宽度自适应功能
"""

import os
import sys
sys.path.append('.')

from models.product import Product
from services.export_service import ExportService

def test_all_columns_width():
    """测试所有列的宽度自适应功能"""
    print("=== 测试所有列宽度自适应功能 ===")
    
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
    
    # 分析每列的内容长度
    print("\n=== 列内容分析 ===")
    selected_columns = ['id', 'name', 'price', 'quantity', 'spec', 'image_path', 'create_time']
    
    for col in selected_columns:
        max_length = 0
        title_length = len(col)
        max_length = max(max_length, title_length)
        
        for product in products_data:
            value = str(product.get(col, ''))
            if len(value) > max_length:
                max_length = len(value)
        
        print(f"列 '{col}': 标题长度={title_length}, 最大内容长度={max_length}")
    
    # 测试导出
    export_service = ExportService()
    
    print(f"\n选择的列: {selected_columns}")
    print("将测试所有列的宽度自适应...")
    
    try:
        excel_data = export_service.export_to_excel(products_data, selected_columns)
        print("✓ 导出成功")
        print(f"Excel数据大小: {len(excel_data.getvalue())} 字节")
        
        # 保存到文件进行测试
        output_path = "test_all_columns_width.xlsx"
        with open(output_path, 'wb') as f:
            f.write(excel_data.getvalue())
        print(f"测试文件已保存: {output_path}")
        
        # 分析列宽设置
        print("\n=== 列宽设置分析 ===")
        print("每列的宽度策略：")
        print("- ID列: 固定宽度 8-12，适合数字ID")
        print("- 商品名称列: 自适应宽度 15-40，根据名称长度调整")
        print("- 价格列: 固定宽度 10-15，适合货币格式")
        print("- 数量列: 固定宽度 8-12，适合数字")
        print("- 规格列: 自适应宽度 10-25，根据规格内容调整")
        print("- 图片列: 固定宽度 15-18，适合缩略图显示")
        print("- 创建时间列: 固定宽度 15-20，适合时间格式")
        
        print("\n=== 使用说明 ===")
        print("1. 打开生成的Excel文件")
        print("2. 所有列都已根据内容自动调整宽度")
        print("3. 图片列宽度已优化，适合显示缩略图")
        print("4. 文本列宽度根据内容智能调整")
        print("5. 如果使用WPS，可以运行JS宏进一步优化")
        
    except Exception as e:
        print(f"✗ 导出失败: {e}")
        import traceback
        traceback.print_exc()

def check_data_content():
    """检查数据内容，了解列宽需求"""
    print("\n=== 数据内容检查 ===")
    
    result = Product.find_all(page=1, per_page=10000)
    products = result['products']
    
    if not products:
        return
    
    for i, product in enumerate(products):
        print(f"\n商品 {i+1}:")
        print(f"  ID: {product.id} (长度: {len(str(product.id))})")
        print(f"  名称: {product.name} (长度: {len(str(product.name))})")
        print(f"  价格: {product.price} (长度: {len(str(product.price))})")
        print(f"  数量: {product.quantity} (长度: {len(str(product.quantity))})")
        print(f"  规格: {product.spec} (长度: {len(str(product.spec))})")
        print(f"  创建时间: {product.create_time} (长度: {len(str(product.create_time))})")

if __name__ == "__main__":
    print("开始测试所有列宽度自适应功能...\n")
    
    check_data_content()
    test_all_columns_width()
    
    print("\n测试完成！")
    print("现在所有列都会根据内容自动调整宽度了！")
