#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品信息管理系统控制器
"""

from flask import Blueprint, request, jsonify, send_file
from models.product import Product
from services.export_service import ExportService
import logging

# 使用主应用的日志配置
logger = logging.getLogger(__name__)

# 创建导出服务实例
export_service = ExportService()

# 创建蓝图
product_bp = Blueprint('product', __name__)

@product_bp.route('/add', methods=['POST'])
def add_product():
    """添加商品"""
    try:
        # 从表单数据中提取字段
        name = request.form.get('name')
        price = request.form.get('price')
        quantity = request.form.get('quantity')
        spec = request.form.get('spec', '')
        image_file = request.files.get('image')
        
        # 直接调用Product模型方法
        result = Product.create(name, price, quantity, spec, image_file)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})

@product_bp.route('/list', methods=['GET'])
def get_products():
    """获取商品列表"""
    try:
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 10, type=int)
        search = request.args.get('search', '')
        
        result = Product.find_all(page, per_page, search)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'获取失败: {str(e)}'})

@product_bp.route('/delete', methods=['POST'])
def delete_product():
    """删除商品"""
    try:
        data = request.get_json()
        product_id = data.get('id')
        
        if not product_id:
            return jsonify({'success': False, 'message': '商品ID不能为空'})
        
        result = Product.delete(product_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})

@product_bp.route('/update', methods=['POST'])
def update_product():
    """更新商品"""
    try:
        # 这里需要实现更新逻辑，暂时返回错误
        return jsonify({'success': False, 'message': '更新功能暂未实现'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})

@product_bp.route('/export', methods=['POST'])
def export_products():
    """导出商品数据到Excel"""
    try:
        data = request.get_json()
        selected_columns = data.get('columns', [])
        
        if not selected_columns:
            return jsonify({'success': False, 'message': '请选择要导出的列'})
        
        logger.info(f"导出请求 - 选择的列: {selected_columns}")
        
        # 获取所有商品数据
        result = Product.find_all(page=1, per_page=10000)  # 获取所有数据
        products = result['products']
        
        logger.info(f"获取到 {len(products)} 条商品数据")
        
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
        
        logger.info(f"转换后的数据: {products_data[:2]}...")  # 显示前两条数据
        
        # 导出到Excel
        excel_data = export_service.export_to_excel(products_data, selected_columns)
        
        if excel_data is None:
            return jsonify({'success': False, 'message': '导出服务返回空数据'})
        
        logger.info(f"导出服务返回数据大小: {len(excel_data)} 字节")
        logger.info(f"文件头信息: {excel_data[:20]}")
        
        # 检查文件类型
        if excel_data.startswith(b'PK\x03\x04'):
            logger.info("✓ 确认文件格式: 标准xlsx格式 (ZIP压缩包)")
        else:
            logger.warning(f"⚠️ 文件格式异常: {excel_data[:10]}")
        
        # 生成文件名 - 使用.xlsx格式
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'products_{timestamp}.xlsx'
        
        logger.info(f"设置文件名: {filename}")
        logger.info(f"设置MIME类型: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # 使用BytesIO创建文件对象
        from io import BytesIO
        excel_file = BytesIO(excel_data)
        excel_file.seek(0)
        
        logger.info(f"✓ 准备返回文件，大小: {len(excel_data)} 字节")
        
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f'导出失败: {str(e)}')
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'message': f'导出失败: {str(e)}'})
