# -*- coding: utf-8 -*-
"""
商品控制器
"""

from flask import Blueprint, request, jsonify, send_file
from services.product_service import ProductService
from services.export_service import ExportService
from models.product import Product
import io

product_bp = Blueprint('product', __name__)
product_service = ProductService()
export_service = ExportService()

@product_bp.route('/add', methods=['POST'])
def add_product():
    """添加商品"""
    try:
        result = product_service.add_product(request)
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
        
        result = product_service.get_products(page, per_page, search)
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
        
        result = product_service.delete_product(product_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})

@product_bp.route('/update', methods=['POST'])
def update_product():
    """更新商品"""
    try:
        result = product_service.update_product(request)
        return jsonify(result)
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
        
        # 获取所有商品数据
        result = Product.find_all(page=1, per_page=10000)  # 获取所有数据
        products = result['products']
        
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
        
        # 导出到Excel
        excel_data = export_service.export_to_excel(products_data, selected_columns)
        
        # 生成文件名 - 使用.xlsm格式以支持宏
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'products_{timestamp}.xlsm'
        
        return send_file(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'导出失败: {str(e)}'})
