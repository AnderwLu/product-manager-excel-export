#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品信息管理系统控制器
"""

from flask import Blueprint, request, jsonify, send_file, session, redirect, url_for
from models.product import Product
from services.export_service import ExportService
from services.product_service import ProductService
from models.user_pref import UserPreference
import logging

# 使用主应用的日志配置
from logging_config import get_logger
logger = get_logger(__name__)

# 创建导出服务实例
export_service = ExportService()
product_service = ProductService()

# 创建蓝图
product_bp = Blueprint('product', __name__)

def ensure_logged_in():
    if not session.get('user_id'):
        return False
    return True

@product_bp.route('/add', methods=['POST'])
def add_product():
    """添加商品"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        # 从表单数据中提取字段
        name = request.form.get('name')
        price = request.form.get('price')
        quantity = request.form.get('quantity')
        spec = request.form.get('spec', '')
        image_file = request.files.get('image')
        # 新字段（可选）
        doc_date = request.form.get('doc_date')
        product_desc = request.form.get('product_desc')
        remark = request.form.get('remark')
        settlement_account = request.form.get('settlement_account')
        description = request.form.get('description')
        freight = request.form.get('freight')
        paid_total = request.form.get('paid_total')
        # 后台自动补录营业员
        salesperson = (session.get('real_name') or session.get('username') or '').strip()

        # 使用服务层处理业务逻辑
        result = product_service.add_product(
            name, price, quantity, spec, image_file,
            salesperson=salesperson,
            doc_date=doc_date,
            product_desc=product_desc,
            remark=remark,
            settlement_account=settlement_account,
            description=description,
            freight=freight,
            paid_total=paid_total
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'添加失败: {str(e)}'})

@product_bp.route('/list', methods=['GET'])
def get_products():
    """获取商品列表"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 10, type=int)
        search = request.args.get('search', '')  # 客户名称模糊
        product_desc = request.args.get('product_desc', '')
        salesperson = request.args.get('salesperson', '')
        date_start = request.args.get('date_start', '')
        date_end = request.args.get('date_end', '')
        
        result = Product.find_all(
            page, per_page, search,
            product_desc=product_desc or None,
            salesperson=salesperson or None,
            date_start=date_start or None,
            date_end=date_end or None
        )
        logger.info(f"获取商品列表: total={result.get('total')}, page={page}, per_page={per_page}")

        if not result or 'products' not in result:
            return jsonify({'success': False, 'message': '数据为空或格式不正确'})

        products = result['products']

        # 使用模型提供的 to_dict 进行序列化
        products_data = [p.to_dict() if hasattr(p, 'to_dict') else p for p in products]

        # 构建与前端期望一致的结构
        response_data = {
            'success': True,
            'data': {
                'products': products_data,
                'page': page,
                'total_pages': result.get('total_pages', 1)
            }
        }

        return jsonify(response_data)
    except Exception as e:
        logger.error(f'获取商品列表失败: {str(e)}')
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'message': f'获取失败: {str(e)}'})

@product_bp.route('/delete', methods=['POST'])
def delete_product():
    """删除商品"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        data = request.get_json()
        product_id = data.get('id')

        if not product_id:
            return jsonify({'success': False, 'message': '商品ID不能为空'})

        # 使用服务层执行删除，统一处理图片与记录
        result = product_service.delete_product(product_id)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'删除失败: {str(e)}'})

@product_bp.route('/update', methods=['POST'])
def update_product():
    """更新商品"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        product_id = request.form.get('id')
        name = request.form.get('name')
        price = request.form.get('price')
        quantity = request.form.get('quantity')
        spec = request.form.get('spec', '')
        image_file = request.files.get('image')

        if not product_id:
            return jsonify({'success': False, 'message': '商品ID不能为空'})

        # 其余可改字段，与录入一致（单据日期禁改）
        product_desc = request.form.get('product_desc')
        remark = request.form.get('remark')
        settlement_account = request.form.get('settlement_account')
        description = request.form.get('description')
        freight = request.form.get('freight')
        paid_total = request.form.get('paid_total')
        delete_image = request.form.get('delete_image')
        result = product_service.update_product(
            product_id, name, price, quantity, spec, image_file,
            product_desc=product_desc,
            remark=remark,
            settlement_account=settlement_account,
            description=description,
            freight=freight,
            paid_total=paid_total,
            delete_image=delete_image
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'message': f'更新失败: {str(e)}'})

@product_bp.route('/export', methods=['POST'])
def export_products():
    """导出商品数据到Excel"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        data = request.get_json()
        selected_columns = data.get('columns', [])
        # 读取筛选条件（与 /list 一致）
        filters = data.get('filters', {})
        
        if not selected_columns:
            return jsonify({'success': False, 'message': '请选择要导出的列'})
        
        logger.info(f"导出请求 - 选择的列: {selected_columns}")
        
        # 获取所有商品数据
        result = Product.find_all(
            page=1, per_page=1000000,
            search=filters.get('search'),
            product_desc=filters.get('product_desc'),
            salesperson=filters.get('salesperson'),
            date_start=filters.get('date_start'),
            date_end=filters.get('date_end')
        )
        products = result['products']
        
        logger.info(f"获取到 {len(products)} 条商品数据")
        
        # 转换为字典格式，确保所有数据都是JSON可序列化的
        products_data = []
        for product in products:
            # 检查product是否为Product对象，如果是则转换为字典
            if hasattr(product, '__dict__'):
                product_dict = product.__dict__.copy()
                # 移除SQLAlchemy内部属性
                product_dict.pop('_sa_instance_state', None)
                # 确保datetime对象转换为字符串
                if 'create_time' in product_dict and product_dict['create_time']:
                    product_dict['create_time'] = str(product_dict['create_time'])
                products_data.append(product_dict)
            else:
                # 如果已经是字典，直接使用
                products_data.append(product)
        
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
        
        # 根据平台动态生成文件名
        from datetime import datetime
        import platform
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Windows下导出为xlsx，Mac/Linux下导出为xlsm
        system_type = platform.system()
        if system_type == 'Windows':
            file_extension = 'xlsx'
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            file_extension = 'xlsm'
            mime_type = 'application/vnd.ms-excel.sheet.macroEnabled.12'
        
        filename = f'{timestamp}.{file_extension}'
        
        logger.info(f"当前系统: {system_type}")
        logger.info(f"设置文件名: {filename}")
        logger.info(f"设置MIME类型: {mime_type}")
        
        # 使用BytesIO创建文件对象
        from io import BytesIO
        excel_file = BytesIO(excel_data)
        excel_file.seek(0)
        
        logger.info(f"✓ 准备返回文件，大小: {len(excel_data)} 字节")
        
        return send_file(
            excel_file,
            mimetype=mime_type,
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f'导出失败: {str(e)}')
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'message': f'导出失败: {str(e)}'})

@product_bp.route('/columns/save', methods=['POST'])
def save_columns_pref():
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        data = request.get_json() or {}
        cfg = data.get('columns')
        if not isinstance(cfg, list):
            return jsonify({'success': False, 'message': '参数不正确'})
        import json
        val = json.dumps(cfg, ensure_ascii=False)
        UserPreference.set_pref(session.get('user_id'), 'export_columns', val)
        return jsonify({'success': True, 'message': '列设置已保存'})
    except Exception as e:
        logger.error(f'保存列设置失败: {str(e)}')
        return jsonify({'success': False, 'message': f'保存失败: {str(e)}'})

@product_bp.route('/columns/load', methods=['GET'])
def load_columns_pref():
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        val = UserPreference.get_pref(session.get('user_id'), 'export_columns')
        import json
        cfg = json.loads(val) if val else []
        return jsonify({'success': True, 'data': cfg})
    except Exception as e:
        logger.error(f'读取列设置失败: {str(e)}')
        return jsonify({'success': False, 'message': f'读取失败: {str(e)}'})

@product_bp.route('/batch_update', methods=['POST'])
def batch_update_products():
    """批量更新商品部分字段"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        data = request.get_json() or {}
        items = data.get('items') or []
        result = product_service.batch_update_products(items)
        return jsonify(result)
    except Exception as e:
        logger.error(f'批量更新失败: {str(e)}')
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'message': f'批量更新失败: {str(e)}'})

@product_bp.route('/update_image', methods=['POST'])
def update_product_image():
    """更新或删除商品图片"""
    try:
        if not ensure_logged_in():
            return jsonify({'success': False, 'message': '未登录'}), 401
        product_id = request.form.get('id') or request.args.get('id')
        if not product_id:
            return jsonify({'success': False, 'message': '商品ID不能为空'})
        image_file = request.files.get('image')
        delete_image = request.form.get('delete_image') in ('1', 'true', 'True', 'on')
        result = product_service.update_product_image(int(product_id), image_file=image_file, delete_image=delete_image)
        return jsonify(result)
    except Exception as e:
        logger.error(f'更新图片失败: {str(e)}')
        import traceback
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'message': f'更新图片失败: {str(e)}'})
