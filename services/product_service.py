# -*- coding: utf-8 -*-
"""
商品业务服务层
"""

from models.product import Product
from utils.file_handler import FileHandler
from utils.validator import ProductValidator

class ProductService:
    """商品业务服务类"""
    
    def __init__(self):
        self.file_handler = FileHandler()
        self.validator = ProductValidator()
    
    def add_product(self, name, price, quantity, spec, image_file, salesperson=None, doc_date=None, product_desc=None, remark=None, settlement_account=None, description=None, freight=None, paid_total=None):
        """添加商品"""
        # 录入必填校验（单据日期、客户名称、品名规格、数量）
        req = self.validator.validate_entry_required(doc_date, name, product_desc, quantity)
        if not req.get('valid'):
            return req

        # 数据验证
        validation_result = self.validator.validate_product_data(name, price, quantity)
        if not validation_result['valid']:
            return validation_result
        
        # 处理图片上传
        image_path = None
        if image_file:
            upload_result = self.file_handler.upload_image(image_file)
            if not upload_result['success']:
                return upload_result
            image_path = upload_result['filename']
        
        # 创建商品对象
        product = Product(
            name=name,
            price=float(price),
            quantity=int(quantity),
            spec=spec,
            image_path=image_path,
            # 新增字段入库（保持后端流程，未改原有SQL列集）
            salesperson=salesperson,
            doc_date=doc_date,
            product_desc=product_desc,
            remark=remark,
            settlement_account=settlement_account,
            description=description,
            freight=float(freight) if freight not in (None, '') else None,
            paid_total=float(paid_total) if paid_total not in (None, '') else None
        )
        
        # 保存到数据库
        try:
            product.save()
            return {
                'success': True,
                'message': '商品添加成功',
                'product_id': product.id
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'保存失败: {str(e)}'
            }
    
    def get_products(self, page=1, per_page=10, search=None):
        """获取商品列表"""
        try:
            result = Product.find_all(page, per_page, search)
            
            # 转换为字典格式
            products_dict = []
            for product in result['products']:
                products_dict.append(product.to_dict())
            
            return {
                'success': True,
                'data': {
                    'products': products_dict,
                    'total': result['total'],
                    'page': result['page'],
                    'per_page': result['per_page'],
                    'total_pages': result['total_pages']
                }
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'查询失败: {str(e)}'
            }
    
    def delete_product(self, product_id):
        """删除商品"""
        try:
            product = Product.find_by_id(product_id)
            if not product:
                return {
                    'success': False,
                    'message': '商品不存在'
                }
            
            # 删除图片文件
            if product.image_path:
                self.file_handler.delete_image(product.image_path)
            
            # 删除数据库记录
            product.delete()
            
            return {
                'success': True,
                'message': '商品删除成功'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'删除失败: {str(e)}'
            }
    
    def update_product(self, product_id, name, price, quantity, spec, image_file,
                       product_desc=None, remark=None, settlement_account=None,
                       description=None, freight=None, paid_total=None, doc_date=None,
                       delete_image=False):
        """更新商品"""
        try:
            product = Product.find_by_id(product_id)
            if not product:
                return {
                    'success': False,
                    'message': '商品不存在'
                }
            
            # 数据验证
            validation_result = self.validator.validate_product_data(name, price, quantity)
            if not validation_result['valid']:
                return validation_result
            
            # 处理图片上传/删除
            if  delete_image and product.image_path:
                self.file_handler.delete_image(product.image_path)
                product.image_path = None
            if image_file:
                upload_result = self.file_handler.upload_image(image_file)
                if not upload_result['success']:
                    return upload_result
                
                # 删除旧图片
                if product.image_path:
                    self.file_handler.delete_image(product.image_path)
                
                product.image_path = upload_result['filename']
            
            # 更新商品信息
            product.name = name
            product.price = float(price)
            product.quantity = int(quantity)
            product.spec = spec
            # 其它字段（与录入一致，允许修改单据日期）
            if product_desc is not None:
                product.product_desc = product_desc
            if doc_date is not None and str(doc_date).strip() != '':
                product.doc_date = doc_date
            if remark is not None:
                product.remark = remark
            if settlement_account is not None:
                product.settlement_account = settlement_account
            if description is not None:
                product.description = description
            if freight not in (None, ''):
                try:
                    product.freight = float(freight)
                except Exception:
                    pass
            if paid_total not in (None, ''):
                try:
                    product.paid_total = float(paid_total)
                except Exception:
                    pass
            
            # 保存到数据库
            product.save()
            
            return {
                'success': True,
                'message': '商品更新成功'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'更新失败: {str(e)}'
            }

    def batch_update_products(self, items):
        """批量按部分字段更新商品
        items: List[{'id': int, 'fields': dict}]
        仅允许更新白名单字段，忽略未知字段。
        """
        from models.database import db_manager
        if not isinstance(items, list) or not items:
            return { 'success': False, 'message': '无有效更新项' }

        # 允许更新的字段（与前端可编辑列一致）
        allowed_fields = {
            'doc_date', 'customer_name', 'product_desc', 'unit',
            'quantity', 'unit_price', 'unit_discount_rate', 'remark',
            'freight', 'order_discount_rate', 'paid_total',
            'settlement_account', 'description'
        }

        success_count = 0
        fail_items = []

        for item in items:
            try:
                pid = int(item.get('id'))
                fields = item.get('fields') or {}
                if not pid or not isinstance(fields, dict) or not fields:
                    fail_items.append({ 'id': item.get('id'), 'error': '参数无效' })
                    continue

                # 过滤字段
                update_fields = {k: v for k, v in fields.items() if k in allowed_fields}
                if not update_fields:
                    fail_items.append({ 'id': pid, 'error': '无可更新字段' })
                    continue

                # 构造SQL
                set_parts = []
                params = []
                for k, v in update_fields.items():
                    set_parts.append(f"{k}=?")
                    params.append(v)
                # 自动更新 update_time
                set_parts.append("update_time=datetime('now')")
                sql = f"UPDATE products SET {', '.join(set_parts)} WHERE id=?"
                params.append(pid)

                affected = db_manager.execute_update(sql, tuple(params))
                if affected > 0:
                    success_count += 1
                else:
                    fail_items.append({ 'id': pid, 'error': '未找到或未更新' })
            except Exception as e:
                fail_items.append({ 'id': item.get('id'), 'error': str(e) })

        return {
            'success': True,
            'message': '批量更新完成',
            'data': {
                'success_count': success_count,
                'fail_count': len(fail_items),
                'fails': fail_items
            }
        }

    def update_product_image(self, product_id, image_file=None, delete_image=False):
        """单独更新商品图片（支持替换或删除）"""
        try:
            product = Product.find_by_id(product_id)
            if not product:
                return { 'success': False, 'message': '商品不存在' }

            # 删除图片
            if delete_image:
                if product.image_path:
                    self.file_handler.delete_image(product.image_path)
                product.image_path = None
                product.save()
                return { 'success': True, 'message': '图片已删除' }

            # 替换图片
            if image_file:
                upload_result = self.file_handler.upload_image(image_file)
                if not upload_result.get('success'):
                    return upload_result
                new_filename = upload_result['filename']
                # 删除旧图
                if product.image_path:
                    self.file_handler.delete_image(product.image_path)
                product.image_path = new_filename
                product.save()
                return { 'success': True, 'message': '图片更新成功', 'filename': new_filename }

            return { 'success': False, 'message': '未提供图片或删除标记' }
        except Exception as e:
            return { 'success': False, 'message': f'图片更新失败: {str(e)}' }
