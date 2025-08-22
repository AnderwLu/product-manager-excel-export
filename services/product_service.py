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
    
    def update_product(self, product_id, name, price, quantity, spec, image_file):
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
            
            # 处理图片上传
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
