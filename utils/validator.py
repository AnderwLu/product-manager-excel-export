# -*- coding: utf-8 -*-
"""
数据验证工具类
"""

class ProductValidator:
    """商品数据验证类"""
    
    @staticmethod
    def validate_product_data(name, price, quantity):
        """验证商品数据"""
        # 检查必填字段
        if not name or not name.strip():
            return {
                'valid': False,
                'message': '商品名称不能为空'
            }
        
        if not price:
            return {
                'valid': False,
                'message': '价格不能为空'
            }
        
        if not quantity:
            return {
                'valid': False,
                'message': '数量不能为空'
            }
        
        # 验证价格
        try:
            price_float = float(price)
            if price_float <= 0:
                return {
                    'valid': False,
                    'message': '价格必须大于0'
                }
        except ValueError:
            return {
                'valid': False,
                'message': '价格必须是有效数字'
            }
        
        # 验证数量
        try:
            quantity_int = int(quantity)
            if quantity_int < 0:
                return {
                    'valid': False,
                    'message': '数量必须大于等于0'
                }
        except ValueError:
            return {
                'valid': False,
                'message': '数量必须是有效整数'
            }
        
        # 验证名称长度
        if len(name.strip()) > 100:
            return {
                'valid': False,
                'message': '商品名称不能超过100个字符'
            }
        
        return {
            'valid': True,
            'message': '数据验证通过'
        }
    
    @staticmethod
    def validate_search_params(page, per_page):
        """验证搜索参数"""
        try:
            page_int = int(page) if page else 1
            per_page_int = int(per_page) if per_page else 10
            
            if page_int < 1:
                return {
                    'valid': False,
                    'message': '页码必须大于0'
                }
            
            if per_page_int < 1 or per_page_int > 100:
                return {
                    'valid': False,
                    'message': '每页数量必须在1-100之间'
                }
            
            return {
                'valid': True,
                'page': page_int,
                'per_page': per_page_int
            }
        except ValueError:
            return {
                'valid': False,
                'message': '页码和每页数量必须是有效数字'
            }
