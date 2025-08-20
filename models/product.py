# -*- coding: utf-8 -*-
"""
商品数据模型
"""

from datetime import datetime
from models.database import db_manager

class Product:
    """商品模型类"""
    
    def __init__(self, id=None, name=None, price=None, quantity=None, 
                 spec=None, image_path=None, create_time=None):
        self.id = id
        self.name = name
        self.price = price
        self.quantity = quantity
        self.spec = spec
        self.image_path = image_path
        self.create_time = create_time or datetime.now()
    
    @classmethod
    def create_table(cls):
        """创建商品表"""
        sql = '''
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                price REAL NOT NULL,
                quantity INTEGER NOT NULL,
                spec TEXT,
                image_path TEXT,
                create_time TEXT DEFAULT (datetime('now'))
            )
        '''
        return db_manager.execute_update(sql)
    
    def save(self):
        """保存商品到数据库"""
        if self.id:
            # 更新
            sql = '''
                UPDATE products 
                SET name=?, price=?, quantity=?, spec=?, image_path=?
                WHERE id=?
            '''
            params = (self.name, self.price, self.quantity, self.spec, 
                     self.image_path, self.id)
            return db_manager.execute_update(sql, params)
        else:
            # 插入
            sql = '''
                INSERT INTO products (name, price, quantity, spec, image_path)
                VALUES (?, ?, ?, ?, ?)
            '''
            params = (self.name, self.price, self.quantity, self.spec, self.image_path)
            self.id = db_manager.execute_insert(sql, params)
            return self.id
    
    @classmethod
    def find_by_id(cls, product_id):
        """根据ID查找商品"""
        sql = "SELECT * FROM products WHERE id = ?"
        result = db_manager.execute_query(sql, (product_id,))
        if result:
            return cls(**result[0])
        return None
    
    @classmethod
    def find_all(cls, page=1, per_page=10, search=None):
        """查找所有商品，支持分页和搜索"""
        offset = (page - 1) * per_page
        
        # 构建查询条件
        where_clause = ""
        params = []
        if search:
            where_clause = "WHERE name LIKE ?"
            params.append(f"%{search}%")
        
        # 查询总数
        count_sql = f"SELECT COUNT(*) as total FROM products {where_clause}"
        count_result = db_manager.execute_query(count_sql, params)
        total = count_result[0]['total'] if count_result else 0
        
        # 查询数据
        data_sql = f'''
            SELECT * FROM products {where_clause}
            ORDER BY create_time DESC
            LIMIT ? OFFSET ?
        '''
        params.extend([per_page, offset])
        products_data = db_manager.execute_query(data_sql, params)
        
        # 转换为Product对象
        products = []
        for data in products_data:
            product = cls(**data)
            products.append(product)
        
        return {
            'products': products,
            'total': total,
            'page': page,
            'per_page': per_page,
            'total_pages': (total + per_page - 1) // per_page
        }
    
    def delete(self):
        """删除商品"""
        if self.id:
            sql = "DELETE FROM products WHERE id = ?"
            return db_manager.execute_update(sql, (self.id,))
        return False
    
    def to_dict(self):
        """转换为字典格式"""
        return {
            'id': self.id,
            'name': self.name,
            'price': float(self.price) if self.price else None,
            'quantity': self.quantity,
            'spec': self.spec,
            'image_path': self.image_path,
            'create_time': str(self.create_time) if self.create_time else None
        }
