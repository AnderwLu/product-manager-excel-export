# -*- coding: utf-8 -*-
"""
商品数据模型
"""

from datetime import datetime
from models.database import db_manager

class Product:
    """商品模型类"""
    
    def __init__(self, id=None, name=None, price=None, quantity=None, 
                 spec=None, image_path=None, create_time=None,
                 # 新增字段（保持为可选，默认值不影响旧流程）
                 doc_date=None,                 # 单据日期
                 customer_name=None,            # 客户名称
                 product_desc=None,             # 品名规格
                 unit=None,                     # 单位（原表：规格）
                 unit_price=None,               # 单价（原表：价格）
                 unit_discount_rate=None,       # 单价折扣率(%)
                 unit_price_discounted=None,    # 折后单价
                 amount=None,                   # 金额
                 remark=None,                   # 备注
                 freight=None,                  # 运费
                 order_discount_rate=None,      # 整单折扣率(%)
                 amount_discounted=None,        # 折后金额
                 receivable=None,               # 应收款
                 payment_current=None,          # 本次收款
                 paid_total=None,               # 已收款
                 balance=None,                  # 尾款
                 settlement_account=None,       # 结算账户
                 description=None,              # 说明
                 salesperson=None,              # 营业员
                 update_time=None               # 修改时间
                 ):
        self.id = id
        self.name = name
        self.price = price
        self.quantity = quantity
        self.spec = spec
        self.image_path = image_path
        self.create_time = create_time or datetime.now()

        # 新增字段
        self.doc_date = doc_date
        self.customer_name = customer_name
        self.product_desc = product_desc
        self.unit = unit
        self.unit_price = unit_price
        self.unit_discount_rate = unit_discount_rate
        self.unit_price_discounted = unit_price_discounted
        self.amount = amount
        self.remark = remark
        self.freight = freight
        self.order_discount_rate = order_discount_rate
        self.amount_discounted = amount_discounted
        self.receivable = receivable
        self.payment_current = payment_current
        self.paid_total = paid_total
        self.balance = balance
        self.settlement_account = settlement_account
        self.description = description
        self.salesperson = salesperson
        self.update_time = update_time
    
    @classmethod
    def create_table(cls):
        """创建/升级商品表（仅新增字段，不改变既有写入逻辑）"""
        sql = '''
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                price REAL NOT NULL,
                quantity INTEGER NOT NULL,
                spec TEXT,
                image_path TEXT,
                create_time TEXT DEFAULT (datetime('now','+8 hours'))
            )
        '''
        db_manager.execute_update(sql)
        # 自动补齐新增列
        cls._ensure_columns()
        return True

    @classmethod
    def _ensure_columns(cls):
        """为已存在表补齐缺失列（只做 ADD COLUMN）"""
        # 期望列: 列名 -> SQL 片段
        expected = {
            'doc_date': "TEXT DEFAULT (date('now','+8 hours'))",
            'customer_name': "TEXT",
            'product_desc': "TEXT",
            'unit': "TEXT",
            'unit_price': "REAL",
            'unit_discount_rate': "REAL DEFAULT 100",
            'unit_price_discounted': "REAL",
            'amount': "REAL",
            'remark': "TEXT",
            'freight': "REAL DEFAULT 0",
            'order_discount_rate': "REAL DEFAULT 100",
            'amount_discounted': "REAL",
            'receivable': "REAL",
            'payment_current': "REAL DEFAULT 0",
            'paid_total': "REAL DEFAULT 0",
            'balance': "REAL",
            'settlement_account': "TEXT",
            'description': "TEXT",
            'salesperson': "TEXT",
            'update_time': "TEXT DEFAULT (datetime('now','+8 hours'))"
        }
        cols = db_manager.execute_query("PRAGMA table_info(products)")
        have = {c['name'] for c in cols} if cols else set()
        for col, ddl in expected.items():
            if col not in have:
                db_manager.execute_update(f"ALTER TABLE products ADD COLUMN {col} {ddl}")

    def save(self):
        """保存商品到数据库"""
        if self.id:
            # 更新（保持原有字段集合，不动业务逻辑）
            sql = '''
                UPDATE products 
                SET name=?, price=?, quantity=?, spec=?, image_path=?,
                    doc_date=?, customer_name=?, product_desc=?, unit=?, unit_price=?,
                    remark=?, settlement_account=?, description=?, salesperson=?, freight=?, paid_total=?,
                    update_time=datetime('now','+8 hours')
                WHERE id=?
            '''
            params = (
                self.name,
                self.price,
                self.quantity,
                self.spec,
                self.image_path,
                self.doc_date,
                (self.customer_name or self.name),
                self.product_desc,
                (self.unit or self.spec),
                (self.unit_price if self.unit_price is not None else self.price),
                self.remark,
                self.settlement_account,
                self.description,
                self.salesperson,
                self.freight,
                self.paid_total,
                self.id
            )
            return db_manager.execute_update(sql, params)
        else:
            # 插入（保持原有字段集合，不动业务逻辑）
            sql = '''
                INSERT INTO products (
                    name, price, quantity, spec, image_path,
                    doc_date, customer_name, product_desc, unit, unit_price,
                    remark, settlement_account, description, salesperson, freight, paid_total,
                    update_time
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now','+8 hours'))
            '''
            params = (
                self.name,
                self.price,
                self.quantity,
                self.spec,
                self.image_path,
                self.doc_date,
                (self.customer_name or self.name),
                self.product_desc,
                (self.unit or self.spec),
                (self.unit_price if self.unit_price is not None else self.price),
                self.remark,
                self.settlement_account,
                self.description,
                self.salesperson,
                self.freight,
                self.paid_total
            )
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
    def find_all(cls, page=1, per_page=10, search=None, product_desc=None, salesperson=None, date_start=None, date_end=None):
        """查找所有商品，支持分页和搜索"""
        offset = (page - 1) * per_page
        
        where_parts = []
        params = []
        # 客户名称模糊（历史保存在 name 列）
        if search:
            where_parts.append("name LIKE ?")
            params.append(f"%{search}%")
        # 品名规格模糊
        if product_desc:
            where_parts.append("product_desc LIKE ?")
            params.append(f"%{product_desc}%")
        # 营业员模糊
        if salesperson:
            where_parts.append("salesperson LIKE ?")
            params.append(f"%{salesperson}%")
        # 单据日期范围
        coalesce_date = "COALESCE(doc_date, substr(create_time,1,10))"
        if date_start:
            where_parts.append(f"{coalesce_date} >= ?")
            params.append(date_start)
        if date_end:
            where_parts.append(f"{coalesce_date} <= ?")
            params.append(date_end)

        where_clause = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
        
        count_sql = f"SELECT COUNT(*) as total FROM products {where_clause}"
        count_result = db_manager.execute_query(count_sql, params)
        total = count_result[0]['total'] if count_result else 0
        
        data_sql = f'''
            SELECT * FROM products {where_clause}
            ORDER BY create_time DESC
            LIMIT ? OFFSET ?
        '''
        params.extend([per_page, offset])
        products_data = db_manager.execute_query(data_sql, params)
        
        products = [cls(**data) for data in products_data]
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
        """转换为字典格式（包含新增字段）"""
        return {
            'id': self.id,
            'name': self.name,
            'price': float(self.price) if self.price is not None else None,
            'quantity': self.quantity,
            'spec': self.spec,
            'image_path': self.image_path,
            'create_time': str(self.create_time) if self.create_time else None,

            'doc_date': self.doc_date,
            'customer_name': self.customer_name,
            'product_desc': self.product_desc,
            'unit': self.unit,
            'unit_price': self.unit_price,
            'unit_discount_rate': self.unit_discount_rate,
            'unit_price_discounted': self.unit_price_discounted,
            'amount': self.amount,
            'remark': self.remark,
            'freight': self.freight,
            'order_discount_rate': self.order_discount_rate,
            'amount_discounted': self.amount_discounted,
            'receivable': self.receivable,
            'payment_current': self.payment_current,
            'paid_total': self.paid_total,
            'balance': self.balance,
            'settlement_account': self.settlement_account,
            'description': self.description,
            'salesperson': self.salesperson,
            'update_time': self.update_time
        }
