# -*- coding: utf-8 -*-
"""
数据库连接和配置管理 - SQLite版本
"""

import sqlite3
import os
from datetime import datetime

class DatabaseConfig:
    """数据库配置类"""
    
    def __init__(self):
        self.database_path = os.getenv('DATABASE_PATH', 'products.db')
    
    def get_database_path(self):
        """获取数据库文件路径"""
        return self.database_path

class DatabaseManager:
    """数据库管理器"""
    
    def __init__(self):
        self.config = DatabaseConfig()
        self.db_path = self.config.get_database_path()
    
    def get_connection(self):
        """获取数据库连接"""
        return sqlite3.connect(self.db_path)
    
    def execute_query(self, sql, params=None):
        """执行查询语句"""
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(sql, params or ())
            columns = [description[0] for description in cursor.description]
            rows = cursor.fetchall()
            
            # 转换为字典格式，兼容原来的代码
            result = []
            for row in rows:
                row_dict = {}
                for i, column in enumerate(columns):
                    row_dict[column] = row[i]
                result.append(row_dict)
            
            return result
        finally:
            connection.close()
    
    def execute_update(self, sql, params=None):
        """执行更新语句"""
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(sql, params or ())
            connection.commit()
            return cursor.rowcount
        finally:
            connection.close()
    
    def execute_insert(self, sql, params=None):
        """执行插入语句，返回插入的ID"""
        connection = self.get_connection()
        try:
            cursor = connection.cursor()
            cursor.execute(sql, params or ())
            connection.commit()
            return cursor.lastrowid
        finally:
            connection.close()

# 全局数据库管理器实例
db_manager = DatabaseManager()

def get_db_connection():
    """获取数据库连接的便捷函数"""
    return db_manager.get_connection()
