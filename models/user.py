# -*- coding: utf-8 -*-
"""
用户数据模型（基于现有SQLite管理器）
"""

from datetime import datetime
from models.database import db_manager
from werkzeug.security import generate_password_hash, check_password_hash


class User:
    """用户模型类"""

    def __init__(self, id=None, username=None, password_hash=None, is_admin=0, real_name=None, create_time=None):
        self.id = id
        self.username = username
        self.password_hash = password_hash
        self.is_admin = int(is_admin) if is_admin is not None else 0
        self.real_name = real_name
        self.create_time = create_time or datetime.now()

    @classmethod
    def create_table(cls):
        sql = '''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                is_admin INTEGER DEFAULT 0,
                real_name TEXT,
                create_time TEXT DEFAULT (datetime('now'))
            )
        '''
        db_manager.execute_update(sql)
        # 兼容旧库：若缺少 real_name 列则新增
        try:
            cols = db_manager.execute_query("PRAGMA table_info(users)")
            col_names = [c.get('name') for c in cols]
            if 'real_name' not in col_names:
                db_manager.execute_update("ALTER TABLE users ADD COLUMN real_name TEXT")
        except Exception:
            pass
        return True

    @classmethod
    def ensure_admin(cls, username='admin', password='admin123', real_name='管理员'):
        """若不存在admin用户，则创建一个默认管理员。"""
        existing = cls.find_by_username(username)
        if existing:
            return existing
        password_hash = generate_password_hash(password)
        sql = 'INSERT INTO users (username, password_hash, is_admin, real_name) VALUES (?, ?, ?, ?)' 
        user_id = db_manager.execute_insert(sql, (username, password_hash, 1, real_name))
        return cls.find_by_id(user_id)

    @classmethod
    def find_by_id(cls, user_id):
        sql = 'SELECT * FROM users WHERE id = ?'
        result = db_manager.execute_query(sql, (user_id,))
        if result:
            return cls(**result[0])
        return None

    @classmethod
    def find_by_username(cls, username):
        sql = 'SELECT * FROM users WHERE username = ?'
        result = db_manager.execute_query(sql, (username,))
        if result:
            return cls(**result[0])
        return None

    @classmethod
    def list_users(cls):
        sql = 'SELECT * FROM users ORDER BY create_time DESC'
        rows = db_manager.execute_query(sql)
        return [cls(**r) for r in rows]

    @classmethod
    def create_user(cls, username, password, real_name, is_admin=False):
        if not username or not password or not real_name:
            return {'success': False, 'message': '用户名、密码、姓名均不能为空'}
        if cls.find_by_username(username):
            return {'success': False, 'message': '用户名已存在'}
        password_hash = generate_password_hash(password)
        user_id = db_manager.execute_insert(
            'INSERT INTO users (username, password_hash, is_admin, real_name) VALUES (?, ?, ?, ?)',
            (username, password_hash, 1 if is_admin else 0, real_name)
        )
        return {'success': True, 'message': '用户创建成功', 'user_id': user_id}

    @classmethod
    def delete_user(cls, user_id):
        # 不允许删除自身admin默认账号（若只有一个admin）
        sql_admin_count = 'SELECT COUNT(*) as c FROM users WHERE is_admin = 1'
        count = db_manager.execute_query(sql_admin_count)[0]['c']
        target = cls.find_by_id(user_id)
        if target and target.is_admin == 1 and count <= 1:
            return {'success': False, 'message': '至少保留一个管理员'}
        rowcount = db_manager.execute_update('DELETE FROM users WHERE id = ?', (user_id,))
        return {'success': True, 'message': '删除成功'} if rowcount else {'success': False, 'message': '用户不存在'}

    @classmethod
    def reset_password(cls, user_id, new_password):
        if not new_password:
            return {'success': False, 'message': '新密码不能为空'}
        password_hash = generate_password_hash(new_password)
        rowcount = db_manager.execute_update('UPDATE users SET password_hash=? WHERE id=?', (password_hash, user_id))
        return {'success': True, 'message': '密码已更新'} if rowcount else {'success': False, 'message': '用户不存在'}

    def verify_password(self, password):
        return check_password_hash(self.password_hash, password)

    def to_dict(self):
        return {
            'id': self.id,
            'username': self.username,
            'is_admin': int(self.is_admin),
            'real_name': self.real_name,
            'create_time': str(self.create_time) if self.create_time else None
        }


