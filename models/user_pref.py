# -*- coding: utf-8 -*-
"""
用户个性化配置（列设置等）
"""

from models.database import db_manager


class UserPreference:
    """简单的键值偏好存储，按用户ID & 键存JSON字符串"""

    @classmethod
    def create_table(cls):
        sql = '''
            CREATE TABLE IF NOT EXISTS user_preferences (
                user_id INTEGER NOT NULL,
                pref_key TEXT NOT NULL,
                pref_value TEXT,
                PRIMARY KEY (user_id, pref_key)
            )
        '''
        db_manager.execute_update(sql)
        return True

    @classmethod
    def get_pref(cls, user_id: int, key: str):
        rows = db_manager.execute_query('SELECT pref_value FROM user_preferences WHERE user_id=? AND pref_key=?', (user_id, key))
        if rows:
            return rows[0].get('pref_value')
        return None

    @classmethod
    def set_pref(cls, user_id: int, key: str, value: str):
        exists = db_manager.execute_query('SELECT 1 FROM user_preferences WHERE user_id=? AND pref_key=?', (user_id, key))
        if exists:
            return db_manager.execute_update('UPDATE user_preferences SET pref_value=? WHERE user_id=? AND pref_key=?', (value, user_id, key))
        else:
            return db_manager.execute_update('INSERT INTO user_preferences (user_id, pref_key, pref_value) VALUES (?, ?, ?)', (user_id, key, value))


