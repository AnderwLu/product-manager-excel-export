#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品信息管理系统主应用文件 - MVC架构版本
"""

from flask import Flask, render_template, send_from_directory, redirect, url_for, session, request
from controllers.product_controller import product_bp
from controllers.auth_controller import auth_bp
from models.product import Product
from models.user import User
from logging_config import setup_logging
import os

# 初始化日志系统
logger = setup_logging()

def create_app():
    app = Flask(__name__)
    # 用于会话管理
    app.secret_key = os.getenv('SECRET_KEY', 'dev-secret-key')
    
    # 注册蓝图
    app.register_blueprint(product_bp, url_prefix='/product')
    app.register_blueprint(auth_bp, url_prefix='/auth')

    @app.before_request
    def require_login():
        path = request.path or ''
        # 放行登录与静态资源
        allowed_prefixes = ['/auth/login', '/static', '/favicon']
        if any(path.startswith(p) for p in allowed_prefixes):
            return None
        # 其余请求需要登录
        if not session.get('user_id'):
            return redirect(url_for('auth.login_page'))
    
    @app.route('/')
    def index():
        if not session.get('user_id'):
            return redirect(url_for('auth.login_page'))
        return redirect(url_for('search_page'))

    @app.route('/search')
    def search_page():
        return render_template('search_export.html')

    @app.route('/search-edit')
    def search_edit_page():
        return render_template('search_edit.html')

    @app.route('/entry')
    def entry_page():
        return render_template('entry.html')
    
    @app.route('/uploads/<filename>')
    def uploaded_file(filename):
        """提供原图文件访问"""
        return send_from_directory('uploads', filename)
    
    @app.route('/uploads/thumb_<filename>')
    def thumbnail_file(filename):
        """提供缩略图文件访问"""
        return send_from_directory('uploads', f"thumb_{filename}")
    
    return app

app = create_app()

# 应用启动时初始化数据库表
with app.app_context():
    try:
        Product.create_table()
        User.create_table()
        # 确保存在admin账号（用户名: admin, 密码: admin, 姓名: admin）
        User.ensure_admin(username='admin', password='admin', real_name='admin')
        logger.info("数据库表初始化完成")
    except Exception as e:
        logger.error(f"数据库表初始化失败: {str(e)}")

if __name__ == '__main__':
    logger.info("应用启动中...")
    app.run(debug=True, host='0.0.0.0', port=5001)
