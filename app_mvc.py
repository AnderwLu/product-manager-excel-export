#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商品信息管理系统主应用文件 - MVC架构版本
"""

from flask import Flask, render_template, send_from_directory
from controllers.product_controller import product_bp
from models.product import Product
import os

def create_app():
    app = Flask(__name__)
    
    # 注册蓝图
    app.register_blueprint(product_bp, url_prefix='/product')
    
    @app.route('/')
    def index():
        return render_template('index.html')
    
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
        print("✅ 数据库表初始化成功")
    except Exception as e:
        print(f"❌ 数据库表初始化失败: {e}")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
