# -*- coding: utf-8 -*-
"""
应用配置文件
"""

import os

class Config:
    """基础配置类"""
    SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key-here')
    DATABASE_PATH = os.getenv('DATABASE_PATH', 'products.db')
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 5 * 1024 * 1024  # 5MB
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

class DevelopmentConfig(Config):
    """开发环境配置"""
    DEBUG = True
    HOST = '0.0.0.0'
    PORT = 5001
    USE_RELOADER = False

class ProductionConfig(Config):
    """生产环境配置"""
    DEBUG = False
    HOST = '0.0.0.0'
    PORT = 5000

# 配置字典
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}
