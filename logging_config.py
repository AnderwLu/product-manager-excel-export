#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日志配置文件
"""

import logging
import os
from datetime import datetime

def setup_logging():
    """设置日志配置"""
    
    # 创建logs目录
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    # 生成日志文件名（按日期）
    log_date = datetime.now().strftime('%Y%m%d')
    log_file = f'logs/app_{log_date}.log'
    
    # 配置根日志记录器
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info(f"日志系统初始化完成，日志文件: {log_file}")
    
    return logger

def get_logger(name):
    """获取指定名称的日志记录器"""
    return logging.getLogger(name)
