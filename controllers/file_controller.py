# -*- coding: utf-8 -*-
"""
文件控制器
"""

import os
from flask import send_file
from utils.file_handler import FileHandler

class FileController:
    """文件控制器类"""
    
    def __init__(self):
        self.file_handler = FileHandler()
    
    def serve_image(self, filename):
        """图片访问接口"""
        try:
            file_path = self.file_handler.get_image_path(filename)
            if os.path.exists(file_path):
                return send_file(file_path)
            else:
                return "文件不存在", 404
        except Exception as e:
            return f"文件访问失败: {str(e)}", 500
    
    def serve_thumbnail(self, filename):
        """缩略图访问接口"""
        try:
            thumb_path = self.file_handler.get_thumb_path(filename)
            if os.path.exists(thumb_path):
                return send_file(thumb_path)
            else:
                # 如果缩略图不存在，返回原图
                return self.serve_image(filename)
        except Exception as e:
            return f"缩略图访问失败: {str(e)}", 500
