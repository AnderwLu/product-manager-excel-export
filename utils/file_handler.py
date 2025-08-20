# -*- coding: utf-8 -*-
"""
文件处理工具类
"""

import os
import uuid
from werkzeug.utils import secure_filename
from PIL import Image
from flask import current_app

class FileHandler:
    """文件处理类"""
    
    def __init__(self, upload_folder='uploads'):
        self.upload_folder = upload_folder
        self.allowed_extensions = {'png', 'jpg', 'jpeg'}
        os.makedirs(self.upload_folder, exist_ok=True)
    
    def allowed_file(self, filename):
        """检查文件扩展名是否允许"""
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in self.allowed_extensions
    
    def upload_image(self, file):
        """上传图片文件"""
        try:
            if not file or not file.filename:
                return {
                    'success': False,
                    'message': '没有选择文件'
                }
            
            if not self.allowed_file(file.filename):
                return {
                    'success': False,
                    'message': '不支持的文件格式，只支持PNG、JPG、JPEG'
                }
            
            # 生成唯一文件名
            filename = secure_filename(file.filename)
            file_ext = filename.rsplit('.', 1)[1].lower()
            unique_filename = f"{uuid.uuid4().hex}.{file_ext}"
            file_path = os.path.join(self.upload_folder, unique_filename)
            
            # 保存原文件
            file.save(file_path)
            
            # 创建缩略图
            try:
                with Image.open(file_path) as img:
                    img.thumbnail((200, 200))
                    thumb_path = os.path.join(self.upload_folder, f"thumb_{unique_filename}")
                    img.save(thumb_path)
            except Exception as e:
                print(f"创建缩略图失败: {e}")
                # 缩略图创建失败不影响主文件上传
            
            return {
                'success': True,
                'filename': unique_filename,
                'message': '文件上传成功'
            }
            
        except Exception as e:
            return {
                'success': False,
                'message': f'文件上传失败: {str(e)}'
            }
    
    def delete_image(self, filename):
        """删除图片文件"""
        try:
            if not filename:
                return True
            
            # 删除原图
            file_path = os.path.join(self.upload_folder, filename)
            if os.path.exists(file_path):
                os.remove(file_path)
            
            # 删除缩略图
            thumb_path = os.path.join(self.upload_folder, f"thumb_{filename}")
            if os.path.exists(thumb_path):
                os.remove(thumb_path)
            
            return True
        except Exception as e:
            print(f"删除文件失败: {e}")
            return False
    
    def get_image_path(self, filename):
        """获取图片完整路径"""
        return os.path.join(self.upload_folder, filename)
    
    def get_thumb_path(self, filename):
        """获取缩略图完整路径"""
        return os.path.join(self.upload_folder, f"thumb_{filename}")
    
    def file_exists(self, filename):
        """检查文件是否存在"""
        if not filename:
            return False
        file_path = os.path.join(self.upload_folder, filename)
        return os.path.exists(file_path)
