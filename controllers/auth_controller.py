#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
认证与用户管理控制器
"""

from flask import Blueprint, request, jsonify, session, render_template, redirect, url_for
from models.user import User
from logging_config import get_logger

logger = get_logger(__name__)

auth_bp = Blueprint('auth', __name__)


def login_required(view_func):
    def wrapper(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('auth.login_page'))
        return view_func(*args, **kwargs)
    wrapper.__name__ = view_func.__name__
    return wrapper


def admin_required(view_func):
    def wrapper(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('auth.login_page'))
        if session.get('is_admin') != 1:
            return jsonify({'success': False, 'message': '需要管理员权限'}), 403
        return view_func(*args, **kwargs)
    wrapper.__name__ = view_func.__name__
    return wrapper


@auth_bp.route('/login', methods=['GET'])
def login_page():
    return render_template('login.html')


@auth_bp.route('/login', methods=['POST'])
def login():
    try:
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.find_by_username(username)
        if not user or not user.verify_password(password):
            return jsonify({'success': False, 'message': '用户名或密码错误'})
        session['user_id'] = user.id
        session['username'] = user.username
        session['real_name'] = user.real_name
        session['is_admin'] = int(user.is_admin)
        return jsonify({'success': True, 'message': '登录成功'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'登录失败: {str(e)}'})


@auth_bp.route('/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success': True, 'message': '已退出'})


@auth_bp.route('/users', methods=['GET'])
@login_required
def users_page():
    return render_template('users.html')


@auth_bp.route('/api/users', methods=['GET'])
@admin_required
def list_users():
    users = [u.to_dict() for u in User.list_users()]
    return jsonify({'success': True, 'data': users})


@auth_bp.route('/api/users', methods=['POST'])
@admin_required
def create_user():
    data = request.get_json() or {}
    username = data.get('username')
    password = data.get('password')
    real_name = data.get('real_name')
    is_admin = bool(data.get('is_admin'))
    result = User.create_user(username, password, real_name, is_admin)
    return jsonify(result)


@auth_bp.route('/api/users/<int:user_id>', methods=['DELETE'])
@admin_required
def delete_user(user_id):
    result = User.delete_user(user_id)
    return jsonify(result)


@auth_bp.route('/api/users/<int:user_id>/password', methods=['POST'])
@admin_required
def reset_user_password(user_id):
    data = request.get_json() or {}
    new_password = data.get('password')
    result = User.reset_password(user_id, new_password)
    return jsonify(result)


