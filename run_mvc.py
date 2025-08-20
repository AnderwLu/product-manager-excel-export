#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
启动MVC架构的商品管理系统
"""

from app_mvc import app

if __name__ == '__main__':
    app.run(
        debug=True,
        host='0.0.0.0',
        port=5001,
        use_reloader=False
    )
