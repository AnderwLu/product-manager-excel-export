# 商品管理系统 - MVC架构版本

## 项目结构

```
product_manager/
├── app_mvc.py              # 主应用文件（MVC架构）
├── run_mvc.py              # 启动脚本
├── config.py                # 配置文件
├── models/                  # 数据模型层
│   ├── __init__.py
│   ├── database.py         # 数据库连接管理
│   └── product.py          # 商品数据模型
├── services/                # 业务服务层
│   ├── __init__.py
│   ├── product_service.py  # 商品业务服务
│   └── export_service.py   # 导出服务
├── controllers/             # 控制器层
│   ├── __init__.py
│   ├── product_controller.py # 商品控制器
│   └── file_controller.py   # 文件控制器
├── utils/                   # 工具函数层
│   ├── __init__.py
│   ├── file_handler.py     # 文件处理工具
│   └── validator.py        # 数据验证工具
├── templates/               # 前端模板
├── uploads/                 # 上传文件目录
└── requirements.txt         # 依赖包
```

## 架构说明

### 1. Models（数据模型层）
- **database.py**: 数据库连接管理，提供统一的数据库操作接口
- **product.py**: 商品数据模型，定义商品的数据结构和数据库操作方法

### 2. Services（业务服务层）
- **product_service.py**: 商品业务逻辑，处理商品的增删改查等业务操作
- **export_service.py**: 导出服务，处理Excel导出等业务需求

### 3. Controllers（控制器层）
- **product_controller.py**: 商品控制器，处理HTTP请求，调用服务层处理业务逻辑
- **file_controller.py**: 文件控制器，处理文件上传和访问

### 4. Utils（工具函数层）
- **file_handler.py**: 文件处理工具，处理图片上传、缩略图生成等
- **validator.py**: 数据验证工具，验证用户输入的数据

### 5. Config（配置管理）
- **config.py**: 应用配置文件，支持开发和生产环境配置

## 优势

1. **代码分离**: 每个模块职责明确，便于维护和扩展
2. **可读性强**: 代码结构清晰，易于理解和修改
3. **可维护性**: 模块化设计，修改某个功能不会影响其他模块
4. **可扩展性**: 新增功能只需在相应层添加代码，不影响现有架构
5. **可测试性**: 各层可以独立测试，提高代码质量

## 启动方式

### 方式1：直接运行
```bash
python app_mvc.py
```

### 方式2：使用启动脚本
```bash
python run_mvc.py
```

### 方式3：设置环境变量
```bash
export FLASK_ENV=production  # 生产环境
python app_mvc.py
```

## 环境要求

- Python 3.6+
- Flask
- PyMySQL
- Pillow (PIL)
- openpyxl

## 数据库配置

可以通过环境变量配置数据库连接：

```bash
export MYSQL_HOST=your_host
export MYSQL_USER=your_user
export MYSQL_PASSWORD=your_password
export MYSQL_DB=your_database
```

## 功能特性

- 商品增删改查
- 图片上传和缩略图生成
- 分页查询和搜索
- Excel导出
- 数据验证
- 错误处理
- 文件管理

## 与原版本的区别

1. **架构清晰**: 采用标准的MVC架构，代码结构更加清晰
2. **职责分离**: 每个模块都有明确的职责，便于维护
3. **代码复用**: 通用功能提取到工具类，避免重复代码
4. **配置管理**: 统一的配置管理，支持多环境部署
5. **错误处理**: 统一的错误处理机制，提高系统稳定性
