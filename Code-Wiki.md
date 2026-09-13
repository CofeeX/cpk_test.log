# 产品功能测试分析 · Code Wiki

> 本文档为 `traecode-测试分析001` 仓库的结构化代码 Wiki，覆盖项目整体架构、各模块职责、关键类与函数、依赖关系、运行方式与核心算法规则。
>
> - 仓库根目录：`D:\Code\CodePython\traecode-测试分析001`
> - 原始需求/设计文档：[测试分析001.md](file:///D:/Code/CodePython/traecode-测试分析001/测试分析001.md)
> - 主入口：[app.py](file:///D:/Code/CodePython/traecode-测试分析001/app.py)

---

## 目录

- [1. 项目概述](#1-项目概述)
- [2. 技术栈与依赖](#2-技术栈与依赖)
- [3. 整体架构](#3-整体架构)
- [4. 工程目录结构](#4-工程目录结构)
- [5. 模块职责详解](#5-模块职责详解)
  - [5.1 app.py — Flask 主应用](#51-apppy--flask-主应用)
  - [5.2 config.py — 应用配置](#52-configpy--应用配置)
  - [5.3 models.py — 数据库模型](#53-modelspy--数据库模型)
  - [5.4 services/excel_parser.py — Excel 解析器](#54-servicesexcel_parserpy--excel-解析器)
  - [5.5 services/analyzer.py — 统计分析器](#55-servicesanalyzerpy--统计分析器)
  - [5.6 services/config_manager.py — 配置管理器](#56-servicesconfig_managerpy--配置管理器)
  - [5.7 templates/ — 前端模板](#57-templates--前端模板)
  - [5.8 static/ — 静态资源](#58-static--静态资源)
  - [5.9 configs/ — 产品配置模板](#59-configs--产品配置模板)
  - [5.10 辅助脚本与打包文件](#510-辅助脚本与打包文件)
- [6. 核心数据流与处理流程](#6-核心数据流与处理流程)
- [7. REST API 接口一览](#7-rest-api-接口一览)
- [8. 关键类与函数](#8-关键类与函数)
- [9. 数据模型](#9-数据模型)
- [10. 产品配置 JSON Schema](#10-产品配置-json-schema)
- [11. 核心算法规则](#11-核心算法规则)
- [12. 运行与部署](#12-运行与部署)
- [13. 关键设计要点](#13-关键设计要点)

---

## 1. 项目概述

**产品功能测试分析** 是一个基于 Flask 的 Web 应用，用于导入 Excel 格式的产品测试记录，按测试项目自动生成曲线图与统计分析，便于工程人员快速评估产品质量。

### 核心功能

| 编号 | 功能 | 说明 |
|------|------|------|
| 功能1 | Web 应用 | 可部署到服务器端，浏览器访问 |
| 功能2 | Excel 输入 | 支持 `.xlsx` / `.xls` 测试记录文件 |
| 功能3 | 多产品解析入口 | 不同产品类别可加载不同解析配置；无配置时从文件自动检测 |
| 功能4 | 曲线 + 统计 | 每个测试项目生成 ECharts 曲线图，含均值/标准差/最值/分位数/合格率 |
| 功能5 | 测试项目筛选 | 左侧树形筛选：一级=图表，二级=图表内曲线，支持勾选/折叠/定位 |
| 功能6 | 配置管理 | 按产品类别保存 JSON 配置，后续同款产品可直接调用 |

### 业务定位

- **目标用户**：硬件/产线测试工程师、品质分析人员
- **典型数据量**：单个 Excel 含数百至数千产品记录（参考 `SC板测试记录01.xlsx` 解析后约 10 万行长格式记录）
- **典型输出**：每文件约 30+ 张图表，支持全屏放大与 PNG 下载

---

## 2. 技术栈与依赖

### 后端

| 包名 | 版本 | 用途 |
|------|------|------|
| Flask | >=3.0.3 | Web 框架 |
| Flask-SQLAlchemy | >=3.1.1 | ORM 封装 |
| SQLAlchemy | >=2.0.30 | 数据库引擎 |
| PyMySQL | >=1.1.0 | MySQL 驱动（可选，生产部署用） |
| pandas | >=2.2.3 | Excel 读取与数据处理 |
| openpyxl | >=3.1.5 | `.xlsx` 解析后端 |
| werkzeug | >=3.0.3 | WSGI 工具库 |

依赖清单见 [requirements.txt](file:///D:/Code/CodePython/traecode-测试分析001/requirements.txt)。

### 前端（CDN 加载，无构建）

| 库 | 版本 | 用途 |
|----|------|------|
| Bootstrap | 5.3.3 | UI 组件与栅格 |
| Bootstrap Icons | 1.11.3 | 图标 |
| ECharts | 5.5.0 | 曲线图渲染 |

### 运行环境

- Python 3.13+（源码运行）
- 操作系统：Windows / Linux / macOS（PyInstaller spec 与 start.bat 针对 Windows）
- 数据库：默认 SQLite（零配置），可切换 MySQL（生产部署）

---

## 3. 整体架构

项目采用经典的 **三层结构**：表现层（Jinja2 + JS）→ 应用层（Flask 路由）→ 服务层（pandas 分析 + JSON 配置 + SQLite 持久化）。

```
┌──────────────────────────────────────────────────────────────┐
│                       浏览器 (前端)                          │
│  Bootstrap 5 + ECharts 5 + main.js (1045 行)                 │
│   - 上传表单 / 配置管理弹窗 / 筛选树 / 图表区 / 放大 Modal   │
└────────────┬─────────────────────────────────────────────────┘
             │ HTTP (REST + 多部分表单)
┌────────────▼─────────────────────────────────────────────────┐
│                   app.py (Flask 主应用)                      │
│   路由层：页面路由 + /api/* REST 端点                        │
│   辅助：_save_df/_load_df (pickle 序列化) / _allowed_file    │
│   会话：Flask session 保存 df_file/filename/product_category │
└────┬──────────┬───────────────┬──────────────┬──────────────┘
     │          │               │              │
┌────▼────┐ ┌───▼────────┐ ┌────▼────────┐ ┌───▼──────────┐
│services.│ │services.   │ │services.     │ │ models.py    │
│excel_   │ │analyzer.py │ │config_       │ │ UploadRecord │
│parser   │ │ Analyzer   │ │manager.py    │ │ (SQLAlchemy) │
│ ExcelP. │ │ 统计/分组  │ │ConfigManager │ └──────┬───────┘
└────┬────┘ └─────┬──────┘ └──────┬───────┘        │
     │            │               │                │
     │       pandas DataFrame     │           SQLite / MySQL
     │            │               │           test_analysis.db
     ▼            ▼               ▼
┌──────────────────────────────────────────────────────────────┐
│  数据存储 (config.py 决定 BASE_DIR / DATA_DIR)              │
│   - uploads/      : 原始 Excel + 序列化 .pkl                 │
│   - configs/      : 产品配置 JSON (只读模板)                 │
│   - {DATA_DIR}/configs/ : 用户保存的新配置                   │
│   - test_analysis.db : 上传记录表                           │
└──────────────────────────────────────────────────────────────┘
```

### 分层职责

| 层 | 职责 | 关键文件 |
|----|------|---------|
| 表现层 | UI 渲染、用户交互、图表绘制 | `templates/*.html`, `static/js/main.js`, `static/css/style.css` |
| 应用层 | 路由分发、参数校验、会话管理、API 序列化 | `app.py` |
| 服务层 | Excel 解析、统计分析、配置 CRUD | `services/*.py` |
| 模型层 | ORM 实体 | `models.py` |
| 配置层 | 路径与模式切换（源码 / frozen） | `config.py` |

---

## 4. 工程目录结构

```
traecode-测试分析001/
├── app.py                          # Flask 主应用（路由 + API）
├── config.py                       # 应用配置（数据目录、上传目录、密钥）
├── models.py                       # 数据库模型（UploadRecord）
├── start.bat                       # Windows 启动脚本
├── requirements.txt                # Python 依赖清单
├── build_exe.spec                  # PyInstaller 打包配置
├── generate_sample.py              # 生成示例 Excel 测试数据
├── _debug_statistics.py            # 规则1~4 统计验证脚本
├── _debug_test.py                  # 解析+分析流程调试脚本
├── 测试分析001.md                  # 原始需求/设计说明
├── 测试数据/                       # 示例数据
│   ├── SC板测试记录01.xlsx
│   ├── W11-25080019-*.xlsx(xls)    # DC 控制板测试记录
│   ├── data/
│   │   ├── test_analysis.db        # SQLite 数据库
│   │   └── uploads/                # 已上传 Excel + .pkl
│   └── 使用方法.docx
├── configs/                        # 产品配置（只读模板）
│   └── SC-POWER板测试记录.json
├── services/                       # 业务服务层
│   ├── __init__.py
│   ├── excel_parser.py             # Excel 解析（含 auto_detect_config）
│   ├── analyzer.py                 # 图表数据生成（按规格自动分组）
│   └── config_manager.py           # 配置管理
├── templates/                      # Jinja2 模板
│   ├── base.html                   # 基础模板（navbar + footer）
│   └── index.html                  # 主页（侧边栏 + 图表区 + 放大 Modal）
├── static/                         # 静态资源
│   ├── css/style.css
│   └── js/main.js                  # 前端逻辑（上传/筛选/渲染/放大）
└── libs/                           # 备用依赖目录（默认为空）
```

---

## 5. 模块职责详解

### 5.1 [app.py](file:///D:/Code/CodePython/traecode-测试分析001/app.py) — Flask 主应用

**职责**：路由分发、会话管理、API 序列化、协调各服务模块。

#### 关键组件

- **Flask 应用实例** `app`：通过 `app.config.from_object(Config)` 载入配置。
- **数据库初始化** `db.init_app(app)` + `with app.app_context(): db.create_all()` 在模块加载时建表。
- **ConfigManager 实例** `config_manager`：在启动时初始化，绑定 `CONFIG_FOLDER`。

#### 辅助函数

| 函数 | 签名 | 职责 |
|------|------|------|
| `_allowed_file` | `(filename: str) -> bool` | 校验文件扩展名 ∈ `{xlsx, xls}` |
| `_save_df` | `(df: pd.DataFrame) -> str` | 将 DataFrame pickle 到 `UPLOAD_FOLDER`，返回文件名（uuid hex） |
| `_load_df` | `() -> pd.DataFrame` | 从 session 中记录的文件名加载 DataFrame，缺失返回空 |

#### 路由表

| 方法 | 路径 | 函数 | 说明 |
|------|------|------|------|
| GET | `/` | `index` | 渲染首页，传入 `configs` 列表 |
| POST | `/api/upload` | `upload` | 上传并解析 Excel，返回测试项目 + 初始图表数据 |
| GET | `/api/chart-data` | `chart_data` | 按 `item` 参数返回单项图表数据 |
| GET | `/api/test-items` | `test_items` | 返回当前数据测试项目列表与每项统计 |
| GET | `/api/configs` | `list_configs` | 列出所有产品配置 |
| GET | `/api/config/<name>` | `get_config` | 获取指定产品配置 |
| POST | `/api/config` | `save_config` | 保存产品配置 |
| DELETE | `/api/config/<name>` | `delete_config` | 删除产品配置 |
| GET | `/api/config/default` | `default_config` | 返回默认配置模板 |
| GET | `/api/sheets` | `get_sheets` | 获取已上传文件的工作表列表 |
| GET | `/api/history` | `history` | 返回最近 50 条上传记录 |

#### 上传流程要点

`/api/upload` 是核心入口，执行以下步骤（见 [app.py:83-208](file:///D:/Code/CodePython/traecode-测试分析001/app.py#L83-L208)）：

1. 校验文件存在性与扩展名。
2. 保存原始 Excel 到 `UPLOAD_FOLDER`。
3. 读取 `product_category` 表单字段；若有则 `config_manager.load()` 加载配置，否则调用 `ExcelParser.auto_detect_config()` 自动检测。
4. 构造 `ExcelParser` 并 `parse(saved_path)`，得到长格式 DataFrame。
5. 提取公共信息：合并 `parser.program_info` + 首条记录的 `fix_work_order/fix_tester/fix_fixture`。
6. `_save_df()` 序列化 DataFrame 到 .pkl，写入 session。
7. 写入 `UploadRecord` 数据库记录。
8. 选择图表合并策略：
   - 若配置含 `merge_items` -> `Analyzer.grouped_chart_data(df, merge_items)`
   - 否则若 `test_items` 为空（白名单全开） -> `Analyzer.auto_grouped_chart_data(df)` 按规格自动分组
   - 否则 -> `Analyzer.all_items_chart_data(df)` 每项一图
9. 统计 `product_count`（按序列号去重）、`analysis_item_count`（有规格限的项目数）、`chart_count`。
10. 返回 JSON：`filename/product_category/total_rows/product_count/analysis_item_count/chart_count/common_info/test_items/charts/record_id`。

---

### 5.2 [config.py](file:///D:/Code/CodePython/traecode-测试分析001/config.py) — 应用配置

**职责**：决定应用资源根目录与可写数据目录，并提供 Flask 配置类 `Config`。

#### 关键函数

| 函数 | 职责 |
|------|------|
| `_get_base_dir()` | 返回资源根目录。frozen 模式下用 `sys._MEIPASS`，否则用脚本所在目录。用于定位只读资源（templates/static/configs）。 |
| `_get_data_dir()` | 返回可写数据目录。优先级：环境变量 `DATA_DIR` > frozen 模式 EXE 同级 `data/` > 脚本所在目录。 |

#### 模块级常量

- `BASE_DIR`：只读资源根。
- `DATA_DIR`：可写数据目录。

#### `Config` 类字段

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `SQLALCHEMY_DATABASE_URI` | `sqlite:///{DATA_DIR}/test_analysis.db` | 数据库 URI；可用环境变量 `DATABASE_URL` 覆盖（MySQL 部署时设置） |
| `SQLALCHEMY_TRACK_MODIFICATIONS` | `False` | 关闭修改追踪以节省内存 |
| `UPLOAD_FOLDER` | `{DATA_DIR}/uploads` | 上传文件目录 |
| `ALLOWED_EXTENSIONS` | `{"xlsx", "xls"}` | 允许的扩展名 |
| `CONFIG_FOLDER` | `{BASE_DIR}/configs` | 产品配置 JSON 目录 |
| `SECRET_KEY` | 环境变量或 `"test-analysis-secret-key-2026"` | Flask 会话密钥 |
| `MAX_CONTENT_LENGTH` | 16 MB | 上传文件大小上限 |

> **设计要点**：通过 `_get_base_dir` 与 `_get_data_dir` 双目录分离，使得 PyInstaller 打包后只读资源（templates/static/configs）从 `sys._MEIPASS` 加载，而可写数据（uploads/db）写到 EXE 同级目录或环境变量指定位置，避免沙箱权限问题。

---

### 5.3 [models.py](file:///D:/Code/CodePython/traecode-测试分析001/models.py) — 数据库模型

**职责**：定义 SQLAlchemy 数据库实例与上传记录实体。

#### 模块级对象

- `db = SQLAlchemy()`：全局 SQLAlchemy 实例，由 `app.py` 调用 `db.init_app(app)` 绑定到 Flask 应用。

#### `UploadRecord` 实体

表名：`upload_records`

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | Integer, PK | 主键 |
| `filename` | String(255), NOT NULL | 原始文件名 |
| `saved_path` | String(512), NOT NULL | 服务器存储路径 |
| `product_category` | String(100) | 产品类别 |
| `total_rows` | Integer, default=0 | 数据总行数 |
| `test_items` | Text | 测试项目列表（JSON 字符串） |
| `created_at` | DateTime, default=utcnow | 创建时间 |

#### 方法

- `to_dict() -> dict`：序列化为字典，`test_items` 反序列化为 list，`created_at` 格式化为 `%Y-%m-%d %H:%M:%S`。

---

### 5.4 [services/excel_parser.py](file:///D:/Code/CodePython/traecode-测试分析001/services/excel_parser.py) — Excel 解析器

**职责**：依据产品配置的多行表头规则读取 Excel，标准化为长格式 DataFrame；并提供无配置场景的自动检测能力。

#### `ExcelParser` 类

##### 构造参数

`ExcelParser(product_config: dict)` —— 从 `product_config["excel_config"]` 读取以下字段：

| 字段 | 默认 | 说明 |
|------|------|------|
| `sheet_name` | `0` | 工作表名或序号 |
| `skip_top_rows` | `0` | 起始跳过的标题行数（SC 板 = 1） |
| `header_rows` | `5` | 测试项目元数据占用行数 |
| `meta_row_mapping` | 见下 | 元数据行索引映射 |
| `fixed_columns_header_row` | `-1` | 固定列名行号（相对跳过行后）；-1 表示不使用 |
| `fixed_columns` | 见下 | 固定列字段映射 |

默认 `meta_row_mapping`：
```python
{"item_name": 0, "item_note": 1, "param_variable": 2, "spec_upper": 3, "spec_lower": 4}
```

默认 `fixed_columns`：
```python
{"sequence": "序列号", "product_model": "产品型号"}
```

构造时还会从 `product_config["test_items"]` 提取 `focus_items` 白名单（非空时仅解析这些项目）。

##### 实例属性

- `self.config`：原始产品配置。
- `self.sheet_name` / `skip_top_rows` / `header_rows` / `meta_row_mapping` / `fixed_cols_header_row` / `fixed_columns`：解析参数。
- `self.focus_items`：白名单列表。
- `self.program_info: dict`：由 `_extract_program_info` 填充的顶部公共信息（如 Program Name）。

##### 核心方法

| 方法 | 签名 | 职责 |
|------|------|------|
| `parse` | `(file_path: str) -> pd.DataFrame` | 解析 Excel，返回长格式 DataFrame（仅数值型测试值） |

`parse` 内部流程（见 [excel_parser.py:65-154](file:///D:/Code/CodePython/traecode-测试分析001/services/excel_parser.py#L65-L154)）：

1. `pd.read_excel(file_path, sheet_name=self.sheet_name, header=None)`。
2. `_extract_program_info(raw)`：从顶部跳过行提取 Program Name 等。
3. 切片跳过 `skip_top_rows`。
4. 取前 `header_rows` 行作为 `header_df`。
5. 确定数据起始行 `data_start_row`（默认 = `header_rows`；若启用 `fixed_cols_header_row` 则取 `max(header_rows, fixed_cols_header_row + 1)`）。
6. 调用 `_find_fixed_col` 定位序列号、产品型号及其他固定列（工单号/测试员/工装编号/结论/时间）索引。
7. `_collect_test_items(header_df, excluded_cols)` 收集测试项目列（跳过固定列），自动组合同名+参数变量为唯一项名 `"项目名 [参数]"`。
8. 逐行解析产品数据：序列号为空则跳过；每个测试项调用 `_try_numeric` 取数值，文本型自动跳过；构造记录并加 `fix_*` 前缀的固定列值。
9. 返回 `pd.DataFrame(records)`。

##### 内部辅助方法

| 方法 | 职责 |
|------|------|
| `_find_fixed_col(field_key, header_df, fixed_cols_row)` | 查找固定列索引；优先匹配 `fixed_cols_row`，回退到 `item_name` 行 |
| `_collect_test_items(header_df, excluded)` | 提取测试项目；同名多列时拼接 `[参数变量]`；应用 `focus_items` 白名单 |
| `_extract_program_info(raw)` | 从顶部跳过行提取 `Key: Value` 或 `Key:` + 右侧值两种写法的公共信息 |
| `_strip(value)` 静态 | 去引号、去空格，处理 `'内容'` 形式 |
| `_cell(header_df, row_idx, col_idx)` 静态 | 安全读取表头单元格 |
| `_has_value(v)` 静态 | 判断单元格非空（含 NaN 处理） |
| `_try_numeric(val)` 静态 | 尝试转 float；文本型/`-`/`*`/`/` 返回 None |
| `_try_spec_numeric(val_str)` 类方法 | 规格上下限转数值；`*`/`-`/`—`/`/` 视为 None（无限） |

##### 静态方法

| 方法 | 职责 |
|------|------|
| `get_sheet_names(file_path)` | 返回 Excel 所有工作表名称 |
| `auto_detect_config(file_path)` | **核心能力**：无配置时从 Excel 自动检测解析配置（见下） |
| `_default_config()` | 返回默认配置模板（与 `ConfigManager.default_config` 一致） |

##### `auto_detect_config` 检测策略

见 [excel_parser.py:338-459](file:///D:/Code/CodePython/traecode-测试分析001/services/excel_parser.py#L338-L459)：

1. **检测 `skip_top_rows`**：扫描前 5 行，查找包含 `program name` / `程序名` / `程序名称` 的行，将其行号+1 作为跳过数。
2. **检测 `meta_row_mapping`**：在跳过后的 10 行内扫描前 8 列，匹配标签关键词集合：
   - `spec_upper`: `规格上限/上限/spec_upper/usl/upper` 等
   - `spec_lower`: `规格下限/下限/spec_lower/lsl/lower` 等
   - `item_name`: `项目名称/测试项目/item_name` 等
   - `param_variable`: `参数变量/参数/param/variable` 等
   - `item_note`: `备注/说明/note` 等
   未匹配的项用默认值补全（item_name=0, item_note=1, ...）。
3. **检测 `fixed_columns_header_row` 与 `fixed_columns`**：从最后一行元数据起向下扫描 6 行，匹配固定列关键词（序列号/工单号/时间等），找到 >=2 个即停止。
4. 构造 `excel_config` 与产品配置 dict，`category_name="自动检测配置"`，`test_items=[]`，`merge_items=[]`。

---

### 5.5 [services/analyzer.py](file:///D:/Code/CodePython/traecode-测试分析001/services/analyzer.py) — 统计分析器

**职责**：使用 Pandas 对长格式测试数据进行统计分析与图表数据生成，输出含规格上下限的 JSON，供前端 ECharts 绘制。

#### 输入 DataFrame 标准列

```
sequence | product_model | item_name | item_value(数值)
| item_note | param_variable | spec_upper | spec_lower
| fix_work_order | fix_tester | fix_fixture | fix_result | fix_time
```

#### `Analyzer` 类（全静态方法）

| 方法 | 签名 | 职责 |
|------|------|------|
| `get_test_items` | `(df) -> list` | 返回所有数值型测试项目列表（按首次出现顺序，去重保序） |
| `get_item_metadata` | `(df, item_name) -> dict` | 提取项目元数据：item_note/param_variable/spec_upper/spec_lower |
| `filter_by_item` | `(df, test_item=None) -> pd.DataFrame` | 按测试项目筛选数据 |
| `statistics` | `(df) -> dict` | 计算 `item_value` 列统计：count/mean/std/min/max/median/q1/q3/pass_rate |
| `chart_data` | `(df, test_item=None) -> dict` | 生成单项 ECharts 数据：item/meta/x_axis/values/stats/spec_upper/spec_lower |
| `all_items_chart_data` | `(df) -> list` | 为每个数值型项目生成图表（每项一图） |
| `grouped_chart_data` | `(df, merge_items) -> list` | 按显式合并组生成图表（多 series） |
| `auto_grouped_chart_data` | `(df) -> list` | **核心算法**：按规则 3+4 自动分组（见下） |

#### `statistics` 输出字段

| 字段 | 计算方式 |
|------|----------|
| `count` | `values.count()` |
| `mean` | `values.mean()`，保留 4 位 |
| `std` | `values.std()`，单点时为 0.0 |
| `min` / `max` | 保留 4 位 |
| `median` | 0.5 分位数 |
| `q1` / `q3` | 0.25 / 0.75 分位数 |
| `pass_rate` | 若有 spec_lower+spec_upper，统计 `(values >= lower) & (values <= upper)` 比例（百分比，保留 2 位） |

#### `chart_data` 返回结构

```python
{
    "item": str,                 # 项目名或 "全部数据"
    "meta": dict,                # 项目元数据
    "x_axis": list[str],         # 优先 fix_time，其次 sequence，兜底序号
    "values": list[float],       # 数值，保留 4 位
    "stats": dict,               # statistics() 输出
    "spec_upper": float | None,
    "spec_lower": float | None,
}
```

#### `grouped_chart_data` 与 `auto_grouped_chart_data` 输出结构

每张图表对象在 `chart_data` 基础上新增 `series` 字段：

```python
{
    "item": str,                 # 组名（如 "4.75 ~ 5.25 (3项)"）
    "meta": dict,                # 组内首个项目元数据
    "x_axis": list[str],         # 基准项（组内首项）的 X 轴
    "values": list[float],       # 兼容字段（首个 series 的值）
    "stats": dict,               # 兼容字段
    "spec_upper": float | None,  # 合并后的规格上限
    "spec_lower": float | None,  # 合并后的规格下限
    "series": [                  # 多条曲线
        {"name": str, "values": list[float], "stats": dict},
        ...
    ],
}
```

#### `auto_grouped_chart_data` 算法

见 [analyzer.py:210-289](file:///D:/Code/CodePython/traecode-测试分析001/services/analyzer.py#L210-L289)：

1. 遍历 `get_test_items(df)`，对每项取 `spec_lower`/`spec_upper`。
2. **规则3 过滤**：若两者均为 None，跳过该项目。
3. **规则4 分组**：以 `(spec_lower, spec_upper)` 元组为 key 分组（None 与 None 视为相同）。
4. 按组内首项在 `all_items` 中的索引排序，保持 Excel 列顺序。
5. 为每组生成图表对象：
   - 组名格式化：双限 -> `"下限 ~ 上限 (N项)"`；仅上限 -> `"<= 上限 (N项)"`；仅下限 -> `">= 下限 (N项)"`。
   - 以首项的 `chart_data` 为基准，其他项追加为 `series` 列表元素。
   - `spec_upper`/`spec_lower` 取组内任一非空值（按设计应一致）。

---

### 5.6 [services/config_manager.py](file:///D:/Code/CodePython/traecode-测试分析001/services/config_manager.py) — 配置管理器

**职责**：以 JSON 文件形式持久化各产品类别的解析配置，提供 CRUD 接口。

#### `ConfigManager` 类

##### 构造

`ConfigManager(config_folder: str)` —— 创建目录（若不存在）并保存路径。

##### 方法

| 方法 | 签名 | 职责 |
|------|------|------|
| `_path` | `(name) -> str` | 根据 `category_name` 生成配置文件路径；`/` 与 `\` 替换为 `_` |
| `list_configs` | `() -> list` | 遍历目录下 `.json` 文件，返回 `[{name, filename, test_items}]` |
| `load` | `(name) -> Optional[dict]` | 读取指定产品配置；不存在返回 None |
| `save` | `(config) -> str` | 保存配置，文件名取 `category_name`；返回保存的 basename |
| `delete` | `(name) -> bool` | 删除配置文件；成功返回 True |
| `default_config` 静态 | `() -> dict` | 返回默认配置模板（与 `ExcelParser._default_config` 一致） |

> 设计上 `_path` 通过替换 `/` 与 `\` 防止路径穿越，但保留了中文等字符，因此配置名可包含中文（如 `SC-POWER板测试记录`）。

---

### 5.7 [templates/](file:///D:/Code/CodePython/traecode-测试分析001/templates) — 前端模板

#### [base.html](file:///D:/Code/CodePython/traecode-测试分析001/templates/base.html)

基础布局模板，提供：

- `<head>`：CDN 引入 Bootstrap 5.3.3、Bootstrap Icons 1.11.3、ECharts 5.5.0，以及本地 `style.css`。
- `<nav>`：navbar，含品牌图标、`#navProgramName` 占位（运行时由 JS 填充 Program Name）、副标题。
- `<main>`：`{% block content %}` 内容占位。
- `<footer>`：版权信息。
- `<script>`：CDN Bootstrap JS、本地 `main.js`，`{% block scripts %}` 额外脚本占位。

#### [index.html](file:///D:/Code/CodePython/traecode-测试分析001/templates/index.html)

主页布局，extends base.html，包含：

- **左侧边栏**（`aside`，固定宽 320px，sticky 卡片）：
  - 上传表单：产品配置管理按钮、产品类别下拉（`#productSelect`，由后端 `configs` 渲染）、文件输入、上传按钮。
  - 测试项目筛选树容器 `#sidebarFilter`（运行时由 JS 填充）。
- **主内容区**：
  - `#statusAlert`：状态提示。
  - `#overviewSection`：数据概览卡片（含 `#commonInfoSection` 与 `#statCards`）。
  - `#chartsSection`：图表区（含网格/列表切换按钮、`#chartsContainer`）。
- **配置管理弹窗** `#configModal`：tab1=已保存配置列表，tab2=编辑/新建表单（含工作表/跳过行/表头行/固定列名行号/各固定列名映射/元数据行映射/关注测试项目/图表标题）。
- **图表放大弹窗** `#chartZoomModal`：全屏 modal，含下载按钮与 `#zoomChartBody` 容器。

---

### 5.8 [static/](file:///D:/Code/CodePython/traecode-测试分析001/static) — 静态资源

#### [static/js/main.js](file:///D:/Code/CodePython/traecode-测试分析001/static/js/main.js)

前端核心逻辑，1045 行，封装在 `DOMContentLoaded` 回调内。模块划分（注释中明确编号）：

| 编号 | 模块 | 职责 |
|------|------|------|
| 1 | 上传并解析 Excel | 表单提交 -> POST `/api/upload` -> 处理 `data.charts` -> 初始化 `visibleChartIdx` / `visibleSeries` |
| 2 | 左侧筛选树 | `renderSidebarTree(charts)`：一级=图表（主项目名+规格 bracket），二级=series；勾选联动 `toggleChartVisible` / `toggleSeriesVisible` -> `applyVisibility` |
| — | 滚动联动 | `setupScrollSync()`：`IntersectionObserver` 监听 `chart-wrapper`，找视口中部最近图表 -> 高亮侧边栏 + 平滑滚动 |
| 3 | 渲染 ECharts | `renderCharts(charts)`：构建 `.chart-wrapper` DOM、`echarts.init` + `buildOption(chart)`、绑定放大按钮 |
| — | 图表放大 Modal | `openZoomModal(chartIdx)` + `renderZoomChart()`：在 modal `shown.bs.modal` 后 init，加 dataZoom 与全屏坐标轴样式 |
| 4 | 数据概览卡片 | `renderOverview(data)`：4 张统计卡 + Program Name 提升 navbar 与 card-header，其余公共信息进公共信息卡片 |
| 5 | 视图切换 | 网格 / 列表模式切换 `view-grid` / `view-list`，触发 `resizeCharts()` |
| 6 | 产品配置管理 | `loadConfigList()` -> `/api/configs` 渲染表格；`loadConfigToForm` / `deleteConfig` / `loadDefaultBtn` / 表单提交保存 |
| 7 | 状态提示 | `showAlert(msg, type)`：顶部 alert 显示 4.5 秒 |

##### 关键内部状态

- `chartInstances`：所有 ECharts 实例（按 chartIdx 对齐）。
- `wrapperElements`：按 chartIdx 对齐的 DOM wrapper。
- `currentCharts` / `currentTestItems`：当前数据。
- `visibleChartIdx: Set<number>`：显示中的图表索引。
- `visibleSeries[chartIdx]: Set<seriesName>`：每个图表显示中的 series 名。
- `scrollIntersectionObserver` / `scrollActiveChartIdx`：滚动联动状态。
- `zoomChartInstance`：放大 modal 的 ECharts 实例。

##### `buildOption(chart)` 关键设计

- **参考线**：规格上限/下限 = 绿色实线 `#198754`（仅画在第一个 series 上，避免重复）；均值 = 粉红色虚线 `#e91e63`，每个 series 一条，多 series 时 label 加 series 名前缀。
- **Y 轴范围**：`yMin = min(数据最小, spec_lower) - 8% pad`，`yMax = max(数据最大, spec_upper) + 8% pad`，确保参考线与标签不被裁剪。
- **调色板**：`PALETTE = ["#5470c6", "#91cc75", ...]`，按 series idx 取模。
- **dataZoom**：`inside`（滚轮）+ `slider`（底部滑动条），全屏版字号加大。
- **toolbox**：`saveAsImage` + `dataZoom`（缩放/还原）。
- **legend**：多 series 时启用 legend 并通过 `selected` 控制 series 显示/隐藏（与侧边栏 `applyVisibility` 联动）。

#### [static/css/style.css](file:///D:/Code/CodePython/traecode-测试分析001/static/css/style.css)

自定义样式，覆盖 Bootstrap 默认：

- `main.container` 改为自适应宽度以容纳侧边栏。
- `#navProgramName` 与 `#headerProgramName` 的 Program Name 容器样式（含紫色 `#f209f6` 文本与固定宽度）。
- 上传区表单元素的精确样式（label 宽度、下拉框自适应、配置按钮蓝底白字）。
- `.chart-box` 卡片样式与 `.chart-body` 高度（320px，列表模式 400px）。
- `#chartZoomModal` 全屏 modal 样式（100vh/vw、flex 列布局、`#zoomChartBody` 填满）。
- 滚动联动高亮 `.sidebar-chart-header-active`（蓝色背景 + 左边框）。
- 网格/列表模式切换：`.view-grid .chart-wrapper { flex: 0 0 50% }` / `.view-list .chart-wrapper { flex: 0 0 100% }`。
- 统计卡片 `.stat-card`、筛选按钮 `.filter-btn`、空状态 `.empty-state`、统计表格 `.stat-table` 等。

---

### 5.9 [configs/](file:///D:/Code/CodePython/traecode-测试分析001/configs) — 产品配置模板

`configs/` 目录下放置只读产品配置 JSON 模板，由 `ConfigManager` 读取。

#### [SC-POWER板测试记录.json](file:///D:/Code/CodePython/traecode-测试分析001/configs/SC-POWER板测试记录.json)

SC-POWER 板测试记录的预置配置：

- `category_name`: `"SC-POWER板测试记录"`
- `excel_config`:
  - `sheet_name`: `"Sheet1"`
  - `skip_top_rows`: `1`（跳过 Program Name 标题行）
  - `header_rows`: `5`
  - `fixed_columns_header_row`: `5`（第 6 行读取固定列名）
  - `fixed_columns`: `sequence=序列号, product_model=名称, work_order=工单号, tester=测试员, fixture=工装编号, result=结论, time=时间`
  - `meta_row_mapping`: 默认（item_name=0, item_note=1, param_variable=2, spec_upper=3, spec_lower=4）
- `test_items`: `[]`（解析全部）
- `chart_config`: `title="SC-POWER板 测试分析曲线图"`, `y_axis_name="测试值"`

> 注意：该配置未声明 `merge_items`，因此上传 SC-POWER 板文件时会走 `Analyzer.auto_grouped_chart_data` 按规格自动分组。

---

### 5.10 辅助脚本与打包文件

#### [start.bat](file:///D:/Code/CodePython/traecode-测试分析001/start.bat)

Windows 启动脚本，设置环境变量后启动 Flask：

- `PYTHON_EXE`：固定路径 `C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe`
- `PYTHONPATH=%TEMP%\ta_libs`：备用依赖目录
- `DATA_DIR=%TEMP%\ta_data`：沙箱可写数据目录
- `PYTHONDONTWRITEBYTECODE=1`：不生成 `.pyc`
- 执行 `"%PYTHON_EXE%" -B app.py`

#### [build_exe.spec](file:///D:/Code/CodePython/traecode-测试分析001/build_exe.spec)

PyInstaller 打包配置：

- 入口：`app.py`
- `datas`：打包 `templates/`、`static/`、`configs/`
- `hiddenimports`：`flask_sqlalchemy`、`openpyxl` 及其子模块、`sqlalchemy` 及其方言、`pandas` 及其 `_libs` 子模块（含 tslibs 时间序列库）
- `excludes`：`tkinter`、`matplotlib`、`PyQt5`、`PySide2`、`IPython`、`notebook`、`jupyter`、`pytest`
- 输出 `测试分析系统.exe`，启用 UPX 压缩

#### [generate_sample.py](file:///D:/Code/CodePython/traecode-测试分析001/generate_sample.py)

生成示例 Excel 测试数据，便于快速验证：

- `build_test_items()`：定义 6 个测试项目（电压/待机电流/工作电流/温度/频率/外观检查），其中"外观检查"为文本型。
- `generate_sample(num_products=40)`：构建 5 行多行表头 + 40 行产品数据，固定种子 42，按规格中心+正态抖动生成数值，每 13 个产品故意越界一次，输出到 `sample_test_data.xlsx`。

#### [_debug_statistics.py](file:///D:/Code/CodePython/traecode-测试分析001/_debug_statistics.py)

按需求规则 1~4 验证解析流程的命令行脚本：

- 调用 `auto_detect_config` + `parse` 解析 `SC板测试记录01.xlsx`。
- 输出规则 1 读取结果、规则 2 公共信息+产品数量、规则 3 有效项目数、规则 4 图表分组数与多曲线组明细、最终汇总。
- 用于回归验证"分析项目数/图表数"是否符合预期。

#### [_debug_test.py](file:///D:/Code/CodePython/traecode-测试分析001/_debug_test.py)

更轻量的调试脚本：

- 打印 `auto_detect_config` 结果（JSON）。
- 解析后打印 rows/unique_items/program_info/columns/fix_time/fix_sequence 抽样。
- 调用 `get_test_items` 与 `auto_grouped_chart_data`，截取前 4 个图表打印 series/x 轴/规格。

---

## 6. 核心数据流与处理流程

### 6.1 上传-解析-渲染主流程

```
用户选 Excel + 产品类别
        |
        v
POST /api/upload  (app.py:upload)
        |
        +-- 保存原始 Excel 到 UPLOAD_FOLDER
        +-- 选择配置:
        |    +-- 有 product_category -> ConfigManager.load(name)
        |    +-- 无 -> ExcelParser.auto_detect_config(path)
        +-- ExcelParser(config).parse(path) --> 长格式 DataFrame
        |       (仅数值型 item_value, 含 fix_* 固定列)
        +-- 提取 common_info:
        |    +-- parser.program_info (Program Name 等)
        |    +-- 首条记录 fix_work_order/fix_tester/fix_fixture
        +-- _save_df(df) --> pickle 到 UPLOAD_FOLDER/uuid.pkl
        +-- 写 session (df_file/filename/product_category)
        +-- UploadRecord 入库
        +-- 选择图表策略:
        |    +-- merge_items 非空  -> Analyzer.grouped_chart_data
        |    +-- test_items 为空   -> Analyzer.auto_grouped_chart_data (规则3+4)
        |    +-- 否则             -> Analyzer.all_items_chart_data
        +-- 统计 product_count/analysis_item_count/chart_count
        +-- 返回 JSON (charts/test_items/common_info/...)
        |
        v
main.js: uploadForm.submit
        |
        +-- currentCharts = data.charts
        +-- 初始化 visibleChartIdx/visibleSeries
        +-- renderSidebarTree(charts)
        +-- renderCharts(charts)  --> echarts.init + buildOption
        +-- renderOverview(data)
        +-- setupScrollSync()  --> IntersectionObserver
```

### 6.2 配置管理流程

```
点击"产品配置管理"按钮
        |
        v
configModal.show.bs.modal -> loadConfigList
        |
        +-- GET /api/configs -> ConfigManager.list_configs
        +-- 渲染表格 (name/test_items/编辑/删除)

编辑/新建 Tab:
        +-- loadDefaultBtn  -> 填充默认模板
        +-- 编辑某配置     -> loadConfigToForm(name)
        |     +-- GET /api/config/<name> -> ConfigManager.load
        +-- 表单提交       -> POST /api/config
              +-- ConfigManager.save -> 写 {CONFIG_FOLDER}/name.json

删除某配置:
        +-- DELETE /api/config/<name> -> ConfigManager.delete
```

### 6.3 数据存储流程

```
原始 Excel   --> UPLOAD_FOLDER/<filename>.xlsx
DataFrame    --> UPLOAD_FOLDER/<uuid>.pkl   (session 记录文件名)
上传元数据   --> SQLite upload_records 表
产品配置     --> configs/<name>.json  (只读模板)
新配置       --> DATA_DIR/configs/<name>.json  (用户保存)
```

---

## 7. REST API 接口一览

| 方法 | 路径 | 请求 | 响应 | 说明 |
|------|------|------|------|------|
| GET | `/` | — | HTML | 首页 |
| POST | `/api/upload` | multipart: `file`, `product_category` | JSON | 上传并解析 |
| GET | `/api/chart-data` | query: `item` | JSON | 单项图表数据 |
| GET | `/api/test-items` | — | JSON `{test_items, stats}` | 测试项目与每项统计 |
| GET | `/api/configs` | — | JSON `{configs:[{name,filename,test_items}]}` | 配置列表 |
| GET | `/api/config/<name>` | — | JSON 配置 / 404 | 获取配置 |
| POST | `/api/config` | JSON 配置 | JSON `{message,filename}` / 400 | 保存配置 |
| DELETE | `/api/config/<name>` | — | JSON / 404 | 删除配置 |
| GET | `/api/config/default` | — | JSON 模板 | 默认配置模板 |
| GET | `/api/sheets` | — | JSON `{sheets}` / 400/404 | 工作表列表 |
| GET | `/api/history` | — | JSON `{history:[record]}` | 最近 50 条上传记录 |

### `/api/upload` 响应字段

```jsonc
{
  "filename": "W11-25080019-...xlsx",
  "product_category": "SC-POWER板测试记录",  // 或 "自动检测配置"
  "total_rows": 106380,                      // 长格式 DataFrame 行数
  "product_count": 647,                      // 按序列号去重
  "analysis_item_count": 53,                 // 有规格限的项目数
  "chart_count": 32,                         // 分组后的图表数
  "common_info": {                           // 公共信息
    "Program Name": "2001-01015-00-EU ...",
    "工单号": "W11-26010037",
    "测试员": "...",
    "工装编号": "..."
  },
  "test_items": ["项目A", "项目B"],
  "charts": [ ],
  "record_id": 1
}
```

---

## 8. 关键类与函数

### 8.1 `ExcelParser`

```python
class ExcelParser:
    def __init__(self, product_config: dict): ...
    def parse(self, file_path: str) -> pd.DataFrame: ...

    # 内部辅助
    def _find_fixed_col(self, field_key, header_df, fixed_cols_row) -> Optional[int]: ...
    def _collect_test_items(self, header_df, excluded) -> list: ...
    def _extract_program_info(self, raw: pd.DataFrame) -> None: ...

    # 静态
    @staticmethod
    def _strip(value) -> str: ...
    @staticmethod
    def _cell(header_df, row_idx, col_idx) -> str: ...
    @staticmethod
    def _has_value(v) -> bool: ...
    @staticmethod
    def _try_numeric(val) -> Optional[float]: ...
    @classmethod
    def _try_spec_numeric(cls, val_str: str) -> Optional[float]: ...
    @staticmethod
    def get_sheet_names(file_path: str) -> list: ...
    @staticmethod
    def auto_detect_config(file_path: str) -> dict: ...
    @staticmethod
    def _default_config() -> dict: ...
```

### 8.2 `Analyzer`

```python
class Analyzer:
    @staticmethod
    def get_test_items(df) -> list: ...
    @staticmethod
    def get_item_metadata(df, item_name) -> dict: ...
    @staticmethod
    def filter_by_item(df, test_item=None) -> pd.DataFrame: ...
    @staticmethod
    def statistics(df) -> dict: ...
    @staticmethod
    def chart_data(df, test_item=None) -> dict: ...
    @staticmethod
    def grouped_chart_data(df, merge_items) -> list: ...
    @staticmethod
    def all_items_chart_data(df) -> list: ...
    @staticmethod
    def auto_grouped_chart_data(df) -> list: ...
```

### 8.3 `ConfigManager`

```python
class ConfigManager:
    def __init__(self, config_folder: str): ...
    def _path(self, name: str) -> str: ...
    def list_configs(self) -> list: ...
    def load(self, name: str) -> Optional[dict]: ...
    def save(self, config: dict) -> str: ...
    def delete(self, name: str) -> bool: ...
    @staticmethod
    def default_config() -> dict: ...
```

### 8.4 `UploadRecord`

```python
class UploadRecord(db.Model):
    __tablename__ = "upload_records"
    id: int
    filename: str
    saved_path: str
    product_category: str
    total_rows: int
    test_items: str  # JSON
    created_at: datetime
    def to_dict(self) -> dict: ...
```

### 8.5 `Config`

```python
class Config:
    SQLALCHEMY_DATABASE_URI: str
    SQLALCHEMY_TRACK_MODIFICATIONS: bool
    UPLOAD_FOLDER: str
    ALLOWED_EXTENSIONS: set
    CONFIG_FOLDER: str
    SECRET_KEY: str
    MAX_CONTENT_LENGTH: int
```

### 8.6 `app.py` 辅助与路由函数

```python
def _allowed_file(filename: str) -> bool: ...
def _save_df(df: pd.DataFrame) -> str: ...
def _load_df() -> pd.DataFrame: ...

# 路由函数
def index(): ...
def upload(): ...
def chart_data(): ...
def test_items(): ...
def list_configs(): ...
def get_config(name): ...
def save_config(): ...
def delete_config(name): ...
def default_config(): ...
def get_sheets(): ...
def history(): ...
```

---

## 9. 数据模型

### 9.1 数据库表 `upload_records`

| 列名 | 类型 | 约束 | 说明 |
|------|------|------|------|
| id | INTEGER | PK, AUTO | 主键 |
| filename | VARCHAR(255) | NOT NULL | 原始文件名 |
| saved_path | VARCHAR(512) | NOT NULL | 服务器存储路径 |
| product_category | VARCHAR(100) | — | 产品类别（含"自动检测配置"） |
| total_rows | INTEGER | DEFAULT 0 | 长格式 DataFrame 行数 |
| test_items | TEXT | — | JSON 字符串，如 `["项目A","项目B"]` |
| created_at | DATETIME | DEFAULT utcnow | 创建时间 |

### 9.2 长格式 DataFrame 列

`ExcelParser.parse()` 输出，`Analyzer` 输入：

| 列名 | 类型 | 说明 |
|------|------|------|
| `sequence` | str | 序列号（产品唯一标识） |
| `product_model` | str | 产品型号 |
| `item_name` | str | 测试项目名（同名多列时含 `[参数变量]` 后缀） |
| `item_value` | float | 数值型测试值（文本型已过滤） |
| `item_note` | str | 项目备注 |
| `param_variable` | str | 参数变量（单位） |
| `spec_upper` | float/None | 规格上限（`*`/`-`/`/` 视为 None） |
| `spec_lower` | float/None | 规格下限 |
| `fix_work_order` | str | 工单号（前缀 `fix_` 表示固定列值） |
| `fix_tester` | str | 测试员 |
| `fix_fixture` | str | 工装编号 |
| `fix_result` | str | 结论 |
| `fix_time` | str | 时间（X 轴优先使用） |

### 9.3 ECharts option 结构（前端 `buildOption`）

```jsonc
{
  "tooltip": {"trigger": "axis"},
  "legend": {"top": 0, "selected": {}},
  "grid": {"left": 80, "right": 80, "top": 25, "bottom": 65},
  "xAxis": {"type": "category", "data": [], "name": "时间"},
  "yAxis": {"type": "value", "name": "测试值", "min": 0, "max": 0},
  "dataZoom": [{"type":"inside"}, {"type":"slider"}],
  "series": [
    {
      "name": "", "type": "line", "smooth": true,
      "data": [], "markLine": {"data": []}
    }
  ],
  "toolbox": {"feature": {"saveAsImage": {}, "dataZoom": {}}}
}
```

---

## 10. 产品配置 JSON Schema

```jsonc
{
  "category_name": "SC-POWER板测试记录",   // 必填，用作文件名
  "excel_config": {
    "sheet_name": "Sheet1",             // 工作表名或序号
    "skip_top_rows": 1,                     // 顶部跳过行数
    "header_rows": 5,                       // 测试项目元数据占用行数
    "fixed_columns_header_row": 5,          // 固定列名行号（相对跳过后）；-1 不用
    "fixed_columns": {                      // 固定列字段 -> Excel 列名
      "sequence": "序列号",
      "product_model": "名称",
      "work_order": "工单号",
      "tester": "测试员",
      "fixture": "工装编号",
      "result": "结论",
      "time": "时间"
    },
    "meta_row_mapping": {                   // 元数据行索引（相对 header_df）
      "item_name": 0,
      "item_note": 1,
      "param_variable": 2,
      "spec_upper": 3,
      "spec_lower": 4
    }
  },
  "test_items": [],                         // 关注项目白名单；空 = 解析全部
  "merge_items": [                           // 显式合并组；空 = 按规格自动分组
    {"name": "组名", "items": ["项A", "项B"]}
  ],
  "chart_config": {
    "title": "SC-POWER板 测试分析曲线图",
    "y_axis_name": "测试值"
  }
}
```

### 字段语义

| 字段 | 作用 |
|------|------|
| `skip_top_rows` | 跳过 Excel 顶部的标题行（如 `Program Name:` 行）。跳过行内的 `Key: Value` 会被 `_extract_program_info` 提取为公共信息 |
| `header_rows` | 测试项目元数据占用的行数。`parse` 取前 `header_rows` 行作为 `header_df` |
| `meta_row_mapping` | 告诉解析器"项目名称"在第几行、"规格上限"在第几行等 |
| `fixed_columns_header_row` | 真实 SC 板格式中固定列名（序列号等）单独占一行，该行号相对跳过行后；若 -1 则从 `item_name` 行匹配 |
| `fixed_columns` | 字段名到 Excel 列名的映射，用于提取每行的固定列值（`fix_*` 前缀） |
| `test_items` | 白名单。空数组 = 解析全部数值型项目；非空 = 仅解析匹配项（支持原始名或带 `[参数]` 的唯一名） |
| `merge_items` | 显式合并组定义。在 `merge_items` 中出现的项会与同组其他项绘制到同一图表；不在任何合并组中的项单独成图 |

### 三种图表合并策略

`app.py` 的 `upload` 路由按以下优先级选择（见 [app.py:157-164](file:///D:/Code/CodePython/traecode-测试分析001/app.py#L157-L164)）：

1. `merge_items` 非空 -> `Analyzer.grouped_chart_data(df, merge_items)`（按配置显式分组）
2. `test_items` 为空（白名单全开） -> `Analyzer.auto_grouped_chart_data(df)`（按 `(spec_lower, spec_upper)` 自动分组）
3. 否则 -> `Analyzer.all_items_chart_data(df)`（每项一图）

---

## 11. 核心算法规则

来源：[测试分析001.md](file:///D:/Code/CodePython/traecode-测试分析001/测试分析001.md) "数据提取规则"小节。

### 规则 1：读取文件

- 读取 Excel 默认第 1 个 sheet（或配置的 `sheet_name`）。
- 若用户已选产品配置，按配置解析；否则 `auto_detect_config` 自动检测。
- 输出**长格式** DataFrame：每行一个 (产品, 测试项目) 数值记录。

### 规则 2：提取公共信息 + 产品数量

- **公共信息**：
  - 顶部跳过行中 `Program Name` 等标题字段（支持 `Key: Value` 同单元格 与 `Key:` + 右侧值两种写法）。
  - 首行产品记录的固定列值：`工单号` / `测试员` / `工装编号` / `结论` / `时间`。
- **产品数量**：按 `序列号`（`sequence` 列）去重计数。

### 规则 3：有上下限的项目列为分析数据 + 统计项目数量

- 候选范围：解析出的所有数值型测试项目。
- **筛选规则**：仅保留同时满足以下条件的测试项目
  - 至少有规格上限或规格下限其中之一（`spec_lower` 或 `spec_upper` 不为 `None`）
  - 对应数据为数值型（非文本/非空）
- 无规格限的项目不参与图表展示。
- **分析项目数量** = 筛选后剩余的测试项目数。

### 规则 4：合并上下限一致的项目为一组 + 统计图表数量

- 按 `(spec_lower, spec_upper)` 元组对分析项目分组。
- 元组相同的项目合并为同一组（`None` 与 `None` 视为相同，可合并）。
- 每组生成 1 张图表，图表内包含 N 条曲线（N = 该组内项目数）。
- **图表数量** = 分组后的组数。
- **曲线总数** = 分析项目数量（与规则 3 输出一致）。

### 规则执行示例（`SC板测试记录01.xlsx`）

| 规则 | 输出 |
|------|------|
| 1 读取文件 | 106,380 行 x 13 列（长格式） |
| 2 公共信息 | Program Name=2001-01015-00-EU SC-POWER板-260304，工单号=W11-26010037，产品数量=647 |
| 3 分析项目数量 | 53 项（有上下限 + 数值列；120 项无规格限被排除） |
| 4 图表数量 | 32 张（9 张多曲线合并图，23 张单曲线图；曲线总数=53） |

### 真实 Excel 结构（SC 板）

- 第 1 行：程序标题行（`Program Name: ...`），需跳过
- 第 2-6 行：测试项目元数据表头（共 5 行）
  - 第 2 行：项目名称（同一项目可能占多列，如 `静态负载测试_16` 占 4 列）
  - 第 3 行：项目备注
  - 第 4 行：参数变量（如 `CH1电压(V)` / `CH1电流(V)` / `量测值(V)`，作为子项区分多列同名项目）
  - 第 5 行：规格上限（`*` 或 `-` 表示无限）
  - 第 6 行：规格下限（`*` 或 `-` 表示无限）
- 第 7 行：固定列名行（序列号/工单号/名称/测试员/工装编号/结论/时间）
- 第 8 行起：产品测试记录（每行一个产品，从序列号开始）
- 单元格内容可能带单引号前缀（如 `'单通道电源设定_1'`），解析时需去除

---

## 12. 运行与部署

### 12.1 方式一：源码运行（开发/调试）

**环境要求**：Python 3.13+，已安装 [requirements.txt](file:///D:/Code/CodePython/traecode-测试分析001/requirements.txt) 中的依赖。

**安装依赖**：

```powershell
& "C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe" -m pip install -r requirements.txt
```

**启动**：

- 双击 [start.bat](file:///D:/Code/CodePython/traecode-测试分析001/start.bat)（推荐），或：

```powershell
$env:DATA_DIR="$env:TEMP\ta_data"
$env:PYTHONDONTWRITEBYTECODE="1"
Set-Location "d:\Code\CodePython\traecode-测试分析001"
& "C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe" -B app.py
```

**访问**：浏览器打开 **http://127.0.0.1:5000/**

**停止**：`Ctrl+C`

### 12.2 方式二：打包 EXE（部署/分发）

**生成 EXE**（使用 [build_exe.spec](file:///D:/Code/CodePython/traecode-测试分析001/build_exe.spec)）：

```powershell
Set-Location "d:\Code\CodePython\traecode-测试分析001"
& "C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe" -m PyInstaller --noconfirm --clean --onefile --name "测试分析系统" --add-data "templates;templates" --add-data "static;static" --add-data "configs;configs" --hidden-import flask_sqlalchemy --hidden-import openpyxl --hidden-import sqlalchemy app.py
```

或直接使用 spec 文件：

```powershell
& "C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe" -m PyInstaller --noconfirm --clean build_exe.spec
```

**运行 EXE**：

1. 进入 `dist/` 目录，双击 `测试分析系统.exe`
2. 控制台显示访问地址 `http://127.0.0.1:5000/`
3. 浏览器打开该地址即可使用
4. 关闭控制台窗口即停止服务

> EXE 启动需要几秒初始化（解压 `_MEIPASS`），请耐心等待控制台出现 "Running on http://..." 字样。

### 12.3 方式三：部署到服务器（生产）

1. **数据库**：修改 [config.py](file:///D:/Code/CodePython/traecode-测试分析001/config.py) 中 `SQLALCHEMY_DATABASE_URI` 为 MySQL 连接串，或设置环境变量 `DATABASE_URL`：
   ```
   mysql+pymysql://用户名:密码@localhost:3306/test_analysis
   ```
2. **WSGI 服务器**：使用生产级 WSGI 服务器（如 waitress）：
   ```powershell
   pip install waitress
   waitress-serve --host=0.0.0.0 --port=5000 app:app
   ```
3. **反向代理**：Nginx + HTTPS 配置。

### 12.4 环境变量

| 变量 | 作用 | 默认 |
|------|------|------|
| `DATA_DIR` | 可写数据目录（沙箱环境下设为 `%TEMP%\ta_data`） | 脚本所在目录（源码）/ EXE 同级 `data/`（frozen） |
| `DATABASE_URL` | 数据库连接字符串（覆盖 SQLite 默认） | `sqlite:///{DATA_DIR}/test_analysis.db` |
| `SECRET_KEY` | Flask 密钥 | `test-analysis-secret-key-2026` |
| `PYTHONDONTWRITEBYTECODE` | 不生成 `.pyc` | `1`（start.bat） |
| `PYTHONPATH` | 备用依赖目录 | `%TEMP%\ta_libs`（start.bat） |

### 12.5 使用流程

1. 启动应用。
2. 浏览器打开 http://127.0.0.1:5000/
3. 左侧边栏：
   - 点击"选择 Excel 文件"，选择 `.xlsx` 测试记录
   - 产品类别选择 "-- 默认配置 --"（自动从文件提取配置）或某个已保存配置
   - 点击"上传并分析"
4. 等待 20-30 秒解析完成，图表自动显示
5. 交互操作：
   - 左侧筛选树：勾选/取消图表或曲线
   - 滚动页面：侧边栏自动高亮+滚动到对应图表
   - 点击放大按钮：全屏查看图表详情，可下载 PNG

---

## 13. 关键设计要点

### 13.1 双目录分离（源码 / frozen 兼容）

`config.py` 通过 `_get_base_dir` 与 `_get_data_dir` 分离只读资源与可写数据：

- 源码运行：两者都指向脚本所在目录。
- PyInstaller frozen：只读资源用 `sys._MEIPASS`（解压临时目录），可写数据用 EXE 同级 `data/` 或 `DATA_DIR` 环境变量。

这避免了打包后无法写 `templates/static/configs` 同级目录的问题。

### 13.2 Excel 解析的多行表头适配

`ExcelParser.parse` 通过 `meta_row_mapping` + `fixed_columns_header_row` + `fixed_columns` 三层配置，适配多种 Excel 格式：

- **简单格式**：固定列名与项目名称在同一表头行（`fixed_columns_header_row=-1`）。
- **真实 SC 板格式**：固定列名单独占一行（`fixed_columns_header_row=5`）。
- **同名多列项目**：通过参数变量行拼接 `项目名 [参数变量]` 生成唯一项名（见 `_collect_test_items`）。

### 13.3 自动配置检测

`ExcelParser.auto_detect_config` 通过关键词集合扫描 Excel 前 25 行，推断 `skip_top_rows` / `meta_row_mapping` / `fixed_columns_header_row` / `fixed_columns`，使应用在零配置下即可工作。

### 13.4 长格式 + 规格分组的核心算法

- **长格式**：`ExcelParser.parse` 将宽表（每产品一行）展开为长表（每 (产品, 测试项目) 一行），便于后续按 `item_name` 聚合统计与分组。
- **规格分组**：`Analyzer.auto_grouped_chart_data` 以 `(spec_lower, spec_upper)` 元组为分组 key，将相同规格区间的项目合并到同一图表（多 series），减少图表数量同时保留对比能力。

### 13.5 前端滚动联动与二级筛选

- **IntersectionObserver**：监听所有 `.chart-wrapper`，找出视口中部最近的图表，高亮侧边栏对应项并平滑滚动到可视区（见 `setupScrollSync`）。
- **二级筛选**：每个图表的 series 单独可控（`visibleSeries[chartIdx]`），通过 `inst.setOption({legend:{selected}})` 控制 ECharts series 显隐，与侧边栏 `applyVisibility` 联动。

### 13.6 DataFrame 序列化与会话

- 大 DataFrame 不直接放 session（受 `SECRET_KEY` 加密 cookie 大小限制），而是 `_save_df` pickle 到磁盘，session 仅记录文件名。
- 后续 `/api/chart-data`、`/api/test-items` 等接口通过 `_load_df` 从磁盘加载。

### 13.7 参考线与 Y 轴范围

- 规格上限/下限只画一次（放在第一个 series 的 markLine），避免多 series 时重复绘制。
- Y 轴范围 = `min(数据最小, spec_lower) - 8% pad` ~ `max(数据最大, spec_upper) + 8% pad`，确保参考线及标签完整显示。

### 13.8 图表放大 Modal 的尺寸处理

`openZoomModal` 在 modal `shown.bs.modal` 事件（动画结束、容器有真实尺寸）后才 `echarts.init`，并监听 `window resize` 同步 `zoomChartInstance.resize()`；`hidden.bs.modal` 时销毁实例与监听，避免尺寸污染。

### 13.9 配置路径安全

`ConfigManager._path` 将 `category_name` 中的 `/` 与 `\` 替换为 `_`，防止路径穿越，同时保留中文字符使配置名可读。

### 13.10 文本型测试值过滤

`ExcelParser._try_numeric` 对 `-` / `—` / `*` / `/` 与无法 `float()` 的值返回 None，使文本型测试项（如 `外观检查 OK/NG`）在 `parse` 阶段被跳过，仅数值型进入长格式 DataFrame。

---

## 附录：模块间依赖关系图

```
app.py
  +-- config.Config           (配置)
  +-- models.db, UploadRecord (ORM)
  +-- services.excel_parser.ExcelParser
  +-- services.analyzer.Analyzer
  +-- services.config_manager.ConfigManager

services/excel_parser.py
  +-- pandas, math, typing

services/analyzer.py
  +-- pandas, typing

services/config_manager.py
  +-- os, json, typing

models.py
  +-- datetime, flask_sqlalchemy

config.py
  +-- os, sys, tempfile

templates/index.html
  +-- templates/base.html
  +-- (运行时) static/js/main.js, static/css/style.css

static/js/main.js
  +-- (运行时) Bootstrap, ECharts (CDN)
```

模块间通过函数参数与返回值传递 DataFrame 与 dict，无循环依赖，**单向数据流**：`ExcelParser` 产出 DataFrame -> `Analyzer` 消费产出 charts JSON -> `app.py` 序列化返回前端 -> `main.js` 渲染 ECharts。
