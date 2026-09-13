"""应用配置"""
import os
import sys
import tempfile


def _get_base_dir():
    """获取应用资源根目录 (兼容 PyInstaller frozen 模式)"""
    if getattr(sys, "frozen", False):
        # PyInstaller 打包后: 资源被解压到 sys._MEIPASS
        return sys._MEIPASS
    # 源码运行: 使用脚本所在目录
    return os.path.abspath(os.path.dirname(__file__))


def _get_data_dir():
    """获取可写数据目录 (兼容 frozen 模式 + 沙箱环境)"""
    # 1. 优先使用环境变量
    env_data = os.environ.get("DATA_DIR")
    if env_data:
        return env_data
    # 2. PyInstaller 打包模式: 写到 EXE 同级的 data 目录 (用户可读写)
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        return os.path.join(exe_dir, "data")
    # 3. 源码运行: 默认写到项目目录
    return os.path.abspath(os.path.dirname(__file__))


BASE_DIR = _get_base_dir()
DATA_DIR = _get_data_dir()


class Config:
    # 数据库配置
    # 开发环境默认使用 SQLite，便于快速运行；部署到服务器时切换为 MySQL：
    #   SQLALCHEMY_DATABASE_URI = "mysql+pymysql://用户名:密码@localhost:3306/test_analysis"
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        "DATABASE_URL",
        f"sqlite:///{os.path.join(DATA_DIR, 'test_analysis.db')}",
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # 文件上传
    UPLOAD_FOLDER = os.path.join(DATA_DIR, "uploads")
    ALLOWED_EXTENSIONS = {"xlsx", "xls"}

    # 产品配置 JSON 目录 (只读配置可放在项目目录, 新配置写入 DATA_DIR)
    CONFIG_FOLDER = os.path.join(BASE_DIR, "configs")

    # 密钥
    SECRET_KEY = os.environ.get("SECRET_KEY", "test-analysis-secret-key-2026")

    # 上传文件最大大小 (16MB)
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
