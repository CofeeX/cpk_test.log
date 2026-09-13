@echo off
chcp 65001 >nul
REM CPK-PCBA - 局域网服务器部署启动脚本
REM 要求: 已安装 Python 3.13 并加入 PATH, 已执行 pip install -r requirements.txt
setlocal
set DATA_DIR=%~dp0data
echo ========================================
echo   CPK-PCBA 测试分析系统 (局域网服务器)
echo ========================================
echo.
echo 正在启动服务...
echo 本机访问:   http://127.0.0.1:5000
echo 局域网访问: http://服务器IP:5000   (ipconfig 查看 IP)
echo 关闭本窗口即停止服务
echo.
python -B app.py --serve
