@echo off
chcp 65001 >nul
REM CPK-PCBA - 启动脚本
REM
REM 首次运行需要安装依赖, 请执行:
REM   pip install -r requirements.txt
REM
REM 若系统沙箱限制写入, 依赖已安装到临时目录 %TEMP%\ta_libs
REM 本脚本会自动设置 PYTHONPATH 指向该目录

set PYTHON_EXE=C:\Users\凯伦\AppData\Local\Programs\Python\Python313\python.exe
set PYTHONPATH=%TEMP%\ta_libs
set DATA_DIR=%TEMP%\ta_data
set PYTHONDONTWRITEBYTECODE=1

echo ========================================
echo   CPK-PCBA
echo ========================================
echo.
echo 正在启动 Flask 应用...
echo 访问地址: http://127.0.0.1:5000
echo 按 Ctrl+C 停止
echo.

"%PYTHON_EXE%" -B app.py
