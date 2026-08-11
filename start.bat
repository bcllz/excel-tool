@echo off
chcp 65001 >nul
title Excel 百宝箱

echo ============================================
echo    Excel 百宝箱 - 启动中...
echo ============================================
echo.

:: 查找 Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    where python3 >nul 2>&1
    if %errorlevel% neq 0 (
        echo [错误] 未找到 Python，请先安装 Python 3.9+
        echo 下载地址: https://www.python.org/downloads/
        echo 安装时请勾选 "Add Python to PATH"
        pause
        exit /b 1
    )
    set PYTHON=python3
) else (
    set PYTHON=python
)

echo Python 版本:
%PYTHON% --version
echo.

:: 创建虚拟环境（首次）
if not exist "venv" (
    echo [1/3] 正在创建虚拟环境...
    %PYTHON% -m venv venv
)

:: 激活虚拟环境
call venv\Scripts\activate.bat

:: 安装依赖（首次）
if not exist "venv\.deps_ok" (
    echo [2/3] 正在安装依赖（首次需要几分钟）...
    pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
    if %errorlevel% neq 0 (
        echo 国内镜像失败，尝试官方源...
        pip install -r requirements.txt
    )
    type nul > venv\.deps_ok
    echo [2/3] 依赖安装完成
)

:: 启动
echo [3/3] 正在启动...
echo.
echo ============================================
echo    浏览器打开 → http://localhost:8501
echo    按 Ctrl+C 停止
echo ============================================
echo.

streamlit run app.py --server.port 8501 --server.headless true

pause
