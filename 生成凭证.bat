@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   生产成本结转凭证生成工具
echo ============================================
echo.

REM 检测 Python 命令
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" where py >nul 2>&1 && set PYTHON=py
if "%PYTHON%"=="" where python3 >nul 2>&1 && set PYTHON=python3

if "%PYTHON%"=="" (
    echo [错误] 未检测到 Python，请先安装 Python：
    echo   https://www.python.org/downloads/
    echo   安装时勾选 "Add Python to PATH"
    echo.
    pause >nul
    exit /b
)

echo 使用 Python：%PYTHON%
echo.

REM 安装依赖
echo 检查依赖...
%PYTHON% -m pip install openpyxl xlwt -q

echo.
echo 当前目录下的 Excel 文件：
dir /b *.xlsx 2>nul
echo.
set /p INPUT_FILE=请输入文件名（如 生产成本0713.xlsx）：

if "%INPUT_FILE%"=="" (
    echo 未输入文件名，退出。
    pause >nul
    exit /b
)

%PYTHON% generate_voucher.py "%INPUT_FILE%"
echo.
echo 按任意键退出...
pause >nul
