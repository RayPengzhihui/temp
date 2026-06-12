@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   生产成本结转凭证生成工具
echo ============================================
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

python generate_voucher.py "%INPUT_FILE%"
echo.
echo 按任意键退出...
pause >nul
