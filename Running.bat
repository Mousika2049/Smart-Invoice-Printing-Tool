@echo off
chcp 65001 > nul  & REM 强制使用 UTF-8 编码，避免中文乱码

set PYTHON_SCRIPT="smart_printer.py"
set PYTHON_PATH="C:\Users\Maverick\AppData\Local\Programs\Python\Python312\python.exe"

echo 正在运行 Python 脚本...
%PYTHON_PATH% %PYTHON_SCRIPT%

if %errorlevel% neq 0 (
    echo 运行失败！错误代码: %errorlevel%
)

pause