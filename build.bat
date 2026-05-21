@echo off
REM ========================================
REM NewsCollector V36_Fixed - Build Script
REM ========================================
REM 版本: V36_Fixed (Cursor优化版)
REM 日期: 2026-05-15
REM ========================================

echo ========================================
echo NewsCollector V36_Fixed - Build Script
echo ========================================
echo.

cd /d "%~dp0"

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found! Please install Python 3.10-3.13
    pause
    exit /b 1
)

echo [OK] Python found
echo.

REM 步骤1: 清理旧构建
echo [1/5] Cleaning old build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
echo [OK] Build directories cleaned
echo.

REM 步骤2: 检查必需文件
echo [2/5] Checking required files...
if not exist "main_gui_standalone.py" (
    echo [ERROR] main_gui_standalone.py not found!
    pause
    exit /b 1
)
if not exist "NewsCollector_V36_Fixed.spec" (
    echo [ERROR] NewsCollector_V36_Fixed.spec not found!
    pause
    exit /b 1
)
if not exist "魔方_最终图标.ico" (
    echo [WARNING] Icon file not found, building without icon...
)
echo [OK] All required files found
echo.

REM 步骤3: 执行打包
echo [3/5] Building executable...
if exist "魔方_最终图标.ico" (
    pyinstaller --noconsole --icon="魔方_最终图标.ico" NewsCollector_V36_Fixed.spec
) else (
    pyinstaller --noconsole NewsCollector_V36_Fixed.spec
)

if errorlevel 1 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)
echo [OK] Build completed
echo.

REM 步骤4: 验证打包结果
echo [4/5] Verifying build...
if not exist "dist\NewsCollector_V36_Fixed.exe" (
    echo [ERROR] Executable not found!
    pause
    exit /b 1
)

for %%F in ("dist\NewsCollector_V36_Fixed.exe") do set SIZE=%%~zF
echo [OK] Executable created: dist\NewsCollector_V36_Fixed.exe
echo [OK] File size: %SIZE% bytes
echo.

REM 步骤5: 复制到发布目录
echo [5/5] Copying to release directory...
if not exist "..\发布" mkdir "..\发布"
copy /Y "dist\NewsCollector_V36_Fixed.exe" "..\发布\" >nul
if exist "README.md" copy /Y "README.md" "..\发布\" >nul
echo [OK] Files copied to release directory
echo.

echo ========================================
echo Build Summary
echo ========================================
echo Output: ..\发布\NewsCollector_V36_Fixed.exe
echo.
echo To run the program:
echo   1. Navigate to release directory
echo   2. Double-click NewsCollector_V36_Fixed.exe
echo.
echo New Features:
echo   - Custom Keywords Filtering
echo   - Cursor Resource Optimization
echo   - Unicode Encoding Fix
echo.
echo ========================================
echo [SUCCESS] Build completed successfully!
echo ========================================
echo.
pause
