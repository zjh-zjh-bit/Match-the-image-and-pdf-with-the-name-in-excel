@echo off
chcp 65001
title 证书分类工具 - 一键打包
echo ========================================
echo   证书分类工具 - 一键打包脚本
echo ========================================
echo.

echo 步骤1: 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3.7+
    pause
    exit /b 1
)

echo 步骤2: 检查必要文件...
if not exist "tessdata\tesseract.exe" (
    echo ❌ 错误: 未找到 tessdata\tesseract.exe
    pause
    exit /b 1
)

if not exist "poppler\Library\bin\pdftoppm.exe" (
    echo ❌ 错误: 未找到 poppler\Library\bin\pdftoppm.exe
    pause
    exit /b 1
)

echo ✅ 必要文件检查通过

echo 步骤3: 清理旧文件...
rmdir /s /q dist 2>nul
rmdir /s /q build 2>nul
del /f /q *.spec 2>nul

echo 步骤4: 安装依赖...
pip install pandas==2.0.3 numpy==1.24.3 pillow==9.5.0 pytesseract==0.3.10 pdf2image==1.16.3 openpyxl==3.1.2 pyinstaller==5.13.0 pywin32==306

if errorlevel 1 (
    echo 使用镜像源重试...
    pip install pandas==2.0.3 numpy==1.24.3 pillow==9.5.0 pytesseract==0.3.10 pdf2image==1.16.3 openpyxl==3.1.2 pyinstaller==5.13.0 pywin32==306 -i https://pypi.tuna.tsinghua.edu.cn/simple/
)

echo 步骤5: 开始打包...
python 打包脚本.py

if exist "dist\证书分类工具\证书分类工具.exe" (
    echo.
    echo ✅ 打包成功！
    echo 📁 程序位置: dist\证书分类工具\证书分类工具.exe
    echo.
    echo 💡 提示：程序已包含所有依赖，可直接使用
) else (
    echo.
    echo ❌ 打包失败
)

pause