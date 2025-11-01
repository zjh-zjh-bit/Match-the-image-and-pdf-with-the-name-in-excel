import os
import subprocess
import sys

def check_resources():
    """检查资源文件"""
    print("检查资源文件...")
    
    required_files = {
        'tessdata': ['tesseract.exe'],
        'poppler': ['Library/bin/pdftoppm.exe']
    }
    
    all_exists = True
    for folder, files in required_files.items():
        for file in files:
            file_path = os.path.join(folder, file)
            if not os.path.exists(file_path):
                print(f"❌ 缺失: {file_path}")
                all_exists = False
            else:
                print(f"✅ 找到: {file_path}")
    
    return all_exists

def create_spec_file():
    """创建PyInstaller spec文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('poppler', 'poppler'),
        ('tessdata', 'tessdata'),
    ],
    hiddenimports=[
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.base',
        'pandas._libs.skiplist',
        'pandas._libs.json',
        'numpy.core._dtype_ctypes',
        'pkg_resources',
        'importlib_metadata',
        'win32timezone',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='证书分类工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
'''
    
    with open('cert_tool.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)

def build_app():
    """构建应用程序"""
    print("开始打包应用程序...")
    
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        'cert_tool.spec',
        '--noconfirm',
        '--clean'
    ]
    
    try:
        result = subprocess.run(cmd, check=True)
        return result.returncode == 0
    except subprocess.CalledProcessError:
        return False
    except FileNotFoundError:
        print("❌ 未找到PyInstaller，请先安装依赖")
        return False

def main():
    print("证书分类工具打包脚本")
    print("=" * 50)
    
    if not check_resources():
        print("❌ 资源文件不完整，无法打包")
        input("按回车键退出...")
        return
    
    print("创建打包配置...")
    create_spec_file()
    
    print("开始打包...")
    if build_app():
        print("✅ 打包成功！")
        print("📁 程序位置: dist/证书分类工具/证书分类工具.exe")
    else:
        print("❌ 打包失败")
    
    input("按回车键退出...")

if __name__ == "__main__":
    main()