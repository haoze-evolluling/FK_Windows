import os
import shutil
import subprocess
import sys

def install_dependencies():
    """安装必要的依赖"""
    print("检查并安装必要的依赖...")
    
    # 安装PyInstaller
    try:
        import PyInstaller
        print("PyInstaller已安装")
    except ImportError:
        print("正在安装PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller安装完成")
    
    # 安装PIL/Pillow
    try:
        from PIL import Image
        print("Pillow已安装")
    except ImportError:
        print("正在安装Pillow...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])
        print("Pillow安装完成")

def build_exe():
    """使用PyInstaller构建EXE文件"""
    print("开始打包程序...")
    
    # 确保dist目录不存在
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    
    # 构建命令 - 使用--onefile选项创建单个EXE文件
    build_cmd = [
        sys.executable, 
        "-m", 
        "PyInstaller",
        "--name=FK_WINDOWS-Windows优化工具",
        "--onefile",
        "--windowed",  # 不显示控制台
        "--icon=webicon.ico",  # 使用图标
        "--add-data=webicon.ico;.",  # 添加图标作为资源文件
        "--uac-admin",  # 请求管理员权限
        "main.py"
    ]
    
    # 执行构建
    subprocess.check_call(build_cmd)
    print("打包完成！EXE文件位于 dist/Windows优化工具.exe")

if __name__ == "__main__":
    install_dependencies()
    build_exe() 