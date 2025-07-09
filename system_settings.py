import winreg
import subprocess
from admin_utils import run_as_admin

def toggle_fast_startup(enable: bool):
    """开启或关闭快速启动"""
    run_as_admin()
    key_path = r"SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    value_name = "HiberbootEnabled"
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, 1 if enable else 0)
        winreg.CloseKey(key)
        status = "开启" if enable else "关闭"
        return f"快速启动已{status}。更改将在下次重启后生效。"
    except Exception as e:
        return f"操作失败: {e}"

def restart_explorer():
    """重启文件管理器"""
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], check=True, capture_output=True)
        subprocess.run(["start", "explorer.exe"], shell=True, check=True)
        return "文件管理器已成功重启。"
    except subprocess.CalledProcessError as e:
        return f"重启文件管理器失败: {e.stderr.decode('gbk', errors='ignore')}" 