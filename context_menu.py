import winreg
from admin_utils import run_as_admin

def toggle_classic_context_menu(enable: bool):
    """恢复或禁用经典右键菜单"""
    run_as_admin()
    key_path = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    if enable:
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
            winreg.CloseKey(key)
            return "经典右键菜单已启用。请重启文件管理器以生效。"
        except Exception as e:
            return f"启用经典右键菜单失败: {e}"
    else:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            return "经典右键菜单已禁用。请重启文件管理器以生效。"
        except FileNotFoundError:
            return "经典右键菜单未启用。"
        except Exception as e:
            return f"禁用经典右键菜单失败: {e}" 