import winreg
import subprocess
from admin_utils import run_as_admin

def toggle_transparency(enable=False):
    """
    启用或禁用Windows透明效果
    
    Args:
        enable: 是否启用透明效果，默认为False（禁用）
    """
    run_as_admin()
    
    try:
        # 设置注册表项
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        
        # 设置透明效果
        # 1 = 启用透明效果，0 = 禁用透明效果
        value = 1 if enable else 0
        winreg.SetValueEx(key, "EnableTransparency", 0, winreg.REG_DWORD, value)
        
        # 设置系统级透明效果
        key_path_adv = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        key_adv = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path_adv, 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key_adv, "UseOLEDTaskbarTransparency", 0, winreg.REG_DWORD, 1 if enable else 0)
        
        # 刷新桌面，使设置立即生效
        try:
            # 使用rundll32重新加载Windows桌面
            subprocess.run(["rundll32.exe", "user32.dll,UpdatePerUserSystemParameters"], check=True, shell=True)
        except:
            # 如果无法重新加载，至少提示用户可能需要注销或重启
            pass
        
        status = "启用" if enable else "禁用"
        return f"Windows透明效果已{status}。您可能需要注销或重启以完全应用更改。"
    
    except Exception as e:
        return f"设置透明效果失败: {str(e)}" 