import winreg
from admin_utils import run_as_admin

def postpone_updates(days=10000):
    """
    设置Windows更新的最大暂停天数
    
    Args:
        days: 可暂停的最大天数，默认为10000天
    """
    run_as_admin()
    
    try:
        # 设置注册表项
        key_path = r"SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
        key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_WRITE)
        
        # 设置最大暂停天数
        winreg.SetValueEx(key, "FlightSettingsMaxPauseDays", 0, winreg.REG_DWORD, days)
        
        return f"Windows更新最大暂停天数已设置为{days}天，请在Windows更新设置中选择暂停更新。"
    
    except Exception as e:
        return f"设置Windows更新暂停天数失败: {str(e)}" 