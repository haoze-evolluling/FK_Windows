import subprocess
import ctypes
from admin_utils import run_as_admin

def speed_up_shutdown():
    """
    加速Win10关闭速度，通过组策略设置自动终止阻止关机的应用程序
    """
    run_as_admin()
    
    try:
        # 使用组策略编辑器自动设置
        cmd_commands = [
            r'REG ADD "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v AutoEndTasks /t REG_DWORD /d 1 /f',
            r'REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control" /v WaitToKillServiceTimeout /t REG_SZ /d 2000 /f',
            r'REG ADD "HKEY_CURRENT_USER\Control Panel\Desktop" /v WaitToKillAppTimeout /t REG_SZ /d 2000 /f',
            r'REG ADD "HKEY_CURRENT_USER\Control Panel\Desktop" /v HungAppTimeout /t REG_SZ /d 2000 /f',
            r'REG ADD "HKEY_CURRENT_USER\Control Panel\Desktop" /v AutoEndTasks /t REG_SZ /d 1 /f',
            r'REG ADD "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" /v ClearPageFileAtShutdown /t REG_DWORD /d 0 /f'
        ]
        
        for cmd in cmd_commands:
            subprocess.run(cmd, shell=True, check=True)
            
        return "关机速度优化已成功应用。系统将自动终止阻止关机的应用程序。"
    
    except subprocess.CalledProcessError as e:
        return f"关机速度优化失败: {str(e)}"
    except Exception as e:
        return f"关机速度优化过程中发生错误: {str(e)}" 