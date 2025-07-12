"""
Windows系统激活工具模块
提供Windows系统KMS激活功能
"""

import subprocess
import random
import platform
import re
from typing import Dict, List, Tuple, Optional


class WindowsActivator:
    """Windows系统激活器"""

    # Windows版本对应的激活密钥
    ACTIVATION_KEYS = {
        "Windows 11 企业版 LTSC 2024": "M7XTQ-FN8P6-TTKYV-9D4CC-J462D",
        "Windows 11 专业版": "W269N-WFGWX-YVC9B-4J6C9-T83GX",
        "Windows 10 专业版": "W269N-WFGWX-YVC9B-4J6C9-T83GX",
        "Windows 10 企业版 LTSC 2019": "M7XTQ-FN8P6-TTKYV-9D4CC-J462D",
        "Windows 10 专业工作站版": "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J",
        "Windows 8 专业版": "NG4HW-VH26C-733KW-K6F98-J8CK4",
        "Windows 7 专业版": "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"
    }

    # KMS激活服务器列表
    KMS_SERVERS = [
        "zh.us.to",
        "kms.loli.beer",
        "kms.loli.best",
        "kms.03k.org",
        "kms-default.cangshui.net",
        "kms.cgtsoft.com"
    ]

    def __init__(self):
        self.current_system_info = self._get_system_info()

    def _get_system_info(self) -> Dict[str, str]:
        """获取当前系统信息"""
        try:
            # 获取Windows版本信息
            system_info = platform.platform()
            version = platform.version()
            release = platform.release()

            return {
                "platform": system_info,
                "version": version,
                "release": release,
                "full_info": f"{system_info} {version}"
            }
        except Exception as e:
            return {
                "platform": "Unknown",
                "version": "Unknown",
                "release": "Unknown",
                "full_info": f"获取系统信息失败: {str(e)}"
            }

    def get_available_versions(self) -> List[str]:
        """获取可用的Windows版本列表"""
        return list(self.ACTIVATION_KEYS.keys())

    def get_recommended_version(self) -> Optional[str]:
        """根据当前系统推荐合适的版本"""
        system_info = self.current_system_info["full_info"].lower()

        # 简单的版本匹配逻辑
        if "windows 11" in system_info:
            if "enterprise" in system_info or "ltsc" in system_info:
                return "Windows 11 企业版 LTSC 2024"
            else:
                return "Windows 11 专业版"
        elif "windows 10" in system_info:
            if "enterprise" in system_info and "ltsc" in system_info:
                return "Windows 10 企业版 LTSC 2019"
            elif "pro for workstations" in system_info:
                return "Windows 10 专业工作站版"
            else:
                return "Windows 10 专业版"
        elif "windows 8" in system_info:
            return "Windows 8 专业版"
        elif "windows 7" in system_info:
            return "Windows 7 专业版"

        # 默认推荐Windows 11专业版
        return "Windows 11 专业版"

    def get_random_kms_server(self) -> str:
        """随机选择一个KMS服务器"""
        return random.choice(self.KMS_SERVERS)

    def _run_powershell_command(self, command: str) -> Tuple[bool, str]:
        """执行PowerShell命令"""
        try:
            # 使用PowerShell执行slmgr命令
            full_command = f"powershell.exe -Command \"{command}\""

            result = subprocess.run(
                full_command,
                shell=True,
                capture_output=True,
                text=True,
                timeout=60  # 60秒超时
            )

            if result.returncode == 0:
                return True, result.stdout.strip()
            else:
                error_msg = result.stderr.strip() if result.stderr else "命令执行失败"
                return False, error_msg

        except subprocess.TimeoutExpired:
            return False, "命令执行超时"
        except Exception as e:
            return False, f"执行命令时发生错误: {str(e)}"

    def install_product_key(self, version: str) -> Tuple[bool, str]:
        """安装产品密钥"""
        if version not in self.ACTIVATION_KEYS:
            return False, f"不支持的Windows版本: {version}"

        key = self.ACTIVATION_KEYS[version]
        command = f"slmgr /ipk {key}"

        success, message = self._run_powershell_command(command)

        if success:
            return True, f"成功安装产品密钥 ({version})"
        else:
            return False, f"安装产品密钥失败: {message}"

    def set_kms_server(self, server: Optional[str] = None) -> Tuple[bool, str]:
        """设置KMS服务器"""
        if server is None:
            server = self.get_random_kms_server()

        command = f"slmgr /skms {server}"

        success, message = self._run_powershell_command(command)

        if success:
            return True, f"成功设置KMS服务器: {server}"
        else:
            return False, f"设置KMS服务器失败: {message}"

    def activate_windows(self) -> Tuple[bool, str]:
        """激活Windows"""
        command = "slmgr /ato"

        success, message = self._run_powershell_command(command)

        if success:
            return True, "Windows激活成功！"
        else:
            return False, f"Windows激活失败: {message}"

    def full_activation_process(self, version: str, server: Optional[str] = None) -> Tuple[bool, List[str]]:
        """完整的激活流程"""
        results = []

        # 步骤1: 安装产品密钥
        success1, msg1 = self.install_product_key(version)
        results.append(f"步骤1 - 安装产品密钥: {msg1}")

        if not success1:
            return False, results

        # 步骤2: 设置KMS服务器
        success2, msg2 = self.set_kms_server(server)
        results.append(f"步骤2 - 设置KMS服务器: {msg2}")

        if not success2:
            return False, results

        # 步骤3: 激活Windows
        success3, msg3 = self.activate_windows()
        results.append(f"步骤3 - 激活Windows: {msg3}")

        return success3, results

    def check_activation_status(self) -> Tuple[bool, str]:
        """检查当前激活状态"""
        command = "slmgr /xpr"

        success, message = self._run_powershell_command(command)

        if success:
            return True, message
        else:
            return False, f"检查激活状态失败: {message}"


# 便捷函数接口
def get_windows_versions() -> List[str]:
    """获取支持的Windows版本列表"""
    activator = WindowsActivator()
    return activator.get_available_versions()


def get_recommended_version() -> Optional[str]:
    """获取推荐的Windows版本"""
    activator = WindowsActivator()
    return activator.get_recommended_version()


def activate_windows_system(version: str, server: Optional[str] = None) -> Tuple[bool, List[str]]:
    """激活Windows系统"""
    activator = WindowsActivator()
    return activator.full_activation_process(version, server)


def check_windows_activation() -> Tuple[bool, str]:
    """检查Windows激活状态"""
    activator = WindowsActivator()
    return activator.check_activation_status()


def get_system_info() -> Dict[str, str]:
    """获取系统信息"""
    activator = WindowsActivator()
    return activator.current_system_info