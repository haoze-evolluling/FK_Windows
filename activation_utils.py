"""
Windows系统激活工具模块
提供Windows系统KMS激活功能
"""

import subprocess
import random
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
            # 使用多种方法获取详细的Windows版本信息
            system_info = {}

            # 方法1: 使用PowerShell获取详细系统信息
            detailed_info = self._get_detailed_windows_info()
            system_info.update(detailed_info)

            # 方法2: 使用wmic获取产品信息
            product_info = self._get_windows_product_info()
            system_info.update(product_info)

            # 构建完整信息字符串
            full_info_parts = []
            if system_info.get("os_name"):
                full_info_parts.append(system_info["os_name"])
            if system_info.get("edition"):
                full_info_parts.append(system_info["edition"])
            if system_info.get("version_number"):
                full_info_parts.append(f"版本 {system_info['version_number']}")
            if system_info.get("build_number"):
                full_info_parts.append(f"内部版本 {system_info['build_number']}")

            system_info["full_info"] = " ".join(full_info_parts) if full_info_parts else "Unknown"

            return system_info

        except Exception as e:
            return {
                "full_info": f"获取系统信息失败: {str(e)}"
            }

    def _get_detailed_windows_info(self) -> Dict[str, str]:
        """使用PowerShell获取详细的Windows信息"""
        try:
            # 使用更简单的PowerShell命令格式
            commands = [
                # 获取操作系统信息
                'powershell.exe -NoProfile -Command "Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber, OSArchitecture | Format-List"',
                # 备用命令
                'powershell.exe -NoProfile -Command "(Get-CimInstance Win32_OperatingSystem).Caption"',
            ]

            info = {}

            for command in commands:
                try:
                    result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=15)

                    if result.returncode == 0 and result.stdout.strip():
                        output = result.stdout.strip()

                        # 解析Format-List输出
                        if ':' in output:
                            lines = output.split('\n')
                            for line in lines:
                                if ':' in line:
                                    key, value = line.split(':', 1)
                                    key = key.strip()
                                    value = value.strip()

                                    if 'Caption' in key and value:
                                        info["os_name"] = value
                                    elif 'Version' in key and 'BuildNumber' not in key and value:
                                        info["version_number"] = value
                                    elif 'BuildNumber' in key and value:
                                        info["build_number"] = value
                                    elif 'OSArchitecture' in key and value:
                                        info["architecture"] = value
                        else:
                            # 如果是简单的Caption输出
                            if "Windows" in output:
                                info["os_name"] = output

                        # 如果获取到了基本信息就跳出循环
                        if info.get("os_name"):
                            break

                except Exception:
                    continue

            return info
        except Exception:
            return {}

    def _get_windows_product_info(self) -> Dict[str, str]:
        """使用wmic获取Windows产品信息"""
        try:
            # 尝试多种wmic命令格式
            commands = [
                'wmic os get Caption,Version,BuildNumber /format:list',
                'wmic os get Caption /format:list',
                'wmic os get Caption,Version,BuildNumber',
            ]

            info = {}

            for command in commands:
                try:
                    result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=15)

                    if result.returncode == 0 and result.stdout.strip():
                        output = result.stdout.strip()

                        # 解析/format:list输出
                        if '=' in output:
                            lines = output.split('\n')
                            for line in lines:
                                if '=' in line and line.strip():
                                    key, value = line.split('=', 1)
                                    key = key.strip()
                                    value = value.strip()

                                    if key == 'Caption' and value:
                                        # 清理Caption信息
                                        clean_caption = value.replace('Microsoft Windows ', '').replace('Microsoft ', '').strip()
                                        info["edition"] = clean_caption
                                        if not info.get("os_name"):
                                            info["os_name"] = value
                                    elif key == 'Version' and value:
                                        info["wmic_version"] = value
                                    elif key == 'BuildNumber' and value:
                                        info["wmic_build"] = value
                                        if not info.get("build_number"):
                                            info["build_number"] = value
                        else:
                            # 解析表格格式输出
                            lines = output.split('\n')
                            for line in lines:
                                line = line.strip()
                                if not line:
                                    continue

                                # 查找包含Windows的行
                                if 'Windows' in line:
                                    info["os_name"] = line
                                    # 尝试提取版本信息
                                    clean_caption = line.replace('Microsoft Windows ', '').replace('Microsoft ', '').strip()
                                    info["edition"] = clean_caption
                                    break

                        # 如果获取到了基本信息就跳出循环
                        if info.get("os_name") or info.get("edition"):
                            break

                except Exception:
                    continue

            return info
        except Exception:
            return {}

    def get_available_versions(self) -> List[str]:
        """获取可用的Windows版本列表"""
        return list(self.ACTIVATION_KEYS.keys())

    def get_recommended_version(self) -> Optional[str]:
        """根据当前系统推荐合适的版本"""
        system_info = self.current_system_info

        # 获取系统信息的各个字段
        full_info = system_info.get("full_info", "").lower()
        os_name = system_info.get("os_name", "").lower()
        edition = system_info.get("edition", "").lower()
        build_number = system_info.get("build_number", "")

        # 调试信息（可在需要时启用）
        # print(f"系统信息调试:")
        # print(f"  完整信息: {full_info}")
        # print(f"  操作系统: {os_name}")
        # print(f"  版本: {edition}")
        # print(f"  内部版本号: {build_number}")

        # 根据内部版本号判断Windows版本
        try:
            build_num = int(build_number) if build_number else 0
        except (ValueError, TypeError):
            build_num = 0

        # Windows 11 (内部版本号 >= 22000)
        if build_num >= 22000 or "windows 11" in full_info or "windows 11" in os_name:
            if any(keyword in edition for keyword in ["enterprise", "企业版"]) or \
               any(keyword in full_info for keyword in ["enterprise", "ltsc", "企业版"]):
                return "Windows 11 企业版 LTSC 2024"
            elif any(keyword in edition for keyword in ["pro", "professional", "专业版"]) or \
                 any(keyword in full_info for keyword in ["pro", "professional", "专业版"]):
                return "Windows 11 专业版"
            else:
                # 默认推荐专业版
                return "Windows 11 专业版"

        # Windows 10 (内部版本号 10240-21999)
        elif build_num >= 10240 or "windows 10" in full_info or "windows 10" in os_name:
            if any(keyword in edition for keyword in ["enterprise", "企业版"]) and \
               any(keyword in full_info for keyword in ["ltsc", "2019"]):
                return "Windows 10 企业版 LTSC 2019"
            elif any(keyword in edition for keyword in ["pro for workstations", "专业工作站"]) or \
                 any(keyword in full_info for keyword in ["workstation", "工作站"]):
                return "Windows 10 专业工作站版"
            elif any(keyword in edition for keyword in ["pro", "professional", "专业版"]) or \
                 any(keyword in full_info for keyword in ["pro", "professional", "专业版"]):
                return "Windows 10 专业版"
            else:
                # 默认推荐专业版
                return "Windows 10 专业版"

        # Windows 8/8.1
        elif "windows 8" in full_info or "windows 8" in os_name:
            return "Windows 8 专业版"

        # Windows 7
        elif "windows 7" in full_info or "windows 7" in os_name:
            return "Windows 7 专业版"

        # 如果无法识别，根据内部版本号推测
        elif build_num > 0:
            if build_num >= 22000:
                return "Windows 11 专业版"
            elif build_num >= 10240:
                return "Windows 10 专业版"
            else:
                return "Windows 10 专业版"  # 保守选择

        # 最后的默认选择
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


