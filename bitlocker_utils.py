import subprocess
import re

def disable_bitlocker(drive_letter):
    """
    解除指定驱动器的BitLocker加密
    
    Args:
        drive_letter: 驱动器盘符，例如 "C:"
    
    Returns:
        str: 操作结果信息
    """
    if not drive_letter.endswith(':'):
        drive_letter = f"{drive_letter}:"
    
    try:
        # 执行PowerShell命令解除BitLocker
        command = f'powershell -Command "Disable-BitLocker -MountPoint \'{drive_letter}\'"'
        subprocess.run(command, shell=True, check=True)
        return f"已开始解除{drive_letter}驱动器的BitLocker加密，此过程可能需要一段时间完成"
    except subprocess.CalledProcessError as e:
        return f"解除BitLocker时发生错误: {str(e)}"

def check_bitlocker_status():
    """
    检查所有驱动器的BitLocker状态，确认是否存在BitLocker加密
    
    Returns:
        tuple: (是否存在BitLocker加密, 已启用BitLocker的驱动器列表)
    """
    try:
        command = 'powershell -Command "Get-BitLockerVolume | Format-List MountPoint, VolumeStatus, EncryptionPercentage"'
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        
        encrypted_drives = []
        drive_info = {}
        current_key = None
        
        # 使用更可靠的方式解析输出
        for line in result.stdout.split('\n'):
            line = line.strip()
            if not line:
                # 空行表示一个驱动器信息结束，处理收集到的信息
                if 'MountPoint' in drive_info and 'VolumeStatus' in drive_info:
                    mount_point = drive_info.get('MountPoint', '')
                    volume_status = drive_info.get('VolumeStatus', '')
                    
                    # 检查是否加密
                    if ('FullyEncrypted' in volume_status or 'EncryptionInProgress' in volume_status) and mount_point:
                        encrypted_drives.append(mount_point)
                
                # 重置收集信息
                drive_info = {}
                continue
                
            if ':' in line:
                parts = line.split(':', 1)
                key = parts[0].strip()
                value = parts[1].strip() if len(parts) > 1 else ""
                
                if key == 'MountPoint':
                    drive_info['MountPoint'] = value
                elif key == 'VolumeStatus':
                    drive_info['VolumeStatus'] = value
        
        # 处理最后一个驱动器信息
        if 'MountPoint' in drive_info and 'VolumeStatus' in drive_info:
            mount_point = drive_info.get('MountPoint', '')
            volume_status = drive_info.get('VolumeStatus', '')
            if ('FullyEncrypted' in volume_status or 'EncryptionInProgress' in volume_status) and mount_point:
                encrypted_drives.append(mount_point)
        
        return len(encrypted_drives) > 0, encrypted_drives
    except subprocess.CalledProcessError as e:
        # 如果命令失败，保守起见假设存在BitLocker
        return True, []

def disable_all_bitlocker():
    """
    解除所有驱动器的BitLocker加密
    
    Returns:
        str: 操作结果信息
    """
    # 检查是否有驱动器启用了BitLocker
    has_bitlocker, encrypted_drives = check_bitlocker_status()
    
    if not has_bitlocker:
        return "所有硬盘均未启用BitLocker加密，无需解除"
    
    # 获取所有驱动器
    drives = encrypted_drives if encrypted_drives else list_available_drives()
    results = []
    
    if not drives:
        return "未找到可用驱动器"
    
    for drive in drives:
        try:
            # 执行PowerShell命令解除BitLocker
            command = f'powershell -Command "Disable-BitLocker -MountPoint \'{drive}\'"'
            subprocess.run(command, shell=True, check=False)
            results.append(f"{drive}驱动器的BitLocker解除已开始")
        except Exception as e:
            results.append(f"{drive}驱动器解除失败: {str(e)}")
    
    return "已开始全局解除BitLocker加密，结果如下:\n" + "\n".join(results)

def get_bitlocker_status(drive_letter=None):
    """
    获取BitLocker状态
    
    Args:
        drive_letter: 可选，指定驱动器盘符
    
    Returns:
        str: BitLocker状态信息
    """
    try:
        if drive_letter:
            if not drive_letter.endswith(':'):
                drive_letter = f"{drive_letter}:"
            command = f'powershell -Command "Get-BitLockerVolume -MountPoint \'{drive_letter}\'"'
        else:
            command = 'powershell -Command "Get-BitLockerVolume"'
        
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        return f"获取BitLocker状态时发生错误: {str(e)}"

def list_available_drives():
    """
    列出系统中所有可用的驱动器
    
    Returns:
        list: 驱动器盘符列表
    """
    try:
        command = 'powershell -Command "Get-PSDrive -PSProvider FileSystem | Select-Object Name, Root"'
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        
        # 改进正则表达式以更可靠地匹配驱动器信息
        drives = []
        pattern = r'([A-Za-z])\s+([A-Za-z]:)'
        for line in result.stdout.split('\n'):
            match = re.search(pattern, line)
            if match:
                drives.append(match.group(2))
        
        # 如果上面的方法没有找到驱动器，尝试另一种解析方式
        if not drives:
            for line in result.stdout.split('\n'):
                if ':' in line:
                    parts = line.split()
                    for part in parts:
                        if re.match(r'^[A-Za-z]:$', part):
                            drives.append(part)
        
        return drives
    except subprocess.CalledProcessError as e:
        # 添加错误日志，但仍然返回空列表
        print(f"获取驱动器列表时发生错误: {str(e)}")
        return [] 