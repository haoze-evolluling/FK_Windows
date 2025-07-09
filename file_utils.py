import os
import shutil
from pathlib import Path

def clear_system_directories():
    """根据清理目录列表清除系统中的各种临时文件和缓存"""
    results = {}
    total_cleared = 0
    all_errors = []
    
    # 定义需要清理的目录列表
    directories_to_clean = [
        # 1. 用户临时文件夹
        os.environ.get("TEMP"),
        # 2. 系统临时文件夹
        r"C:\Windows\Temp",
        # 3. 预读取文件夹
        r"C:\Windows\Prefetch",
        # 4. Windows 更新缓存
        r"C:\Windows\SoftwareDistribution\Download",
        # 5. 旧的 Windows 安装文件
        r"C:\Windows.old",
        # 6. 浏览器缓存
        # Edge
        os.path.join(os.environ.get("LOCALAPPDATA"), r"Microsoft\Edge\User Data\Default\Cache"),
        # Chrome
        os.path.join(os.environ.get("LOCALAPPDATA"), r"Google\Chrome\User Data\Default\Cache"),
        # Firefox
        os.path.join(os.environ.get("LOCALAPPDATA"), r"Mozilla\Firefox\Profiles"),
        # 7. 回收站
        r"C:\$Recycle.Bin"
    ]
    
    # 过滤掉不存在的目录
    directories_to_clean = [dir_path for dir_path in directories_to_clean if dir_path and os.path.exists(dir_path)]
    
    for directory in directories_to_clean:
        cleared_count = 0
        errors = []
        
        try:
            # 遍历目录中的所有项目
            for item in os.listdir(directory):
                path = os.path.join(directory, item)
                try:
                    if os.path.isfile(path) or os.path.islink(path):
                        os.unlink(path)
                        cleared_count += 1
                    elif os.path.isdir(path):
                        shutil.rmtree(path)
                        cleared_count += 1
                except Exception as e:
                    errors.append(f"{path}: {str(e)}")
            
            # 保存该目录的清理结果
            results[directory] = {
                "cleared_count": cleared_count,
                "errors": errors
            }
            total_cleared += cleared_count
            all_errors.extend([f"[{directory}] {err}" for err in errors])
        
        except Exception as e:
            all_errors.append(f"无法访问目录 {directory}: {str(e)}")
    
    # 构建返回消息
    message = f"总共清除了 {total_cleared} 个项目。\n\n"
    
    for directory, result in results.items():
        message += f"目录: {directory}\n"
        message += f"- 清除了 {result['cleared_count']} 个项目\n"
        if result['errors']:
            message += f"- 错误: {len(result['errors'])} 个\n"
    
    if all_errors:
        message += "\n错误详情:\n" + "\n".join(all_errors)
    
    return message

# 保留原始函数作为向后兼容
def clear_temp_files():
    """清除临时文件 (旧函数，为保持兼容性而保留)"""
    return clear_system_directories() 