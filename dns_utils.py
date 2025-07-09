import subprocess
import os

def flush_dns_cache():
    """
    清理本地DNS缓存
    使用ipconfig /flushdns命令
    
    Returns:
        str: 操作结果信息
    """
    try:
        # 运行命令清理DNS缓存
        result = subprocess.run(
            ['ipconfig', '/flushdns'], 
            capture_output=True, 
            text=True, 
            check=True
        )
        
        # 检查命令是否成功执行
        if result.returncode == 0:
            return "DNS缓存清理成功！"
        else:
            return f"DNS缓存清理失败: {result.stderr}"
    except subprocess.CalledProcessError as e:
        return f"执行命令时出错: {str(e)}"
    except Exception as e:
        return f"清理DNS缓存时发生错误: {str(e)}" 