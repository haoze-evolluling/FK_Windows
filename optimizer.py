# 主模块 - 导入并暴露所有功能
from admin_utils import is_admin, run_as_admin
from context_menu import toggle_classic_context_menu
from system_settings import toggle_fast_startup, restart_explorer
from file_utils import clear_temp_files
# 导入新功能
from shutdown_optimizer import speed_up_shutdown
from update_manager import postpone_updates
from visual_effects import toggle_transparency
from bitlocker_utils import disable_bitlocker, disable_all_bitlocker, check_bitlocker_status, get_bitlocker_status, list_available_drives
from dns_utils import flush_dns_cache

# 导出所有功能，保持与原来相同的接口
__all__ = [
    'is_admin',
    'run_as_admin',
    'toggle_classic_context_menu',
    'toggle_fast_startup',
    'clear_temp_files',
    'restart_explorer',
    # 添加新功能
    'speed_up_shutdown',
    'postpone_updates',
    'toggle_transparency',
    'disable_bitlocker',
    'disable_all_bitlocker',
    'check_bitlocker_status',
    'get_bitlocker_status',
    'list_available_drives',
    'flush_dns_cache'
]

