import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import optimizer
import os
from PIL import Image, ImageTk  # 添加PIL库用于加载ICO文件

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Windows 优化工具")
        self.geometry("600x500")
        # 设置应用图标 - 使用PIL库加载ico文件
        try:
            icon_path = "webicon.ico"
            # 检查资源文件路径（PyInstaller打包后）
            if not os.path.exists(icon_path):
                # 如果是PyInstaller打包的应用，图标可能在不同位置
                base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
                icon_path = os.path.join(base_path, "webicon.ico")
            
            # 加载图标（如果可用）
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception:
            # 如果加载图标失败，忽略错误
            pass

        # 创建主框架
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 功能列表
        features = [
            ("WIN10右键菜单", lambda: self.run_feature(optimizer.toggle_classic_context_menu, True)),
            ("WIN11右键菜单", lambda: self.run_feature(optimizer.toggle_classic_context_menu, False)),
            ("开启快速启动", lambda: self.run_feature(optimizer.toggle_fast_startup, True)),
            ("关闭快速启动", lambda: self.run_feature(optimizer.toggle_fast_startup, False)),
            ("清除临时文件", lambda: self.run_feature(optimizer.clear_temp_files)),
            ("重启文件管理器", lambda: self.run_feature(optimizer.restart_explorer)),
            # 新增功能
            ("加速系统关机", lambda: self.run_feature(optimizer.speed_up_shutdown)),
            ("设置Windows更新暂停天数", lambda: self.run_feature(optimizer.postpone_updates)),
            ("禁用透明效果", lambda: self.run_feature(optimizer.toggle_transparency, False)),
            ("启用透明效果", lambda: self.run_feature(optimizer.toggle_transparency, True)),
            # BitLocker功能
            ("解除BitLocker加密", self.disable_bitlocker),
            # DNS缓存清理功能
            ("清理DNS缓存", lambda: self.run_feature(optimizer.flush_dns_cache)),
        ]

        # 创建并放置按钮
        for i, (text, command) in enumerate(features):
            button = ttk.Button(main_frame, text=text, command=command)
            button.grid(row=i // 2, column=i % 2, padx=5, pady=5, sticky="ew")

        # 配置网格布局
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def disable_bitlocker(self):
        # 先检查是否有驱动器启用了BitLocker
        try:
            has_bitlocker, encrypted_drives = optimizer.check_bitlocker_status()
            
            if not has_bitlocker:
                messagebox.showinfo("BitLocker状态", "所有硬盘均未启用BitLocker加密，无需解除")
                return
            
            # 如果有BitLocker加密的驱动器，显示确认对话框
            drives_text = "\n".join(encrypted_drives) if encrypted_drives else "检测到加密驱动器"
            confirm = messagebox.askyesno(
                "确认", 
                f"检测到以下驱动器启用了BitLocker加密:\n{drives_text}\n\n您确定要解除这些驱动器的BitLocker加密吗？\n此操作可能需要较长时间完成，且不应中断。"
            )
            
            if confirm:
                result = optimizer.disable_all_bitlocker()
                messagebox.showinfo("操作结果", result)
        except Exception as e:
            messagebox.showerror("错误", f"检查BitLocker状态时发生错误: {str(e)}")

    def run_feature(self, func, *args):
        try:
            result = func(*args)
            messagebox.showinfo("成功", result)
        except Exception as e:
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    import sys
    if not optimizer.is_admin():
        optimizer.run_as_admin()
    else:
        app = App()
        app.mainloop()
