import sys
import os
import optimizer
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                           QHBoxLayout, QWidget, QGridLayout, QMessageBox, QLabel,
                           QGroupBox, QInputDialog, QFrame)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize

class WindowsOptimizerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 设置窗口属性
        self.setWindowTitle("Windows优化工具")
        self.setFixedSize(700, 700)  # 固定窗口大小，使其无法调整
        
        # 设置应用图标
        try:
            icon_path = "webicon.ico"
            # 检查资源文件路径（PyInstaller打包后）
            if not os.path.exists(icon_path):
                # 如果是PyInstaller打包的应用，图标可能在不同位置
                base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
                icon_path = os.path.join(base_path, "webicon.ico")
            
            # 加载图标（如果可用）
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception:
            # 如果加载图标失败，忽略错误
            pass
        
        # 设置中央窗口部件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # 创建主布局
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(10)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 添加标题
        self.setup_header()
        
        # 创建功能区域
        self.create_feature_groups()
        
        # 添加底部状态信息
        self.setup_footer()
        
        # 应用样式
        self.apply_styles()
    
    def setup_header(self):
        """创建头部标题区域"""
        header = QLabel("Windows优化工具")
        header.setAlignment(Qt.AlignCenter)
        header.setFont(QFont("Microsoft YaHei UI", 16, QFont.Bold))
        
        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        
        self.main_layout.addWidget(header)
        self.main_layout.addWidget(separator)
    
    def create_feature_groups(self):
        """创建分组功能区域"""
        # 创建分组容器布局
        groups_layout = QVBoxLayout()
        
        # 右键菜单组
        menu_group = self.create_group_box("右键菜单设置", [
            ("WIN10经典右键菜单", lambda: self.run_feature(optimizer.toggle_classic_context_menu, True)),
            ("WIN11新版右键菜单", lambda: self.run_feature(optimizer.toggle_classic_context_menu, False))
        ])
        
        # 系统性能组
        performance_group = self.create_group_box("系统性能优化", [
            ("开启快速启动", lambda: self.run_feature(optimizer.toggle_fast_startup, True)),
            ("关闭快速启动", lambda: self.run_feature(optimizer.toggle_fast_startup, False)),
            ("加速系统关机", lambda: self.run_feature(optimizer.speed_up_shutdown)),
            ("重启文件管理器", lambda: self.run_feature(optimizer.restart_explorer))
        ])
        
        # 系统清理组
        cleanup_group = self.create_group_box("系统清理", [
            ("清除临时文件", lambda: self.run_feature(optimizer.clear_temp_files)),
            ("清理DNS缓存", lambda: self.run_feature(optimizer.flush_dns_cache))
        ])
        
        # 视觉效果组
        visual_group = self.create_group_box("视觉效果", [
            ("禁用透明效果", lambda: self.run_feature(optimizer.toggle_transparency, False)),
            ("启用透明效果", lambda: self.run_feature(optimizer.toggle_transparency, True))
        ])
        
        # 系统安全组
        security_group = self.create_group_box("系统安全", [
            ("设置Windows更新暂停天数", lambda: self.run_feature(optimizer.postpone_updates)),
            ("解除BitLocker加密", self.disable_bitlocker)
        ])

        # 系统激活组
        activation_group = self.create_group_box("系统激活", [
            ("Windows系统激活", self.activate_windows),
            ("检查激活状态", self.check_activation_status)
        ])

        # 添加所有组到布局
        groups_layout.addWidget(menu_group)
        groups_layout.addWidget(performance_group)
        groups_layout.addWidget(cleanup_group)
        groups_layout.addWidget(visual_group)
        groups_layout.addWidget(security_group)
        groups_layout.addWidget(activation_group)
        
        self.main_layout.addLayout(groups_layout)
    
    def create_group_box(self, title, buttons):
        """创建分组框并添加按钮"""
        group_box = QGroupBox(title)
        group_box.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        
        # 创建网格布局
        grid = QGridLayout()
        grid.setSpacing(10)
        
        # 添加按钮到网格
        for i, (text, command) in enumerate(buttons):
            button = QPushButton(text)
            button.setMinimumHeight(40)
            button.clicked.connect(command)
            button.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            
            row = i // 2
            col = i % 2
            grid.addWidget(button, row, col)
        
        # 设置网格布局属性
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        
        group_box.setLayout(grid)
        return group_box
    
    def setup_footer(self):
        """创建底部状态信息"""
        footer = QLabel("请以管理员身份运行此工具以确保所有功能正常工作")
        footer.setAlignment(Qt.AlignCenter)
        footer.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        
        self.main_layout.addWidget(footer)
    
    def apply_styles(self):
        """应用全局样式表"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f6fa;
            }
            
            QGroupBox {
                border: 1px solid #dcdde1;
                border-radius: 5px;
                margin-top: 15px;
                background-color: #ffffff;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #2f3542;
            }
            
            QPushButton {
                background-color: #f1f2f6;
                border: 1px solid #dfe4ea;
                border-radius: 4px;
                padding: 5px;
                color: #2f3542;
            }
            
            QPushButton:hover {
                background-color: #dfe4ea;
                border: 1px solid #ced6e0;
            }
            
            QPushButton:pressed {
                background-color: #c8d6e5;
            }
            
            /* 消息框样式 */
            QMessageBox {
                font-size: 12px;
            }
            
            QMessageBox QPushButton {
                min-width: 80px;
                min-height: 30px;
                font-size: 12px;
                font-weight: bold;
            }
        """)
    
    def disable_bitlocker(self):
        """解除BitLocker加密功能"""
        # 先检查是否有驱动器启用了BitLocker
        try:
            has_bitlocker, encrypted_drives = optimizer.check_bitlocker_status()
            
            if not has_bitlocker:
                QMessageBox.information(self, "BitLocker状态", 
                                       "所有硬盘均未启用BitLocker加密，无需解除")
                return
            
            # 如果有BitLocker加密的驱动器，显示确认对话框
            drives_text = "\n".join(encrypted_drives) if encrypted_drives else "检测到加密驱动器"
            reply = QMessageBox.question(
                self, "确认", 
                f"检测到以下驱动器启用了BitLocker加密:\n{drives_text}\n\n您确定要解除这些驱动器的BitLocker加密吗？\n此操作可能需要较长时间完成，且不应中断。",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                result = optimizer.disable_all_bitlocker()
                QMessageBox.information(self, "操作结果", result)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"检查BitLocker状态时发生错误: {str(e)}")

    def activate_windows(self):
        """Windows系统激活功能"""
        try:
            # 获取可用的Windows版本
            versions = optimizer.get_windows_versions()
            recommended = optimizer.get_recommended_version()

            # 创建版本选择对话框
            from PyQt5.QtWidgets import QDialog, QVBoxLayout, QComboBox, QLabel, QPushButton, QHBoxLayout

            dialog = QDialog(self)
            dialog.setWindowTitle("选择Windows版本")
            dialog.setFixedSize(400, 200)

            layout = QVBoxLayout(dialog)

            # 添加说明标签
            info_label = QLabel("请选择您的Windows版本进行激活：")
            info_label.setFont(QFont("Microsoft YaHei UI", 10))
            layout.addWidget(info_label)

            # 添加版本选择下拉框
            version_combo = QComboBox()
            version_combo.addItems(versions)

            # 设置推荐版本为默认选择
            if recommended and recommended in versions:
                version_combo.setCurrentText(recommended)
                recommended_label = QLabel(f"推荐版本: {recommended}")
                recommended_label.setFont(QFont("Microsoft YaHei UI", 9))
                recommended_label.setStyleSheet("color: #27ae60; font-weight: bold;")
                layout.addWidget(recommended_label)

            version_combo.setFont(QFont("Microsoft YaHei UI", 10))
            layout.addWidget(version_combo)

            # 添加按钮
            button_layout = QHBoxLayout()

            activate_btn = QPushButton("开始激活")
            activate_btn.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
            activate_btn.clicked.connect(dialog.accept)

            cancel_btn = QPushButton("取消")
            cancel_btn.setFont(QFont("Microsoft YaHei UI", 10))
            cancel_btn.clicked.connect(dialog.reject)

            button_layout.addWidget(activate_btn)
            button_layout.addWidget(cancel_btn)
            layout.addLayout(button_layout)

            # 显示对话框
            if dialog.exec_() == QDialog.Accepted:
                selected_version = version_combo.currentText()

                # 显示确认对话框
                reply = QMessageBox.question(
                    self, "确认激活",
                    f"您选择激活的版本是: {selected_version}\n\n"
                    f"激活过程将执行以下步骤:\n"
                    f"1. 安装产品密钥\n"
                    f"2. 设置KMS服务器\n"
                    f"3. 激活Windows\n\n"
                    f"确定要继续吗？",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )

                if reply == QMessageBox.Yes:
                    # 执行激活
                    success, results = optimizer.activate_windows_system(selected_version)

                    # 显示结果
                    result_text = "\n".join(results)
                    if success:
                        QMessageBox.information(self, "激活成功", f"Windows激活完成！\n\n详细信息:\n{result_text}")
                    else:
                        QMessageBox.critical(self, "激活失败", f"Windows激活失败。\n\n详细信息:\n{result_text}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"激活过程中发生错误: {str(e)}")

    def check_activation_status(self):
        """检查Windows激活状态"""
        try:
            # 调用激活状态检查，Windows会自动弹出系统对话框显示状态
            optimizer.check_windows_activation()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"检查激活状态时发生错误: {str(e)}")

    def run_feature(self, func, *args):
        """运行优化功能并显示结果"""
        try:
            result = func(*args)

            # 特殊处理返回元组的函数（如激活状态检查）
            if isinstance(result, tuple) and len(result) == 2:
                success, message = result
                if success:
                    QMessageBox.information(self, "成功", message)
                else:
                    QMessageBox.critical(self, "错误", message)
            else:
                # 普通字符串结果
                QMessageBox.information(self, "成功", result)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))


def set_message_box_defaults():
    """设置消息框默认样式"""
    # 设置默认按钮大小
    font = QFont("Microsoft YaHei UI", 10, QFont.Bold)
    QApplication.setFont(font, "QMessageBox")
    QApplication.setFont(font, "QPushButton")

if __name__ == "__main__":
    # 检查管理员权限
    if not optimizer.is_admin():
        optimizer.run_as_admin()
    else:
        # 创建应用和窗口
        app = QApplication(sys.argv)
        
        # 设置消息框默认样式
        set_message_box_defaults()
        
        window = WindowsOptimizerApp()
        window.show()
        sys.exit(app.exec_())
