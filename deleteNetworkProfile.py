import os
import sys
import ctypes
import winreg
from PySide6.QtWidgets import (QApplication, QWidget, QListWidget, QPushButton,
                               QVBoxLayout, QMessageBox, QListWidgetItem, QHBoxLayout, 
                               QLabel, QMenuBar, QMenu)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette, QColor, QIcon, QAction

# 全局变量&常量
versionNumber = "1.3"
versionDate = "2025-03-24"

def is_admin():
    """检查当前是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
    
def resource_path(relative_path):
    """ 动态获取资源路径（兼容开发环境和 PyInstaller 打包环境） """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 临时目录
        base_path = sys._MEIPASS
    else:
        # 开发环境当前目录
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class RegistryCleaner(QWidget):
    def __init__(self):
        super().__init__()
        # 窗口设置
        self.setWindowTitle("Windows 网络配置清理工具 - " + versionNumber)
        self.resize(800, 600)
        self.setMinimumSize(800, 600)
        # 设置窗口背景色
        palette = self.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))  # 设置窗口背景色
        self.setPalette(palette)
        
        # 创建菜单栏
        self.menu_bar = QMenuBar(self)
        
        # 创建菜单
        file_menu = QMenu("文件(&F)", self)
        help_menu = QMenu("帮助(&H)", self)
        
        # 创建菜单项
        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        
        # 将菜单项添加到菜单
        file_menu.addAction(exit_action)
        help_menu.addAction(about_action)
        
        # 将菜单添加到菜单栏
        self.menu_bar.addMenu(file_menu)
        self.menu_bar.addMenu(help_menu)

        # 设置菜单栏样式
        menu_bar_style = """
            QMenuBar {
                background-color: #ffffff;  /* 设置菜单栏背景色 */
            }
            QMenuBar::item {
                background-color: #ffffff;  /* 设置菜单项背景色 */
                color: #000000;  /* 设置菜单项文字颜色 */
                padding: 5px 10px;  /* 设置菜单项内边距 */
            }
            QMenuBar::item:selected {
                background-color: #e0e0e0;  /* 设置菜单项选中背景色 */
            }
        """
        self.menu_bar.setStyleSheet(menu_bar_style)
        

        # 创建组件
        self.label_profiles = QLabel("🍟位于 Profiles 注册表的连接过的网络")
        self.label_signatures = QLabel("🍕位于 Signatures 注册表的连接过的网络")
        self.list_widget_profiles = QListWidget()
        self.list_widget_signatures = QListWidget()
        self.btn_refresh_profiles = QPushButton("刷新")
        self.btn_refresh_signatures = QPushButton("刷新")
        self.btn_delete_profiles = QPushButton("删除选中 Profiles")
        self.btn_delete_signatures = QPushButton("删除选中 Signatures")
        self.btn_select_all_profiles = QPushButton("全选")
        self.btn_select_all_signatures = QPushButton("全选")
        
        # 按钮样式
        self.btn_refresh_profiles.setObjectName("refreshBtn")
        self.btn_refresh_signatures.setObjectName("refreshBtn")
        self.btn_delete_profiles.setObjectName("delBtn1")
        self.btn_delete_signatures.setObjectName("delBtn2")
        self.btn_select_all_profiles.setObjectName("selectAllBtn")
        self.btn_select_all_signatures.setObjectName("selectAllBtn")
        button_style = """
            QLabel {
                font-size: 14px;  /* 设置表头文字大小 */
                color: #595959;
            }
            QListWidget {
                background-color: #f5f5f5;
                border-radius: 5px;
                font-size: 16px;  /* 设置列表项文字大小 */
            }
            QListWidget::item {
                padding: 8px;  /* 设置列表项内边距 */
            }
            QPushButton {
                min-width: 60px;
                min-height: 40px;
                padding: 8px 15px;
                font-size: 14px;
                border-radius: 4px;
                background-color: #f9f9f9;
            }
            QPushButton#delBtn1 {
                background-color: #c7565b;
                color: white;
                width: 120px;
            }
            QPushButton#delBtn1:hover {
                background-color: #8f233f;
            }
            QPushButton#delBtn2 {
                background-color: #598ec4;
                color: white;
                width: 130px;
            }
            QPushButton#delBtn2:hover {
                background-color: #234c8f;
            }
            QPushButton#refreshBtn {
                background-color: #fcfcfc;
            }
            QPushButton#refreshBtn:hover {
                background-color: #ebebeb;
            }
            QPushButton#selectAllBtn {
                background-color: #fffff3;
            }
            QPushButton#selectAllBtn:hover {
                background-color: #ddddd3;
            }
            
        """
        self.setStyleSheet(button_style)
        
        # 布局设置
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 15)
        main_layout.setSpacing(15)
        
        # 添加菜单栏到布局
        main_layout.setMenuBar(self.menu_bar)

        # 水平布局容器
        list_layout = QHBoxLayout()
        
        # Profiles 列表控件及标签
        profiles_layout = QVBoxLayout()
        profiles_layout.addWidget(self.label_profiles)
        profiles_layout.addWidget(self.list_widget_profiles)
        list_layout.addLayout(profiles_layout)
        
        # Signatures 列表控件及标签
        signatures_layout = QVBoxLayout()
        signatures_layout.addWidget(self.label_signatures)
        signatures_layout.addWidget(self.list_widget_signatures)
        list_layout.addLayout(signatures_layout)
        
        main_layout.addLayout(list_layout)
        
        # 按钮容器
        btn_container = QHBoxLayout()
        btn_container.addStretch()
        btn_container.addWidget(self.btn_refresh_profiles)
        btn_container.addWidget(self.btn_select_all_profiles)
        btn_container.addWidget(self.btn_delete_profiles)
        btn_container.addSpacing(10)
        btn_container.addWidget(self.btn_refresh_signatures)
        btn_container.addWidget(self.btn_select_all_signatures)
        btn_container.addWidget(self.btn_delete_signatures)
        btn_container.addStretch()
        
        main_layout.addLayout(btn_container)
        self.setLayout(main_layout)
        
        # 连接信号槽
        self.btn_refresh_profiles.clicked.connect(self.load_registry_items_profiles)
        self.btn_refresh_signatures.clicked.connect(self.load_registry_items_signatures)
        self.btn_delete_profiles.clicked.connect(self.delete_selected_items_profiles)
        self.btn_delete_signatures.clicked.connect(self.delete_selected_items_signatures)
        self.btn_select_all_profiles.clicked.connect(self.toggle_select_all_profiles)
        self.btn_select_all_signatures.clicked.connect(self.toggle_select_all_signatures)
        
        # 初始加载数据
        self.load_registry_items_profiles()
        self.load_registry_items_signatures()

    def show_about_dialog(self):
        """显示关于对话框"""
        QMessageBox.about(self, "关于", 
                          f"这是一个用于清理 Windows 网络配置的工具。\n\n你可以用它清理多余的“网络 2”、“网络 3”\n"
                          f"以及“本地连接 2”、“本地连接 3”等条目\n\n至于为啥会有两个注册表项，咱也不清楚~🤣\n\n"
                          f"版本号：{versionNumber} ({versionDate})\n作者 Github：@manshaoco")


    def load_registry_items_profiles(self):
        """加载Profiles注册表项到列表"""
        self.list_widget_profiles.clear()
        try:
            # 打开注册表主键
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles",
                0, winreg.KEY_READ
            )
            
            # 枚举所有子键
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                try:
                    # 打开子键并读取Description
                    subkey = winreg.OpenKey(key, subkey_name)
                    description, _ = winreg.QueryValueEx(subkey, "Description")
                    winreg.CloseKey(subkey)
                    
                    # 添加带复选框的列表项
                    item = QListWidgetItem(description)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    item.setData(Qt.UserRole, subkey_name)  # 保存GUID
                    self.list_widget_profiles.addItem(item)
                except WindowsError:
                    continue
            
            winreg.CloseKey(key)
        except WindowsError as e:
            QMessageBox.critical(self, "错误", f"无法访问注册表: {str(e)}")

    def load_registry_items_signatures(self):
        """加载Signatures注册表项到列表"""
        self.list_widget_signatures.clear()
        try:
            # 打开注册表主键
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\Unmanaged",
                0, winreg.KEY_READ
            )
            
            # 枚举所有子键
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                try:
                    # 打开子键并读取Description
                    subkey = winreg.OpenKey(key, subkey_name)
                    description, _ = winreg.QueryValueEx(subkey, "Description")
                    winreg.CloseKey(subkey)
                    
                    # 添加带复选框的列表项
                    item = QListWidgetItem(description)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    item.setData(Qt.UserRole, subkey_name)  # 保存GUID
                    self.list_widget_signatures.addItem(item)
                except WindowsError:
                    continue
            
            winreg.CloseKey(key)
        except WindowsError as e:
            QMessageBox.critical(self, "错误", f"无法访问注册表: {str(e)}")

    def delete_selected_items_profiles(self):
        """删除选中的Profiles注册表项"""
        selected = []
        for i in range(self.list_widget_profiles.count()):
            item = self.list_widget_profiles.item(i)
            if item.checkState() == Qt.Checked:
                selected.append((item.text(), item.data(Qt.UserRole)))
        
        if not selected:
            QMessageBox.warning(self, "提示", "请先选择要删除的项")
            return
        
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除 {len(selected)} 个注册表项吗？此操作不可恢复！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            success = []
            failed = []
            try:
                main_key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles",
                    0, winreg.KEY_WRITE
                )
                
                for name, guid in selected:
                    try:
                        winreg.DeleteKey(main_key, guid)
                        success.append(name)
                    except WindowsError as e:
                        failed.append((name, str(e)))
                
                winreg.CloseKey(main_key)
                self.load_registry_items_profiles()  # 刷新列表
                
                # 显示操作结果
                result_msg = []
                if success:
                    result_msg.append(f"已成功删除：{', '.join(success)}")
                if failed:
                    result_msg.append(f"删除失败：")
                    result_msg.extend([f"{name}: {reason}" for name, reason in failed])
                QMessageBox.information(self, "结果", "\n".join(result_msg))
                
            except WindowsError as e:
                QMessageBox.critical(self, "错误", 
                    f"需要管理员权限！请以管理员身份运行本程序。\n错误详情：{str(e)}")

    def delete_selected_items_signatures(self):
        """删除选中的Signatures注册表项"""
        selected = []
        for i in range(self.list_widget_signatures.count()):
            item = self.list_widget_signatures.item(i)
            if item.checkState() == Qt.Checked:
                selected.append((item.text(), item.data(Qt.UserRole)))
        
        if not selected:
            QMessageBox.warning(self, "提示", "请先选择要删除的项")
            return
        
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除 {len(selected)} 个注册表项吗？此操作不可恢复！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            success = []
            failed = []
            try:
                main_key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\Unmanaged",
                    0, winreg.KEY_WRITE
                )
                
                for name, guid in selected:
                    try:
                        winreg.DeleteKey(main_key, guid)
                        success.append(name)
                    except WindowsError as e:
                        failed.append((name, str(e)))
                
                winreg.CloseKey(main_key)
                self.load_registry_items_signatures()  # 刷新列表
                
                # 显示操作结果
                result_msg = []
                if success:
                    result_msg.append(f"已成功删除：{', '.join(success)}")
                if failed:
                    result_msg.append(f"删除失败：")
                    result_msg.extend([f"{name}: {reason}" for name, reason in failed])
                QMessageBox.information(self, "结果", "\n".join(result_msg))
                
            except WindowsError as e:
                QMessageBox.critical(self, "错误", 
                    f"需要管理员权限！请以管理员身份运行本程序。\n错误详情：{str(e)}")

    def toggle_select_all_profiles(self):
        """全选/全不选 Profiles 列表项"""
        all_selected = all(self.list_widget_profiles.item(i).checkState() == Qt.Checked for i in range(self.list_widget_profiles.count()))
        for i in range(self.list_widget_profiles.count()):
            self.list_widget_profiles.item(i).setCheckState(Qt.Unchecked if all_selected else Qt.Checked)

    def toggle_select_all_signatures(self):
        """全选/全不选 Signatures 列表项"""
        all_selected = all(self.list_widget_signatures.item(i).checkState() == Qt.Checked for i in range(self.list_widget_signatures.count()))
        for i in range(self.list_widget_signatures.count()):
            self.list_widget_signatures.item(i).setCheckState(Qt.Unchecked if all_selected else Qt.Checked)

if __name__ == "__main__":
    # UAC提权逻辑
    if not is_admin():
        # 请求管理员权限
        ctypes.windll.shell32.ShellExecuteW(
            None, 
            "runas",  # 请求管理员权限
            sys.executable, 
            " ".join([f'"{arg}"' for arg in sys.argv]),  # 处理带空格的路径
            None,
            1
        )
        sys.exit()  # 退出当前非管理员进程
    
    # 管理员权限下正常启动程序
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    icon_path = resource_path("icon.png")
    app.setWindowIcon(QIcon(icon_path))
    window = RegistryCleaner()
    window.show()
    sys.exit(app.exec())