import sys
import os
import winreg
import time
import psutil
import threading
import wmi
import pythoncom

from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
                               QCheckBox, QLabel, QMessageBox, QHeaderView, QFileDialog,
                               QLineEdit, QSystemTrayIcon, QMenu)
from PySide6.QtGui import QIcon, QAction


def get_resource_path(relative_path):
    """ 获取资源绝对路径，兼容 PyInstaller 打包后的资源定位 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class GuardSignals(QObject):
    """ 跨线程信号：用于 WMI 监听线程通知主界面弹出对话框 """
    request_ask = Signal(str)


class GPUSwitch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPUSwitch")
        self.setMinimumSize(1000, 650)

        # 1. 变量初始化
        self.reg_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        self.ask_list = set()  # 存放需要运行时询问的程序路径
        self.cooldowns = {}  # 防止短时间内重复触发弹窗
        self.pending_changes = set()  # 暂存用户在界面上修改但未保存的项
        self.is_running = True  # 控制子线程运行标志
        self.really_quit = False  # 是否彻底退出程序的标志

        # 2. 信号绑定
        self.signals = GuardSignals()
        self.signals.request_ask.connect(self.show_ask_dialog)

        # 3. UI 初始化
        self.setup_ui()
        self.load_apps()

        # 4. 托盘初始化
        self.init_tray()

        # 5. 设置图标（兼容打包路径）
        icon_path = get_resource_path("app_icon.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.tray_icon.setIcon(QIcon(icon_path))

        # 6. 开启后台监听线程
        self.start_guard_thread()

    def setup_ui(self):
        """ 构建主界面布局 """
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)

        # --- 顶部工具栏 ---
        top_bar = QHBoxLayout()
        self.btn_add = QPushButton("➕ 添加新程序")
        self.btn_add.clicked.connect(self.add_new_app)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 搜索程序名称或路径...")
        self.search_input.setFixedWidth(300)
        self.search_input.textChanged.connect(self.filter_table)

        self.btn_refresh = QPushButton("🔄 刷新列表")
        self.btn_refresh.clicked.connect(self.load_apps)

        top_bar.addWidget(self.btn_add)
        top_bar.addWidget(self.search_input)
        top_bar.addStretch()
        top_bar.addWidget(self.btn_refresh)
        self.main_layout.addLayout(top_bar)

        # --- 程序列表表格 ---
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["程序名称", "文件完整路径", "显卡偏好设置", "运行时询问"])
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 450)
        self.table.setColumnWidth(2, 150)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.main_layout.addWidget(self.table)

        # --- 底部状态栏 ---
        bottom_bar = QHBoxLayout()
        self.status_label = QLabel("列表已更新")
        self.status_label.setStyleSheet("color: #31C950;")

        # 最小化到托盘复选框
        self.check_minimize_to_tray = QCheckBox("最小化到系统托盘")
        self.check_minimize_to_tray.setChecked(True)

        # 开机自启动复选框
        self.check_autostart = QCheckBox("开机自动启动")
        self.check_autostart.setChecked(self.is_autostart_enabled())
        self.check_autostart.stateChanged.connect(self.toggle_autostart)

        self.btn_apply = QPushButton("💾 应用所有更改")
        self.btn_apply.setEnabled(False)
        self.btn_apply.clicked.connect(self.apply_all_changes)

        bottom_bar.addWidget(self.status_label)
        bottom_bar.addSpacing(20)
        bottom_bar.addWidget(self.check_minimize_to_tray)
        bottom_bar.addWidget(self.check_autostart)
        bottom_bar.addStretch()
        bottom_bar.addWidget(self.btn_apply)
        self.main_layout.addLayout(bottom_bar)

    def init_tray(self):
        """ 初始化系统托盘及菜单 """
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setToolTip("GPUSwitch - 显卡切换守护者")

        tray_menu = QMenu()
        show_action = QAction("显示主界面", self)
        show_action.triggered.connect(self.show_and_activate)
        quit_action = QAction("彻底退出", self)
        quit_action.triggered.connect(self.quit_app)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def show_and_activate(self):
        """ 恢复窗口并置顶 """
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def on_tray_icon_activated(self, reason):
        """ 处理托盘图标点击事件 """
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            if self.isVisible():
                self.hide()
            else:
                self.show_and_activate()

    def closeEvent(self, event):
        """ 关闭按钮逻辑改进 """
        if self.really_quit:
            event.accept()
            return

        # 逻辑：如果勾选了最小化到托盘，直接隐藏，不再弹窗
        if self.check_minimize_to_tray.isChecked():
            event.ignore()
            self.hide()
        else:
            # 未勾选时，询问用户是退出还是去托盘
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("退出确认")
            msg_box.setText("您点击了关闭按钮，请选择操作：")
            btn_tray = msg_box.addButton("最小化到托盘", QMessageBox.AcceptRole)
            btn_quit = msg_box.addButton("完全退出程序", QMessageBox.DestructiveRole)
            msg_box.setDefaultButton(btn_tray)
            msg_box.exec()

            if msg_box.clickedButton() == btn_tray:
                self.check_minimize_to_tray.setChecked(True)  # 自动勾选，下次不询问
                event.ignore()
                self.hide()
            else:
                self.quit_app(ask=False)  # 直接执行退出清理

    def quit_app(self, ask=True):
        """ 执行彻底退出流程 """
        if ask:
            reply = QMessageBox.question(self, '确认退出', "退出后将停止所有进程监听，确定吗？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        self.is_running = False  # 停止子线程循环
        self.really_quit = True
        self.tray_icon.hide()
        QApplication.quit()

    def load_apps(self):
        """ 从注册表加载应用列表 """
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.ask_list.clear()
        self.pending_changes.clear()
        self.btn_apply.setEnabled(False)

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_READ)
            i = 0
            while True:
                try:
                    path, value, _ = winreg.EnumValue(key, i)
                    if not isinstance(value, str):
                        i += 1
                        continue

                    is_ask = "ASK" in value
                    if is_ask: self.ask_list.add(path.lower())
                    self.add_row(path, value, is_ask)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except OSError:
            pass

        self.status_label.setText("列表已更新")
        self.status_label.setStyleSheet("color: #31C950;")
        self.table.setSortingEnabled(True)

    def add_row(self, path, val, is_ask):
        """ 向表格插入一行数据 """
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(os.path.basename(path)))
        self.table.setItem(row, 1, QTableWidgetItem(path))

        combo = QComboBox()
        combo.addItems(["系统默认", "省电 (集显)", "高性能 (独显)"])
        clean_val = val.replace("ASK", "")
        m = {"GpuPreference=0;": 0, "GpuPreference=1;": 1, "GpuPreference=2;": 2}
        combo.setCurrentIndex(m.get(clean_val, 0))
        combo.currentIndexChanged.connect(lambda: self.mark_as_changed(path))
        self.table.setCellWidget(row, 2, combo)

        check = QCheckBox()
        check.setChecked(is_ask)
        check.stateChanged.connect(lambda: self.mark_as_changed(path))
        self.table.setCellWidget(row, 3, check)

    def mark_as_changed(self, path):
        """ 记录被修改的行 """
        self.pending_changes.add(path)
        self.status_label.setText(f"提示：有 {len(self.pending_changes)} 处待修改的项")
        self.status_label.setStyleSheet("color: #e67e22; font-weight: bold;")
        self.btn_apply.setEnabled(True)

    def apply_all_changes(self):
        """ 批量保存更改到注册表 """
        for i in range(self.table.rowCount()):
            path = self.table.item(i, 1).text()
            if path in self.pending_changes:
                gpu_idx = self.table.cellWidget(i, 2).currentIndex()
                is_ask = self.table.cellWidget(i, 3).isChecked()
                val = f"GpuPreference={gpu_idx};"
                if is_ask:
                    val += "ASK"
                    self.ask_list.add(path.lower())
                else:
                    self.ask_list.discard(path.lower())

                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
                winreg.SetValueEx(key, path, 0, winreg.REG_SZ, val)
                winreg.CloseKey(key)

        QMessageBox.information(self, "成功", "所有更改已成功应用！")
        self.load_apps()

    def start_guard_thread(self):
        """ 开启 WMI 守护线程 """

        def watch_procs():
            pythoncom.CoInitialize()
            c = wmi.WMI()
            # 监听进程创建事件
            watcher = c.watch_for(notification_type="Creation", wmi_class="Win32_Process")
            while self.is_running:
                try:
                    # 增加超时，以便循环能检查 is_running 标志
                    proc = watcher(timeout_ms=1000)
                    if proc.ExecutablePath:
                        full_path = proc.ExecutablePath.lower()
                        if full_path in self.ask_list:
                            curr = time.time()
                            # 10秒冷却，防止弹窗轰炸
                            if curr - self.cooldowns.get(full_path, 0) > 10:
                                self.cooldowns[full_path] = curr
                                self.signals.request_ask.emit(proc.ExecutablePath)
                except wmi.x_wmi_timed_out:
                    continue
                except:
                    break

        threading.Thread(target=watch_procs, daemon=True).start()

    def show_ask_dialog(self, exe_path):
        """ 弹出显卡模式询问对话框 """
        dialog = QWidget()
        dialog.setWindowTitle("模式切换确认")
        dialog.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        dialog.setFixedSize(380, 220)
        l = QVBoxLayout(dialog)
        l.setContentsMargins(25, 25, 25, 25)

        msg = QLabel(
            f"<b>检测到受监控程序启动：</b><br>{os.path.basename(exe_path)}<br><br>请选择显卡模式 (程序将自动重启生效)：")
        msg.setWordWrap(True)
        l.addWidget(msg)

        def do_restart(pref):
            # 1. 修改注册表
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, f"GpuPreference={pref};ASK")
            winreg.CloseKey(key)
            # 2. 终止现有进程
            for p in psutil.process_iter(['exe']):
                if p.info['exe'] and p.info['exe'].lower() == exe_path.lower():
                    try:
                        p.kill()
                    except:
                        pass
            # 3. 冷却与重新拉起
            self.cooldowns[exe_path.lower()] = time.time()
            time.sleep(0.8)
            os.startfile(exe_path)
            dialog.close()

        btn_high = QPushButton("🚀 高性能 (独显) 模式并重启")
        btn_high.setFixedHeight(40)
        btn_high.clicked.connect(lambda: do_restart(2))

        btn_low = QPushButton("🍃 省电 (集显) 模式并重启")
        btn_low.setFixedHeight(40)
        btn_low.clicked.connect(lambda: do_restart(1))

        l.addWidget(btn_high)
        l.addWidget(btn_low)
        dialog.show()
        self.active_dialog = dialog  # 保持引用防止销毁

    def filter_table(self):
        """ 表格实时搜索过滤 """
        text = self.search_input.text().lower()
        for i in range(self.table.rowCount()):
            name = self.table.item(i, 0).text().lower()
            path = self.table.item(i, 1).text().lower()
            self.table.setRowHidden(i, text not in name and text not in path)

    def add_new_app(self):
        """ 手动添加 exe 到列表 """
        f, _ = QFileDialog.getOpenFileName(self, "选择程序", "", "EXE (*.exe)")
        if f:
            p = f.replace("/", "\\")
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
            winreg.SetValueEx(key, p, 0, winreg.REG_SZ, "GpuPreference=0;")
            winreg.CloseKey(key)
            self.load_apps()

    def is_autostart_enabled(self):
        """ 检查开机自启动状态 """
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0,
                                 winreg.KEY_READ)
            winreg.QueryValueEx(key, "GPUSwitch")
            winreg.CloseKey(key)
            return True
        except:
            return False

    def toggle_autostart(self, state):
        """ 切换开机自启动逻辑 """
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "GPUSwitch"
        app_path = os.path.realpath(sys.executable)

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            if state == 2:  # Qt.Checked
                # 启动时带上参数以便静默进入托盘
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}" --minimized')
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            QMessageBox.warning(self, "权限错误", f"设置自启动失败: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 允许窗口关闭后应用不立即退出（由我们逻辑接管）
    app.setQuitOnLastWindowClosed(False)

    window = GPUSwitch()

    # 根据启动参数决定是否显示主界面
    if "--minimized" in sys.argv:
        window.hide()
    else:
        window.show()

    sys.exit(app.exec())