import os
import sys
import threading
import time
import winreg

import psutil
import pythoncom
import wmi
from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QAction, QIcon
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (QApplication,QCheckBox,QComboBox,QFileDialog,QHeaderView,QHBoxLayout, QLabel,
                               QLineEdit,QMainWindow,QMenu,QMessageBox,QPushButton,QSystemTrayIcon,QTableWidget,
                               QTableWidgetItem,QVBoxLayout,QWidget)


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)


class GuardSignals(QObject):
    request_ask = Signal(str, int)


class GPUSwitch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GPUSwitch')
        self.setMinimumSize(1000, 650)

        self.reg_path = r'Software\Microsoft\DirectX\UserGpuPreferences'
        self.ask_list = set()
        self.cooldowns = {}
        self.pending_changes = set()
        self.is_running = True
        self.really_quit = False
        self.server_name = 'GPUSwitchSingleton'
        self.activation_server = None
        self.active_dialog = None
        self.is_loading = False

        self.signals = GuardSignals()
        self.signals.request_ask.connect(self.show_ask_dialog)

        self.setup_ui()
        self.load_apps()
        self.init_tray()

        icon_path = get_resource_path('app_icon.ico')
        self.setWindowIcon(QIcon(icon_path))
        self.tray_icon.setIcon(QIcon(icon_path))

        self.setup_single_instance_server()
        self.start_guard_thread()

    def setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)

        top_bar = QHBoxLayout()
        self.btn_add = QPushButton('添加新程序')
        self.btn_add.clicked.connect(self.add_new_app)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('搜索程序名称或路径...')
        self.search_input.setFixedWidth(300)
        self.search_input.textChanged.connect(self.filter_table)

        self.btn_refresh = QPushButton('刷新列表')
        self.btn_refresh.clicked.connect(self.load_apps)

        top_bar.addWidget(self.btn_add)
        top_bar.addWidget(self.search_input)
        top_bar.addStretch()
        top_bar.addWidget(self.btn_refresh)
        self.main_layout.addLayout(top_bar)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(['程序名称', '文件完整路径', '显卡偏好设置', '运行时询问', '删除程序'])
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.setColumnWidth(0, 160)
        self.table.setColumnWidth(1, 450)
        self.table.setColumnWidth(2, 150)
        self.table.setColumnWidth(3, 90)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.itemChanged.connect(self.on_table_item_changed)
        self.main_layout.addWidget(self.table)

        bottom_bar = QHBoxLayout()
        self.status_label = QLabel('列表已更新')
        self.status_label.setStyleSheet('color: #31C950;')

        self.check_minimize_to_tray = QCheckBox('最小化到系统托盘')
        self.check_minimize_to_tray.setChecked(True)

        self.check_autostart = QCheckBox('开机自动启动')
        self.check_autostart.setChecked(self.is_autostart_enabled())
        self.check_autostart.stateChanged.connect(self.toggle_autostart)

        self.btn_apply = QPushButton('应用所有更改')
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
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setToolTip('GPUSwitch')

        tray_menu = QMenu()
        show_action = QAction('显示主界面', self)
        show_action.triggered.connect(self.show_and_activate)
        quit_action = QAction('彻底退出', self)
        quit_action.triggered.connect(self.quit_app)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def setup_single_instance_server(self):
        QLocalServer.removeServer(self.server_name)
        self.activation_server = QLocalServer(self)
        self.activation_server.newConnection.connect(self.handle_activation_connection)
        if not self.activation_server.listen(self.server_name):
            self.activation_server = None

    def handle_activation_connection(self):
        if not self.activation_server:
            return

        while self.activation_server.hasPendingConnections():
            socket = self.activation_server.nextPendingConnection()
            socket.readyRead.connect(lambda s=socket: self.handle_activation_socket(s))
            socket.disconnected.connect(socket.deleteLater)

    def handle_activation_socket(self, socket):
        try:
            socket.readAll()
        except Exception:
            pass
        self.show_and_activate()
        socket.disconnectFromServer()

    def show_and_activate(self):
        self.show()
        self.showNormal()
        if self.isMinimized():
            self.setWindowState(self.windowState() & ~Qt.WindowMinimized)
        self.activateWindow()
        self.raise_()

    def on_tray_icon_activated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            if self.isVisible():
                self.hide()
            else:
                self.show_and_activate()

    def closeEvent(self, event):
        if self.really_quit:
            event.accept()
            return

        if self.check_minimize_to_tray.isChecked():
            event.ignore()
            self.hide()
            return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('退出确认')
        msg_box.setText('您点击了关闭按钮，请选择操作：')
        btn_tray = msg_box.addButton('最小化到托盘', QMessageBox.AcceptRole)
        msg_box.addButton('完全退出程序', QMessageBox.DestructiveRole)
        msg_box.setDefaultButton(btn_tray)
        msg_box.exec()

        if msg_box.clickedButton() == btn_tray:
            self.check_minimize_to_tray.setChecked(True)
            event.ignore()
            self.hide()
        else:
            self.quit_app(ask=False)
            event.accept()

    def quit_app(self, ask=True):
        if ask:
            reply = QMessageBox.question(
                self,
                '确认退出',
                '退出后将停止所有进程监听，确定吗？',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return

        self.is_running = False
        self.really_quit = True
        if self.activation_server is not None:
            self.activation_server.close()
            QLocalServer.removeServer(self.server_name)
        self.tray_icon.hide()
        QApplication.quit()

    def load_apps(self):
        self.is_loading = True
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
                    i += 1
                    if not isinstance(value, str):
                        continue

                    is_ask = 'ASK' in value
                    if is_ask:
                        self.ask_list.add(path.lower())
                    self.add_row(path, value, is_ask)
                except OSError:
                    break
            winreg.CloseKey(key)
        except OSError:
            pass

        self.status_label.setText('列表已更新')
        self.status_label.setStyleSheet('color: #31C950;')
        self.table.setSortingEnabled(True)
        self.is_loading = False

    def add_row(self, path, val, is_ask):
        row = self.table.rowCount()
        self.table.insertRow(row)

        self.table.setItem(row, 0, QTableWidgetItem(os.path.basename(path)))
        self.table.setItem(row, 1, QTableWidgetItem(path))

        combo = QComboBox()
        combo.addItems(['系统默认', '省电 (集显)', '高性能 (独显)'])
        clean_val = val.replace('ASK', '')
        value_map = {'GpuPreference=0;': 0, 'GpuPreference=1;': 1, 'GpuPreference=2;': 2}
        combo.setCurrentIndex(value_map.get(clean_val, 0))
        combo.currentIndexChanged.connect(lambda _=None, p=path: self.mark_as_changed(p))
        self.table.setCellWidget(row, 2, combo)

        check_container = QWidget()
        check_layout = QHBoxLayout(check_container)
        check_layout.setContentsMargins(0, 0, 0, 0)
        check_layout.setAlignment(Qt.AlignCenter)

        check_box = QCheckBox()
        check_box.setChecked(is_ask)
        check_box.stateChanged.connect(lambda _=None, p=path: self.mark_as_changed(p))
        check_layout.addWidget(check_box)

        self.table.setCellWidget(row, 3, check_container)

        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setAlignment(Qt.AlignCenter)
        btn_layout.setContentsMargins(0, 0, 0, 0)

        btn_del = QPushButton('删除')
        btn_del.setFixedSize(60, 28)
        btn_del.setStyleSheet(
            '''
            QPushButton {
                background-color: #ff4d4f;
                color: white;
                border-radius: 4px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #ff7875; }
            QPushButton:pressed { background-color: #d9363e; }
            '''
        )
        btn_del.clicked.connect(lambda: self.delete_app_confirm(path))

        btn_layout.addWidget(btn_del)
        self.table.setCellWidget(row, 4, btn_container)

    def on_table_item_changed(self, item):
        if self.is_loading:
            return
        if item.column() == 3:
            path_item = self.table.item(item.row(), 1)
            if path_item is not None:
                self.mark_as_changed(path_item.text())

    def mark_as_changed(self, path):
        self.pending_changes.add(path)
        self.status_label.setText(f'提示：有 {len(self.pending_changes)} 处待修改的项')
        self.status_label.setStyleSheet('color: #e67e22; font-weight: bold;')
        self.btn_apply.setEnabled(True)

    def apply_all_changes(self):
        for i in range(self.table.rowCount()):
            path = self.table.item(i, 1).text()
            if path not in self.pending_changes:
                continue

            gpu_idx = self.table.cellWidget(i, 2).currentIndex()
            check_box = self.table.cellWidget(i, 3).findChild(QCheckBox)
            is_ask = check_box.isChecked() if check_box else False
            val = f'GpuPreference={gpu_idx};'
            if is_ask:
                val += 'ASK'
                self.ask_list.add(path.lower())
            else:
                self.ask_list.discard(path.lower())

            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
            winreg.SetValueEx(key, path, 0, winreg.REG_SZ, val)
            winreg.CloseKey(key)

        QMessageBox.information(self, '成功', '所有更改已成功应用。')
        self.load_apps()

    def start_guard_thread(self):
        def watch_procs():
            pythoncom.CoInitialize()
            try:
                watcher = wmi.WMI().watch_for(notification_type='Creation', wmi_class='Win32_Process')
                while self.is_running:
                    try:
                        proc = watcher(timeout_ms=1000)
                        if not proc.ExecutablePath:
                            continue

                        full_path = proc.ExecutablePath.lower()
                        if full_path not in self.ask_list:
                            continue

                        curr = time.time()
                        if curr - self.cooldowns.get(full_path, 0) <= 10:
                            continue

                        self.cooldowns[full_path] = curr
                        self.signals.request_ask.emit(proc.ExecutablePath, int(proc.ProcessId))
                    except wmi.x_wmi_timed_out:
                        continue
                    except Exception:
                        break
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=watch_procs, daemon=True).start()

    def delete_app_confirm(self, path):
        reply = QMessageBox.question(
            self,
            '确认删除',
            f'确定要从列表中移除该程序吗？\n{os.path.basename(path)}',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply != QMessageBox.Yes:
            return

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, path)
            winreg.CloseKey(key)
            self.load_apps()
            QMessageBox.information(self, '成功', '已成功移除该项。')
        except Exception as e:
            QMessageBox.warning(self, '错误', f'删除失败: {e}')

    def show_ask_dialog(self, exe_path, pid):
        dialog = QWidget()
        dialog.setWindowTitle('模式切换确认')
        dialog.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        dialog.setFixedSize(380, 220)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 25, 25, 25)

        msg = QLabel(
            f'<b>检测到受监控程序启动：</b><br>{os.path.basename(exe_path)}<br><br>'
            '请选择显卡模式（程序将自动重启生效）：'
        )
        msg.setWordWrap(True)
        layout.addWidget(msg)

        def do_restart(pref):
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, f'GpuPreference={pref};ASK')
            winreg.CloseKey(key)

            matching_procs = []
            launch_cmd = None
            launch_cwd = None
            try:
                target_proc = psutil.Process(pid)
                launch_cmd = target_proc.cmdline()
                launch_cwd = target_proc.cwd()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

            for proc in psutil.process_iter(['pid', 'exe']):
                try:
                    if proc.info['exe'] and proc.info['exe'].lower() == exe_path.lower():
                        matching_procs.append(proc)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            for proc in matching_procs:
                try:
                    proc.terminate()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            _, alive = psutil.wait_procs(matching_procs, timeout=3)

            for proc in alive:
                try:
                    proc.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

            if alive:
                psutil.wait_procs(alive, timeout=2)

            self.cooldowns[exe_path.lower()] = time.time()

            try:
                if launch_cmd:
                    launch_kwargs = {}
                    if launch_cwd:
                        launch_kwargs['cwd'] = launch_cwd
                    psutil.Popen(launch_cmd, **launch_kwargs)
                else:
                    os.startfile(exe_path)
            except OSError as e:
                QMessageBox.warning(self, '错误', f'无法重启程序: {e}')
                return

            dialog.close()

        btn_high = QPushButton('高性能 (独显) 模式并重启')
        btn_high.setFixedHeight(40)
        btn_high.clicked.connect(lambda: do_restart(2))

        btn_low = QPushButton('省电 (集显) 模式并重启')
        btn_low.setFixedHeight(40)
        btn_low.clicked.connect(lambda: do_restart(1))

        layout.addWidget(btn_high)
        layout.addWidget(btn_low)
        dialog.show()
        self.active_dialog = dialog

    def filter_table(self):
        text = self.search_input.text().lower()
        for i in range(self.table.rowCount()):
            name = self.table.item(i, 0).text().lower()
            path = self.table.item(i, 1).text().lower()
            self.table.setRowHidden(i, text not in name and text not in path)

    def add_new_app(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择程序', '', 'EXE (*.exe)')
        if not file_path:
            return

        path = file_path.replace('/', '\\')
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
        winreg.SetValueEx(key, path, 0, winreg.REG_SZ, 'GpuPreference=0;')
        winreg.CloseKey(key)
        self.load_apps()

    def is_autostart_enabled(self):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r'Software\Microsoft\Windows\CurrentVersion\Run',
                0,
                winreg.KEY_READ,
            )
            winreg.QueryValueEx(key, 'GPUSwitch')
            winreg.CloseKey(key)
            return True
        except OSError:
            return False

    def toggle_autostart(self, state):
        key_path = r'Software\Microsoft\Windows\CurrentVersion\Run'
        app_name = 'GPUSwitch'
        app_path = os.path.realpath(sys.executable)

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            if state == Qt.Checked:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}" --minimized')
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except OSError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            QMessageBox.warning(self, '权限错误', f'设置自启动失败: {e}')


def notify_existing_instance(server_name='GPUSwitchSingleton'):
    socket = QLocalSocket()
    socket.connectToServer(server_name)
    if not socket.waitForConnected(300):
        return False

    socket.write(b'show')
    socket.flush()
    socket.waitForBytesWritten(300)
    socket.disconnectFromServer()
    return True


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    if notify_existing_instance():
        sys.exit(0)

    window = GPUSwitch()

    if '--minimized' in sys.argv:
        window.hide()
    else:
        window.show()

    sys.exit(app.exec())
