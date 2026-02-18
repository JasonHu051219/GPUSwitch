import sys
import os
import winreg
import time
import psutil
import threading
import wmi
import pythoncom
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QTableWidget, QTableWidgetItem, QComboBox,
                               QCheckBox, QLabel, QMessageBox, QHeaderView, QFileDialog, QLineEdit)
from PySide6.QtCore import Qt, Signal, QObject

# --- Win11 风格 QSS 样式表 ---
WIN11_STYLE = """
QMainWindow { background-color: #f3f3f3; }
QTableWidget {
    background-color: white;
    border: 1px solid #dcdcdc;
    gridline-color: #f0f0f0;
    border-radius: 8px;
    outline: 0;
}
QHeaderView::section {
    background-color: #f9f9f9;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #dcdcdc;
    font-weight: bold;
}
QPushButton {
    background-color: #ffffff;
    border: 1px solid #dcdcdc;
    border-radius: 6px;
    padding: 6px 15px;
}
QPushButton:hover { background-color: #f9f9f9; border-color: #c0c0c0; }
QPushButton#applyBtn {
    background-color: #2ecc71;
    color: white;
    font-weight: bold;
    border: none;
    padding: 8px 25px;
}
QPushButton#applyBtn:hover { background-color: #27ae60; }
QPushButton#applyBtn:disabled { background-color: #bdc3c7; }
QComboBox {
    border: 1px solid #dcdcdc;
    border-radius: 4px;
    padding: 3px 10px;
    background: white;
}
QLineEdit {
    border: 1px solid #dcdcdc;
    border-radius: 6px;
    padding: 6px;
    background: white;
}
QComboBox QAbstractItemView {
    background-color: white;
    border: 1px solid #dcdcdc;
    selection-background-color: #0000CD; /* 鼠标悬停时的浅灰色背景 */
    selection-color: black;            /* 鼠标悬停时的文字颜色，强制为黑色 */
    outline: 0;
}

/* 下拉列表内每个选项的高度和间距 */
QComboBox QAbstractItemView::item {
    min-height: 30px;
    padding-left: 10px;
}

/* 鼠标悬停在选项上时的状态 */
QComboBox QAbstractItemView::item:hover {
    background-color: #f0f0f0;
    color: black;
}
"""


class GuardSignals(QObject):
    request_ask = Signal(str)


class GPUSwitch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPUSwitch (GPUS) - 进程监听重启版")
        self.setMinimumSize(1000, 650)
        self.setStyleSheet(WIN11_STYLE)

        self.reg_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        self.ask_list = set()
        self.cooldowns = {}
        self.pending_changes = set()  # 记录被修改过的行路径

        self.signals = GuardSignals()
        self.signals.request_ask.connect(self.show_ask_dialog)

        self.setup_ui()
        self.load_apps()
        self.start_guard_thread()

    def setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)

        # 顶部工具栏
        top_bar = QHBoxLayout()
        self.btn_add = QPushButton("➕ 添加新程序")
        self.btn_add.clicked.connect(self.add_new_app)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 搜索程序名称或路径...")
        self.search_input.setFixedWidth(300)  # 设置默认搜索栏宽度
        self.search_input.textChanged.connect(self.filter_table)

        self.btn_refresh = QPushButton("🔄 刷新列表")
        self.btn_refresh.clicked.connect(self.load_apps)

        top_bar.addWidget(self.btn_add)
        top_bar.addWidget(self.search_input)
        top_bar.addStretch()
        top_bar.addWidget(self.btn_refresh)
        self.layout.addLayout(top_bar)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["程序名称", "文件完整路径", "显卡偏好设置", "运行时询问 (自动重启)"])

        # 启用点击表头排序功能
        self.table.setSortingEnabled(True)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        # 设置默认列宽分配
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 450)
        self.table.setColumnWidth(2, 150)
        header.setStretchLastSection(True)

        self.layout.addWidget(self.table)

        # 底部栏
        bottom_bar = QHBoxLayout()
        self.status_label = QLabel("列表已更新")
        self.status_label.setStyleSheet("color: #7f8c8d;")

        self.btn_apply = QPushButton("💾 应用所有更改")
        self.btn_apply.setObjectName("applyBtn")
        self.btn_apply.setEnabled(False)
        self.btn_apply.clicked.connect(self.apply_all_changes)

        bottom_bar.addWidget(self.status_label)
        bottom_bar.addStretch()
        bottom_bar.addWidget(self.btn_apply)
        self.layout.addLayout(bottom_bar)

    def load_apps(self):
        # 刷新时临时关闭排序，防止插入数据时乱跳
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        self.ask_list.clear()
        self.pending_changes.clear()
        self.btn_apply.setEnabled(False)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_READ)
            i = 0
            while True:
                path, value, _ = winreg.EnumValue(key, i)
                is_ask = "ASK" in value
                if is_ask: self.ask_list.add(path.lower())
                self.add_row(path, value, is_ask)
                i += 1
        except OSError:
            pass
        self.status_label.setText("列表已更新")
        self.table.setSortingEnabled(True)

    def add_row(self, path, val, is_ask):
        row = self.table.rowCount()
        self.table.insertRow(row)

        # 使用自定义 QTableWidgetItem 以支持更好的排序逻辑
        name_item = QTableWidgetItem(os.path.basename(path))
        path_item = QTableWidgetItem(path)

        self.table.setItem(row, 0, name_item)
        self.table.setItem(row, 1, path_item)

        combo = QComboBox()
        combo.addItems(["系统默认", "省电 (集显)", "高性能 (独显)"])
        clean_val = val.replace("ASK", "")
        m = {"GpuPreference=0;": 0, "GpuPreference=1;": 1, "GpuPreference=2;": 2}
        combo.setCurrentIndex(m.get(clean_val, 0))
        combo.currentIndexChanged.connect(lambda: self.mark_as_changed(path))
        self.table.setCellWidget(row, 2, combo)

        check = QCheckBox("启用弹窗重启")
        check.setChecked(is_ask)
        check.stateChanged.connect(lambda: self.mark_as_changed(path))
        self.table.setCellWidget(row, 3, check)

    def mark_as_changed(self, path):
        self.pending_changes.add(path)
        self.status_label.setText(f"提示：有 {len(self.pending_changes)} 处待修改的项")
        self.status_label.setStyleSheet("color: #e67e22; font-weight: bold;")
        self.btn_apply.setEnabled(True)

    def apply_all_changes(self):
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

        QMessageBox.information(self, "成功", "所有更改已成功应用！")
        self.load_apps()

    def start_guard_thread(self):
        def watch_procs():
            pythoncom.CoInitialize()
            c = wmi.WMI()
            watcher = c.watch_for(notification_type="Creation", wmi_class="Win32_Process")
            while True:
                try:
                    proc = watcher(timeout_ms=1000)
                    if proc.ExecutablePath:
                        full_path = proc.ExecutablePath.lower()
                        if full_path in self.ask_list:
                            curr = time.time()
                            if curr - self.cooldowns.get(full_path, 0) > 10:
                                self.cooldowns[full_path] = curr
                                self.signals.request_ask.emit(proc.ExecutablePath)
                except:
                    continue

        threading.Thread(target=watch_procs, daemon=True).start()

    def show_ask_dialog(self, exe_path):
        dialog = QWidget()
        dialog.setStyleSheet(WIN11_STYLE)
        dialog.setWindowTitle("GPUSwitch 模式切换")
        dialog.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        dialog.setFixedSize(360, 200)
        l = QVBoxLayout(dialog)
        l.setContentsMargins(20, 20, 20, 20)

        msg = QLabel(f"<b>检测到启动：</b><br>{os.path.basename(exe_path)}<br><br>请选择显卡模式 (程序将自动重启)：")
        msg.setWordWrap(True)
        l.addWidget(msg)

        def do_restart(pref):
            # 修改注册表
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_path, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, f"GpuPreference={pref};ASK")
            winreg.CloseKey(key)
            # 杀进程
            procs = [p for p in psutil.process_iter(['exe']) if
                     p.info['exe'] and p.info['exe'].lower() == exe_path.lower()]
            for p in procs: p.kill()
            # 冷却与重启
            self.cooldowns[exe_path.lower()] = time.time()
            time.sleep(1.0)
            os.startfile(exe_path)
            dialog.close()

        b1 = QPushButton("🚀 独立显卡模式并重启")
        b1.setObjectName("applyBtn")
        b1.clicked.connect(lambda: do_restart(2))

        b2 = QPushButton("🍃 集成显卡模式并重启")
        b2.clicked.connect(lambda: do_restart(1))

        l.addWidget(b1)
        l.addWidget(b2)
        dialog.show()
        self.active_dialog = dialog

    def filter_table(self):
        text = self.search_input.text().lower()
        for i in range(self.table.rowCount()):
            name = self.table.item(i, 0).text().lower()
            path = self.table.item(i, 1).text().lower()
            self.table.setRowHidden(i, text not in name and text not in path)

    def add_new_app(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择程序", "", "EXE (*.exe)")
        if f:
            p = f.replace("/", "\\")
            # 默认添加一个“系统默认”且不询问的配置
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, self.reg_path)
            winreg.SetValueEx(key, p, 0, winreg.REG_SZ, "GpuPreference=0;")
            winreg.CloseKey(key)
            self.load_apps()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = GPUSwitch()
    window.show()
    sys.exit(app.exec())