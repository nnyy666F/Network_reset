import winreg
from PySide6.QtWidgets import QApplication, QWidget, QListWidgetItem, QMessageBox, QLabel, QLineEdit, QPushButton, QListWidget, QDialog, QDialogButtonBox, QVBoxLayout, QComboBox
from PySide6.QtCore import Qt, QDateTime
from PySide6.QtGui import QKeySequence

class RegistryEditor(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("注册表路径：HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles")
        layout.addWidget(self.label)

        self.listWidget = QListWidget()
        layout.addWidget(self.listWidget)

        self.lineEdit = QLineEdit()
        self.lineEdit.setPlaceholderText("默认：网络 1")
        layout.addWidget(self.lineEdit)

        self.button = QPushButton("重置")
        self.button.clicked.connect(self.resetNetworkNames)
        layout.addWidget(self.button)

        # 添加刷新按钮
        self.refreshButton = QPushButton("刷新")
        self.refreshButton.clicked.connect(self.traverseRegistry)
        layout.addWidget(self.refreshButton)

        self.setLayout(layout)

    def traverseRegistry(self):
        """遍历注册表，获取包含'网络'的子键信息并添加到列表中。"""
        try:
            self.updateRegistryList()
        except WindowsError as e:
            self.showError(f"遍历注册表时出现错误：{e}")

    def updateRegistryList(self):
        """实际执行遍历注册表并更新列表的操作。"""
        self.listWidget.clear()
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles')
        items = []
        index = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, index)
                subkey = winreg.OpenKey(key, subkey_name)

                value_name = "ProfileName"
                value_type = winreg.QueryValueEx(subkey, value_name)[1]

                if value_type == winreg.REG_SZ and "网络" in winreg.QueryValueEx(subkey, value_name)[0]:
                    last_modified_time = winreg.QueryInfoKey(subkey)[2]
                    # 将时间戳转换为可读时间格式
                    seconds = last_modified_time / 10000000
                    try:
                        date_time = QDateTime.fromSecsSinceEpoch(int(seconds))
                        # 减去 369 年
                        adjusted_date_time = date_time.addYears(-369)
                        formatted_time = adjusted_date_time.toString("yyyy年MM月dd日 hh:mm:ss")
                        items.append(f"{subkey_name}: {winreg.QueryValueEx(subkey, value_name)[0]} ({formatted_time})")
                    except OverflowError:
                        # 处理可能的溢出情况，可以显示一个默认时间或者其他提示信息
                        formatted_time = "时间转换错误"
                        items.append(f"{subkey_name}: {winreg.QueryValueEx(subkey, value_name)[0]} ({formatted_time})")

                winreg.CloseKey(subkey)
                index += 1
            except WindowsError:
                break

        # 按照时间倒序排序
        sorted_items = sorted(items, key=lambda x: QDateTime.fromString(x.split('(')[-1].strip(')'), "yyyy年MM月dd日 hh:mm:ss"), reverse=True)

        # 倒序插入列表
        for item in sorted_items:
            self.listWidget.addItem(item)

    def resetNetworkNames(self):
        #selected_keys = []
        dialog = QDialog(self)
        dialog.setWindowTitle("选择要修改的网络项")

        layout = QVBoxLayout()

        # 使用下拉选择框代替复选框列表
        self.comboBox = QComboBox(dialog)
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            self.comboBox.addItem(item.text())
        layout.addWidget(self.comboBox)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, Qt.Horizontal, dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        dialog.setLayout(layout)

        if dialog.exec() == QDialog.Accepted:
            selected_key = self.comboBox.currentText()
            if selected_key:
                key_part = selected_key.split('}')[0] + '}'
                delete_all_flag = False
                if not delete_all_flag:
                    reply = QMessageBox.question(self, '确认删除', '确定要删除选中的网络项吗？', QMessageBox.Yes | QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        delete_all_flag = True

                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles')
                index = 0
                to_delete_keys = []
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, index)
                        subkey = winreg.OpenKey(key, subkey_name)

                        value_name = "ProfileName"
                        value_type = winreg.QueryValueEx(subkey, value_name)[1]

                        profile_name = winreg.QueryValueEx(subkey, value_name)[0]
                        if value_type == winreg.REG_SZ and profile_name.startswith("网络"):
                            if subkey_name!= key_part:
                                to_delete_keys.append(subkey_name)

                        winreg.CloseKey(subkey)
                        index += 1
                    except WindowsError:
                        break

                for subkey_name in to_delete_keys:
                    try:
                        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, f'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\NetworkList\\Profiles\\{subkey_name}')
                        self.listWidget.takeItem(self.listWidget.row(QListWidgetItem(subkey_name)))
                    except WindowsError as e:
                        self.showError(f"删除子键时出现错误：{e}")

                # 设置选中的项的 ProfileName 为输入框中的内容
                try:
                    subkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles', 0, winreg.KEY_ALL_ACCESS)
                    for i in range(winreg.QueryInfoKey(subkey)[0]):
                        temp_subkey_name = winreg.EnumKey(subkey, i)
                        if temp_subkey_name.startswith(key_part):
                            temp_subkey = winreg.OpenKey(subkey, temp_subkey_name, 0, winreg.KEY_ALL_ACCESS)
                            new_name = self.lineEdit.text() if self.lineEdit.text() else "网络 1"
                            winreg.SetValueEx(temp_subkey, "ProfileName", 0, winreg.REG_SZ, new_name)
                            winreg.CloseKey(temp_subkey)
                            break
                    winreg.CloseKey(subkey)
                except WindowsError as e:
                    self.showError(f"设置 ProfileName 时出现错误：{e}")

            # 重置完成后重新遍历注册表更新列表
            self.listWidget.clear()
            self.traverseRegistry()

    def showError(self, message):
        """显示错误消息。"""
        QMessageBox.critical(self, '错误', message)

    def keyPressEvent(self, event):
        # 监听键盘 F5 按键事件进行刷新
        if event.key() == Qt.Key_F5 or event.matches(QKeySequence.Refresh):
            self.traverseRegistry()

if __name__ == '__main__':
    app = QApplication([])
    window = RegistryEditor()
    window.traverseRegistry()
    window.show()
    app.exec()
