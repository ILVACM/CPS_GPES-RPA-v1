import os
import json
import subprocess

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtGui import QStandardItemModel, QStandardItem  # 新增这一行
from PyQt5 import QtCore, QtGui, QtWidgets

from QT.qt2 import Ui_MainWindow

def get_python_interpreter_path():
        """
        获取当前运行的 Python 解释器的绝对路径。

        返回:
            str: 当前 Python 解释器的绝对路径。
            如果出错则返回 None。
        """
        try:
            interpreter_path = sys.executable
            if not os.path.exists(interpreter_path):
                raise FileNotFoundError(f"解释器路径不存在: {interpreter_path}")
            return os.path.abspath(interpreter_path)
        except Exception as e:
            print(f"[错误] 获取 Python 解释器路径时发生异常: {e}")
            return None

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.load_list_view()
        self.load_table_view()
        self.load_list_view_2()

        # 绑定按钮事件
        self.pushButton_1.clicked.connect(self.business_logic1)
        self.pushButton.clicked.connect(self.business_logic2)

    def load_list_view(self):
        """加载历史配置文件列表"""
        config_dir = "config"
        if not os.path.exists(config_dir):
            self.show_error("目录不存在", f"无法找到配置目录：{config_dir}")
            return

        files = [f for f in os.listdir(config_dir) if os.path.isfile(os.path.join(config_dir, f))]
        model = QStandardItemModel()
        for file in files:
            item = QStandardItem(file)
            model.appendRow(item)
        self.listView.setModel(model)

    def load_table_view(self):
        """加载规则配置的JSON数据"""
        json_path = "config/temp.json"
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError) as e:
            self.show_error("JSON文件错误", f"加载temp.json失败: {str(e)}")
            return

        operate_list = data.get("OperateList", [])
        model = QStandardItemModel()
        model.setColumnCount(4)
        model.setHorizontalHeaderLabels(["项1", "项2", "项3", "项4"])  # 可自定义列标题
        self.tableView.setModel(model)

        for item in operate_list:
            row_data = []
            keys = list(item.keys())[:4]
            for key in keys:
                value = str(item[key])
                row_data.append(QStandardItem(value))
            # 补充空项到4个
            while len(row_data) < 4:
                row_data.append(QStandardItem(""))
            model.appendRow(row_data)

    def load_list_view_2(self):
        """加载日志文件内容"""
        log_path = "log/202503172122/record.log"
        try:
            with open(log_path, 'r', encoding='utf-8') as f:
                lines = f.read().splitlines()
        except FileNotFoundError:
            lines = ["日志文件未找到"]
        except Exception as e:
            lines = [f"读取日志失败: {str(e)}"]

        model = QStandardItemModel()
        for line in lines:
            item = QStandardItem(line.strip())
            model.appendRow(item)
        self.listView_2.setModel(model)

    def show_error(self, title, message):
        """显示错误信息"""
        msg_box = QtWidgets.QMessageBox(self)
        msg_box.setIcon(QtWidgets.QMessageBox.Critical)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg_box.exec_()

    def business_logic1(self):
        pass
        # 按钮1的业务逻辑
        # print(1)

        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑1...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的Python程序路径
        script1 = r"QT\qt3_0.py"

        # 启动程序
        process1 = subprocess.Popen([venv_python, script1])
        process1.wait()

        self.show()  # 执行完成后恢复显示

    def business_logic2(self):
        pass
        # 按钮2的业务逻辑
        # print(2)

        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑2...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的两个Python程序路径
        script1 = r"QT\qt2.py"
        script2 = r"test.py"

        # 同步启动两个程序（并行执行）
        process1 = subprocess.Popen([venv_python, script1])
        process2 = subprocess.Popen([venv_python, script2])

        # 等待两个程序执行完成（若需同步等待）
        process1.wait()
        process2.wait()

        self.show()  # 执行完成后恢复显示


# 按钮2的业务逻辑
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())