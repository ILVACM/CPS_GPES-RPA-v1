from time import sleep

from functions import WindowsUI_OPSimulation
from functions import office
from functions import Init

import os
import json
import time
import sys
import subprocess

import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from QT.qt import Ui_MainWindow  # 导入生成的UI类

# class MainWindow(QtWidgets.QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.ui = Ui_MainWindow()
#         self.ui.setupUi(self)
#         self.model = QtGui.QStandardItemModel()
#
#         # 初始化表格模型
#         self.init_table()
#
#         # 绑定按钮点击事件
#         # self.ui.pushButton.clicked.connect(self.load_excel)
#         # self.ui.pushButton.clicked.connect(self.load_json)
#         self.ui.pushButton.clicked.connect(self.start)
#
#         # 窗口显示后自动加载数据（比直接__init__调用更安全）
#         QtCore.QTimer.singleShot(0, self.load_and_display_json)
#
#     def init_table(self):
#         self.ui.TableView.setModel(self.model)
#         self.ui.TableView.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
#         self.ui.TableView.verticalHeader().hide()
#
#     def load_excel(self):
#         path, _ = QtWidgets.QFileDialog.getOpenFileName(
#             self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
#         )
#         if not path: return
#
#         try:
#             df = pd.read_excel(path)
#             self.model.clear()
#             self.model.setHorizontalHeaderLabels(df.columns.tolist())
#
#             for row in df.itertuples(index=False):
#                 items = [QtGui.QStandardItem(str(value)) for value in row]
#                 self.model.appendRow(items)
#
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(self, "错误", f"加载失败：{str(e)}")
#
#     def load_json(self):
#         path, _ = QtWidgets.QFileDialog.getOpenFileName(
#             self, "选择JSON文件", "", "JSON文件 (*.json)"
#         )
#         if not path:
#             return
#
#         try:
#             with open(path, 'r', encoding='utf-8') as f:
#                 data = json.load(f)
#                 self.model.clear()
#
#                 # 假设JSON是数组结构，如：[{"name":"张三","age":25},...]
#                 if isinstance(data, list) and len(data) > 0:
#                     # 设置表头（取第一条数据的键）
#                     headers = list(data[0].keys())
#                     self.model.setHorizontalHeaderLabels(headers)
#
#                     # 填充数据行
#                     for item in data:
#                         row = [QtGui.QStandardItem(str(value)) for value in item.values()]
#                         self.model.appendRow(row)
#                 else:
#                     QtWidgets.QMessageBox.warning(self, "警告", "JSON格式应为数组且包含有效数据")
#
#         except Exception as e:
#             QtWidgets.QMessageBox.critical(self, "错误", f"加载失败：{str(e)}")
#
#     def load_and_display_json(self):
#         try:
#             # 加载JSON文件
#             with open('config/temp_1.json', 'r', encoding='utf-8') as file:
#                 data = json.load(file)
#                 # 提取OperateList数组
#                 operate_list = data.get('OperateList', [])
#
#                 # 创建数据模型并设置表头
#                 model = QStandardItemModel()
#                 model.setHorizontalHeaderLabels(['模式', '代码', '坐标X', '坐标Y'])
#
#                 # 填充数据到模型
#                 for item in operate_list:
#                     coord = item.get('coord', [])
#                     row = [
#                         QStandardItem(str(item.get('mode', ''))),
#                         QStandardItem(item.get('code', '')),
#                         QStandardItem(str(coord[0]) if len(coord) > 0 else ''),
#                         QStandardItem(str(coord[1]) if len(coord) > 1 else '')
#                     ]
#                     model.appendRow(row)
#
#                 # 设置模型到TableView并调整列宽
#                 self.ui.TableView.setModel(model)
#                 self.ui.TableView.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
#
#         except Exception as e:
#             print(f"错误: {e}")
#
#     def start(self):
#         """按钮点击后触发的新函数"""
#         self.hide()  # 立即隐藏窗口
#
#         # 这里添加点击按钮后的具体业务逻辑
#         print("启动业务逻辑...")
#         # sleep(3)
#
#         # 直接执行命令，返回退出码（0为成功）
#         # os.system(r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe source/qt2.py")
#         # exit_code = os.system(r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe test.py")
#         # print("退出码:", exit_code)
#
#         # 虚拟环境Python解释器路径
#         venv_python = r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe"
#
#         # 需要启动的两个Python程序路径
#         script1 = r"source/qt3.py"
#         script2 = r"test.py"
#
#         # 同步启动两个程序（并行执行）
#         process1 = subprocess.Popen([venv_python, script1])
#         process2 = subprocess.Popen([venv_python, script2])
#
#         # 等待两个程序执行完成（若需同步等待）
#         process1.wait()
#         process2.wait()
#
#         self.show()  # 执行完成后恢复显示

# if __name__ == '__main__':
#     pass
#
#     # Init.Init()
#     #
#     # with open('config\default.json', 'r', encoding='utf-8') as file:
#     #     source = json.load(file)
#     #
#     # OPList = source['OperateList']
#     #
#     # target_window = WindowsUI_OPSimulation.find_activate_window("监理一体化")
#     # if target_window:
#     #     print(f"窗口标题: {target_window.title}")
#     #     print(f"窗口位置: {target_window.left}, {target_window.top}")
#     #     print(f"窗口大小: {target_window.width}x{target_window.height}")
#     #     WindowsUI_OPSimulation.force_maximize_and_foreground(target_window)  # 最大化并置顶
#     # else:
#     #     print("未找到目标窗口！")
#     #
#     # for OP in OPList:
#     #     pass
#     #
#     #     WindowsUI_OPSimulation.smart_click( OP['coord'][0], OP['coord'][1])
#     #     time.sleep(1)
#
#     #==================== QT =================
#
#     app = QtWidgets.QApplication(sys.argv)
#     window = MainWindow()
#     window.show()
#     sys.exit(app.exec_())

import os
import json
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtGui import QStandardItemModel, QStandardItem  # 新增这一行
from PyQt5 import QtCore, QtGui, QtWidgets

from QT.qt2 import Ui_MainWindow

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
        print("启动业务逻辑...")

        # 虚拟环境Python解释器路径
        venv_python = r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe"

        # 需要启动的Python程序路径
        script1 = r"source/qt3_0.py"

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
        print("启动业务逻辑...")

        # 直接执行命令，返回退出码（0为成功）
        # os.system(r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe source/qt2.py")
        # exit_code = os.system(r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe test.py")
        # print("退出码:", exit_code)

        # 虚拟环境Python解释器路径
        venv_python = r"C:\Users\AdmOA\.conda\envs\python311-auto\python.exe"

        # 需要启动的两个Python程序路径
        script1 = r"source/qt2.py"
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