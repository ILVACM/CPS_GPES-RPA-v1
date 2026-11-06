import sys
import os
import subprocess
from PyQt5 import QtWidgets
from QT.qt0 import Ui_mainWindow  # 导入生成的UI类

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

class MainWindow(QtWidgets.QMainWindow, Ui_mainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)  # 初始化UI
        self.init_buttons()
    
    def init_buttons(self):
        """按钮信号连接"""
        self.pushButton.clicked.connect(self.start_0)
        self.pushButton_2.clicked.connect(self.start_1)
        self.pushButton_3.clicked.connect(self.start_2)
        self.pushButton_4.clicked.connect(self.start_3)
        
    
    def start_0(self):
        pass
        
        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑0...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的Python程序路径
        script1 = r"test3.py"

        # 启动程序
        process1 = subprocess.Popen([venv_python, script1])
        process1.wait()

        self.show()  # 执行完成后恢复显示
    
    def start_1(self):
        pass
    
        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑1...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的Python程序路径
        script1 = r"test4.py"

        # 启动程序
        process1 = subprocess.Popen([venv_python, script1])
        process1.wait()

        self.show()  # 执行完成后恢复显示
    
    def start_2(self):
        pass
    
        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑2...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的Python程序路径
        script1 = r"test1.py"

        # 启动程序
        process1 = subprocess.Popen([venv_python, script1])
        process1.wait()

        self.show()  # 执行完成后恢复显示
    
    def start_3(self):
        pass
    
        """按钮点击后触发的新函数"""
        self.hide()  # 立即隐藏窗口

        # 这里添加点击按钮后的具体业务逻辑
        print("启动业务逻辑3...")

        # 虚拟环境Python解释器路径
        venv_python = get_python_interpreter_path()

        # 需要启动的Python程序路径
        script1 = r"test2.py"

        # 启动程序
        process1 = subprocess.Popen([venv_python, script1])
        process1.wait()

        self.show()  # 执行完成后恢复显示

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())