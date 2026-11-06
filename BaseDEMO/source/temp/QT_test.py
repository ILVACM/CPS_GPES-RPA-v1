import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QRadioButton, QGroupBox, QProgressBar, QDateTimeEdit
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("本地应用示例")
        self.setGeometry(100, 100, 800, 600)

        # 主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)

        # 输入部分
        input_group = QGroupBox("设置")
        input_layout = QVBoxLayout()
        input_group.setLayout(input_layout)

        # 用户名输入
        input_layout.addWidget(QLabel("个人用户名"))
        self.username_input = QLineEdit()
        input_layout.addWidget(self.username_input)

        # 添加操作按钮
        action_group = QGroupBox("操作设置")
        action_layout = QHBoxLayout()
        action_group.setLayout(action_layout)
        self.add_button = QPushButton("增")
        self.delete_button = QPushButton("删")
        action_layout.addWidget(self.add_button)
        action_layout.addWidget(self.delete_button)

        # 日期时间选择
        send_group = QGroupBox("发送设置")
        send_layout = QVBoxLayout()
        send_group.setLayout(send_layout)
        self.datetime_edit = QDateTimeEdit()
        self.datetime_edit.setCalendarPopup(True)
        send_layout.addWidget(QLabel("定时发送时间"))
        send_layout.addWidget(self.datetime_edit)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(11)

        # 将所有组添加到主布局
        main_layout.addWidget(input_group)
        main_layout.addWidget(action_group)
        main_layout.addWidget(send_group)
        main_layout.addWidget(self.progress_bar)

        self.setCentralWidget(main_widget)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
