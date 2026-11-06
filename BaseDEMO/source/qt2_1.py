import sys
import time
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QDesktopWidget, QProgressBar, QLabel, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt, QTimer, QRectF
from PyQt5.QtGui import QPainter, QColor, QPainterPath,  QFont

class MovingBlocksWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.blocks = []  # 用于存储方块的位置和大小
        self.speed = 5    # 方块移动速度
        self.block_height = 30  # 方块高度
        self.initBlocks()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.updateBlocks)
        self.timer.start(30)  # 每30毫秒更新一次

    def initBlocks(self):
        # 初始化多个方块，设置它们的初始位置和大小
        for i in range(6):
            x = 50 + i * 60  # 每个方块的初始x坐标
            y = self.height() // 2 - self.block_height // 2  # 垂直居中
            width = 40       # 方块宽度
            height = self.block_height  # 方块高度
            self.blocks.append({'rect': QRectF(x, y, width, height), 'color': QColor(243, 199, 82)})

    def updateBlocks(self):
        # 更新每个方块的位置
        for block in self.blocks:
            new_x = block['rect'].x() + self.speed
            if new_x > self.width():
                new_x = -block['rect'].width()  # 重置到左边界
            block['rect'].moveTo(new_x, block['rect'].y())
        self.update()  # 触发重绘

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)  # 使绘制更平滑
        
        # 设置背景颜色和圆角边框
        background_color = QColor(212, 217, 232)
        path = QPainterPath()
        # 创建一个与方块高度一致的矩形区域，并添加圆角
        rect = QRectF(0, self.height() // 2 - self.block_height // 2, self.width(), self.block_height)
        path.addRoundedRect(rect, 10, 10)  # 圆角半径为10
        painter.fillPath(path, background_color)
        
        # 设置裁剪路径，确保方块只在背景区域内绘制
        painter.setClipPath(path)

        #绘制方块
        for block in self.blocks:
            painter.fillRect(block['rect'], block['color'])

    def resizeEvent(self, event):
        # 重置方块位置以适应窗口大小变化
        self.initBlocks()
        super().resizeEvent(event)

class NoTitleBarWindow(QWidget):
    # 定义一个静态变量来保存实例引用
    instance = None
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        
        # 设置窗口为无标题栏
        self.setWindowFlags( self.windowFlags() |Qt.FramelessWindowHint| Qt.WindowStaysOnTopHint)  # 移除标题栏，去除默认边框
        self.setAttribute(Qt.WA_TranslucentBackground)  # 去掉窗口默认装饰，使背景透明，这意味着窗口的背景不会被绘制，可以用来实现自定义的窗口形状。
        
        # 获取屏幕尺寸
        desktop = QDesktopWidget().screenGeometry()
        screen_width = desktop.width()
        screen_height = desktop.height()
        
        # 计算窗口尺寸（1/4 宽度，1/8 高度）
        window_width = int(screen_width / 5)
        window_height = int(screen_height / 10)
        
        # 设置窗口大小
        self.resize(window_width, window_height)
        
        # 获取屏幕的几何信息
        screen_geometry = QDesktopWidget().screenGeometry()
        screen_width = screen_geometry.width()

        # 计算窗口左上角坐标
        x = (screen_width - self.width()) // 2  # 水平居中
        y = 0  # 垂直置顶

        # 设置窗口位置
        self.move(x, y)
        
        
        # 创建主布局（垂直布局）
        main_layout = QVBoxLayout()

        # 创建顶部状态标签
        status_label = QLabel("GDUT_CPS-Auto is Running")
        status_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        status_label.setStyleSheet("font-size: 26px; font-weight: bold;font-family: 'Tahoma';")
        # 设置字体
        # font = QFont("Roboto", 10, QFont.Bold)  # 字体名称、大小、粗细
        # status_label.setFont(font)
        main_layout.addWidget(status_label)
        

        # 创建动画方块部件（取代进度条）
        moving_blocks = MovingBlocksWidget()
        moving_blocks.setFixedSize(window_width - 20, 50)  # 设置动画部件的大小
        main_layout.addWidget(moving_blocks)

        # 创建底部控制栏（水平布局）
        control_layout = QHBoxLayout()
        
        # 添加伸缩空间，使按钮靠右对齐
        spacer = QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        control_layout.addSpacerItem(spacer)

        # 创建暂停按钮
        pause_button = QPushButton("||")
        pause_button.setFixedSize(30, 30)
        pause_button.setStyleSheet("font-size: 22px; font-weight: bold; color: black;border-width: 2px; border-style: solid; border-color: black;border-radius: 5px;")
        control_layout.addWidget(pause_button)

        # 添加分割符
        separator = QLabel("|")
        separator.setAlignment(Qt.AlignCenter)
        separator.setFixedSize(20, 20)
        separator.setStyleSheet("font-size: 24px;font-weight: bold;")
        control_layout.addWidget(separator)

        # 创建终止按钮
        stop_button = QPushButton("\u25CF")#原本设为 方块符号 ■，现在改成实心圆，更美观
        stop_button.setFixedSize(30, 30)
        stop_button.setStyleSheet("font-size: 11px; font-weight: bold; color: black; border-width: 2px; border-style: solid; border-color: black;border-radius: 5px;")
        stop_button.clicked.connect(self.close)#点击则关闭程序
        control_layout.addWidget(stop_button)

        # 将水平布局添加到主布局
        main_layout.addLayout(control_layout)

        # 设置窗口布局
        self.setLayout(main_layout)
        
        self.show()
        
    #自定义绘制
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制白色背景
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), 10, 10)  # 圆角半径为10
        painter.fillPath(path, QColor(255, 255, 255))  # 填充白色背景

        # 绘制边框
        painter.setPen(Qt.black)  # 设置边框颜色
        painter.drawPath(path)
        
    @staticmethod
    def open_window():
        app = QApplication(sys.argv)
        if NoTitleBarWindow.instance is None:
            NoTitleBarWindow.instance = NoTitleBarWindow()
        NoTitleBarWindow.instance.show()
        sys.exit(app.exec_())
    
    # 定义一个静态方法用于关闭窗口
    @staticmethod
    def close_window():
        if NoTitleBarWindow.instance is not None:
            NoTitleBarWindow.instance.close()


# 如果直接运行此文件，则打开窗口
if __name__ == '__main__':
    NoTitleBarWindow.open_window()
    
    