from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QLabel, QWidget
import websockets
import asyncio
import json

class TimerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('计时器')
        self.setGeometry(100, 100, 300, 200)

        self.layout = QVBoxLayout()

        # 标签显示计时
        self.timer_label = QLabel("00:00", self)
        self.layout.addWidget(self.timer_label)

        # 按钮控制计时
        self.start_button = QPushButton('开始/暂停', self)
        self.start_button.clicked.connect(self.toggle_timer)
        self.layout.addWidget(self.start_button)

        # 创建一个定时器，每秒更新计时
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer)
        self.elapsed_time = 0  # 计时秒数

        # WebSocket连接标志
        self.websocket = None
        self.is_running = False

        # 设置主窗口
        container = QWidget()
        container.setLayout(self.layout)
        self.setCentralWidget(container)

    async def connect_websocket(self):
        uri = "ws://localhost:8765"
        self.websocket = await websockets.connect(uri)

    async def send_message(self, message):
        if self.websocket:
            await self.websocket.send(json.dumps(message))

    def toggle_timer(self):
        """切换计时器的暂停/继续状态"""
        if not self.is_running:
            self.is_running = True
            self.timer.start(1000)  # 启动计时器
            asyncio.ensure_future(self.send_message({"action": "start_timer"}))  # 使用ensure_future来启动协程
        else:
            self.is_running = False
            self.timer.stop()  # 停止计时器
            asyncio.ensure_future(self.send_message({"action": "pause_timer"}))

    def update_timer(self):
        """更新计时显示"""
        self.elapsed_time += 1
        minutes = self.elapsed_time // 60
        seconds = self.elapsed_time % 60
        time_str = f"{minutes:02}:{seconds:02}"
        self.timer_label.setText(time_str)

    def closeEvent(self, event):
        """窗口关闭时断开WebSocket连接"""
        if self.websocket:
            asyncio.ensure_future(self.websocket.close())
        event.accept()

# 启动PyQt应用，并运行事件循环
def run_app():
    app = QApplication([])
    window = TimerApp()
    window.show()
    asyncio.ensure_future(window.connect_websocket())  # 连接WebSocket服务器
    app.exec_()

if __name__ == "__main__":
    run_app()
