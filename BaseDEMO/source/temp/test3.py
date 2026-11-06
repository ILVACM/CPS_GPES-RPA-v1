import pyautogui as pag
import time

try:
    while True:
        x, y = pag.position()  # 获取当前坐标
        screen_width, screen_height = pag.size()  # 获取屏幕分辨率
        print(f"屏幕尺寸：({screen_width}, {screen_height})，当前坐标：({x}, {y})")
        time.sleep(0.5)  # 控制刷新频率，避免高CPU占用
except KeyboardInterrupt:
    print("监控已终止")