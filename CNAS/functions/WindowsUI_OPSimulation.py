from typing import List, Optional, Tuple
from enum import Enum

import pyautogui as pag
import pyautogui
import pygetwindow as gw
import win32api
import win32gui
import win32con

import subprocess
import ctypes
import keyboard
import time
import re
import os

def monitor_mouse_position( interval=0.5, show_screen_size=True):
    """
    实时监控并输出鼠标位置信息

    :param interval: 刷新间隔时间（秒），默认0.5秒
    :param show_screen_size: 是否显示屏幕尺寸，默认True

    示例用法：
    monitor_mouse_position(interval=0.3, show_screen_size=False)
    """

    # 动态生成输出信息
    output = []

    try:
        # 获取坐标数据
        x, y = pag.position()

        if show_screen_size:
            screen_size = pag.size()
            # output.append(f"Screen: {screen_size[0]}x{screen_size[1]}")

        output.append(x)
        output.append(y)
        # output.append(f"Cursor: ({x}, {y})")
        # print(" | ".join(output))

        # time.sleep(interval)

    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        print("Monitoring cleanup completed")

    return output

def monitor_mouse_position_1(interval=0.5, show_screen_size=True, exit_message="Monitoring terminated"):
    """
    实时监控并输出鼠标位置信息

    :param interval: 刷新间隔时间（秒），默认0.5秒
    :param show_screen_size: 是否显示屏幕尺寸，默认True
    :param exit_message: 退出时提示消息，默认"Monitoring terminated"

    示例用法：
    >>> monitor_mouse_position(interval=0.3, show_screen_size=False)
    """
    try:
        while True:
            # 获取坐标数据
            x, y = pag.position()

            # 动态生成输出信息
            output = []
            if show_screen_size:
                screen_size = pag.size()
                output.append(f"Screen: {screen_size[0]}x{screen_size[1]}")

            output.append(f"Cursor: ({x}, {y})")
            print(" | ".join(output))

            time.sleep(interval)

    except KeyboardInterrupt:
        print(f"\n{exit_message}")
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        print("Monitoring cleanup completed")


class ClickType(Enum):
    LEFT = 'left'
    RIGHT = 'right'
    DOUBLE = 'double'
    MIDDLE = 'middle'

def smart_click(
        x: Optional[int] = None,
        y: Optional[int] = None,
        *,
        image_path: Optional[str] = None,
        click_type: ClickType = ClickType.LEFT,
        offset: Tuple[int, int] = (0, 0),
        move_duration: float = 0.2,
        retries: int = 3,
        confidence: float = 0.9,
        fail_safe: bool = True,
        log_prefix: str = ""
) -> bool:
    """
    智能鼠标点击函数（支持坐标定位和图像识别两种模式）

    :param x: 目标X坐标（与image_path二选一）
    :param y: 目标Y坐标（与image_path二选一）
    :param image_path: 图像匹配路径（优先使用）
    :param click_type: 点击类型，默认左键
    :param offset: 相对于目标位置的偏移量，默认(0,0)
    :param move_duration: 鼠标移动时间（秒），0表示瞬间移动
    :param retries: 最大重试次数
    :param confidence: 图像匹配置信度（0~1）
    :param fail_safe: 启用安全保护（触发移动到角落终止）
    :param log_prefix: 日志前缀标识
    :return: 是否成功执行点击
    """
    pag.FAILSAFE = fail_safe

    def validate_coordinates(coord_x: int, coord_y: int) -> bool:
        """验证坐标是否在屏幕范围内"""
        width, height = pag.size()
        return 0 <= coord_x <= width and 0 <= coord_y <= height

    def perform_click(Target_x: int, Target_y: int) -> None:
        """执行具体的点击操作"""
        if move_duration > 0:
            pag.moveTo(Target_x, Target_y, duration=move_duration)
        else:
            pag.moveTo(Target_x, Target_y)

        if click_type == ClickType.LEFT:
            pag.click()
        elif click_type == ClickType.RIGHT:
            pag.rightClick()
        elif click_type == ClickType.DOUBLE:
            pag.doubleClick()
        elif click_type == ClickType.MIDDLE:
            pag.middleClick()

    for attempt in range(1, retries + 1):
        try:
            # 图像识别模式优先
            if image_path:
                if match := pag.locateOnScreen(image_path, confidence=confidence):
                    base_x, base_y = pag.center(match)
                    target_x = base_x + offset[0]
                    target_y = base_y + offset[1]
                    print(f"{log_prefix}图像匹配成功：{image_path} → ({target_x}, {target_y})")
                else:
                    raise pag.ImageNotFoundException(f"未找到匹配图像: {image_path}")
            else:
                if x is None or y is None:
                    raise ValueError("必须提供坐标或图像路径")
                target_x, target_y = x + offset[0], y + offset[1]

            if not validate_coordinates(target_x, target_y):
                raise pag.FailSafeException(f"无效坐标: ({target_x}, {target_y})")

            perform_click(target_x, target_y)
            return True

        except pag.ImageNotFoundException as e:
            print(f"{log_prefix}图像识别失败（尝试 {attempt}/{retries}）: {str(e)}")
            time.sleep(1)
        except pag.FailSafeException as e:
            print(f"{log_prefix}安全保护触发: {str(e)}")
            return False
        except Exception as e:
            print(f"{log_prefix}未知错误: {str(e)}")
            time.sleep(0.5)

    return False

def clean_title(title: str) -> str:
    """深度清理窗口标题中的非常规字符"""
    # 扩展清理规则：移除零宽字符、控制字符、非常用符号，保留常规文本
    return re.sub(
        r'[\u200b-\u200f\u202a-\u202e\ufeff\x00-\x1f\\:：<>*?|]',
        '',
        str(title).strip().lower()
    )

# def get_browser_window(title_part: str) -> gw.Window:
#     """
#     根据标题关键字查找浏览器窗口（模糊匹配）
#     """
#     windows = gw.getWindowsWithTitle(title_part)
#     for window in windows:
#         # 放宽条件：检查标题是否包含 "edge" 或 "chrome" 或 "firefox"，不区分大小写
#         lower_title = window.title.lower()
#         if "edge" in lower_title or "chrome" in lower_title or "firefox" in lower_title:
#             return window
#     return None

def get_browser_window(title_part: str):
    """
    模糊匹配浏览器窗口（支持处理特殊字符）
    """
    cleaned_target = clean_title(title_part)
    for window in gw.getAllWindows():
        window_title_clean = clean_title(window.title)
        # 检查标题是否包含目标关键词，且是浏览器窗口
        if (cleaned_target in window_title_clean) and \
           ("edge" in window_title_clean or "chrome" in window_title_clean or "firefox" in window_title_clean):
            return window
    return None

def find_activate_window(
        keyword: str,
        max_retries: int = 3,
        match_mode: str = 'contains'
) -> Optional[gw.Window]:
    """
    查找并激活第一个匹配的Windows窗口

    :param keyword: 搜索关键词（支持模糊匹配）
    :param max_retries: 最大重试次数（应对窗口加载延迟）
    :param match_mode: 匹配模式 ('contains', 'startswith', 'endswith')
    :return: 窗口对象或None
    """
    cleaned_keyword = clean_title(keyword)

    for _ in range(max_retries):
        windows = gw.getAllWindows()
        for window in windows:
            if not window.title:
                continue

            processed_title = clean_title(window.title)

            # 根据匹配模式检查
            match match_mode:
                case 'contains' if cleaned_keyword in processed_title:
                    pass
                case 'startswith' if processed_title.startswith(cleaned_keyword):
                    pass
                case 'endswith' if processed_title.endswith(cleaned_keyword):
                    pass
                case _:
                    continue

            try:
                if window.isMinimized:
                    window.restore()
                window.activate()
                time.sleep(0.3)  # 等待窗口前置
                return window
            except gw.PyGetWindowException as e:
                print(f"窗口操作失败: {e}")
                continue

        time.sleep(1)  # 重试间隔

    return None

def list_all_browsers():
    browsers = []
    for window in gw.getAllWindows():
        title = window.title.lower()
        browsers.append({
            "title": window.title,
            "position": (window.left, window.top),
            "size": (window.width, window.height)
        })
    return browsers

def force_maximize_and_foreground(window: gw.Window) -> None:
    """
    最大化窗口并强制置顶
    """
    if window.isMinimized:
        window.restore()  # 先恢复窗口（如果已最小化）
    window.maximize()
    # 通过窗口句柄调用 Win32 API 强制置顶
    hwnd = window._hWnd  # 获取窗口句柄（pygetwindow 内部属性）
    win32gui.SetForegroundWindow(hwnd)  # 置顶窗口

def safe_type(text: str, interval=0.01):
    """安全输入函数（防输入法干扰）"""
    # 切换英文输入法（需系统已添加英文输入法）
    # pyautogui.hotkey('ctrl', 'shift')  # 切换输入法组合键
    # time.sleep(0.5)

    # 逐字符输入
    for char in text:
        keyboard.write(char)
        time.sleep(interval)

def smart_input(content: str, enter=True):
    """智能输入函数"""
    # 处理特殊符号
    special_chars = {
        '\n': 'enter',
        '\t': 'tab',
        ' ': 'space'
    }

    # 执行输入
    for char in content:
        if char in special_chars:
            pyautogui.press(special_chars[char])
        else:
            safe_type(char)

    if enter:
        pyautogui.press('enter')

def click_image( image_path, confidence=0.8, region=(0, 0, 0, 0), grayscale=False, timeout=10):
    start_time = time.time()
    while True:

        # 查找图片位置
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)

        if location:

            center = pyautogui.center(location)
            pyautogui.click(center)

            print(f"成功点击：{image_path}")
            return True
        else:

            # 超时处理
            if time.time() - start_time > timeout:
                print(f"超时未找到图片：{image_path}")
                return False
            time.sleep(0.5)  # 每0.5秒重试

def click_on_screen( image_path, confidence=0.8, timeout=10):
    """
    在指定时间内循环匹配目标图像并执行点击

    参数：
    image_path : str   目标图像文件路径
    confidence : float 匹配可信度阈值（0-1），默认0.8
    timeout    : int   匹配超时时间（秒），默认10秒

    返回：
    bool 是否成功执行点击
    """
    # 启用ImageNotFoundException异常模式
    pyautogui.useImageNotFoundException()  #

    try:
        # 获取屏幕参数
        physical_width = win32api.GetSystemMetrics(0)
        physical_height = win32api.GetSystemMetrics(1)
        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100  #

        # 主屏幕匹配区域定义
        region = (0, 0, physical_width, physical_height)

        # 记录开始时间
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 带置信度的图像匹配
                target_pos = pyautogui.locateOnScreen(
                    image_path,
                    region=region,
                    confidence=confidence,
                    grayscale=True  # 灰度匹配提升稳定性
                )

                if target_pos:
                    # 坐标转换处理
                    center_x, center_y = pyautogui.center(target_pos)
                    physical_x = int(center_x * dpi_scale)
                    physical_y = int(center_y * dpi_scale)

                    # 执行点击动作
                    pyautogui.click(physical_x, physical_y)
                    return True

                # 降低CPU占用
                time.sleep(0.5)

            except pyautogui.ImageNotFoundException:
                # 本轮未找到继续循环
                continue

        # 超时未找到
        print(f"Timeout: Image not found within {timeout} seconds")
        return False

    except pyautogui.PyAutoGUIException as e:
        print(f"PyAutoGUI Error: {str(e)}")
        return False
    except Exception as e:
        print(f"Unexpected Error: {str(e)}")
        return False

def click_on_screen_0( image_path, direction=6, distance=0, confidence=0.8, timeout=10):
    """
    在指定时间内循环匹配目标图像并执行点击

    参数：
    image_path : str   目标图像文件路径
    confidence : float 匹配可信度阈值（0-1），默认0.8
    timeout    : int   匹配超时时间（秒），默认10秒

    返回：
    bool 是否成功执行点击
    """
    # 启用ImageNotFoundException异常模式
    pyautogui.useImageNotFoundException()  #

    try:
        # 获取屏幕参数
        physical_width = win32api.GetSystemMetrics(0)
        physical_height = win32api.GetSystemMetrics(1)
        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100  #

        # 主屏幕匹配区域定义
        region = (0, 0, physical_width, physical_height)

        # 记录开始时间
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 带置信度的图像匹配
                target_pos = pyautogui.locateOnScreen(
                    image_path,
                    region=region,
                    confidence=confidence,
                    grayscale=True  # 灰度匹配提升稳定性
                )

                if target_pos:
                    # 坐标转换处理
                    center_x, center_y = pyautogui.center(target_pos)
                    physical_x = int(center_x * dpi_scale)
                    physical_y = int(center_y * dpi_scale)

                    if  direction == 12:
                        physical_y = physical_y - distance

                    elif direction == 3:
                        physical_x = physical_x + distance

                    elif direction == 6:
                        physical_y = physical_y + distance

                    elif direction == 9:
                        physical_x = physical_x - distance

                    # 执行点击动作
                    pyautogui.click(physical_x, physical_y)
                    return True

                # 降低CPU占用
                time.sleep(0.5)

            except pyautogui.ImageNotFoundException:
                # 本轮未找到继续循环
                continue

        # 超时未找到
        print(f"Timeout: Image not found within {timeout} seconds")
        return False

    except pyautogui.PyAutoGUIException as e:
        print(f"PyAutoGUI Error: {str(e)}")
        return False
    except Exception as e:
        print(f"Unexpected Error: {str(e)}")
        return False

def click_on_screen_1( image_x, image_y, confidence=0.8, timeout=10):
    """
    在指定时间内循环匹配目标图像并执行点击

    参数：
    image_x    : str   目标图像文件路径(水平)
    image_x    : str   目标图像文件路径(竖直)
    confidence : float 匹配可信度阈值（0-1），默认0.8
    timeout    : int   匹配超时时间（秒），默认10秒

    返回：
    bool 是否成功执行点击
    """
    # 启用ImageNotFoundException异常模式
    pyautogui.useImageNotFoundException()  #

    try:
        # 获取屏幕参数
        physical_width = win32api.GetSystemMetrics(0)
        physical_height = win32api.GetSystemMetrics(1)
        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100  #

        # 主屏幕匹配区域定义
        region = (0, 0, physical_width, physical_height)

        # 记录开始时间
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 带置信度的图像匹配
                target_pos_1 = pyautogui.locateOnScreen(
                    image_x,
                    region=region,
                    confidence=confidence,
                    grayscale=True  # 灰度匹配提升稳定性
                )
                # 带置信度的图像匹配
                target_pos_2 = pyautogui.locateOnScreen(
                    image_y,
                    region=region,
                    confidence=confidence,
                    grayscale=True  # 灰度匹配提升稳定性
                )

                if target_pos_1 and target_pos_2:
                    # 坐标转换处理
                    center_x, temp = pyautogui.center(target_pos_1)
                    print( center_x, temp)
                    temp, center_y = pyautogui.center(target_pos_2)
                    print( temp, center_y)
                    physical_x = int(center_x * dpi_scale)
                    physical_y = int(center_y * dpi_scale)
                    print( physical_x, physical_y)

                    # 执行点击动作
                    # pyautogui.click(physical_x, physical_y)
                    return True

                # 降低CPU占用
                time.sleep(0.5)

            except pyautogui.ImageNotFoundException:
                # 本轮未找到继续循环
                continue

        # 超时未找到
        print(f"Timeout: Image not found within {timeout} seconds")
        return False

    except pyautogui.PyAutoGUIException as e:
        print(f"PyAutoGUI Error: {str(e)}")
        return False
    except Exception as e:
        print(f"Unexpected Error: {str(e)}")
        return False

def scroll_on_screen( image_path, confidence=0.8, timeout=10, distance=-1000):
    pass

    """
        在指定时间内循环匹配目标图像并执行点击

        参数：
        image_path : str   目标图像文件路径
        confidence : float 匹配可信度阈值（0-1），默认0.8
        timeout    : int   匹配超时时间（秒），默认10秒

        返回：
        bool 是否成功执行点击
        """
    # 启用ImageNotFoundException异常模式
    pyautogui.useImageNotFoundException()  #

    try:
        # 获取屏幕参数
        physical_width = win32api.GetSystemMetrics(0)
        physical_height = win32api.GetSystemMetrics(1)
        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100  #

        # 主屏幕匹配区域定义
        region = (0, 0, physical_width, physical_height)

        # 记录开始时间
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                # 带置信度的图像匹配
                target_pos = pyautogui.locateOnScreen(
                    image_path,
                    region=region,
                    confidence=confidence,
                    grayscale=True  # 灰度匹配提升稳定性
                )

                if target_pos:
                    # 坐标转换处理
                    center_x, center_y = pyautogui.center(target_pos)
                    physical_x = int(center_x * dpi_scale)
                    physical_y = int(center_y * dpi_scale)

                    pyautogui.moveTo( physical_x, physical_y, duration=0.5)

                    # 执行滚动动作
                    print("action")
                    pyautogui.scroll(distance)

                    return True

                # 降低CPU占用
                time.sleep(0.5)

            except pyautogui.ImageNotFoundException:
                # 本轮未找到继续循环
                continue

        # 超时未找到
        print(f"Timeout: Image not found within {timeout} seconds")
        return False


    except pyautogui.PyAutoGUIException as e:
        print(f"PyAutoGUI Error: {str(e)}")
        return False
    except Exception as e:
        print(f"Unexpected Error: {str(e)}")
        return False

def set_primary_dpi_to_100():
    try:
        # 检测当前缩放比例
        current_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        if current_scale == 1.0:
            print("Already 100% scaling.")
            return True

        # 调用 SetDPI.exe
        script_dir = os.path.dirname(os.path.realpath(__file__))
        setdpi_path = os.path.join( script_dir, r"SetDPI.exe")
        subprocess.run([setdpi_path, "100"], check=True, shell=True)

        # 强制刷新系统设置
        ctypes.windll.user32.SystemParametersInfoW(win32con.SPI_SETDESKWALLPAPER, 0, None, 0)
        print("DPI set to 100% successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"SetDPI failed: {e}")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
    return False

def move_window_to_primary_screen(window: gw.Window) -> bool:
    """
    将指定窗口移动到主屏幕并置顶最大化

    参数：
    window : gw.Window  pygetwindow获取的窗口对象

    返回：
    bool  是否操作成功
    """
    try:
        # 获取主屏幕参数（物理分辨率）
        primary_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        primary_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

        # 恢复最小化窗口（需先恢复才能移动）
        if window.isMinimized:
            window.restore()
            time.sleep(0.2)

        # 移动窗口到主屏幕左上角
        window.moveTo(0, 0)

        # 置顶窗口（通过Win32 API实现永久置顶）
        win32gui.SetWindowPos(
            window._hWnd,  # 从pygetwindow对象获取原生句柄
            win32con.HWND_TOPMOST,  # 置顶标志
            0, 0, 0, 0,  # 位置尺寸由后续参数控制
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE  # 保持当前位置和大小
        )

        # 最大化窗口（兼容不同缩放模式）
        window.maximize()

        # 二次验证窗口位置
        if (window.left, window.top) != (0, 0):
            window.moveTo(0, 0)

        return True

    except (gw.PyGetWindowException, win32gui.error) as e:
        print(f"窗口操作失败: {e}")
        return False

# 示例：操作浏览器窗口
if __name__ == "__main__":
    pass

    # 基础用法
    monitor_mouse_position()

    # 高级用法示例
    # monitor_mouse_position(
    #     interval=0.2,
    #     show_screen_size=False
    # )

    monitor_mouse_position_1(
        interval=0.2,
        show_screen_size=False
    )

    # 示例：使用更简洁的关键词匹配
    # target_window_1 = get_browser_window("deepseek")
    # target_window = find_activate_window("记事本")
    # if target_window:
    #     print(f"窗口标题: {target_window.title}")
    #     print(f"窗口位置: {target_window.left}, {target_window.top}")
    #     print(f"窗口大小: {target_window.width}x{target_window.height}")
    #     force_maximize_and_foreground(target_window)  # 最大化并置顶
    # else:
    #     print("未找到目标窗口！")
    #
    # smart_input("Hello World! 你好，世界！", enter=True)
    #
    # # 列出所有浏览器窗口
    # browser_list = list_all_browsers()
    # for idx, browser in enumerate(browser_list):
    #     print(f"窗口 {idx+1}: {browser['title']}")

    # set_primary_dpi_to_100()

    # image = r"D:\CPS-OA_RPA-20250312\config\IMG\1\2025-03-25_022244.png"
    # image_1 = r"D:\CPS-OA_RPA-20250312\config\IMG\1\2025-03-25_022406.png"
    # image_2 = r"D:\CPS-OA_RPA-20250312\config\IMG\1\2025-03-25_022417.png"

    # target_window = find_activate_window("监理一体化")

    # status = move_window_to_primary_screen(target_window)
    # status = scroll_on_screen( image, confidence=0.5, distance=-800)
    # status = click_on_screen_0( image_2, direction=6, confidence=0.5, distance=523)
    # status = click_on_screen_1( image_2, image_1, confidence=0.5)
    # print(status)