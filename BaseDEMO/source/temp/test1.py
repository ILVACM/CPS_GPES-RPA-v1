import pygetwindow as gw
import win32gui
import win32con

import re

def clean_title(title: str) -> str:
    """清理标题中的特殊字符（如零宽空格）"""
    return re.sub(r'[\u200b:]', '', title).strip().lower()

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

def get_browser_window(title_part: str) -> gw.Window:
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

def list_all_browsers():
    browsers = []
    for window in gw.getAllWindows():
        title = window.title.lower()
        if "chrome" in title or "firefox" in title or "edge" in title:
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

# 示例：操作浏览器窗口
if __name__ == "__main__":        
    # 示例：使用更简洁的关键词匹配
    target_window = get_browser_window("教学管理系统")
    if target_window:
        print(f"窗口标题: {target_window.title}")
        print(f"窗口位置: {target_window.left}, {target_window.top}")
        print(f"窗口大小: {target_window.width}x{target_window.height}")
        force_maximize_and_foreground(target_window)  # 最大化并置顶
    else:
        print("未找到目标窗口！")

    # 列出所有浏览器窗口
    browser_list = list_all_browsers()
    for idx, browser in enumerate(browser_list):
        print(f"窗口 {idx+1}: {browser['title']}")