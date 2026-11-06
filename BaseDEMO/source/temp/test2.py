import win32gui
import win32con

def get_window_info(window_title: str) -> dict:
    """
    获取指定标题窗口的句柄、位置和大小
    """
    hwnd = win32gui.FindWindow(None, window_title)  # 根据窗口标题查找句柄
    if not hwnd:
        return None
    rect = win32gui.GetWindowRect(hwnd)  # 获取窗口坐标和大小 (left, top, right, bottom)
    return {
        "hwnd": hwnd,
        "position": (rect[0], rect[1]),  # 窗口左上角坐标
        "size": (rect[2] - rect[0], rect[3] - rect[1])  # 窗口宽高
    }

def maximize_window(hwnd: int) -> None:
    """
    最大化指定窗口
    """
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

def force_maximize(hwnd: int) -> None:
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)  # 先恢复窗口（防止最小化）
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(hwnd)  # 置顶窗口

# 示例：操作 Chrome 窗口
if __name__ == "__main__":
    target_title = "教学管理"  # 浏览器窗口标题（需完整或部分匹配）
    window_info = get_window_info(target_title)
    
    if window_info:
        print(f"窗口句柄: {window_info['hwnd']}")
        print(f"窗口位置: {window_info['position']}")
        print(f"窗口大小: {window_info['size']}")
        maximize_window(window_info['hwnd'])  # 最大化窗口
    else:
        print("未找到目标窗口！")