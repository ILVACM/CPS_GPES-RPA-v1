import pandas as pd
import os
import sys
import win32api
import pyautogui
import time
import ctypes

def check_excel_deps():
    """检查必要的依赖库是否安装"""
    required = {
        '.xlsx': ['openpyxl'],
        '.xls': ['xlwt', 'xlrd']
    }

    missing = []
    for ext, libs in required.items():
        for lib in libs:
            try:
                __import__(lib)
            except ImportError:
                missing.append(f"{lib} (用于{ext}格式)")

    if missing:
        print("❌ 缺少必要依赖库:")
        print("\n".join(f"  - {lib}" for lib in missing))
        print("\n请执行以下命令安装:")
        print("pip install xlwt xlrd openpyxl")
        sys.exit(1)

def pandas_excel_copy(source_path: str, target_path: str) -> bool:
    """
    增强版Excel复制函数（自动处理格式兼容性问题）

    功能改进：
    1. 自动处理旧版Excel格式限制
    2. 数据类型兼容性转换
    3. 详细的错误提示
    """
    try:
        # 前置依赖检查
        check_excel_deps()

        # 路径验证
        if not os.path.exists(source_path):
            raise FileNotFoundError(f"源文件不存在: {source_path}")

        # 创建目标目录
        target_dir = os.path.dirname(target_path)
        if target_dir and not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)

        # 自动检测文件格式
        src_ext = os.path.splitext(source_path)[1].lower()
        dst_ext = os.path.splitext(target_path)[1].lower()

        # 读取数据
        with pd.ExcelFile(source_path, engine='xlrd' if src_ext == '.xls' else None) as src:
            sheets_data = {}
            for sheet_name in src.sheet_names:
                df = pd.read_excel(src, sheet_name=sheet_name)

                # 数据类型兼容处理
                df = df.applymap(lambda x: x.strftime('%Y-%m-%d %H:%M:%S')
                if pd.isna(x) == False and isinstance(x, pd.Timestamp)
                else x)

                sheets_data[sheet_name] = df

        # 写入引擎配置
        engine_config = {
            '.xlsx': {'engine': 'openpyxl', 'options': {}},
            '.xls': {
                'engine': 'xlwt',
                'options': {
                    'datetime_format': 'YYYY-MM-DD',
                    'encoding': 'utf-8'
                }
            }
        }

        if dst_ext not in engine_config:
            raise ValueError(f"不支持的输出格式: {dst_ext}，仅支持.xls/.xlsx")

        config = engine_config[dst_ext]

        # 处理旧版Excel限制
        if dst_ext == '.xls':
            for sheet_name, df in sheets_data.items():
                if len(df) > 65535:
                    print(f"⚠️ 警告: {sheet_name} 超过65535行限制，自动截断")
                    sheets_data[sheet_name] = df.iloc[:65535]

        # 写入文件
        with pd.ExcelWriter(target_path,
                            engine=config['engine'],
                            **config['options']) as writer:

            for sheet_name, df in sheets_data.items():
                df.to_excel(
                    writer,
                    sheet_name=sheet_name[:31],  # 旧版工作表名称31字符限制
                    index=False,
                    header=not df.empty,
                    na_rep=''  # 处理空值
                )
                print(f"✅ 已复制: {sheet_name} ({len(df)}行x{len(df.columns)}列)")

        print(f"\n🎉 成功导出到: {target_path}")
        print(f"   源格式: {src_ext} → 目标格式: {dst_ext}")
        return True

    except PermissionError as e:
        print(f"\n❌ 文件访问错误: {str(e)}")
        print("可能原因：")
        print("- 文件被其他程序打开")
        print("- 没有写入权限")
    except Exception as e:
        print(f"\n❌ 操作失败: {str(e)}")
        print("调试建议：")
        print("- 检查文件路径是否包含特殊字符")
        print("- 尝试用Excel打开源文件确认是否损坏")

    return False

# 使用示例
# if __name__ == "__main__":
#     pass

    # 测试不同格式转换
    # test_cases = [
    #     ("风险管控到位列表.xls", "output.xls"),
    #     ("风险管控到位列表1.xlsx", "output1.xls"),
    #     ("风险管控到位列表.xls", "output2.xlsx"),
    #     ("风险管控到位列表1.xlsx", "output3.xlsx")
    # ]
    #
    # for src, dst in test_cases:
    #     print(f"\n{'=' * 30}")
    #     print(f"测试转换: {src} → {dst}")
    #     pandas_excel_copy(src, dst)

def click_image( image_path, confidence=0.8, region=(0, 0, 0, 0), grayscale=False, timeout=10):
    start_time = time.time()
    while True:

        # 查找图片位置
        location = pyautogui.locateOnScreen(image_path, confidence=confidence)

        if location:

            center = pyautogui.center(location)
            print(center)

            pyautogui.click(center)

            print(f"成功点击：{image_path}")
            return True
        else:

            # 超时处理
            if time.time() - start_time > timeout:
                print(f"超时未找到图片：{image_path}")
                return False
            time.sleep(0.5)  # 每0.5秒重试

def get_main_screen_info():
    """获取主屏幕的尺寸和位置"""
    user32 = ctypes.windll.user32
    left = 0
    top = 0
    width = user32.GetSystemMetrics(0)  # 主屏幕宽度
    height = user32.GetSystemMetrics(1)  # 主屏幕高度

    print( left, top, width, height)

    return left, top, width, height

def click_on_main_screen(image_path, confidence=0.9, timeout=12):
    main_region = get_main_screen_info()
    start_time = time.time()

    while True:
        # 仅在主屏幕区域内搜索
        location = pyautogui.locateOnScreen(
            image_path,
            region=main_region,
            confidence=confidence
        )

        if location:

            center = pyautogui.center(location)
            x, y = center
            print( x, y)

            # 获取系统缩放比例（如 1.25 表示 125%）
            user32 = ctypes.windll.user32
            scale = user32.GetDpiForSystem() / 96.0  # 96 DPI 为默认值

            # 调整坐标（假设目标图像在主屏幕）
            adjusted_x = x * scale
            adjusted_y = y * scale
            print( adjusted_x, adjusted_y)

            # 点击
            # pyautogui.moveTo( adjusted_x, adjusted_y)
            pyautogui.click(center)
            # pyautogui.click( adjusted_x, adjusted_y)

            print(f"成功点击：{image_path}")

            return True
        else:
            if time.time() - start_time > timeout:
                print(f"超时未找到：{image_path}")
                return False
            time.sleep(0.5)

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

# 使用示例
if __name__ == "__main__":
    # 1. 定义要匹配的图片路径（确保图片与屏幕上的元素完全或部分匹配）
    image_path = r'D:\test.png'  # 替换为你的图片路径
    # click_image(image_path, confidence=0.5, timeout=15, region=( 0, 0, 200, 150))
    # click_on_main_screen( image_path, confidence=0.5)

    # print("2")

    # compare
    # pyautogui.click( 2200, 200)

    # # 获取主屏幕的物理分辨率和DPI缩放率
    # physical_width = win32api.GetSystemMetrics(0)
    # physical_height = win32api.GetSystemMetrics(1)
    # print( physical_width, physical_height)

    # dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
    # print(dpi_scale)

    # 限定搜索区域为主屏幕
    # region = (0, 0, physical_width, physical_height)
    # target_pos = pyautogui.locateOnScreen( image_path, region=region, confidence=0.5)

    # if target_pos:
    #     # 转换逻辑坐标到物理坐标
    #     center_x, center_y = pyautogui.center(target_pos)
    #     print( center_x, center_y)
    #
    #     physical_x = int(center_x * dpi_scale)
    #     physical_y = int(center_y * dpi_scale)
    #
    #     # 执行点击
    #     pyautogui.click( center_x, center_y)
    #     # pyautogui.click( physical_x, physical_y)

    success = click_on_screen(
        image_path=image_path,
        confidence=0.9,
        timeout=30
    )
    print(f"Click result: {success}")