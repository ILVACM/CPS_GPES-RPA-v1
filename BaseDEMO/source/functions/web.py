import winreg
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# # 第二种是打开指定网页，下载excel文件到指定文件目录都是全自动化（该封装函数需两个传参）
def download_file2(url,download_dir):
    #用于获取用户电脑的默认浏览器类型
    def get_default_browser():
        try:
            # 查询默认浏览器注册表项
            reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.htm\UserChoice"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
                value, _ = winreg.QueryValueEx(key, "Progid")
                if value:
                    return value
        except FileNotFoundError:
            return "无法获取默认浏览器"
    # 获取并输出默认浏览器类型
    default_browser = get_default_browser()
    print(f"默认浏览器类型：{default_browser}")
    
    
    ##查找QQ 浏览器的可执行路径
    def check_qq_browser_path():
        # QQ 浏览器的常见安装路径
        common_paths = [
            r"C:\Program Files\QQBrowser\QQBrowser.exe",
            r"C:\Program Files (x86)\QQBrowser\QQBrowser.exe",
            r"C:\Program Files\Tencent\QQBrowser\QQBrowser.exe",
            r"C:\Program Files (x86)\Tencent\QQBrowser\QQBrowser.exe"
        ]
    
        for path in common_paths:
            if os.path.exists(path):
                print(f"QQ 浏览器的可执行路径: {path}")
                return path
        print("未找到 QQ 浏览器的安装路径")


    #获取到默认浏览器类型后进行浏览器初始化，分情况：如果是chrome；如果是edge；如果是qq浏览器
    if 'chrome' in default_browser.lower():
            # 配置 chrome 浏览器选项
            options = webdriver.ChromeOptions()
            # options.add_argument("--headless")  # 可选，启动无头模式
            # （无头模式：无头模式（Headless mode）是指在没有图形界面（GUI）的情况下运行浏览器。在无头模式下，浏览器会以后台的形式运行，而不会显示界面。它常用于自动化测试、抓取网页和其他无需用户交互的场景。）
            # Edge 浏览器默认就是有头模式。如果不添加 --headless 参数，浏览器界面会正常显示。
            options.add_experimental_option("prefs", {
                "download.default_directory": download_dir,  # 设置下载目录
                "download.prompt_for_download": False,  # 不弹出下载对话框
                "download.directory_upgrade": True,  # 允许文件路径的升级
                "savefile.default_directory": download_dir  # 设置文件保存路径
            })
            
            # 使用 webdriver-manager 自动管理驱动
            driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager().install()), options=options)
            time.sleep(1)

    elif 'edge' in default_browser.lower():
        # 配置 edge 浏览器选项
        options = webdriver.EdgeOptions()
        # options.add_argument("--headless")  # 可选，启动无头模式
        # （无头模式：无头模式（Headless mode）是指在没有图形界面（GUI）的情况下运行浏览器。在无头模式下，浏览器会以后台的形式运行，而不会显示界面。它常用于自动化测试、抓取网页和其他无需用户交互的场景。）
        # Edge 浏览器默认就是有头模式。如果不添加 --headless 参数，浏览器界面会正常显示。
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,  # 设置下载目录
            "download.prompt_for_download": False,  # 不弹出下载对话框
            "download.directory_upgrade": True,  # 允许文件路径的升级
            "savefile.default_directory": download_dir  # 设置文件保存路径
        })

        # 使用 webdriver-manager 自动管理驱动
        driver = webdriver.Edge(service=webdriver.EdgeService(EdgeChromiumDriverManager().install()), options=options)
        time.sleep(1)
    elif 'qqbrowser' in default_browser.lower():
         # 配置 qq 浏览器选项
        options = webdriver.ChromeOptions()
        # options.add_argument("--headless")  # 可选，启动无头模式
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,  # 设置下载目录
            "download.prompt_for_download": False,  # 不弹出下载对话框
            "download.directory_upgrade": True,  # 允许文件路径的升级
            "savefile.default_directory": download_dir  # 设置文件保存路径
        })
        qq_browser_path=check_qq_browser_path()
        options.binary_location = qq_browser_path  # 设置为 QQ 浏览器的路径
        # 使用 webdriver-manager 自动管理驱动
        driver = webdriver.Chrome(service=webdriver.ChromeService(ChromeDriverManager(driver_version = "116.0.5845.97" ).install()), options=options)
        time.sleep(1)
    else:
        print("未知浏览器类型")
        
    
    driver.maximize_window()#最大化页面
    time.sleep(1)
    driver.get(url) #调用driver对象的get方法，可以让浏览器打开指定网址（试试），原设网址为'https://www.gdut.edu.cn/info/1715/21966.htm'
    time.sleep(1)

    
    # 尝试的下载链接的CSS选择器和部分匹配链接文本
    selectors = [
        'a[href$=".docx"]',  # 匹配以 ".docx" 结尾的链接
        'a[href*="xlsx"]',  # 匹配包含 "xlsx" 的链接
        'a[href*="xls"]',  # 匹配包含 "xls" 的链接
        'a[href*="csv"]',  # 匹配包含 "csv" 的链接
        'a[href*="doc"]',  # 匹配包含 "doc" 的链接
        'a[href*="docx"]',  # 匹配包含 "docx" 的链接
    ]
    link_texts = [
        "xls",  # 匹配包含 "xls" 的链接文本
        "xlsx",  # 匹配包含 "xlsx" 的链接文本
        "csv",  # 匹配包含 "csv" 的链接文本
        "doc",  # 匹配包含 "doc" 的链接文本
        "docx",  # 匹配包含 "docx" 的链接文本
    ]

    # 尝试通过CSS选择器查找并点击文件下载链接
    for selector in selectors:
        try:
            # 等待并点击第一个可用的下载链接（使用CSS选择器）
            docx_button = WebDriverWait(driver, 10).until(
             EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
            )
            docx_button.click()
            print(f"开始下载文件... 使用CSS选择器匹配: {selector}")
            break  # 找到并点击后就退出循环
        except (NoSuchElementException, TimeoutException):
            continue  # 如果当前选择器未找到，继续尝试下一个

    # 如果没有通过CSS选择器找到文件链接，尝试通过部分链接文本查找
    else:
        for text in link_texts:
            try:
                # 等待并点击第一个可用的下载链接（使用部分链接文本）
                docx_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, text))
                )
                docx_button.click()
                print(f"开始下载文件... 匹配到链接文本: {text}")
                break  # 找到并点击后就退出循环
            except (NoSuchElementException, TimeoutException):
                continue  # 如果当前链接文本未找到，继续尝试下一个
        else:
            # 如果所有选择器和部分链接文本都未找到
            print("下载失败")
            driver.quit()

        
    #等待文件下载完成
    time.sleep(5)
    # 检查下载的文件是否存在
    files = [f for f in os.listdir(download_dir) if os.path.isfile(os.path.join(download_dir, f))]# 获取目录中的所有文件和文件夹
    full_paths = [os.path.join(download_dir, f) for f in files]# 获取文件的完整路径
    latest_file_path = max(full_paths, key=os.path.getmtime)# 修改时间并按修改时间排序，降序排列（最新的文件在前），获取最新下载文件的完整路径
    latest_file = os.path.basename(latest_file_path)#获取最新文件的文件名
    if os.path.getsize(latest_file_path) > 0:  # 文件大小大于 0 字节表示文件下载完成
        print(f"文件 {latest_file} 下载成功。")
        driver.quit()  # 关闭浏览器
        return True  # 返回 True，表示文件下载成功
    else:
        print("下载失败")
        return False


# url = 'https://tjj.gz.gov.cn/datav/admin/home/www_nj/2023/directory/19/html/19-02.htm'  # 要访问的网页
# url = 'https://www.stats.gov.cn/sj/tjzd/202305/t20230523_1939969.html'# 要访问的网页
# url = 'https://www.gdut.edu.cn/info/1715/21966.htm'# 要访问的网页
# url = 'https://www.zgtjnj.org/navipage-n3024122306000449.html'# 要访问的网页
# url = 'https://www.stats.gov.cn/sj/tjbz/gjtjbz/202302/t20230213_1902784.html'# 要访问的网页
# url = 'https://www.stats.gov.cn/sj/tjbz/gjtjbz/202307/t20230728_1941612.html'# 要访问的网页


# download_dir = r"C:\WorkArea\webfile"  # 下载文件的目标目录
# download_successful = download_file2(url, download_dir)
# if download_successful:
#     print("验证所下载的文件存在。")
# else:
#     print("文件下载失败。") 

if __name__ == "__main__":
    url = 'https://www.gdut.edu.cn/info/1715/21966.htm'# 要访问的网页
    download_dir = r"C:\Users\AdmOA\Downloads"  # 下载文件的目标目录
    download_successful = download_file2( url, download_dir)
