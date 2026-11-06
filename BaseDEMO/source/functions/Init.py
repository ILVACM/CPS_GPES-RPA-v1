import os
import sys
import json
import psutil
import wxauto
import chardet
import requests
import logging
from datetime import datetime

#=================================================================================

Name        = None
Download    = None
URL         = None
OperateList = []
SendList    = []

wx = None

friends = []
friends_num = 0

LogFiles_Path = None
LogFiles = None

AutoSend = None
SendType = None

Tasklogger = None

#=================================================================================

def add_main():
    # 将 main_folder 的路径添加到 sys.path 中
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), 'Z:\CPS-Auto\py')))

def init():
    
    ##### 检查操作系统环境 ########
    # print(os.name)
    
    if (os.name == 'nt') or (os.name == 'ce'):
        pass
    else:
        sys.exit("当前操作系统不支持程序运行，请在 Windows系统 中运行本程序。\n>>>>>程序已终止<<<<<")
        
    ##### 关键变量全局属性 ##########
    
    global Name, Download, URL, OperateList, SendList, friends, friends_num, wx, LogFiles, LogFiles_Path, AutoSend, SendType, Tasklogger
        
    ##### 读取参数文件 ###########
    
    try:
        with open('config\default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)
            
            # print(data)
            
            Name = source['name']
            
            URL = source['url']
            Download = source['download']
            
            LogFiles_Path = source['log']
            OperateList = source['OperateList']       
            
            AutoSend = source['autosend']
            SendType = source['sendtype']
            SendList = source['SendList']    
            
            # print(Name)
            # print(URL)
            # print(Download)
            # print(LogFiles_Path)
            # print(OperateList)
            # print(SendList)
    
    except json.JSONDecodeError as e:
        print(f"JSON 解码错误: {e}")
        sys.exit("文件解析异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except FileNotFoundError as e:
        print(f"文件未找到: {e}")
        sys.exit("文件解析异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
        
    ##### 日志配置 ########
    
    # 获取当前系统时间
    current_time = datetime.now()

    # 将当前时间格式化为字符串
    time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

    print("当前时间:", time_str)
    
    LogFiles = current_time.strftime('%Y%m%d%H%M')
    
    # # 配置日志格式和日志级别
    # logging.basicConfig(
    #     level=logging.DEBUG,  # 日志级别
    #     format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',  # 自定义日志格式
    #     datefmt='%Y-%m-%d %H:%M:%S',  # 日期格式
    # )
    
    # 配置日志格式和日志级别，日志输出到文件
    # logging.basicConfig(
    #     level=logging.DEBUG,
    #     format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',
    #     datefmt='%Y-%m-%d %H:%M:%S',
    #     filename='app.log',  # 输出到文件
    #     filemode='w'  # 文件模式，'w'表示每次启动时覆盖，'a'表示追加
    # )
    
    # 检查文件夹是否存在，如果不存在则创建
    if not os.path.exists( LogFiles_Path + "\\" + LogFiles):
        os.makedirs( LogFiles_Path + "\\" + LogFiles)  # 创建文件夹，exist_ok=True 会避免已存在的文件夹引发错误
    
    try:
        # 配置日志格式和日志级别
        logging.basicConfig(
            level=logging.DEBUG,  # 日志级别
            format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',  # 自定义日志格式
            datefmt='%Y-%m-%d %H:%M:%S',  # 日期格式
            filename= LogFiles_Path + "\\" + LogFiles + "\\" + 'record.log',
            filemode='a',
            encoding='utf-8'
        )
        
        # 生成日志信息
        logging.debug('这是一条调试记录测试')
        logging.info('这是一条记录测试')
        logging.warning('这是一条警告记录测试')
        logging.error('这是一条错误记录测试')
        logging.critical('这是一条严重错误记录测试')
        
        #===========任务运行记录
        
        # # 创建一个新的日志记录器
        # Tasklogger = logging.getLogger('Tasks')
        
        # # 添加一个新的文件处理器，追加日志到新的日志文件
        # file_handler = logging.FileHandler( LogFiles_Path + "\\" + LogFiles + "\\" + 'Tasks.log', mode='a', encoding='utf-8')
        # file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s'))
        
        # Tasklogger.addHandler(file_handler)
        
        Tasklogger = open( LogFiles_Path + "\\" + LogFiles+'\\Tasks.txt', 'a', encoding='utf-8')
        
        Tasklogger.write( time_str + "\n")
        Tasklogger.flush()
        
        # print( time_str, file=Tasklogger)
        # print( time_str, file=Tasklogger)
        
    except Exception as e:
        logging.error(f"日志设置过程发生异常: {e}")
    
    ##### 文件读写测试 ###########
    
    try:
        # 创建文件夹（如果文件夹不存在）
        if not os.path.exists(Download):
            os.makedirs(Download)

        # 打开文件并写入内容
        with open( Download+'\\test.txt', 'w', encoding='utf-8') as file:
            file.write("文件读写测试(File read and write I\O test)")
        
        # 任务日志
        # with open( LogFiles_Path + "\\" + LogFiles+'\\Tasks.txt', 'w', encoding='utf-8') as file:
        # Tasklogger.write(time_str)
        # Tasklogger.write(time_str)
            
        with open( Download+'\\test.txt', 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            
        with open( Download+'\\test.txt', 'r', encoding=encoding) as file:
            content = file.read()
            # print(content)

        # print(f"文件已成功创建并写入内容: {Download}")

    except PermissionError:
        print(f"权限错误：无法在 {Download} 创建文件。请检查文件夹权限。")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
    
    except FileNotFoundError:
        print(f"错误：指定的路径 {Download} 无法找到。请检查路径是否正确。")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
    
    except OSError as e:
        print(f"操作系统错误：{e}")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
    
    except Exception as e:
        print(f"发生未知错误：{e}")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
    
    ##### 网络状态检测 ########
    
    try:
        response = requests.get( 'https://www.baidu.com', timeout=5)  # 设置超时为5秒
        # response = requests.get( URL, timeout=5)  # 设置超时为5秒
        
        if response.status_code == 200:
            # print(f"网址 {URL} 可以正常访问。")
            pass
        
        else:
            print(f"Abnormal network status, status code: {response.status_code}")
    
    except requests.RequestException as e:
        
        print(f"发生错误: {e}")
    
    ##### 微信检测初始化 ########
    
    for process in psutil.process_iter(['pid', 'name']):
        try:
            # 如果进程名为 WeChat.exe，说明微信在运行
            if 'WeChat.exe' in process.info['name']:
                break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 忽略访问权限错误、僵尸进程等
            sys.exit("程序解析异常，请检查微信是否启动并登录账号。\n>>>>>程序已终止<<<<<")
            
    wx = wxauto.WeChat()
    
    if wx.nickname == None:
        sys.exit("程序解析异常，请检查微信是否启动并登录账号。\n>>>>>程序已终止<<<<<")
        
    # print("微信已登录，初始化成功。")

    friends = wx.GetAllFriends()
    # print(friends)
    
    friends_num = len(friends)
            
if __name__ == '__main__':
    init()
    