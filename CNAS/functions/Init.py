import os
import sys

import json

import socket
import random

import chardet

import logging
from datetime import datetime

import requests

import psutil

from wxauto import WeChat

#---------------------------

from functions import config

#=================================================================================

FirstInit = False

UserName = None

# port = None

DownLoad = None
URL = None

LogType = None
LogFiles_Path = None
Tasklogger = None

OperateList = []
NewFile = None
NewFile_Suffix = None

SendType = None
SendList = []

wx = None

#=================================================================================

def is_port_in_use(port_id):
    """检测端口是否被占用"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)  # 设置超时，避免长时间等待
    try:
        # 尝试绑定到指定的端口
        sock.bind(('0.0.0.0', port_id))
        sock.listen(1)  # 启动监听

        print(f"端口 {port_id} 可用！")
        config.permission_change(3, True)

        return False  # 端口未被占用

    except (socket.error, OSError) as e:
        # 如果端口绑定失败，表示端口已经被占用

        print(f"端口 {port_id} 已被占用！")
        config.permission_change(3, False)

        return True
    finally:
        sock.close()

def find_available_port():
    """在指定范围内随机选择一个端口，并检测端口是否可用"""
    port_num = random.randint(1024, 65535)
    return port_num
    # for _ in range(100):  # 尝试最多100次，避免死循环
    #     port_num = random.randint(1024, 65535)
    #     if not is_port_in_use(port_num):
    #         # print(f"端口 {port_num} 可用！")
    #         return port_num
    #     else:
    #         # print(f"端口 {port_num} 已被占用，重新选择...")
    # return None  # 如果没有找到可用端口，返回 None

def Init():
    ##### 检查操作系统环境 ########
    # print(os.name)

    if (os.name == 'nt') or (os.name == 'ce'):
        # pass
        config.permission_change(1, True)

    else:
        config.permission_change(1, False)
        sys.exit("当前操作系统不支持程序运行，请在 Windows系统 中运行本程序。\n>>>>>程序已终止<<<<<")

    ##### 关键变量全局属性 ##########
    global FirstInit
    global UserName
    # global port
    global DownLoad
    global URL
    global LogType
    global LogFiles_Path
    global Tasklogger
    global OperateList
    global NewFile
    global NewFile_Suffix
    global SendType
    global SendList
    global wx

    ##### 读取参数文件 ###########

    try:
        with open('config\default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

            # print(source)

        FirstInit = source['first init']

        UserName = source['name']
        port = source['port']

        DownLoad = source['download']
        URL = source['url']

        LogType = source['log type']
        LogFiles_Path = source['log file']

        OperateList = source['OperateList']
        NewFile = source['new file']
        NewFile_Suffix = source['suffix']

        SendType = source['send type']
        SendList = source['SendList']

        # 权限检测记录
        config.permission_change(2, True)

    except json.JSONDecodeError as e:
        print(f"JSON 解码错误: {e}")
        config.permission_change(2, False)

        sys.exit("文件解析异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except FileNotFoundError as e:
        print(f"文件未找到: {e}")
        config.permission_change(2, False)

        sys.exit("文件解析异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except Exception as e:
        print(f"发生错误: {e}")
        config.permission_change(2, False)
        sys.exit("程序发生异常，已终止。")

    #### 初始化端口判断 #####

    i = 0
    while is_port_in_use(port):
        port = find_available_port()

        i += 1
        if i > 100:
            config.permission_change(3, False)
            sys.exit("端口占用异常，请检查端口占用情况。\n>>>>>程序已终止<<<<<")

    ##### 日志配置 ###############

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
    if not os.path.exists(LogFiles_Path + "\\" + LogFiles):
        os.makedirs(LogFiles_Path + "\\" + LogFiles)  # 创建文件夹，exist_ok=True 会避免已存在的文件夹引发错误

    try:
        # 配置系统日志格式和日志级别
        logging.basicConfig(
            level=logging.DEBUG,  # 日志级别
            format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',  # 自定义日志格式
            datefmt='%Y-%m-%d %H:%M:%S',  # 日期格式
            filename=LogFiles_Path + "\\" + LogFiles + "\\" + 'record.log',
            filemode='a',
            encoding='utf-8'
        )

        # 生成日志测试信息
        logging.debug('这是一条调试记录测试')
        logging.info('这是一条记录测试')
        logging.warning('这是一条警告记录测试')
        logging.error('这是一条错误记录测试')
        logging.critical('这是一条严重错误记录测试')

        # ===========任务运行日志记录===========

        # # 创建一个新的日志记录器
        # Tasklogger = logging.getLogger('Tasks')

        # # 添加一个新的文件处理器，追加日志到新的日志文件
        # file_handler = logging.FileHandler( LogFiles_Path + "\\" + LogFiles + "\\" + 'Tasks.log', mode='a', encoding='utf-8')
        # file_handler.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s'))

        # Tasklogger.addHandler(file_handler)

        Tasklogger = open(LogFiles_Path + "\\" + LogFiles + '\\Tasks' + LogType, 'a', encoding='utf-8')

        Tasklogger.write(time_str + "\n")
        Tasklogger.flush()

        # print( time_str, file=Tasklogger)
        # print( time_str, file=Tasklogger)

        config.permission_change( 3, True)

    except Exception as e:
        config.permission_change( 3, False)
        logging.error(f"日志设置过程发生异常: {e}")

    #### 初始化操作必要性判断 #####

    if FirstInit:
        pass
    else:
        return

    ######### 文件读写测试 ###########

    try:
        # 创建文件夹（如果文件夹不存在）
        if not os.path.exists(DownLoad):
            os.makedirs(DownLoad)

        # 打开文件并写入内容
        with open(DownLoad + '\\test.txt', 'w', encoding='utf-8') as file:
            # file.write(time_str + "\n")
            file.writelines(time_str)
            # file.write("文件读写测试(File read and write I\O test)\n")
            file.writelines("文件读写测试(File read and write I\O test)\n")

            logging.info("文件写入内容测试成功")

        # 任务日志
        # with open( LogFiles_Path + "\\" + LogFiles+'\\Tasks.txt', 'w', encoding='utf-8') as file:
        # Tasklogger.write(time_str)
        # Tasklogger.write(time_str)

        with open(DownLoad + '\\test.txt', 'rb') as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']

        with open(DownLoad + '\\test.txt', 'r', encoding=encoding) as file:
            content = file.readline()
            if content is not None:
                # print("文件读写测试成功")

                # 权限检测记录
                config.permission_change(4, True)

                logging.info("文件读取测试成功")

            else:
                # print("文件读写测试失败")
                # 权限检测记录
                config.permission_change(4, False)

                logging.error("文件读取测试失败")

                sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")
            # print(content)

        # print(f"文件已成功创建并写入内容: {DownLoad}")

    except PermissionError:
        config.permission_change( 4, False)

        print(f"权限错误：无法在 {DownLoad} 创建文件。请检查文件夹权限。")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except FileNotFoundError:
        config.permission_change( 4, False)

        print(f"错误：指定的路径 {DownLoad} 无法找到。请检查路径是否正确。")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except OSError as e:
        config.permission_change( 4, False)

        print(f"操作系统错误：{e}")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    except Exception as e:
        config.permission_change( 4, False)

        print(f"发生未知错误：{e}")
        sys.exit("文件读写测试异常，请检查系统环境以及程序配置文件。\n>>>>>程序已终止<<<<<")

    ######## 网络状态检测 ########

    try:
        response = requests.get('https://www.baidu.com', timeout=5)  # 设置超时为5秒
        # response = requests.get( URL, timeout=5)  # 设置超时为5秒

        if response.status_code == 200:
            print(f"网址 {URL} 可以正常访问。")
            # pass
            config.permission_change( 5, True)

        else:
            print(f"Abnormal network status, status code: {response.status_code}")
            config.permission_change( 5, False)

    except requests.RequestException as e:
        config.permission_change( 5, False)
        print(f"发生错误: {e}")

    ##### 微信检测初始化 ########

    # 检查微信是否已经启动运行
    for process in psutil.process_iter(['pid', 'name']):
        try:
            # 如果进程名为 WeChat.exe，说明微信在运行
            if 'WeChat.exe' in process.info['name']:
                break

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            # 忽略访问权限错误、僵尸进程等
            sys.exit("程序解析异常，请检查微信是否启动并登录账号。\n>>>>>程序已终止<<<<<")

    wx = WeChat()

    try:
        if wx.nickname == UserName:
            print("微信已登录，初始化成功。")
            config.permission_change(6, True)

        else:
            config.permission_change(6, False)
            raise ValueError("程序解析异常，请检查微信是否启动并登录账号。")

    except Exception as e:
        sys.exit(f"发生异常：{e}\n>>>>>程序已终止<<<<<")

    ###### 初始化完成 ########

    FirstInit = False

    # 将修改后的数据保存回文件
    with open('config/default.json', 'w', encoding='utf-8') as file:
        json.dump(source, file, ensure_ascii=False, indent=4)  # ensure_ascii=False 用于保存中文字符

if __name__ == '__main__':
    # Init()
    is_port_in_use(443)