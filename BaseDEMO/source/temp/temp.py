# import pandas as pd

# # 读取同时有行列标的Excel文件
# df = pd.read_excel('example.xlsx', header=0, index_col=0)

# # 打印结果
# print(df.head())

##########################################################################

# import pandas as pd

# # 读取表头跨越多行的Excel文件
# df = pd.read_excel('example.xlsx', header=[0, 1], index_col=0)

# # 打印结果
# print(df.head())

##########################################################################

# from openpyxl import load_workbook

# # 加载 Excel 文件
# wb = load_workbook('example.xlsx')

# # 选择活动工作表
# sheet = wb.active

# # 读取前两行作为多行表头
# columns = []
# for row in sheet.iter_rows(min_row=1, max_row=2):
#     columns.extend([cell.value for cell in row])

# # 读取数据内容
# data = []
# for row in sheet.iter_rows(min_row=3, min_col=1, max_col=sheet.max_column, max_row=sheet.max_row):
#     data.append([cell.value for cell in row])

# # 将数据整理成字典或DataFrame（可以进一步处理）
# import pandas as pd

# df = pd.DataFrame(data, columns=columns)
# print(df.head())

##########################################################################

# import pandas as pd
# from openpyxl import load_workbook

# # 使用 openpyxl 读取表头行列信息
# wb = load_workbook('example.xlsx')
# sheet = wb.active

# # 读取表头（假设表头跨越2行）
# columns = []
# for row in sheet.iter_rows(min_row=1, max_row=2):
#     columns.extend([cell.value for cell in row])

# # 读取数据内容
# data = []
# for row in sheet.iter_rows(min_row=3, min_col=1, max_col=sheet.max_column, max_row=sheet.max_row):
#     data.append([cell.value for cell in row])

# # 使用 pandas 创建 DataFrame
# df = pd.DataFrame(data, columns=columns)
# print(df.head())

##########################################################################

# from wxpy import *

# # 初始化机器人并登录
# bot = Bot(cache_path=True)

# # 获取所有好友
# friends = bot.friends()
# print("好友列表:")
# for friend in friends:
#     print(friend.nick_name, friend.user_name)

# # 获取所有群聊
# groups = bot.groups()
# print("群组列表:")
# for group in groups:
#     print(group.name, group.user_name)

##########################################################################

# import json

# # 读取 JSON 文件
# with open('config.json', 'r') as file:
#     data = json.load(file)

# print(data['name'])  # 输出 name 键的值

##########################################################################

# try:
#     # 一些可能导致异常的代码
#     num1 = int(input("请输入数字: "))
#     num2 = int(input("请输入另一个数字: "))
#     result = num1 / num2
# except Exception as e:
#     # 捕获所有异常并打印异常信息
#     print(f"发生了异常: {e}")

##########################################################################

# import os
# import pandas as pd

# # 指定目标文件夹路径
# folder_path = 'C:/path/to/your/folder'

# # 遍历文件夹中的所有文件
# for filename in os.listdir(folder_path):
#     # 检查文件是否是 Excel 格式文件
#     if filename.endswith('.xlsx'):
#         file_path = os.path.join(folder_path, filename)
        
#         # 使用 pandas 打开并读取 Excel 文件
#         try:
#             df = pd.read_excel(file_path)
#             print(f"文件 {filename} 读取成功，内容如下：")
#             print(df.head())  # 打印前几行内容
#         except Exception as e:
#             print(f"无法读取文件 {filename}: {e}")

###

# import glob
# from openpyxl import load_workbook

# # 指定目标文件夹路径
# folder_path = 'C:/path/to/your/folder'

# # 使用 glob 获取所有的 Excel 文件
# excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# for file_path in excel_files:
#     try:
#         # 使用 openpyxl 打开 Excel 文件
#         wb = load_workbook(file_path)
#         print(f"文件 {file_path} 读取成功，sheet 名称：{wb.sheetnames}")
        
#         # 假设我们读取第一个 sheet
#         sheet = wb.active
#         for row in sheet.iter_rows(values_only=True):
#             print(row)  # 打印每一行数据
        
#         wb.close()
#     except Exception as e:
#         print(f"无法读取文件 {file_path}: {e}")

#####

import pandas as pd

# def check_cell_value(file_path, row_index, column_name, condition):
#     # 读取 Excel 文件
#     df = pd.read_excel(file_path)

#     # 获取指定单元格的值
#     cell_value = df.at[row_index, column_name]
    
#     # 检查该单元格是否符合预设的条件
#     if condition(cell_value):
#         print(f"单元格 ({row_index}, {column_name}) 的值 '{cell_value}' 符合条件！")
#     else:
#         print(f"单元格 ({row_index}, {column_name}) 的值 '{cell_value}' 不符合条件。")

# # 文件路径
# file_path = 'data.xlsx'

# # 检查第2行（索引从0开始）`Age` 列是否大于 25
# check_cell_value(file_path, 1, 'Age', lambda x: x > 25)

# # 检查第3行（索引从0开始）`City` 列是否为 'Paris'
# check_cell_value(file_path, 2, 'City', lambda x: x == 'Paris')

######

# from openpyxl import load_workbook

# def check_cell_value(file_path, row_index, column_index, condition):
#     # 加载 Excel 文件
#     wb = load_workbook(file_path)
#     ws = wb.active  # 获取当前活动的工作表
    
#     # 获取指定单元格的值
#     cell_value = ws.cell(row=row_index, column=column_index).value
    
#     # 检查该单元格是否符合预设的条件
#     if condition(cell_value):
#         print(f"单元格 ({row_index}, {column_index}) 的值 '{cell_value}' 符合条件！")
#     else:
#         print(f"单元格 ({row_index}, {column_index}) 的值 '{cell_value}' 不符合条件。")

# # 文件路径
# file_path = 'data.xlsx'

# # 检查第2行（索引从1开始）第2列（'Age'）的值是否大于 25
# check_cell_value(file_path, 2, 2, lambda x: x > 25)

# # 检查第3行（索引从1开始）第3列（'City'）的值是否为 'Paris'
# check_cell_value(file_path, 3, 3, lambda x: x == 'Paris')

#=================================================================================

# import openpyxl

# # 打开 Excel 文件
# file_path = 'C:\\Users\\AdmOA\\Downloads\\test.xlsx'
# workbook = openpyxl.load_workbook(file_path)

# # 获取指定表单
# sheet = workbook['Sheet1']  # 替换为你要操作的表单名称

# # 设置指定单元格内容
# sheet.cell(row=2, column=2, value="Hello World")

# # 保存文件
# workbook.save(file_path)

# choice = input("Please enter your choice:")
# choice = choice.split()
# for i in choice:
#     print(i)

# from datetime import datetime
    
# # 获取当前系统时间
# current_time = datetime.now()

# # 将当前时间格式化为字符串
# time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

# print("当前时间:", time_str)

# LogFiles = current_time.strftime('%Y%m%d%H%M')

# print(LogFiles)

import threading
import asyncio

# # 创建事件用于控制主程序的暂停和继续
# pause_event = threading.Event()
# exit_flag = threading.Event()  # 退出标志

# # pause_event.set()
# pause_event.clear()

# # 定义一个函数来检测键盘输入
# def listen_for_input(pause_event, exit_flag):
#     while True:
#         user_input = input("\n请输入内容 (按 'q' 停止程序，按 'c' 继续程序): \n")
        
#         if user_input.lower() == 'q':
#             print("检测到 'q'，程序中断。")
#             exit_flag.set()  # 设置退出标志
#             break
        
#         elif user_input.lower() == 'c':
#             print("检测到 'c'，继续执行主程序任务。")
#             pause_event.set()  # 设置事件，解除主程序的暂停状态
            
#         elif user_input.lower() == 'p':
#             print("检测到 'p'，暂停执行主程序任务。")
#             pause_event.clear()  # 设置事件，暂停主程序的执行

#         else:
#             print("无效输入，请输入 'q' 或 'c'")

# # 主程序的任务
# async def main_task(pause_event, exit_flag):
#     i = 0
    
#     # 初始时设置 pause_event，允许主程序运行
#     pause_event.set()
    
#     while True:
#         # 如果退出标志被设置，结束主线程
#         if exit_flag.is_set():
#             print("检测到退出标志，主程序退出。")
#             break
        
#         # 主程序循环任务
#         print(f"主程序运行中... {i}")
#         i += 1
#         await asyncio.sleep(1)  # 模拟主程序正在执行某些任务

#         # 在每次循环时检查是否需要暂停
#         if not pause_event.is_set():
#             print("主程序暂停，等待输入...")
#             pause_event.wait()  # 阻塞直到事件被 set()，即恢复执行

# # 启动输入线程
# input_thread = threading.Thread(target=listen_for_input, args=(pause_event, exit_flag))
# input_thread.daemon = True  # 设置为守护线程，程序退出时自动关闭
# input_thread.start()

# # 主程序开始执行任务
# print("主程序启动，正在执行任务...\n")
# # 使用 asyncio 来运行主任务
# asyncio.run(main_task(pause_event, exit_flag))

# print("程序已退出。")

import logging
import functions.Init
import functions.LogEdit
import functions.WeChat
import functions.config
import functions.excel

####==========================================

functions.Init.init()
functions.excel.operate()

def my_func(): pass
print(my_func.__name__)  # 输出："my_func"

# 在列表中获取
func_list = [my_func]
print(func_list[0].__name__)  # 输出："my_func"

#===============================================

import asyncio
import json
import websockets

operate ={
    "OperateType": None,
    "command": None
}

response = {"status": "success"}

async def start_operate( test):
    # 格式逻辑测试
    print( test)

class WSS:
    def __init__(self, host="localhost"):
        with open(r'Z:\CPS-OA_RPA-20250312\config\default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        self.host = host
        self.port = source['port']
        self.clients = set()
        self.lock = asyncio.Lock()
        self.server = None  # 存储WebSocket服务器实例
        self.running = False  # 标志服务器是否在运行

    async def handle_client(self, websocket):
        self.clients.add(websocket)
        try:
            async for message in websocket:
                print(f"收到消息: {message}")
                data = json.loads(message)

                # 处理退出信号
                if data.get("type") == "signal" and data.get("data") == "exit":
                    print("接收到退出信号，停止服务器...")
                    await self.stop_server()
                    return  # 退出当前连接处理

                # 更新操作信息
                async with self.lock:
                    operate["OperateType"] = data.get("type")
                    operate["command"] = data.get("data")

                start_operate(data)
                # 示例响应（需根据实际逻辑修改）
                response = {"status": "success"}
                await websocket.send(json.dumps(response))

        except json.JSONDecodeError:
            error_response = {"error": "无效的JSON格式"}
            await websocket.send(json.dumps(error_response))
        except websockets.ConnectionClosed:
            print("客户端断开连接")
        finally:
            self.clients.remove(websocket)

    async def send_client(self, message):
        """广播消息给所有客户端"""
        if not self.clients:
            print("无客户端连接")
            return

        disconnected = []
        for client in self.clients:
            try:
                await client.send(message)
            except websockets.ConnectionClosed:
                disconnected.append(client)
        # 清理断开连接的客户端
        for client in disconnected:
            self.clients.remove(client)

    async def stop_server(self):
        """安全关闭服务器"""
        if self.running:
            print("停止服务器...")
            self.running = False
            if self.server:
                self.server.close()
                await self.server.wait_closed()
                print("服务器已停止")

    async def start_server(self):
        """启动服务器并保持运行"""
        print(f"启动WebSocket服务器: ws://{self.host}:{self.port}")
        async with websockets.serve(self.handle_client, self.host, self.port) as server:
            self.server = server
            self.running = True
            print("服务器已启动，等待连接...")
            while self.running:  # 保持服务器运行
                await asyncio.sleep(1)  # 避免CPU占用过高

async def main():
    server = WSS()
    try:
        await server.start_server()
    except KeyboardInterrupt:
        print("手动中断: 程序终止")
    finally:
        await server.stop_server()

if __name__ == "__main__":
    asyncio.run(main())