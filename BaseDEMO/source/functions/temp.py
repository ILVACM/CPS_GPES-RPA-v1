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

from datetime import datetime
    
# 获取当前系统时间
current_time = datetime.now()

# 将当前时间格式化为字符串
time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

print("当前时间:", time_str)

LogFiles = current_time.strftime('%Y%m%d%H%M')

print(LogFiles)