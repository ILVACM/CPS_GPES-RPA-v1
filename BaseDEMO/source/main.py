import os
import time
# import threading
# import asyncio
import keyboard
import glob
import pandas as pd
import openpyxl as xl

import functions.Init
import functions.excel
import functions.LogEdit
import functions.config
import functions.WeChat
import functions.web

# import threading
# import time

# # 定义一个函数来检测键盘输入
# def listen_for_input(pause_event, exit_flag):
#     while True:
        
#         user_input = input("\nPleace Input ( 'q' -> stop/ 'c' -> continue): \n")
        
#         if user_input.lower() == 'q':
#             print("检测到 'q'，程序中断。")
#             exit_flag.set()  # 设置退出标志
#             break
        
#         elif user_input.lower() == 'c':
#             print("检测到 'c'，继续执行主程序任务。")
#             pause_event.set()  # 设置事件，解除主程序的暂停状态

#         else:
#             print("无效输入，请输入 'q' 或 'c'")

Total_OP = 0
Completed_OP = 0

# # 定义一个函数来检测键盘输入
# def listen_for_input( pause_event, exit_flag):
    # while True:
    #     user_input = input("\nPleace Input ( 'q' -> stop/ 'c' -> continue/ 'p' -> pause): \n")
        
    #     if user_input.lower() == 'q':
    #         print(" 'q' detected, program task interrupt. ")
    #         exit_flag.set()  # 设置退出标志
    #         break
        
    #     elif user_input.lower() == 'c':
    #         print(" 'c' detected, program task continues.")
    #         pause_event.set()  # 设置事件，解除主程序的暂停状态
            
    #     elif user_input.lower() == 'p':
    #         print(" 'p' detected, program task pause.")
    #         pause_event.clear()  # 设置事件，暂停主程序的执行

    #     else:
    #         print("无效输入，请输入 'q' 或 'c'")

# def main_task( pause_event, exit_flag):
def main_task():
    
    Total_OP = 0
    Completed_OP = 0
    
    sign = True
    
    # 获取所有 Excel 文件路径
    ExcelFiles = glob.glob(os.path.join( functions.Init.Download, '*.xlsx'))
    for excel in ExcelFiles:
        try:
            # openpyxl库 读取 Excel 文件
            wb = xl.load_workbook(excel)
            Sheets = wb.sheetnames
            
            Total_OP += len(Sheets) * len(functions.Init.OperateList)
            # pandas库 读取 Excel 文件
            # Data = pd.ExcelFile(excel)
            # Sheets = Data.sheet_names
        
        except Exception as e:
            print(f"无法读取文件 {excel}:{e}")
    
    # # 初始时设置 pause_event，允许主程序运行
    # pause_event.set()
    
    for excel in ExcelFiles:
        # print(excel)
        if not sign: break
        
        # # 如果退出标志被设置，结束主线程
        # if exit_flag.is_set():
        #     print("检测到退出标志，主程序退出。")
        #     break
        
        # # 在每次循环时检查是否需要暂停
        # if not pause_event.is_set():
        #     print("主程序暂停，等待输入...")
        #     pause_event.wait()  # 阻塞直到事件被 set()，即恢复执行

        
        try:
            # openpyxl库 读取 Excel 文件
            wb = xl.load_workbook(excel)
            Sheets = wb.sheetnames
            
            # pandas库 读取 Excel 文件
            # Data = pd.ExcelFile(excel)
            # Sheets = Data.sheet_names
        
        except Exception as e:
            print(f"无法读取文件 {excel}:{e}")
            
        for sheet in Sheets:
            
            if not sign: break
            
            # # 如果退出标志被设置，结束主线程
            # if exit_flag.is_set():
            #     print("检测到退出标志，主程序退出。")
            #     break
            
            # # 在每次循环时检查是否需要暂停
            # if not pause_event.is_set():
            #     print("主程序暂停，等待输入...")
            #     pause_event.wait()  # 阻塞直到事件被 set()，即恢复执行
            
            # # pandas库 读取 Excel 文件
            # data = pd.read_excel( excel, sheet_name=sheet)
                
            # openpyxl库 读取 Excel 文件指定表单数据
            data = wb[sheet]
            
            # test
            # data.insert_rows(1)
            # data.insert_cols(1)
            # cell_value_c3 = sheet.cell(row=3, column=3).value
                            
            for OperateCode in functions.Init.OperateList:
                
                # # 如果退出标志被设置，结束主线程
                # if exit_flag.is_set():
                #     print("检测到退出标志，主程序退出。")
                #     break
                
                # # 在每次循环时检查是否需要暂停
                # if not pause_event.is_set():
                #     print("主程序暂停，等待输入...")
                #     pause_event.wait()  # 阻塞直到事件被 set()，即恢复执行
                
                Completed_OP += 1
                
                
                if keyboard.is_pressed('\n'):
                    print("1. continue \n2. stop\n")
                    choice = input("")
                    if choice == '2':
                        # print("continue")
                        # pause_event.set()  # 设置事件，解除主程序的暂停状态
                        sign = False
                        break
                        
                    # elif keyboard.is_pressed('2'):
                        # print("stop")
                        # exit_flag.set()  # 设置退出标志
                        # break
                
                if OperateCode['mode'] == '1':
                    functions.excel.add( OperateCode, data)
                
                elif OperateCode['mode'] == '2':
                    functions.excel.search( OperateCode, data)
                
                elif OperateCode['mode'] == '3':
                    functions.excel.change( OperateCode, data)
                
                elif OperateCode['mode'] == '4':
                    functions.excel.delete( OperateCode, data)
                    
                else:
                    pass
            #
            
        wb.save(excel)
    
    functions.LogEdit.TaskLog( Task_num, "Run", Total_OP, Completed_OP)
    functions.LogEdit.Record( f"Task {Task_num} Completed.", 1)
    # i = 0
    
    # while True:
    #     # 如果退出标志被设置，结束主线程
    #     if exit_flag.is_set():
    #         print("检测到退出标志，主程序退出。")
    #         break
        
    #     # 主程序循环任务
    #     print(f"主程序运行中... {i}")
    #     i += 1
    #     await asyncio.sleep(1)  # 模拟主程序正在执行某些任务

    #     # 在每次循环时检查是否需要暂停
    #     if not pause_event.is_set():
    #         print("主程序暂停，等待输入...")
    #         pause_event.wait()  # 阻塞直到事件被 set()，即恢复执行


if __name__ == '__main__':
    functions.Init.init()
    
    # functions.excel.operate()
    
    # functions.WeChat.Auto_send()
    
    ### 创建事件用于控制主程序的暂停和继续 ###
    # pause_event = threading.Event()
    # exit_flag = threading.Event()  # 退出标志
    
    Task_num = 0

    while True:
        os.system("cls") # clear screen
        print(" 1. config\n 2. Download only\n 3. Run only\n 4. Download And Run\n 5. fresh config\n 0. Exit\n")
        
        choice = input()
        if choice == "1":
            # config
            functions.config.set()
            
        elif choice == "2":
            # download
            Task_num += 1
            
            if functions.web.download_file2( functions.Init.URL, functions.Init.Download):
                functions.LogEdit.Record( "Files Download Successful", 1)
                functions.LogEdit.TaskLog( Task_num, "Download", 1, 1)
            
            else:
                functions.LogEdit.Record( "Files Download Failed", 2)
                functions.LogEdit.TaskLog( Task_num, "Download", 1, 0)
            
            # Send Task log
            functions.WeChat.Auto_send()
            
            # pass
            
        elif choice == "3":
            # run
            Task_num += 1
            
            # # 启动输入线程
            # input_thread = threading.Thread(target=listen_for_input, args=(pause_event, exit_flag))
            # input_thread.daemon = True  # 设置为守护线程，程序退出时自动关闭
            # input_thread.start()

            # pause_event.clear()
            
            # asyncio.run(main_task(pause_event, exit_flag))
            
            # exit_flag.set()  # 手动触发输入线程退出
            
            main_task()
            
            # Send Task log
            functions.WeChat.Auto_send()
            
            # pass
            
        # elif choice == "4":
        #     # download and run
        #     Task_num += 1
            
        #     if functions.web.download_file2( functions.Init.URL, functions.Init.Download):
        #         functions.LogEdit.Record( "Files Download Successful", 1)
                
        #     else:
        #         functions.LogEdit.Record( "Files Download Failed, Task stop", 2)
                
        #     # Send Task log
        #     functions.WeChat.Auto_send()
            
        #     # pass
        
        elif choice == "5":
            # fresh config
            functions.Init.init()
            
        elif choice== "0":
            functions.LogEdit.Record( "Exit", 1)
            print("Exit")
            break
        
        else:
            print("Invalid Choice.Input Again.")
            time.sleep(1)