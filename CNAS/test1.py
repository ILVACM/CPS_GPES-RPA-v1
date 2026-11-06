import pandas as pd
from openpyxl import load_workbook
import json
import time
import pyautogui

# -------------------------------------

from functions import Init
from functions import WindowsUI_OPSimulation

from pptx import Presentation

from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32gui
import win32con
import keyboard

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 数据提取工具")
        self.setup_ui()
        self.loop_running = True
        keyboard.add_hotkey('esc', self.exit_program)

    def exit_program(self):
        print("\n用户按下 ESC，程序终止！")
        self.loop_running = False  # 通知主线程退出

    def setup_ui(self):
        """ 界面布局 """
        # 文件选择部分
        ttk.Label(self.root, text="Excel文件路径：").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.file_entry = ttk.Entry(self.root, width=40)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.root, text="浏览", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        # 行号输入部分
        ttk.Label(self.root, text="行号（N）：").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.n_entry = ttk.Entry(self.root, width=10)
        self.n_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 执行按钮
        ttk.Button(self.root, text="提取数据", command=self.execute).grid(row=2, column=0, pady=10)
        ttk.Button(self.root, text="开始运行", command=self.web_operation).grid(row=2, column=1, pady=10)

        # 添加勾选框（是否保存）
        self.save_var = tk.BooleanVar(value=False)  # 默认不勾选
        save_check = ttk.Checkbutton(
            self.root,
            text="是否保存",
            variable=self.save_var,
            onvalue=True,  # 勾选时的值
            offvalue=False  # 未勾选时的值
        )
        save_check.grid(row=2, column=2, pady=10, sticky="w")

        # 结果显示
        self.result_text = tk.Text(self.root, height=20, width=100, state="disabled")
        self.result_text.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

    def get_excel_row(self, file_path, n):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', header=1)
            if n < 1 or n > len(df):
                raise ValueError(f"有效范围为 1 到 {len(df)}")
            return df.iloc[n - 1].to_dict()
        except Exception as e:
            raise

    def browse_file(self):
        """ 文件选择对话框 """
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def execute(self):
        """ 执行数据提取 """
        file_path = self.file_entry.get()
        n_str = self.n_entry.get()

        # 输入验证
        if not file_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        if not n_str.isdigit():
            messagebox.showerror("错误", "行号必须为正整数")
            return

        n = int(n_str)

        try:
            # 调用核心函数
            result = self.get_excel_row(file_path, n)
            self.show_result(result)
            return result, n
        except Exception as e:
            messagebox.showerror("错误", str(e))

    def show_result(self, data):
        """ 显示结果 """
        self.result_text.config(state="normal")
        self.result_text.delete(1.0, tk.END)

        # 格式化输出
        formatted = "\n".join([f"{k}: {v}" for k, v in data.items()])
        self.result_text.insert(tk.END, f"第 {self.n_entry.get()} 行数据：\n{formatted}")
        self.result_text.config(state="disabled")

    def get_tar_index(self,tar,data):
        try:
            index = data.index(tar)  # 获取索引
            return index + 1  # 索引为1、2、3...
        except ValueError:  # 如果不存在返回0
            return 0

    def web_operation(self):
        data, total = self.execute()
        self.loop_running = True

        # ============================================================================

        WindowsUI_OPSimulation.set_primary_dpi_to_100()

        # ============================================================================

        # 启动脚本前将鼠标移至安全区域
        pyautogui.moveTo(100, 100)  # 移动到屏幕非边缘位置
        pyautogui.FAILSAFE = False  # 禁用安全保护
        time.sleep(1.5)

        # ============================================================================
        Init.Init()

        with open(r'config\temp_4_IMG.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        OPList = source['OperateList']

        target_window = WindowsUI_OPSimulation.find_activate_window("监理一体化")
        if target_window:
            print(f"窗口标题: {target_window.title}")
            print(f"窗口位置: {target_window.left}, {target_window.top}")
            print(f"窗口大小: {target_window.width}x{target_window.height}")

            # WindowsUI_OPSimulation.force_maximize_and_foreground(target_window)  # 最大化并置顶
            status = WindowsUI_OPSimulation.move_window_to_primary_screen(target_window)
            print(status)

            print(f"total: {total}")
            print(f"save: {self.save_var.get()}")
            for i in range(1,total+1):
                if not self.loop_running:
                    print("用户主动终止")
                    break
                print(f"i: {i}")
                data = self.get_excel_row(self.file_entry.get(), i)
                print(data["姓名"])
                for image in OPList:
                    if not self.loop_running:
                        print("用户主动终止")
                        break
                    if image['code'] == '1' and i == 1:
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=10)
                        print(f"Click result: {success}")
                        time.sleep(1.5)
                    if image['code'] == '2':
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=10)
                        print(f"Click result: {success}")
                        time.sleep(1.5)
                    if image['code'] == '3':
                        WindowsUI_OPSimulation.paste_text(data[image['type']])
                        time.sleep(1.5)
                        WindowsUI_OPSimulation.click_offset(120, 0)
                        WindowsUI_OPSimulation.click_offset(-35, 0)
                        WindowsUI_OPSimulation.click_offset(0, 60)
                    if image['code'] == '4':
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=2)
                        print(f"Click result: {success}")
                        time.sleep(0.2)
                        if success:
                            pyautogui.press("tab")
                            time.sleep(0.2)
                            WindowsUI_OPSimulation.paste_text(data[image['type']])
                            time.sleep(0.2)
                    if image['code'] == '8':
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=2)
                        print(f"Click result: {success}")
                        time.sleep(0.2)
                        if success:
                            pyautogui.press("tab")
                            time.sleep(0.2)
                            pyautogui.hotkey("ctrl", "a")
                            time.sleep(0.2)
                            pyautogui.press("delete")
                            time.sleep(0.2)
                            WindowsUI_OPSimulation.paste_text(data[image['type']])
                            time.sleep(0.2)
                    if image['code'] == '5':
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=2)
                        print(f"Click result: {success}")
                        time.sleep(0.2)
                        if success:
                            pyautogui.press("tab")
                            time.sleep(0.2)
                            index = self.get_tar_index(data[image['type']], image['list'])
                            for keystrokes in range(index+1):
                                pyautogui.press("down")
                                time.sleep(0.2)
                            pyautogui.press("enter")
                            time.sleep(0.2)
                    if image['code'] == '6':
                        success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.9, timeout=2)
                        print(f"Click result: {success}")
                        time.sleep(0.2)
                        if success:
                            pyautogui.press("tab")
                            time.sleep(0.2)
                            if "全日制" in data[image['major']]:
                                if "本科" in data[image['educationalbackground']]:
                                    print("本科 = 学士")
                                    index = self.get_tar_index("学士", image['list'])
                                if "硕士研究生" in data[image['educationalbackground']]:
                                    print("硕士研究生 = 硕士")
                                    index = self.get_tar_index("硕士", image['list'])
                                if "博士研究生" in data[image['educationalbackground']]:
                                    print("博士研究生 = 博士")
                                    index = self.get_tar_index("博士", image['list'])
                            else:
                                print("学历不包含'全日制'")
                                index = self.get_tar_index("无", image['list'])
                            for keystrokes in range(index+1):
                                pyautogui.press("down")
                                time.sleep(0.2)
                            pyautogui.press("enter")
                            time.sleep(0.2)
                    if image['code'] == '7' and self.save_var.get():
                        #success = WindowsUI_OPSimulation.click_on_screen(image_path=image['path'], confidence=0.8, timeout=10)
                        print(f"点击保存！")
                        time.sleep(1.5)
            else:
                print("所有数据处理完成")
                messagebox.showinfo("提示", f"共导入 {total} 份信息！")
        else:
            print("未找到目标窗口！")

        # 取消窗口置顶
        win32gui.SetWindowPos(
            target_window._hWnd,  # 从pygetwindow对象获取原生句柄
            win32con.HWND_NOTOPMOST,  # 取消置顶标志
            0, 0, 0, 0,  # 位置尺寸参数（这里不使用）
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE  # 保持当前位置和大小
        )
        if not self.loop_running:
            print("用户主动终止")
            messagebox.showinfo("提示", f"用户终止运行！")

if __name__ == '__main__':
    root = tk.Tk()
    root.geometry("550x950+10+10")
    app = ExcelReaderApp(root)
    root.mainloop()