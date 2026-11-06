from openpyxl import load_workbook
import os
import json
import time
import pyautogui
import datetime

# -------------------------------------

from pptx import Presentation

import os
from datetime import datetime


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import keyboard
import pyperclip
from openpyxl import Workbook

class WebTableToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("网页表格抓取工具")


        # 初始化配置
        self.loop_running = True
        keyboard.add_hotkey('esc', self.exit_program)



        self.config_path = "config/coordinates.json"
        self.current_company = ""  # 存储当前公司名
        self.debug_mode = False
        self.default_config = {
            "enterprise_coords": {"x": 100, "y": 200},
            "next_step_coords": {"x": 300, "y": 400},
            "table_coords": {"x": 500, "y": 950},  # 表格坐标
            "default_save_path": os.path.expanduser("~/Desktop"),
            "loading_time": 3  # 默认加载时间
        }

        # 加载或创建配置文件
        self.load_config()
        self.setup_ui()
    def load_config(self):
        """加载或创建配置文件"""
        try:
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = self.default_config
                self.save_config()
        except Exception as e:
            messagebox.showerror("配置错误", f"加载配置文件失败: {str(e)}")
            self.config = self.default_config

    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            messagebox.showerror("配置错误", f"保存配置文件失败: {str(e)}")

    def setup_ui(self):
        """设置用户界面（美化版）"""
        # 设置全局字体
        default_font = ('Microsoft YaHei', 10)
        title_font = ('Microsoft YaHei', 10, 'bold')
        self.root.option_add('*Font', default_font)

        # 主框架
        main_frame = tk.Frame(self.root, padx=20, pady=15)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # ==================== 使用说明区域 ====================
        instruction_frame = tk.LabelFrame(main_frame, text="使用说明", font=title_font, padx=10, pady=10)
        instruction_frame.pack(fill=tk.X, pady=(0, 15))

        instructions = [
            "1. 设置坐标参数",
            "2. 选择Excel保存路径",
            "3. 设置需要抓取的页数",
            "4. 点击开始作业按钮",
            "5. 按ESC键可随时终止程序运行",
            "",
            f"版本号: 20250514"
        ]

        for i, text in enumerate(instructions):
            tk.Label(instruction_frame,
                     text=text,
                     anchor='w',
                     justify=tk.LEFT).pack(fill=tk.X, pady=2)

        # ==================== 坐标设置区域 ====================
        coord_frame = tk.Frame(main_frame)
        coord_frame.pack(fill=tk.X, pady=5)

        # 第一组坐标：企业名坐标
        enterprise_frame = tk.LabelFrame(coord_frame, text="企业名坐标", font=title_font, padx=10, pady=10)
        enterprise_frame.pack(fill=tk.X, pady=5, padx=5)

        tk.Label(enterprise_frame, text="X坐标:").grid(row=0, column=0, sticky='e', padx=5)
        self.enterprise_x = tk.Entry(enterprise_frame, width=8)
        self.enterprise_x.grid(row=0, column=1, sticky='w')
        self.enterprise_x.insert(0, str(self.config["enterprise_coords"]["x"]))

        tk.Label(enterprise_frame, text="Y坐标:").grid(row=0, column=2, sticky='e', padx=5)
        self.enterprise_y = tk.Entry(enterprise_frame, width=8)
        self.enterprise_y.grid(row=0, column=3, sticky='w')
        self.enterprise_y.insert(0, str(self.config["enterprise_coords"]["y"]))

        # 第二组坐标：下一步坐标
        next_step_frame = tk.LabelFrame(coord_frame, text="下一步坐标", font=title_font, padx=10, pady=10)
        next_step_frame.pack(fill=tk.X, pady=5, padx=5)

        tk.Label(next_step_frame, text="X坐标:").grid(row=0, column=0, sticky='e', padx=5)
        self.next_step_x = tk.Entry(next_step_frame, width=8)
        self.next_step_x.grid(row=0, column=1, sticky='w')
        self.next_step_x.insert(0, str(self.config["next_step_coords"]["x"]))

        tk.Label(next_step_frame, text="Y坐标:").grid(row=0, column=2, sticky='e', padx=5)
        self.next_step_y = tk.Entry(next_step_frame, width=8)
        self.next_step_y.grid(row=0, column=3, sticky='w')
        self.next_step_y.insert(0, str(self.config["next_step_coords"]["y"]))

        # 第三组坐标：表格坐标
        table_frame = tk.LabelFrame(coord_frame, text="表格坐标", font=title_font, padx=10, pady=10)
        table_frame.pack(fill=tk.X, pady=5, padx=5)

        tk.Label(table_frame, text="X坐标:").grid(row=0, column=0, sticky='e', padx=5)
        self.table_x = tk.Entry(table_frame, width=8)
        self.table_x.grid(row=0, column=1, sticky='w')
        self.table_x.insert(0, str(self.config["table_coords"]["x"]))

        tk.Label(table_frame, text="Y坐标:").grid(row=0, column=2, sticky='e', padx=5)
        self.table_y = tk.Entry(table_frame, width=8)
        self.table_y.grid(row=0, column=3, sticky='w')
        self.table_y.insert(0, str(self.config["table_coords"]["y"]))
        # ==================== 保存设置区域 ====================
        setting_frame = tk.Frame(main_frame)
        setting_frame.pack(fill=tk.X, pady=10)

        # 文件保存路径
        path_frame = tk.LabelFrame(setting_frame, text="文件保存设置", font=title_font, padx=10, pady=10)
        path_frame.pack(fill=tk.X, pady=5, padx=5)

        tk.Label(path_frame, text="保存路径:").grid(row=0, column=0, sticky='e', padx=5)
        self.save_path_var = tk.StringVar(value=self.config["default_save_path"])
        path_entry = tk.Entry(path_frame, textvariable=self.save_path_var, state='readonly')
        path_entry.grid(row=0, column=1, sticky='ew', padx=5)

        browse_btn = tk.Button(path_frame, text="浏览...", command=self.select_save_path)
        browse_btn.grid(row=0, column=2, padx=5)

        # 抓取设置
        page_frame = tk.LabelFrame(setting_frame, text="抓取设置", font=title_font, padx=10, pady=10)
        page_frame.pack(fill=tk.X, pady=5, padx=5)

        # 抓取页数
        tk.Label(page_frame, text="抓取页数:").grid(row=0, column=0, sticky='e', padx=5)
        self.page_count = tk.Spinbox(page_frame, from_=1, to=100, width=5)
        self.page_count.grid(row=0, column=1, sticky='w')
        self.page_count.delete(0, tk.END)
        self.page_count.insert(0, "1")

        # 加载等待时间 (新增)
        tk.Label(page_frame, text="加载等待时间(秒):").grid(row=0, column=2, sticky='e', padx=5)
        self.loading_time_entry = tk.Spinbox(page_frame, from_=1, to=30, width=5)
        self.loading_time_entry.grid(row=0, column=3, sticky='w')
        self.loading_time_entry.delete(0, tk.END)
        self.loading_time_entry.insert(0, str(self.config.get("loading_time", 3)))

        # 调整列权重使输入框可以扩展
        path_frame.columnconfigure(1, weight=1)

        # ==================== 操作按钮区域 ====================
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=15)

        start_btn = tk.Button(btn_frame,
                              text="开始作业",
                              command=self.start_process,
                              bg="#4CAF50",
                              fg="white",
                              font=('Microsoft YaHei', 10, 'bold'),
                              padx=25,
                              pady=5)
        start_btn.pack()

        # 设置窗口最小大小
        self.root.minsize(550, 950)

        # 让所有元素居中
        for child in main_frame.winfo_children():
            child.pack_configure(anchor='center')

    def select_save_path(self):
        """选择文件保存路径"""
        path = filedialog.askdirectory(initialdir=self.save_path_var.get())
        if path:
            self.save_path_var.set(path)
            self.config["default_save_path"] = path
            self.save_config()

    def start_process(self):
        """支持分页的处理流程"""
        try:
            # 更新配置
            self.config.update({
                "enterprise_coords": {
                    "x": int(self.enterprise_x.get()),
                    "y": int(self.enterprise_y.get())
                },
                "next_step_coords": {
                    "x": int(self.next_step_x.get()),
                    "y": int(self.next_step_y.get())
                },
                "table_coords": {
                    "x": int(self.table_x.get()),
                    "y": int(self.table_y.get())
                },
                "loading_time": int(self.loading_time_entry.get()),
                "default_save_path": self.save_path_var.get()
            })
            self.save_config()
            self.loop_running = True

            # 获取并保存公司名
            self.current_company = self.copy_entire_table(is_company_name=True)
            if not self.current_company:
                return
            if self.debug_mode:
                messagebox.showinfo("提示", f"当前公司: {self.current_company}")

            # 获取页数
            total_pages = int(self.page_count.get())
            if total_pages < 1:
                raise ValueError("页数必须大于0")

            # 执行分页抓取
            self.multi_page_capture(total_pages)

            # 如果用户主动终止则弹窗
            if not self.loop_running:
                print("用户主动终止")
                messagebox.showinfo("提示", f"用户终止运行！")

        except ValueError as ve:
            messagebox.showerror("输入错误", f"无效的输入值: {str(ve)}")
        except Exception as e:
            messagebox.showerror("处理错误", f"操作失败: {str(e)}")

    def multi_page_capture(self, total_pages):
        """执行多页抓取和保存"""
        # 初始化Excel文件
        wb = Workbook()
        ws = wb.active

        # 修改表头（添加"公司"列）
        headers = ["序号", "姓名", "身份证号", "职业资格", "证书编号", "备注", "公司"]
        ws.append(headers)

        # 生成唯一文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{self.current_company}_{timestamp}.xlsx"
        save_path = os.path.join(self.save_path_var.get(), filename)

        # 分页处理
        for page in range(1, total_pages + 1):
            if not self.loop_running:
                print("用户主动终止")
                break
            try:
                # 1. 抓取当前页数据
                data = self.process_single_page()

                if data:
                    # 2. 追加数据到Excel（添加公司列）
                    for row in data:
                        # 确保每行数据有6个字段（根据你的数据结构）
                        if len(row) < 6:
                            row += [""] * (6 - len(row))
                        # 添加公司名列
                        row.append(self.current_company)
                        ws.append(row)

                    # 3. 临时保存（防止中途出错丢失数据）
                    wb.save(save_path)
                    print(f"第 {page} 页数据已临时保存")

                # 如果不是最后一页，点击下一页
                if page < total_pages:
                    self.click_next_button()

            except Exception as e:
                messagebox.showerror("错误", f"第 {page} 页处理失败: {str(e)}")
                continue

        # 最终保存
        wb.save(save_path)
        messagebox.showinfo("完成", f"所有数据已保存到:\n{save_path}")

    def process_single_page(self):
        """处理单页数据"""
        clipboard_text = self.copy_entire_table()
        if not clipboard_text:
            return None

        print(f"=== 第 {getattr(self, 'current_page', 1)} 页原始数据 ===")
        print(repr(clipboard_text[:500]))

        return self.parse_clipboard_data(clipboard_text)

    def click_next_button(self):
        """点击下一步按钮"""
        next_x = self.config["next_step_coords"]["x"]
        next_y = self.config["next_step_coords"]["y"]

        pyautogui.moveTo(next_x, next_y, duration=0.5)
        pyautogui.click()
        print(f"已点击下一步按钮 ({next_x}, {next_y})")

        # 等待页面加载
        time.sleep(self.config["loading_time"])  # 可根据实际需要调整等待时间

        # 添加加载完成检测（可选）
        # pyautogui.moveTo(next_x + 50, next_y)  # 移动鼠标避免遮挡

    def copy_entire_table(self, is_company_name=False):
        """三连击复制表格数据
        :param is_company_name: 是否为复制公司名模式
        """
        try:
            # 根据模式选择坐标
            if is_company_name:
                x = self.config["enterprise_coords"]["x"]
                y = self.config["enterprise_coords"]["y"]
            else:
                # 使用配置的表格坐标
                x = self.config["table_coords"]["x"]
                y = self.config["table_coords"]["y"]

            # 移动到指定位置并模拟鼠标左键三连击
            pyautogui.moveTo(x, y, duration=0.5)
            for _ in range(3):
                pyautogui.click()
                time.sleep(0.15)

            # 执行复制操作
            for _ in range(3):  # 最多尝试3次
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.5)
                if pyperclip.paste():
                    break

            clipboard_data = pyperclip.paste().strip()
            if not clipboard_data:
                raise ValueError("剪贴板内容为空")

            # 如果是公司名模式，只返回第一行（去除可能的换行符）
            if is_company_name:
                pyautogui.press("pagedown")
                return clipboard_data.split('\n')[0].strip()
            return clipboard_data

        except Exception as e:
            error_msg = f"复制{'公司名' if is_company_name else '表格'}失败: {str(e)}"
            messagebox.showerror("错误", error_msg)
            return None

    def parse_clipboard_data(self, clipboard_text):
        """专用解析器（适用于你的数据结构）"""
        text = clipboard_text.replace('\r\n', '\n').replace('\r', '\n').strip()
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # 调试打印原始行数
        print(f"原始数据总行数: {len(lines)}")

        # 按6行一组重组（根据你的实际字段数调整）
        table_data = []
        field_count = 6  # 每组记录的字段数（根据你的数据调整）

        for i in range(0, len(lines), field_count):
            record = lines[i:i + field_count]
            # 补全缺失字段
            record += [""] * (field_count - len(record))
            table_data.append(record)

        # 调试输出
        print(f"重组后表格: {len(table_data)}行 x {field_count}列")
        print("样例数据:", table_data[:1])  # 打印第一行样例
        return table_data

    def save_to_excel(self, filepath=None):
        """单页保存方法（保持兼容性）"""
        try:
            # 获取表格数据
            clipboard_text = self.copy_entire_table()
            if not clipboard_text:
                return False

            # 解析数据
            table_data = self.parse_clipboard_data(clipboard_text)
            if not table_data:
                messagebox.showwarning("警告", "未获取到有效数据")
                return False

            # 处理文件路径
            if filepath is None:
                filepath = f"{self.current_company}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            else:
                if os.path.isdir(filepath):
                    filename = f"{self.current_company}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    filepath = os.path.join(filepath, filename)
                elif not filepath.endswith('.xlsx'):
                    filepath += '.xlsx'

            # 创建Excel文件
            wb = Workbook()
            ws = wb.active
            headers = ["序号", "姓名", "身份证号", "职业资格", "证书编号", "备注", "公司"]
            ws.append(headers)

            # 添加数据（含公司名）
            for row in table_data:
                if len(row) < 6:
                    row += [""] * (6 - len(row))
                row.append(self.current_company)
                ws.append(row)

            wb.save(filepath)
            messagebox.showinfo("成功", f"数据已保存到:\n{filepath}")
            return True

        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")
            return False

    def exit_program(self):
        print("\n用户按下 ESC，程序终止！")
        self.loop_running = False  # 通知主线程退出

if __name__ == '__main__':
    root = tk.Tk()
    root.geometry("550x950+10+10")
    app = WebTableToExcelApp(root)
    root.mainloop()