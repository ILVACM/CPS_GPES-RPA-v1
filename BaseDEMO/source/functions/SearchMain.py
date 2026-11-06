import sys
import os

def add_main():
    # 将 main_folder 的路径添加到 sys.path 中
    sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../main.py')))

# # 现在可以导入 main.py 中的内容了
# import main

# # 访问 main.py 中的变量和函数
# print(main.my_variable)   # 输出: Hello from main.py
# main.print_message()      # 输出: Message from main.py
