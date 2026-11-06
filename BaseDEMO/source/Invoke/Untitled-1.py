import importlib
import json
from typing import List, Callable

import inspect

def load_functions_from_file(file_path: str) -> List[Callable]:
    # 读取文件内容
    if file_path.endswith(".json"):
        with open(file_path, "r") as f:
            data = json.load(f)
            function_strings = data.get("functions", [])
    else:  # 默认为TXT
        with open(file_path, "r") as f:
            function_strings = [line.strip() for line in f if line.strip()]
    
    function_refs = []
    for func_str in function_strings:
        try:
            # 分割模块路径和函数名（如 "utils.helper.parse_data" → 模块路径="utils.helper", 函数名="parse_data"）
            module_path, func_name = func_str.rsplit(".", 1)
            
            # 动态导入模块（支持跨层级）
            module = importlib.import_module(module_path)
            
            # 获取函数引用
            func = getattr(module, func_name)
            if callable(func):
                function_refs.append(func)
        except (ImportError, AttributeError, ValueError) as e:
            print(f"加载失败: {func_str}, 错误: {e}")
    
    return function_refs

# 从JSON文件加载
# function_list = load_functions_from_file(r"Z:\Invoke\functions.json")

# 从TXT文件加载
function_list = load_functions_from_file(r"Z:\Invoke\functions.txt")

# 调用所有函数
for func in function_list:
    print("########--------------------")
    result = func()  # 假设输入参数统一
    print(inspect.signature(func))  # 输出参数类型
    print(f"{func.__name__} 返回结果: {result}")
    print("------------------------\n")