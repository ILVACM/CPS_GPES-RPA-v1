import os
import json
from typing import List, Dict

def permission_change(Permission_ID: int, status: bool):
    try:
        with open( 'config\permission.json', 'r') as f:
            config = json.load(f)

        if Permission_ID == 1:
            config["OSEnv"] = status

        elif Permission_ID == 2:
            config["Config"] = status

        elif Permission_ID == 3:
            config["Port"] = status

        elif Permission_ID == 4:
            config["Log"] = status

        elif Permission_ID == 5:
            config["FileIO"] = status

        elif Permission_ID == 6:
            config["Network"] = status

        elif Permission_ID == 7:
            config["WeChat"] = status

        else:
            # ??? 日志记录
            return False

        with open( 'config\permission.json', 'w', encoding='utf-8') as file:
            json.dump( config, file, ensure_ascii=False, indent=4)  # 写回 JSON 文件

    except Exception as e:
        print(e)
        # ??? 日志记录
        return

def process_folder_to_json(input_folder: str, output_file: str) -> None:
    """
    将指定文件夹内所有文件路径结构化存储为JSON

    参数：
    input_folder : 需要扫描的文件夹路径
    output_file  : 输出的JSON文件路径

    返回：
    None (结果直接写入文件)
    """
    # 初始化操作列表
    operate_list: List[Dict] = []

    # 遍历文件系统
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            # 构造绝对路径
            abs_path = os.path.join(root, filename)

            # 构建标准字典结构
            item = {
                "mode": 0,
                "code": "1",
                "path": abs_path
            }

            operate_list.append(item)

    # 写入JSON文件
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({"OperateList": operate_list}, f,
                  ensure_ascii=False, indent=4)

# 使用示例
if __name__ == "__main__":
    process_folder_to_json(
        input_folder=r"D:\CPS-OA_RPA-20250312\config\IMG\1",
        output_file=r"D:\CPS-OA_RPA-20250312\config\temp_1_IMG.json"
    )