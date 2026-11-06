import pandas as pd
import openpyxl as xl
import re
import os

#-----------------------------

from functions import Init
from functions import Log

#=======================================

def test():
    import openpyxl

    # 打开 Excel 文件
    file_path = 'test.xlsx'
    workbook = openpyxl.load_workbook(file_path)

    # 获取指定表单
    sheet = workbook['Sheet1']  # 替换为你要读取的表单名称

    # 逐行读取数据
    for row in sheet.iter_rows(values_only=True):
        print(row)

    import openpyxl

    # 打开 Excel 文件
    file_path = 'test.xlsx'
    workbook = openpyxl.load_workbook(file_path)

    # 获取指定表单
    sheet = workbook['Sheet1']  # 替换为你要读取的表单名称

    # 逐列读取数据
    for col in sheet.iter_cols(values_only=True):
        print(col)

        # 只读取前2行和前2列的数据
    for row in sheet.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2, values_only=True):
        print(row)

    # 只读取前2列的数据
    for col in sheet.iter_cols(min_col=1, max_col=2, min_row=1, max_row=2, values_only=True):
        print(col)

#=======================================

def create_file(file_path) -> bool:
    """
    通用文件创建函数（自动跳过已存在文件）
    增强功能：为Office文档创建有效空白文件

    参数：
        file_path (str): 要创建的文件路径（含文件名和扩展名）
    """
    # 检查文件是否已存在
    if os.path.exists(file_path):
        print(f"文件已存在，跳过创建: {file_path}")
        return True

    # 创建父目录（如果不存在）
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)

    # 获取文件扩展名
    file_ext = os.path.splitext(file_path)[1].lower()

    try:
        # 根据文件类型选择创建方式
        if file_ext == '.docx':
            from docx import Document
            Document().save(file_path)
        elif file_ext == '.xlsx':
            from openpyxl import Workbook
            Workbook().save(file_path)
        elif file_ext == '.pptx':
            from pptx import Presentation
            Presentation().save(file_path)
        else:
            # 非Office文件使用基础创建方式
            with open(file_path, 'w') as f:
                pass

        print(f"文件创建成功: {file_path}")
        return True

    except ImportError as ie:
        print(f"创建失败：{file_ext} 文件需要额外依赖库")
    except Exception as e:
        print(f"创建文件失败: {str(e)}")

    finally:
        return False

############### 文件保存函数 ###############

def SaveFile(source, Type: str, path: str) -> bool:
    """
    将 source 数据按指定类型保存到文件
    :param source: 数据对象（如 openpyxl.Workbook 或 python-docx.Document）
    :param Type: 文件类型（'excel' 或 'word'）
    :param path: 保存路径
    :return: 成功返回 True，失败返回 False
    """
    try:
        # Excel 文件
        if Type == 'excel':
            try:
                source.save(path)
                print(f"Excel 文件已成功保存至: {path}")
                return True
            except PermissionError:
                print(f"错误: 无写入权限或文件被占用 - {path}")
            except IOError as e:
                print(f"IO错误: 路径无效或磁盘空间不足 - {e}")
            except Exception as e:
                print(f"保存Excel文件时发生未预期错误: {e}")
            return False

        # Word 文件（预留实现）
        elif Type == 'word':
            try:
                # 假设使用 python-docx 的保存逻辑
                source.save(path)
                print(f"Word 文件已成功保存至: {path}")
                return True
            except AttributeError:
                print("错误: 传入的Word对象不支持 save() 方法")
            except PermissionError:
                print(f"错误: 无写入权限或文件被占用 - {path}")
            except Exception as e:
                print(f"保存Word文件时发生未预期错误: {e}")
            return False

        # 未知文件类型
        else:
            print(f"错误: 不支持的文件类型 '{Type}'")
            return False

    # 捕获全局异常（如 Type 参数类型错误等）
    except Exception as e:
        print(f"保存过程中发生全局异常: {e}")
        return False

############### 两种匹配函数 ##############
def re_match( match, obj):
    """
    判断目标字符串中是否包含指定的字面量子串（自动处理正则表达式特殊字符）

    参数:
        match_str (str): 要检查的普通字符串（非正则表达式）
        obj (str): 被搜索的目标字符串

    返回:
        bool: 如果目标字符串包含指定子串则返回 True，否则返回 False

    示例:
        >>> re_match("file.txt", "Opening file.txt...")
        True
        >>> re_match("a.b", "a+b")  # 不会误将 "." 当作正则表达式通配符
        False
    """
    try:
        return match in str(obj)
    except TypeError:
        # 处理输入非字符串的情况（如数字输入）
        return False

def whole_match( match, obj):
        if match == obj:
            return True
        return False

############### Excel 操作 ##############
def Excel_operate(source: xl.Workbook, OperateCode: list):
    # pass

    try:
        # 读取 Excel 文件中所有工作表的名字
        Sheets = source.sheetnames

        for sheet in Sheets:
            try:
                data = source[sheet]
                Log.Record( "Reading data from the active worksheet of the current file", 1)

            except Exception as e:
                Log.Record( "读取工作表 '{sheet}' 时发生错误: {e}", 1)
                print(f"处理工作表 '{sheet}' 时发生错误: {e}")

                continue  # 继续处理下一个工作表

            try:
                for code in OperateCode:
                    if   code['mode'] == '1':
                        Excel_add(data, code)
                    elif code['mode'] == '2':
                        Excel_search( data, code)
                    elif code['mode'] == '3':
                        Excel_delete( data, code)
                    elif code['mode'] == '4':
                        Excel_change( data, code)
                    else:
                        print(f"错误：未知的操作模式 '{code['mode']}'。")

            except KeyError:
                print("错误：操作码中缺少 'mode' 键。")
            except Exception as e:
                print(f"执行操作时发生未预期错误: {e}")

    except AttributeError as e:
        print(f"文件对象错误，请确认已正确打开Excel文件: {e}")
    except Exception as e:
        print(f"程序运行发生未捕获异常: {e}")

def Excel_add( SheetData, OperateCode):
    # pass
    print('Excel_add')

    row = SheetData.max_row
    cols = SheetData.max_column
    # print( row, cols) # 获取行数和列数 test

    # Add a new row to row i of the sheet data table in the excel file test.xlsx
    # Operate on the active worksheet of the current file

    ##########
    if   OperateCode['code'][0] == '1':
        # 列模式

        ##########
        if   OperateCode['code'][1] == '1':
            # 插入空白列
            SheetData.insert_cols(int(OperateCode['column']))

            # 对空白列进行赋值
            for i in range(row):
                SheetData.cell(row=i + 1, column=int(OperateCode['column']), value=OperateCode['value'])

            Log.Record( "Add a new column to column [" + OperateCode['column'] + "] of the active worksheet of the current file",
                        1)

        ##########
        elif OperateCode['code'][1] == '2':

            i = 0
            while i < cols :
                if re.search(OperateCode['match'], str(SheetData.cell(row=1, column=i + 1).value)):
                    # 插入空白列
                    SheetData.insert_cols( idx= i+1, amount= 1)

                    # 对空白列进行赋值
                    for j in range(row):
                        SheetData.cell(row=j + 1, column=i+1, value=OperateCode['value'])

                    i += 1 # 跳过插入后的列

                    Log.Record(
                        "Add a new column to column [" + str(i) + "] of the active worksheet of the current file",
                        1)

                # 操作步数累加
                i += 1

        ##########
        elif OperateCode['code'][1] == '3':

            i = 0
            while i < cols:
                if OperateCode['match'] == str(SheetData.cell(row=1, column=i + 1).value):
                    # 插入空白列
                    SheetData.insert_cols( idx= i+1, amount= 1)

                    # 对空白列进行赋值
                    for j in range(row):
                        SheetData.cell(row=j + 1, column=i+1, value=OperateCode['value'])

                    i += 1  # 跳过插入后的列

                    Log.Record("Add a new row to row [" + str(i) + "] of the active worksheet of the current file", 1)

                # 操作步数累加
                i += 1
        # pass

    ##########
    elif OperateCode['code'][0] == '2':
        # 行模式

        ##########
        if OperateCode['code'][1] == '1':
            # pass
            SheetData.insert_rows(int(OperateCode['row']))

            Log.Record("Add a new row to row [" + OperateCode['row'] + "] of the active worksheet of the current file", 1)

            for i in range(cols):
                SheetData.cell(row=int(OperateCode['row']), column=i + 1, value=OperateCode['value'])

        ##########
        elif OperateCode['code'][1] == '2':
            # pass

            i = 0
            while i < row:
                if re.search(OperateCode['match'], str(SheetData.cell(row=i + 1, column=1).value)):
                    SheetData.insert_rows( idx= i+1, amount= 1)

                    for j in range(cols):
                        SheetData.cell( row= i+1, column= j+1, value=OperateCode['value'])

                    i += 1  # 跳过插入后的行

                    Log.Record("Add a new row to row [" + str(i) + "] of the active worksheet of the current file",
                               1)

                i += 1

        elif OperateCode['code'][1] == '3':
            # pass

            i = 0
            while i < row:
                if OperateCode['match'] == str(SheetData.cell(row=i + 1, column=1).value):
                    SheetData.insert_rows( idx= i+1, amount= 1)

                    Log.Record("Add a new row to row [" + str(i) + "] of the active worksheet of the current file",
                               1)

                    for j in range(cols):
                        SheetData.cell( row= i+1, column= j+1, value=OperateCode['value'])

                    i += 1  # 跳过插入后的行

                i += 1

        #######

        # SheetData.insert_rows(int(OperateCode['row']))

        # for i in range(cols):
        #     SheetData.cell( row = int(OperateCode['row']), column = i+1, value = OperateCode['value'])

    ##########
    elif OperateCode['code'][0] == '3':
        # pass

        if OperateCode['code'][1] == '1':
            # pass
            i = OperateCode['row']
            j = OperateCode['column']

            x = row + 1

            # 腾个空格来添加值
            while i < x :
                SheetData.cell( row= x, column= j).value = SheetData.cell( row= x-1, column= j).value
                x -= 1

            SheetData.cell( row= i, column= j).value = OperateCode['value']

def Excel_search( SheetData, OperateCode):
    # pass

    print('Excel_search')

    row = SheetData.max_row
    cols = SheetData.max_column
    # print( row, cols) # 获取行数和列数 test

    if OperateCode['code'][1] == '1':
        # pass

        if OperateCode['match'] == str(SheetData.cell( row= int(OperateCode['row']), column= int(OperateCode['column'])).value):
            # pass

            Log.Record( "Complete a search at the specified coordinates" , 1)

    if OperateCode['code'][1] == '2':
        # pass

        i = 0
        j = 0

        while i < row:
            while j < cols:
                if re.search(OperateCode['match'], str(SheetData.cell( row= i+1, column= j+1).value)):
                    # pass

                    Log.Record( "Complete a search at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)

    if OperateCode['code'][1] == '3':
        # pass

        i = 0
        j = 0

        while i < row:
            while j < cols:
                if OperateCode['match'] == str(SheetData.cell( row= i+1, column= j+1).value):
                    # pass

                    Log.Record( "Complete a search at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)

def Excel_insert( SheetData, OperateCode: list):
    pass

def Excel_delete( SheetData, OperateCode: list):
    pass

def Excel_change( SheetData, OperateCode: list):
    pass

def Excel_Sort(self):
    pass

def Excel_Copy( OperateCode: dict):
    # pass

    if OperateCode['code'] == '1':
        # pass
        source_data = pd.ExcelFile( OperateCode['source_file'], engine='openpyxl')


    else:
        pass

def Excel_ConditionalFiltrate( OperateCode: dict, ):
    pass

    if OperateCode['code'][2] == '0':
        pass

    if OperateCode['code'][2] == '1':
        pass

    if OperateCode['code'][2] == '2':
        pass

def search_specific_row(ws, row_num, target_value):
    matches = []
    # 限定遍历范围（仅指定行）
    for cell in ws.iter_rows(
            min_row=row_num, max_row=row_num,
            min_col=1, max_col=ws.max_column,
            values_only=False
    ):
        for c in cell:
            if c.value == target_value:  # 精确匹配
                # matches.append((c.row, c.column))  # 返回数字坐标，如(5,3)
                # 或返回字母坐标：matches.append(c.coordinate)
                return c.column
    # return matches

def search_specific_column(ws, col_num, target_value):
    matches = []
    # 限定遍历范围（仅指定列）
    for cell in ws.iter_cols(
            min_col=col_num, max_col=col_num,
            min_row=1, max_row=ws.max_row,
            values_only=False
    ):
        for c in cell:
            if c.value == target_value:
                # matches.append((c.row, c.column))
                return c.row
    # return matches

def search_by_range(ws, target_value, rows=None, cols=None):
    matches = []
    # 设置行列范围（默认全表）
    min_row = min(rows) if rows else 1
    max_row = max(rows) if rows else ws.max_row
    min_col = min(cols) if cols else 1
    max_col = max(cols) if cols else ws.max_column

    # 遍历指定区域
    for row in ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
            values_only=False
    ):
        for cell in row:
            if cell.value == target_value:
                matches.append((cell.row, cell.column))  # 返回数字坐标
    return matches

#---------------------------------------

# 使用示例
if __name__ == "__main__":
    pass

    # 测试不同格式文件创建
    # files_to_create = [
    #     "Z:\\demo.docx",  # Word文档
    #     "Z:\\presentation.pptx",  # PPT文档
    #     "Z:\\data.xlsx",  # Excel文档
    #     "Z:\\notes.txt",  # 文本文件
    #     "Z:\\script.py",  # Python代码
    #     "Z:\\program.cpp",  # C++代码
    #     "Z:\\binary.bin",  # 二进制文件
    #     "Z:\\subdir\\image.jpg"  # 带目录结构的文件
    # ]

    # for file in files_to_create:
    #     create_file(file)