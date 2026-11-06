import os
import pandas as pd
import openpyxl as xl
import glob
import re

import functions.Init
import functions.LogEdit
   
#============================================================================

def read_excel_with_dynamic_headers(file_path):
    # 尝试读取 Excel 文件
    try:
        # 第一种情况：只有行表头（header=0表示第一行是表头）
        df = pd.read_excel(file_path, header=0)  # 默认读取第一行作为列名
        print("使用行表头读取：")
        print(df)
    except Exception as e:
        print(f"无法使用行表头读取：{e}")

    try:
        # 第二种情况：只有列表头（假设列表头在第一行，数据从第二行开始）
        df = pd.read_excel(file_path, header=None, skiprows=1)  # 跳过第一行作为列头
        print("使用列表头读取：")
        print(df)
    except Exception as e:
        print(f"无法使用列表头读取：{e}")

    try:
        # 第三种情况：行列表头都有（multi-level header）
        df = pd.read_excel(file_path, header=[0, 1])  # 读取前两行作为多级表头
        print("使用行列表头读取：")
        print(df)
    except Exception as e:
        print(f"无法使用行列表头读取：{e}")
    
#============================================================================

# from openpyxl import load_workbook

# # 打开已有的 Excel 文件
# file_path = 'Z:\\CPS-Auto\\updated_excel_file.xlsx'
# wb = load_workbook(file_path)
# ws = wb.active  # 获取活动工作表

# # 假设我们要在第2行第3列（C2单元格）插入一个新值
# insert_row = 2  # 指定插入行
# insert_col = 3  # 指定插入列
# new_value = "e"  # 要插入的值

# # 向下插入新数据，移动原来位置的数据
# ws.insert_rows(insert_row)

# # 向右插入新数据，移动原来位置的数据
# ws.insert_cols(insert_col)

# # 在指定位置写入新值
# ws.cell(row=insert_row, column=insert_col, value=new_value)

# # 保存更改
# wb.save('updated_excel_file.xlsx')

# print(f"已成功在 ({insert_row}, {insert_col}) 插入值 {new_value}")

#===============================================================================

# 两种匹配函数

def re_match( pattern, string):
    # 使用正则表达式匹配字符串
    match = re.search(pattern, string)
    
    # 如果匹配成功，返回匹配结果；否则，返回None
    return match.group(1) if match else None

##xxxx
def whole_match( pattern, string):
    # 使用正则表达式匹配字符串
    match = re.search(pattern, string)
    
    # 如果匹配成功，返回匹配结果；否则，返回None
    return match.group(0) if match else None

#=============================================================================

def test():
    import openpyxl

    # 打开 Excel 文件
    file_path = 'your_excel_file.xlsx'
    workbook = openpyxl.load_workbook(file_path)

    # 获取指定表单
    sheet = workbook['Sheet1']  # 替换为你要读取的表单名称

    # 逐行读取数据
    for row in sheet.iter_rows(values_only=True):
        print(row)
    
    import openpyxl

    # 打开 Excel 文件
    file_path = 'your_excel_file.xlsx'
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


#=============================================================================
    
def operate():
    # print('operate')
    
    # 获取所有 Excel 文件路径
    ExcelFiles = glob.glob(os.path.join( functions.Init.Download, '*.xlsx'))
    
    for excel in ExcelFiles:
        # print(excel)
        
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
            
            # # pandas库 读取 Excel 文件
            # data = pd.read_excel( excel, sheet_name=sheet)
                
            # openpyxl库 读取 Excel 文件指定表单数据
            data = wb[sheet]
            
            # test
            # data.insert_rows(1)
            # data.insert_cols(1)
            # cell_value_c3 = sheet.cell(row=3, column=3).value
                            
            for OperateCode in functions.Init.OperateList:
                
                if OperateCode['mode'] == '1':
                    add( OperateCode, data)
                
                elif OperateCode['mode'] == '2':
                    search( OperateCode, data)
                
                elif OperateCode['mode'] == '3':
                    change( OperateCode, data)
                
                elif OperateCode['mode'] == '4':
                    delete( OperateCode, data)
                    
                else:
                    pass
                
            #
            
        wb.save(excel)
    
def add( code, data):
    print('add')
    
    row = data.max_row
    cols = data.max_column
    # print( row, cols) # 获取行数和列数 test
    
    # Add a new row to row i of the sheet data table in the excel file test.xlsx
    # Operate on the active worksheet of the current file
    
    ##########
    if   code['OperateCode'][0] == '1':
        # 列模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # 插入空白列
            data.insert_cols(int(code['column']))
            
            # 对空白列进行赋值
            for i in range(row):
                data.cell( row = i+1, column = int(code['column']), value = code['value'])
                
            functions.LogEdit.Record( "Add a new column to column [" + code['column'] + "] of the active worksheet of the current file", 1)
        
        ##########
        elif code['OperateCode'][1] == '2':
            
            for i in range(cols):
                if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                    # 插入空白列
                    data.insert_cols(i)
                    
                    functions.LogEdit.Record( "Add a new column to column [" + str(i) + "] of the active worksheet of the current file", 1)
                    
                    # 对空白列进行赋值
                    for j in range(row):
                        data.cell( row = j+1, column = i, value = code['value'])
        
        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            for i in range(cols):
                if code['match'] == data.cell( row = 1, column = i+1).value:
                    # 插入空白列
                    data.insert_cols(i)
                    
                    functions.LogEdit.Record( "Add a new row to row [" + str(i) + "] of the active worksheet of the current file", 1)
                    
                    # 对空白列进行赋值
                    for j in range(row):
                        data.cell( row = j+1, column = i, value = code['value'])
        
        # pass
    
    ##########
    elif code['OperateCode'][0] == '2':
        # 行模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            data.insert_rows(int(code['row']))
            
            functions.LogEdit.Record( "Add a new row to row [" + code['row'] + "] of the active worksheet of the current file", 1)
            
            for i in range(cols):
                data.cell( row = int(code['row']), column = i+1, value = code['value'])
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
           for i in range(row):
               if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                    data.insert_rows(i)
                    
                    functions.LogEdit.Record( "Add a new row to row [" + str(i) + "] of the active worksheet of the current file", 1)
                   
                    for j in range(cols):
                        data.cell( row = int(code['row']), column = i+1, value = code['value'])
        
        elif code['OperateCode'][1] == '3':
            # pass
            for i in range(row):
                if code['match'] == data.cell( row = i+1, column = 1).value:
                    data.insert_rows(i)
                    
                    functions.LogEdit.Record( "Add a new row to row [" + str(i) + "] of the active worksheet of the current file", 1)
                   
                    for j in range(cols):
                        data.cell( row = int(code['row']), column = i+1, value = code['value'])
        
        #######
        
        # data.insert_rows(int(code['row']))
            
        # for i in range(cols):
        #     data.cell( row = int(code['row']), column = i+1, value = code['value'])
    
    elif code['OperateCode'][0] == '3':
        
        # # 单个单元格
        # if   code['OperateCode'][1] == '1':
        #     # pass
        
        # elif code['OperateCode'][1] == '2':
        #     # pass

        # elif code['OperateCode'][1] == '3':
        #     # pass

        pass
    
def search( code, data):
    print('search')
    
    row = data.max_row
    cols = data.max_column
    # print( row, cols) # 获取行数和列数 test
    
    ##########
    if   code['OperateCode'][0] == '1':
        # pass # 列模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            
            if code['match'] == data.cell( row = 1, column = int(code['column'])).value:
                
                functions.LogEdit.Record( "Complete a match at the specified column coordinates:" + code['column'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified column coordinates:" + code['column'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        functions.LogEdit.Record( "Complete a match at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        functions.LogEdit.Record( "Complete a match at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        
                    # else:
                        
                    #     functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        functions.LogEdit.Record( "Complete a match at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                    
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        functions.LogEdit.Record( "Complete a match at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        # return True
                return signal
    
    ##########
    elif code['OperateCode'][0] == '2':
        # pass # 行模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = 1).value:
                
                functions.LogEdit.Record( "Complete a match at the specified row coordinates:" + code['row'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified row coordinates:" + code['row'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        functions.LogEdit.Record( "Complete a match at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                    
                return False
                
            else:
                
                signal = False
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        functions.LogEdit.Record( "Complete a match at the row coordinates:" + code['row'], 1)
                        # return True
                    
                return signal

        #########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        functions.LogEdit.Record( "Complete a match at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        functions.LogEdit.Record( "Complete a match at the row coordinates:" + code['row'], 1)
                        
                        signal = True
                        
                        # return True
                    # else:
                return signal
    
    ##########
    elif code['OperateCode'][0] == '3':
        # pass # 单个单元格
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = int(code['column'])).value:
                
                functions.LogEdit.Record( "Complete a match at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                #
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            functions.LogEdit.Record( "Complete a match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                
                # functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            functions.LogEdit.Record( "Complete a match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                            # return True
                        # else:

                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            
                            functions.LogEdit.Record( "Complete a match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                        # else:
                            
                functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                            
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            # return True
                            
                            functions.LogEdit.Record( "Complete a match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                        # else:
                        
                return signal
    
def change( code, data):
    print('change')
    
    row = data.max_row
    cols = data.max_column
    # print( row, cols) # 获取行数和列数 test
    
    ##########
    if   code['OperateCode'][0] == '1':
        # pass # 列模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            
            if code['match'] == data.cell( row = 1, column = int(code['column'])).value:
                
                data.cell( row = 1, column = int(code['column'])).value = code['value']
                
                functions.LogEdit.Record( "Complete a change at the specified column coordinates:" + code['column'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO change at the specified column coordinates:" + code['column'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        
                    # else:
                        
                    #     functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                    
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        # return True
                return signal
    
    ##########
    elif code['OperateCode'][0] == '2':
        # pass # 行模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = 1).value:
                
                data.cell( row = int(code['row']), column = 1).value = code['value']
                
                functions.LogEdit.Record( "Complete a change at the specified row coordinates:" + code['row'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified row coordinates:" + code['row'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        data.cell( row = i+1, column = 1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                    
                return False
                
            else:
                
                signal = False
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        data.cell( row = i+1, column = 1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        # return True
                    
                return signal

        #########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        data.cell( row = i+1, column = 1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        data.cell( row = i+1, column = 1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        signal = True
                        
                        # return True
                    # else:
                return signal
    
    ##########
    elif code['OperateCode'][0] == '3':
        # pass # 单个单元格
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = int(code['column'])).value:
                
                data.cell( row = int(code['row']), column = int(code['column'])).value = code['value']
                
                functions.LogEdit.Record( "Complete a change at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                #
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            data.cell( row = i+1, column = j+1).value = code['value']
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                
                # functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            data.cell( row = i+1, column = j+1).value = code['value']
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                            # return True
                        # else:

                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            
                            data.cell( row = i+1, column = j+1).value = code['value']
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                        # else:
                            
                functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                            
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            # return True
                            
                            data.cell( row = i+1, column = j+1).value = code['value']
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                        # else:
                        
                return signal
    
def delete( code, data):
    print('delete')
    
    row = data.max_row
    cols = data.max_column
    # print( row, cols) # 获取行数和列数 test
    
    ##########
    if   code['OperateCode'][0] == '1':
        # pass # 列模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            
            if code['match'] == data.cell( row = 1, column = int(code['column'])).value:
                
                data.cell( row = 1, column = int(code['column'])).value = code['value']
                
                functions.LogEdit.Record( "Complete a change at the specified column coordinates:" + code['column'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO change at the specified column coordinates:" + code['column'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if re.search( code['match'], str(data.cell( row = 1, column = i+1).value)):
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        
                    # else:
                        
                    #     functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                        
                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        return True
                    
                functions.LogEdit.Record( "NO match at the column coordinates:" + code['column'], 2)
                    
                return False
                
            else:
                
                signal = False
                
                for i in range(cols):
                    if code['match'] == data.cell( row = 1, column = i+1).value:
                        
                        data.cell( row = 1, column = i+1).value = code['value']
                        
                        functions.LogEdit.Record( "Complete a change at the column coordinates:" + code['column'], 1)
                        
                        signal = True
                        # return True
                return signal
    
    ##########
    elif code['OperateCode'][0] == '2':
        # pass # 行模式
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = 1).value:
                
                data.cell( row = int(code['row']), column = 1).value = None
                
                functions.LogEdit.Record( "Complete a change at the specified row coordinates:" + code['row'], 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified row coordinates:" + code['row'], 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        data.cell( row = i+1, column = 1).value = None
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                    
                return False
                
            else:
                
                signal = False
                for i in range(row):
                    if re.search( code['match'], str(data.cell( row = i+1, column = 1).value)):
                        
                        data.cell( row = i+1, column = 1).value = None
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        # return True
                    
                return signal

        #########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        data.cell( row = i+1, column = 1).value = None
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        return True
                
                functions.LogEdit.Record( "NO match at the row coordinates:" + code['row'], 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    if code['match'] == data.cell( row = i+1, column = 1).value:
                        
                        data.cell( row = i+1, column = 1).value = None
                        
                        functions.LogEdit.Record( "Complete a change at the row coordinates:" + code['row'], 1)
                        
                        signal = True
                        
                        # return True
                    # else:
                return signal
    
    ##########
    elif code['OperateCode'][0] == '3':
        # pass # 单个单元格
        
        ##########
        if   code['OperateCode'][1] == '1':
            # pass
            if code['match'] == data.cell( row = int(code['row']), column = int(code['column'])).value:
                
                data.cell( row = int(code['row']), column = int(code['column'])).value = None
                
                functions.LogEdit.Record( "Complete a change at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 1)
                
                return True
            else:
                
                functions.LogEdit.Record( "NO match at the specified Unit coordinates:[ " + code['column'] + ", " + code['row'] +"]", 2)
                
                return False
        
        ##########
        elif code['OperateCode'][1] == '2':
            # pass
            
            if code['OperateCode'][2] == '1':
                #
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            data.cell( row = i+1, column = j+1).value = None
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                
                # functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if re.search( code['match'], str(data.cell( row = i+1, column = j+1).value)):
                            
                            data.cell( row = i+1, column = j+1).value = None
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                            # return True
                        # else:

                return signal

        ##########
        elif code['OperateCode'][1] == '3':
            # pass
            
            if code['OperateCode'][2] == '1':
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            
                            data.cell( row = i+1, column = j+1).value = None
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            return True
                        # else:
                            
                functions.LogEdit.Record( "NO match at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 2)
                            
                return False
                
            else:
                signal = False
                for i in range(row):
                    for j in range(cols):
                        if code['match'] == data.cell( row = i+1, column = j+1).value:
                            # return True
                            
                            data.cell( row = i+1, column = j+1).value = None
                            
                            functions.LogEdit.Record( "Complete a change at the Unit coordinates:[ " + str(j+1) + ", " + str(i+1) +"]", 1)
                            
                            signal = True
                        # else:
                        
                return signal
    
#################################################################################
def excel_csv( file_path):

    file = pd.read_excel(file_path)
    
    print(file.head())

#===============================================================================

# path = 'Z:\\test.xlsx'

if __name__ == '__main__':
    functions.Init.init()
    operate()
    
#     # excel_csv(path)
#     read_excel_with_dynamic_headers("Z:\\CPS-Auto\\updated_excel_file.xlsx")
    
#===============================================================================
