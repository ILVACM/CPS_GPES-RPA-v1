import pandas as pd
import openpyxl as xl
from openpyxl import load_workbook
import os
import json
import win32com.client
import win32com.client as win32
import pythoncom
import datetime

# -------------------------------------

from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt
from pptx import Presentation

import os
import shutil
from datetime import datetime

from pathlib import Path

def copy_files_to_dated_folder(src_path, dest_root):
    """
    将源文件夹所有内容复制到目标路径下按时间命名的文件夹中

    参数:
    src_path (str): 源文件夹路径
    dest_root (str): 目标根目录路径

    返回:
    str: 成功返回目标文件夹绝对路径，失败返回None
    """
    try:
        # 生成时间格式并处理Windows非法字符
        timestamp = datetime.now().strftime("%Y.%m.%d_%H-%M")
        dest_folder = os.path.abspath(  # 新增路径标准化
            os.path.join(dest_root, timestamp)
        )

        # 创建目标文件夹并复制内容
        shutil.copytree(src_path, dest_folder)
        print(f"成功复制: {src_path} -> {dest_folder}")
        return dest_folder  # 直接返回绝对路径

    except FileNotFoundError as e:
        print(f"路径不存在错误: {e}")
    except PermissionError as e:
        print(f"权限不足: {e}")
    except shutil.Error as e:
        print(f"复制错误: {e}")
    except Exception as e:
        print(f"未知错误: {e}")

    return None


TestList = [
    {
        "type": "excel",
        "mode": "1",
        "operate code": "311",
        "row": 2,
        "column": 2,
        "match": "",
        "value": "test2"
    },
    {
        "type": "excel",
        "mode": "1",
        "operate code": "121",
        "row": -1,
        "column": -1,
        "match": "a",
        "value": "test1"
    },
    {
        "type": "excel",
        "mode": "1",
        "operate code": "231",
        "row": -1,
        "column": -1,
        "match": "3",
        "value": "test3"
    }
]

def analyze_excel_column(file_path, sheet_name, target_column):
    try:
        # 读取指定列（优化性能）
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=[target_column], engine='openpyxl')

        # 检查列是否存在
        if target_column not in df.columns:
            raise ValueError(f"列名 '{target_column}' 不存在")

        # 统计行数
        total_rows = len(df)
        non_null_rows = df[target_column].count()

        print(f"列 '{target_column}' 总行数（含空值）: {total_rows}")
        print(f"列 '{target_column}' 有效数据量（非空值）: {non_null_rows}")

        return df[target_column]

    except Exception as e:
        print(f"操作失败：{str(e)}")

# 读取并合并文件
def merge_excel(file1, file2, output):
    df1 = pd.read_excel(file1, sheet_name="作业明细表", header=0)
    df2 = pd.read_excel(file2, sheet_name="作业明细表", header=0)
    combined_df = pd.concat([df1, df2], ignore_index=True)
    combined_df.to_excel(output, sheet_name="Sheet1", index=False)


def insert_image_to_ppt(ppt_path, slide_index, img_path, position='center'):
    """ 企业级PPT图片插入方案（严格匹配原图规范） """
    # 验证文件可达性
    if not os.access(img_path, os.R_OK):
        raise FileNotFoundError(f"图片路径不可访问: {img_path}")

    prs = Presentation(ppt_path) if os.path.exists(ppt_path) else Presentation()

    # 智能版式管理系统
    while len(prs.slides) <= slide_index:
        new_slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
        # 精准还原原图渐变背景
        background = new_slide.background
        fill = background.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = RGBColor(0, 32, 96)  # 深蓝
        fill.gradient_stops[1].color.rgb = RGBColor(0, 64, 128)  # 渐变蓝

    slide = prs.slides[slide_index]

    # 专业定位算法（匹配原图布局）
    slide_width = prs.slide_width
    base_top = Cm(3.8)  # 精确测量原图标题下沿位置

    position_config = {
        'center': (int((slide_width - Cm(15)) / 2), base_top),
        'left': (Cm(2.5), base_top)  # 左侧边距微调
    }

    try:
        left, top = position_config[position.lower()]
    except KeyError:
        raise ValueError("定位参数仅支持 'center' 或 'left'")

    # 插入图片（保持原图比例）
    try:
        pic = slide.shapes.add_picture(img_path, left, top, width=Cm(20))
        # 应用原图样式规范
        pic.line.color.rgb = RGBColor(0, 112, 192)  # 主题蓝边框
        pic.line.width = Pt(1.2)  # 精确到0.1pt
    except Exception as e:
        raise RuntimeError(f"图形引擎错误: {str(e)}")

    prs.save(ppt_path)


def set_signal_to_true(file_path):
    try:
        # 1. 读取JSON文件
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)  # 解析为字典

        # 2. 修改signal字段
        data['signal'] = True

        # 3. 写回文件（格式化缩进提升可读性）
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        print("signal已成功设置为True")

    except FileNotFoundError:
        print(f"错误：文件 {file_path} 不存在")
    except json.JSONDecodeError:
        print("错误：JSON格式不合法")
    except KeyError:
        print("错误：JSON中不存在signal字段")
    except PermissionError:
        print("错误：文件被占用或无写入权限")


def convert_xls_to_xlsx(src_path):
    # 新增COM库初始化
    pythoncom.CoInitialize()

    # 转换为绝对路径并规范化
    src_path = os.path.abspath(src_path)
    src_path = os.path.normpath(src_path)

    # 检查文件是否存在
    if not os.path.exists(src_path):
        print(f"错误：文件不存在 - {src_path}")
        pythoncom.CoUninitialize()  # 新增：反初始化
        return None

    excel = None
    wb = None
    try:
        # 修改点1：使用Dispatch替代gencache
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False  # 新增：禁用警告弹窗

        # 打开工作簿
        wb = excel.Workbooks.Open(src_path)

        # 构造新路径（将.xls替换为.xlsx）
        base, ext = os.path.splitext(src_path)
        new_path = f"{base}.xlsx"

        # 改进点：更安全的文件存在检查
        if os.path.exists(new_path):
            # 新增：可选择修改为时间戳后缀
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            new_path = f"{base}_{timestamp}.xlsx"
            print(f"提示：目标文件已存在，已自动添加时间戳后缀")

        # 保存为新格式
        wb.SaveAs(new_path, FileFormat=51)  # 51 = xlsx
        print(f"转换成功：{new_path}")
        return new_path

    except Exception as e:
        print(f"转换失败：{str(e)}")
        return None
    finally:
        # 改进点：增强资源释放
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Application.Quit()
                # 新增：强制释放COM对象
                del wb
                del excel
        except:
            pass

        # 新增：彻底终止Excel进程
        os.system('taskkill /IM EXCEL.EXE /F 2>NUL')
        pythoncom.CoUninitialize()  # 必须的COM库反初始化

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

def rename_excel_objects(
        root_path,
        keyword_file,
        keyword_sheet,
        new_filename,
        new_sheetname
):
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False

    try:
        for foldername, _, filenames in os.walk(root_path):
            for filename in filenames:
                if not filename.lower().endswith('.xlsx'):
                    continue

                original_path = os.path.join(foldername, filename)

                # ===== 处理工作表 =====
                if keyword_sheet:  # 新增条件判断
                    wb = excel.Workbooks.Open(original_path)
                    sheets = wb.Sheets
                    sheets_to_rename = [
                        sheet for sheet in sheets
                        if keyword_sheet in sheet.Name
                    ]

                    for sheet in sheets_to_rename:
                        # ...（同前文的唯一命名逻辑）
                        sheet.Name = new_sheetname

                    if sheets_to_rename:
                        wb.SaveAs(original_path, ConflictResolution=2)
                        print(f"重命名工作表：{original_path} → {new_sheetname}")

                # ===== 处理文件名 =====
                if keyword_file and keyword_file in filename:
                    base, ext = os.path.splitext(filename)
                    new_base = new_filename.strip()
                    new_name = f"{new_base}{ext}"
                    new_path = os.path.join(foldername, new_name)

                    try:
                        # 强制覆盖已有文件
                        if os.path.exists(new_path):
                            os.remove(new_path)  # 先删除已存在的文件
                        os.rename(original_path, new_path)
                        print(f"【文件重命名（覆盖）】{original_path} → {new_path}")
                    except PermissionError as e:
                        print(f"文件被占用无法覆盖：{original_path}\n\t错误类型：{type(e).__name__}")
                    except Exception as e:
                        print(f"文件重命名失败：{original_path}\n\t错误类型：{type(e).__name__}")

    finally:
        if excel:
            excel.Quit()

    # 转换为绝对路径并规范化
    src_path = os.path.abspath(root_path)
    src_path = os.path.normpath(root_path)

    # 检查文件是否存在
    if not os.path.exists(src_path):
        print(f"错误：文件不存在 - {src_path}")
        return None

    # 初始化Excel应用
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    try:
        # 打开工作簿
        wb = excel.Workbooks.Open(src_path)

        # 构造新路径（将.xls替换为.xlsx）
        base, ext = os.path.splitext(src_path)
        new_path = f"{base}.xlsx"

        # 避免覆盖已有文件（可选）
        if os.path.exists(new_path):
            os.remove(new_path)

        # 保存为新格式
        wb.SaveAs(new_path, FileFormat=51)  # 51 = xlsx
        wb.Close()
        print(f"转换成功：{new_path}")
        return new_path

    except Exception as e:
        print(f"全局错误：{str(e)}")
    finally:
        if excel:
            excel.Quit()
            del excel
        pythoncom.CoUninitialize()  # 反初始化
        os.system('taskkill /IM EXCEL.EXE /F 2>NUL')


def rename_files_and_sheets(root_path, keyword_file, new_filename, keyword_sheet, new_sheetname):
    if not new_filename.endswith(".xlsx"):
        new_filename += ".xlsx"

    for filename in os.listdir(root_path):
        if filename.endswith(".xlsx") and keyword_file in filename:
            if filename == new_filename:
                continue

            src_path = os.path.abspath(os.path.join(root_path, filename))
            dst_path = os.path.abspath(os.path.join(root_path, new_filename))

            try:
                if os.path.exists(dst_path):
                    os.remove(dst_path)
                os.rename(src_path, dst_path)
            except Exception as e:
                print(f"重命名失败: {e}")
                continue

            wb = None
            try:
                # 加载工作簿并禁用样式警告
                wb = load_workbook(dst_path, data_only=True)
                modified = False

                # 处理同名表
                if new_sheetname in wb.sheetnames:
                    del wb[new_sheetname]

                # 筛选需重命名的表
                sheets_to_rename = [s for s in wb.sheetnames if keyword_sheet in s]
                for sheet_name in sheets_to_rename:
                    wb[sheet_name].title = new_sheetname
                    modified = True

                if modified:
                    wb.save(dst_path)
            except Exception as e:
                print(f"工作表修改失败: {e}")
            finally:
                if wb:
                    wb.close()  # 确保释放资源

def merge_excel_sheets(root_path,  # 要扫描的根目录
                       file_keyword,  # 文件名关键词(模糊匹配)
                       sheet_keyword,  # 工作表名关键词(模糊匹配)
                       output_path,  # 合并结果保存路径
                       output_filename,  # 合并结果文件名(不含扩展名)
                       header_row=0):  # 数据起始行(默认第1行为表头)

    # 创建保存目录（如果不存在）
    os.makedirs(output_path, exist_ok=True)

    # 合并结果保存路径
    output_file = os.path.join(output_path, f"{output_filename}.xlsx")

    # 收集所有需要合并的数据
    all_data = []

    # 遍历目录树
    for foldername, _, filenames in os.walk(root_path):
        for filename in filenames:
            # 仅处理xlsx文件且文件名包含关键词
            if filename.lower().endswith('.xlsx') and (file_keyword in filename):
                filepath = os.path.join(foldername, filename)

                try:
                    # 使用ExcelFile对象提高性能
                    with pd.ExcelFile(filepath, engine='openpyxl') as xls:
                        # 遍历所有工作表
                        for sheet_name in xls.sheet_names:
                            # 匹配包含关键词的工作表
                            if sheet_keyword in sheet_name:
                                # 读取数据
                                df = pd.read_excel(xls,
                                                   sheet_name=sheet_name,
                                                   header=header_row)

                                # 添加来源信息
                                df['来源文件'] = filename
                                df['来源工作表'] = sheet_name

                                all_data.append(df)

                                print(f"已加载：{filename} -> {sheet_name} ({len(df)}行)")
                except Exception as e:
                    print(f"错误：无法读取文件 {filename}\n错误信息：{str(e)}")
                    continue

    # 如果没有找到数据
    if not all_data:
        print("警告：未找到符合条件的工作表！")
        return

    # 合并数据
    merged_df = pd.concat(all_data, ignore_index=True)

    # 保存结果
    try:
        # 创建新工作簿
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            merged_df.to_excel(writer,
                               sheet_name='作业明细表',
                               index=False,
                               header=True)

            # 设置自适应列宽
            worksheet = writer.sheets['作业明细表']
            for column_cells in worksheet.columns:
                max_length = 0
                column = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                worksheet.column_dimensions[column].width = adjusted_width

        print(f"\n合并完成！共合并 {len(all_data)} 个工作表")
        print(f"总数据量：{len(merged_df)} 行")
        print(f"保存路径：{output_file}")

    except PermissionError:
        print("错误：保存失败，请关闭目标文件或检查写入权限！")

def copy_excel_files(src_path, dst_path):
    """
    复制源路径根目录下的Excel文件到目标路径（不处理子目录）

    参数:
    src_path (str): 源文件夹路径
    dst_path (str): 目标文件夹路径

    返回:
    list: 成功复制的文件路径列表
    """
    copied_files = []

    try:
        # 创建目标目录（自动处理目录不存在的情况）
        Path(dst_path).mkdir(parents=True, exist_ok=True)

        # 遍历源目录根层文件
        for filename in os.listdir(src_path):
            src_file = os.path.join(src_path, filename)

            # 仅处理文件且后缀匹配
            if os.path.isfile(src_file) and filename.lower().endswith(('.xls', '.xlsx')):
                dst_file = os.path.join(dst_path, filename)

                # 执行复制并保留元数据
                shutil.copy2(src_file, dst_file)
                copied_files.append(dst_file)
                print(f"成功复制: {src_file} -> {dst_file}")

        return copied_files

    except FileNotFoundError as e:
        print(f"路径不存在错误: {e}")
    except PermissionError as e:
        print(f"权限不足: {e}")
    except shutil.Error as e:
        print(f"复制错误: {e}")
    except Exception as e:
        print(f"未知错误: {e}")

    return []

def save_to_excel_sheet(
        df: pd.DataFrame,
        original_path: str,
        new_sheet_name: str,
        if_sheet_exists: str = "replace"
) -> bool:
    """
    将DataFrame保存到指定Excel文件的新工作表中（支持覆盖同名工作表）

    参数：
    df : 要保存的DataFrame数据
    original_path : 目标Excel文件路径（必须为.xlsx格式）
    new_sheet_name : 新工作表的名称
    if_sheet_exists : 同名工作表处理策略（replace/new/overlay）

    返回：
    bool : 操作是否成功
    """

    try:
        with pd.ExcelWriter(
                original_path,
                engine="openpyxl",
                mode="a",  # 追加模式
                if_sheet_exists="replace"  # 覆盖同名工作表
        ) as writer:
            df.to_excel(writer, sheet_name=new_sheet_name, index=False)

        print(f"✅ 成功保存到工作表: {new_sheet_name}")
        return True

    except FileNotFoundError:
        print(f"❌ 文件不存在: {original_path}")
        return False
    except PermissionError:
        print("⛔ 文件被占用，请关闭Excel后重试")
        return False
    except ValueError as ve:
        print(f"⚠️ 参数错误: {str(ve)}")
        return False
    except Exception as e:
        print(f"❌ 未知错误: {str(e)}")
        return False

if __name__ == '__main__':

    # ====================================================
    path = copy_files_to_dated_folder(src_path=r"env\demo", dest_root=r"./")

    # 文件格式转换
    folder = path
    for file in os.listdir(folder):
        if file.endswith(".xls"):
            convert_xls_to_xlsx(os.path.join(folder, file))
    
    rename_files_and_sheets(
        root_path=path,
        keyword_file="风险管控",
        new_filename="风险管控到位列表",
        keyword_sheet="风险管控",
        new_sheetname="风险管控到位情况"
    )
    
    merge_excel_sheets(
        root_path=path,  # 要扫描的根目录
        file_keyword="作业统计",  # 文件名需包含此关键词
        sheet_keyword="作业",  # 工作表名需包含此关键词
        output_path=path,  # 保存目录
        output_filename="作业统计详情",  # 合并文件名(自动加.xlsx)
        header_row=0  # 表头所在行号(从0开始计数)
    )
    
    # ====================================================
    df_0 = pd.read_excel(path + r"/风险管控到位列表.xlsx", sheet_name="风险管控到位情况")
    
    # 获取总行数和列数
    total_rows = df_0.shape[0]  # 总行数
    total_cols = df_0.shape[1]  # 总列数
    print(f"总行数：{total_rows}，总列数：{total_cols}")
    
    # ====================================================
    target_header = "所属项目部"
    temp = 0
    match = ['设备监造总监项目部', '棠下柔直监理项目部']
    if target_header in df_0.columns:
        col_index = df_0.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_0.iat[i, temp] in match:
                df_0 = df_0.drop(index=df_0.index[i])
        except IndexError:
            continue
    
    # ================================================================
    # 单Sheet保存
    df_0.to_excel(path + r"/风险管控到位列表(去设备监造-).xlsx", sheet_name="Sheet1", index=False)
    
    # ====================================================
    temp_df = df_0
    
    target_header = "项目类型"
    match = ["主网项目"]
    temp = 0
    if target_header in temp_df.columns:
        col_index = temp_df.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if temp_df.iat[i, temp] not in match:
                temp_df = temp_df.drop(index=temp_df.index[i])
        except IndexError:
            continue
    
    # ====================================================
    
    save_to_excel_sheet(df=temp_df, original_path=path + r"/风险管控到位列表.xlsx", new_sheet_name="sheet1",
                        if_sheet_exists="replace")
    temp_df = None
    
    # ====================================================
    target_header = "工程类型"
    match = ["配网搬迁", "配网基建", "线路搬迁", "小型基建", "主网基建", "主网项目(新)"]
    temp = 0
    if target_header in df_0.columns:
        col_index = df_0.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_0.iat[i, temp] in match:
                df_0 = df_0.drop(index=df_0.index[i])
        except IndexError:
            continue
    
    # ====================================================
    
    save_to_excel_sheet(df=df_0, original_path=path + r"/风险管控到位列表.xlsx", new_sheet_name="去基建",
                        if_sheet_exists="replace")
    
    # ====================================================
    target_header = "项目名称"
    temp = 0
    if target_header in df_0.columns:
        col_index = df_0.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    record = []
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_0.iat[i, temp] in record:
                df_0 = df_0.drop(index=df_0.index[i])
            else:
                record.append(df_0.iat[i, temp])
        except IndexError:
            continue
    
    # ================================================================
    # 单Sheet保存
    df_0.to_excel(path + r"/风险管控到位列表(去设备监造-去基建-去重-).xlsx", sheet_name="Sheet1", index=False)
    
    # 多Sheet保存
    # df_0.to_excel(r"Z-Temp/风险管控到位列表.xlsx", sheet_name="去基建", index=False)
    
    # ====================================================
    target_header = "风险等级"
    temp = 0
    if target_header in df_0.columns:
        col_index = df_0.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_0.iat[i, temp] == '低风险' or df_0.iat[i, temp] == '可接受风险':
                df_0 = df_0.drop(index=df_0.index[i])
        except IndexError:
            continue
    
    # ================================================================
    # 单Sheet保存
    df_0.to_excel(path + r"/风险管控到位列表(去设备监造-去基建-去重-统计-).xlsx", sheet_name="Sheet1", index=False)
      
    # ====================================================
    df_1 = pd.read_excel(path + r"/作业统计详情.xlsx", sheet_name="作业明细表")
    
    # ====================================================
    # 加载工作簿（支持.xlsx格式）
    wb = xl.load_workbook(path + r'/监理人员整体应用数据列表.xlsx')  #
    # 选择工作表（通过名称或活动表）
    ws = wb['监理人员整体应用数据']  # 或 ws = wb.active
    
    max_row = ws.max_row  # 最大行数
    max_col = ws.max_column  # 最大列数
    
    print(max_col, max_row)
    
    # ====================================================
    # 获取总行数和列数
    total_rows = df_1.shape[0]  # 总行数
    total_cols = df_1.shape[1]  # 总列数
    print(f"总行数：{total_rows}，总列数：{total_cols}")
    
    # ====================================================
    target_header = "监理单位名称"
    temp = 0
    if target_header in df_1.columns:
        col_index = df_1.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_1.iat[i, temp] != '广州电力工程监理有限公司':
                df_1 = df_1.drop(index=df_1.index[i])
        except IndexError:
            continue
    
    # ====================================================
    new_header = "总监项目部"
    target_header = "监理人员姓名"
    
    # 示例：检索指定行列单元格坐标
    match_x = search_specific_row(ws, row_num=1, target_value="姓名")
    print(match_x)
    
    # 检查目标列是否存在
    if target_header in df_1.columns:
        # 3. 获取目标列的索引位置
        target_index = df_1.columns.get_loc(target_header)
    
        # 4. 在目标列的右侧插入新列（初始值为空）
        # 使用pd.NA表示空值（适用于Pandas 0.24+）
        df_1.insert(target_index + 1, new_header, pd.NA)
        # 或者使用None（兼容旧版本）
        # df_1.insert(target_index + 1, new_header, None)
        print(f"成功在列 '{target_header}' 右侧插入新列 '{new_header}'！")
    
        # 提取目标列数据（两种方式任选）
        # target_column = df_1.iloc[:, target_index]  # 按索引提取列
        target_column = df_1[target_header]  # 直接按列名提取
    
        # 遍历处理每个值
        # results = []
        for idx, value in target_column.items():  # 带索引遍历
            # 示例操作：打印值并收集非空结果
            print(f"行 {idx + 2}（第{idx + 2}行）: {value}")
            # if pd.notna(value):
            #     results.append(value)
    
            match_y = search_specific_column(ws, col_num=match_x, target_value=value)
            print(match_y)
    
            # 对target_column右侧赋值
            # 根据当前行的索引idx，将match_y值写入右侧新列
            df_1.at[idx, new_header] = ws.cell(row=match_y, column=3).value  # 根据实际返回结构调整
    
        print(f"成功对列 '{target_column}' 右侧的列 进行赋值！")
    
    else:
        print(f"警告：目标列 '{target_header}' 不存在于工作表中！")
    
    # ================================================================
    # 单Sheet保存
    df_1.to_excel(path + r"/作业统计详情(项目部匹配-).xlsx", sheet_name="Sheet1", index=False)
    
    # ====================================================
    
    save_to_excel_sheet(df=df_1, original_path=path + r"/作业统计详情.xlsx", new_sheet_name="作业明细表",
                        if_sheet_exists="replace")
    
    # ====================================================
    target_header = "总监项目部"
    temp = 0
    match = ['设备监造总监项目部', '棠下柔直监理项目部']
    if target_header in df_1.columns:
        col_index = df_1.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if df_1.iat[i, temp] in match:
                df_1 = df_1.drop(index=df_1.index[i])
        except IndexError:
            continue
    
    # ====================================================
    
    save_to_excel_sheet(df=df_1, original_path=path + r"/作业统计详情.xlsx", new_sheet_name="Sheet1",
                        if_sheet_exists="replace")
    
    # ====================================================
    
    temp_df_1 = df_1
    temp_df_2 = df_1
    
    # ====================================================
    target_header = "项目名称"
    temp = 0
    if target_header in temp_df_1.columns:
        col_index = temp_df_1.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    record = []
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if temp_df_1.iat[i, temp] in record:
                temp_df_1 = temp_df_1.drop(index=temp_df_1.index[i])
            else:
                record.append(temp_df_1.iat[i, temp])
        except IndexError:
            continue
    
    # ================================================================
    # 单Sheet保存
    temp_df_1.to_excel(path + r"/作业统计详情(项目部匹配-去重1-).xlsx", sheet_name="Sheet1", index=False)
    
    # ====================================================
    
    save_to_excel_sheet(df=temp_df_1, original_path=path + r"/作业统计详情.xlsx", new_sheet_name="去重",
                        if_sheet_exists="replace")
    
    # ====================================================
    target_header = "风险等级"
    temp = 0
    if target_header in temp_df_2.columns:
        col_index = temp_df_2.columns.get_loc(target_header)  # 获取列索引
        temp = col_index
        print(temp)
    else:
        raise ValueError(f"未找到表头 '{target_header}'")
    
    for i in range(total_rows - 1, -1, -1):  # 从最后一行倒序检查
        try:
            if temp_df_2.iat[i, temp] == '低':
                temp_df_2 = temp_df_2.drop(index=temp_df_2.index[i])
        except IndexError:
            continue
    
    # ================================================================
    # 单Sheet保存
    temp_df_2.to_excel(path + r"/作业统计详情(项目部匹配-去低风险-).xlsx", sheet_name="Sheet1", index=False)
    
    # ====================================================
    
    save_to_excel_sheet(df=temp_df_2, original_path=path + r"/作业统计详情.xlsx", new_sheet_name="去重 (2)",
                        if_sheet_exists="replace")
