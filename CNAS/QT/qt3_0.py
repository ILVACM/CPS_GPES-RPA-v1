import sys
import json
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from qt3 import Ui_Dialog  # 你的UI文件生成的类


class JsonTableDialog(QDialog, Ui_Dialog):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        # self.json_file = 'config/temp.json'
        self.json_file = r'Z:\CPS-OA_RPA-20250312\config\temp.json'
        self.operate_list = []

        # 初始化表格设置
        self.tableWidget.setColumnCount(4)
        self.tableWidget.setHorizontalHeaderLabels(['Mode', 'Code', 'Coord', 'Value'])
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)

        # 加载数据
        self.load_data()

        # 连接关闭事件
        self.buttonBox.accepted.connect(self.save_data)
        self.buttonBox.rejected.connect(self.close)

    def load_data(self):
        """加载JSON数据到表格"""
        try:
            with open(self.json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.operate_list = data.get('OperateList', [])
        except FileNotFoundError:
            self.operate_list = []

        # 清空表格
        self.tableWidget.setRowCount(0)

        for item in self.operate_list:
            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)

            # 提取前4个键，不足则补空字符串
            keys = list(item.keys())[:4]
            keys += [''] * (4 - len(keys))  # 确保keys长度为4

            for col in range(4):
                key = keys[col]
                value = str(item.get(key, ''))  # 直接使用key

                # 处理坐标列表的显示
                if key and isinstance(item.get(key, list), list):
                    value = ', '.join(map(str, item[key]))

                item_widget = QTableWidgetItem(value)
                self.tableWidget.setItem(row_position, col, item_widget)

        # 调整列宽
        self.tableWidget.resizeColumnsToContents()

    def save_data(self):
        """保存修改后的数据到JSON"""
        # 收集表格数据
        new_operate_list = []
        for row in range(self.tableWidget.rowCount()):
            row_data = {}

            # 获取当前行数据
            for col in range(4):
                item = self.tableWidget.item(row, col)
                key = self.tableWidget.horizontalHeaderItem(col).text().lower()  # 使用列标题作为键
                value = item.text().strip() if item else ''

                if key == 'coord':
                    # 尝试将坐标字符串转换为列表
                    try:
                        coords = list(map(int, value.split(',')))
                        row_data[key] = coords
                    except:
                        row_data[key] = value
                else:
                    row_data[key] = value

            new_operate_list.append(row_data)

        # 读取原有数据并更新
        try:
            with open(self.json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except:
            data = {'OperateList': []}

        # 追加保存（如果需要覆盖则直接用 new_operate_list）
        data['OperateList'].extend(new_operate_list)

        with open(self.json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)

        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    dialog = JsonTableDialog()
    dialog.show()
    sys.exit(app.exec_())