import tkinter as tk
import openpyxl


class ExcelChange(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 数据修改器")
        self.root.resizable(True, True)  # 自由缩放窗口大小
        self.root.geometry('1000x600+400+0')  # 首次打开屏幕定位

        # 单元格输入值列表：全局变量, 获取输入框的值并传递给表格.
        self.input_label_B3 = None
        self.input_label_C3 = None

    def ui_input(self):
        # 创建两个组块, 然后根据行列定位.
        frame_1 = tk.Frame(self.root, background="gray")  # gray 就是背景颜色, 可以自行更改.
        frame_2 = tk.Frame(self.root, pady=200)  # pady:y轴定位位置，x位置默认居中

        frame_1_name = tk.Label(frame_1, text="第一个布局块", background="gray")
        frame_1_name.grid(row=0, column=0)  # 第一行第一列(0 就是起始序列)
        frame_2_name = tk.Label(frame_2, text="第二个布局块")  # 修改文本就可以改变组块的名称
        frame_2_name.grid(row=0, column=0)  # 第一行第一列(0 就是起始序列)

        # B3单元格值输入, 方便后续获取.
        cell_label_B3 = tk.Label(frame_1, text="中间腹板轴孔中心距侧边长度L1")  # 修改文本改变的说明
        cell_label_B3.grid(row=1, column=1)  # 第二行第二列
        self.input_label_B3 = tk.Entry(frame_1, width=10)
        self.input_label_B3.grid(row=2, column=1)  # 第三行第二列

        # C3单元格值输入, 方便后续获取.
        cell_label_C3 = tk.Label(frame_2, text="中间腹板轴孔中心距侧边长度L2")
        cell_label_C3.grid(row=1, column=1)  # 第二行第二列
        self.input_label_C3 = tk.Entry(frame_2, width=10)
        self.input_label_C3.grid(row=2, column=1)  # 第三行第二列

        # 设置修改按钮, 并且关联到修改表格事件.
        button_next = tk.Button(self.root, text='确认修改', width=7, height=3, fg='black', command=self.change_excel)
        button_next.place(x=830, y=450)

        frame_1.pack()  # 组块激活
        frame_2.pack()

        self.root.mainloop()  # 创建Tk循环

    def change_excel(self):
        """
        改变表格内的数据, 批量修改, 有多少需要修改的单元格, 就写多少个 if.
        然后做批量修改, 如果输入框内的数据没有被清除, 表格内的数据会再次更新,
        只不过前后数值相同, 不影响最终效果.
        :return:
        """
        # B3 输入框内有值, 则 if 成立, 修改表格内容.
        if self.input_label_B3.get():
            self.change_cell(cell_name="B3", cell_value=self.input_label_B3.get())
            print("Done B3")  # 打印确认修改完成.
        if self.input_label_C3.get():
            self.change_cell(cell_name="C3", cell_value=self.input_label_C3.get())
            print("Done C3")

    def change_cell(self, cell_name=None, cell_value=None):
        """
        可以自主替换表格的名称, 相应的表格路径在 excel_editor.py 文件的根目录中.
        最终保存的表格会覆盖之前的表格, 相当于在一个表格内做修改.
        :param cell_name: 单元格的名称: A3 B3 C3...
        :param cell_value: 需要填入到单元格内的数据, 数据来自于 ui 界面输入框, 支持实时更新.
        :return:
        """
        wb = openpyxl.load_workbook("大平衡梁.xlsx")
        sheet = wb.active
        cell = sheet[f"{cell_name}"]
        cell.value = f"{cell_value}"
        wb.save(filename="大平衡梁.xlsx")


if __name__ == '__main__':
    ExcelChange().ui_input()
