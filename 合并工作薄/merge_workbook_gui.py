#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2020/3/13 9:52
# @Author  : karinlee
# @FileName: merge_workbook_gui.py
# @Software: PyCharm
# @Blog    ：https://blog.csdn.net/weixin_43972976
 
import os
import xlrd
import pandas as pd
import tkinter as tk
import tkinter.filedialog

class MergeWorkbook(object):
    """
    用于合并工作薄
    """
    def __init__(self):
        """
        创建界面
        """
        # 新建窗口
        self.master = tk.Tk()
        # 在界面顶部添加横幅图片
        self.photo = tk.PhotoImage(file="img\\doge4.gif")
        # self.path 用于存放选择的文件路径
        # self.flag 当程序运行完成后给用户提示信息
        # self.v1 self.v2接收用户参数 待选用，不够可以加
        # 注意！这些都是tk.StringVar()对象，不是str，其他地方要用的话要用get()方法获取str
        self.path = tk.StringVar()
        self.flag = tk.StringVar()
        self.sheet_index = tk.StringVar()
        self.title_row = tk.StringVar()
        # 窗口图片横幅
        self.img_pack()
        # 主界面布局
        self.window()
        # 框架布局
        self.frame()
        # 保持运行
        self.master.mainloop()

    def img_pack(self):
        """
        窗口图片横幅
        :return:
        :rtype:
        """
        # 图片贴上去
        img_lable = tk.Label(self.master, image=self.photo)
        img_lable.pack()

    def window(self):
        """
        主界面布局设置
        :return:
        :rtype:
        """
        # 调整窗口默认大小及在屏幕上的位置
        self.master.geometry("800x650+600+100")
        # 窗口的标题栏，自己修改
        self.master.title("合并工作薄by李加林v1.0")
        # 把标题贴上去，自己修改
        tk.Label(self.master,text="合并工作薄v1.0",font=("黑体",20)).pack()
        tk.Label(self.master,text="------注意！待合并的工作薄必须是.xlsx格式------",font=("黑体",16)).pack()


    def frame(self):
        """
        框架布局设置
        :return:
        :rtype:
        """
        # 框架贴上去，再在框架里添加Lable，Entry，Button等控件
        frame1 = tk.Frame(self.master)
        frame1.pack()
        # 输入框，标记，按键
        tk.Label(frame1, text="工作薄所在文件夹路径:", font=("黑体", 16)).grid(row=1, column=0)
        tk.Entry(frame1, textvariable=self.path, width=50).grid(row=1, column=1)
        tk.Button(frame1, text="选择文件夹", command=self.select_path, font=("黑体", 16)).grid(row=1, column=2)
        tk.Label(frame1, text="工作表序号：", font=("黑体", 16)).grid(row=2, column=0)
        tk.Entry(frame1, textvariable=self.sheet_index).grid(row=2, column=1)
        tk.Label(frame1, text="表头行数:", font=("黑体", 16)).grid(row=3, column=0)
        tk.Entry(frame1, textvariable=self.title_row).grid(row=3, column=1)
        # 按这个按钮执行主程序
        tk.Button(frame1, text="开始处理", command=self.main, font=("黑体", 16)).grid(row=4, column=1,pady=66)
        tk.Entry(frame1, textvariable=self.flag,state="readonly").grid(row=4, column=2)

    def select_path(self):
        """
        选择文件，并获取文件的绝对路径
        :return:
        :rtype:
        """
        # 选择文件，path_select变量接收文件地址
        # 注意：self.path 是tk.StringVar()对象，而path_select是str变量
        path_select = tkinter.filedialog.askdirectory()
        # 通过replace函数替换绝对文件地址中的/来使文件可被程序读取
        # 注意：\\转义后为\，所以\\\\转义后为\\
        path_select = path_select.replace("/", "\\\\")
        # self.path设置path_select的值
        self.path.set(path_select)

    def xlsx_filename_list(self,file_dir):
        """
        构造获取所有需要合并的工作簿路径并生成路径列表的函数
        :param file_dir: 要合并的工作薄所在的文件夹路径
        :type file_dir:  str
        :return: 文件夹下所有.xlsx文件名列表
        :rtype:   list
        """
        xlsx_list = []  # 构造一个用于存放文件名(包括扩展名)的空列表
        for file in os.listdir(file_dir):  # 遍历文件夹file_dir下的所有文件
            if os.path.splitext(file)[1] == '.xlsx':  # 筛选出扩展名是.xlsx的所有文件
                xlsx_list.append(file)  # 将文件扩展名是.xlsx的所有文件的文件名存放到列表list中
        return xlsx_list



    def merge_workbook(self,path,sheet_index=1,title_row=1):
        """
        合并工作薄
        :param path: 待合并工作薄文件夹
        :type path: str
        :param title_row: 表头行数（从1开始）
        :type title_row: int
        :return: None
        :rtype:  None
        """
        title = None  # 先定义title变量避免pycharm提示错误
        sheet_index = int(sheet_index)
        title_row = int(title_row)
        wks = self.xlsx_filename_list(path)  # 通过file_name函数获取path路径下所有xlsx文件的文件名
        data = []  # 定义一个空列表对象
        for i in range(len(wks)):
            read_xlsx = xlrd.open_workbook(path + '\\' + wks[i])  # 根据path和文件名合并每个待合并工作簿的路径
            sheet1 = read_xlsx.sheets()[sheet_index-1]  # 找到工作簿中的第一个工作表  ## 序号都是从0开始，注意！
            nrow = sheet1.nrows  # 提取出第一个工作表中的数据行数
            title = sheet1.row_values(title_row-1)  # 提取出第一个工作表中的表头(第1行） ## 序号都是从0开始，注意！
            for j in range(title_row, nrow):  # 逐行将工作表中的数据添加到空列表data中 数据从第2行开始。## 序号都是从0开始，注意！
                data.append(sheet1.row_values(j))
        # 将列表data转化为DataFrame对象
        content = pd.DataFrame(data)
        # 修改DataFrame对象的标题
        content.columns = title
        # 将DataFrame对象content写入新的Excel工作簿中
        content.to_excel(path + '\\合并完成.xlsx', header=True, index=False)


    def main(self):
        """
        这个是主程序，最好将非gui版的程序封装成类，接收该GUI类的参数后实例化运行
        :return:
        :rtype:
        """
        self.merge_workbook(path=self.path.get(),sheet_index=self.sheet_index.get(),title_row=self.title_row.get())
        # 标志设置为处理完成
        self.flag.set("处理完成！")

if __name__ == '__main__':
    running = MergeWorkbook()

