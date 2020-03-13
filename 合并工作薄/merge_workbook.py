#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2020/3/13 9:35
# @Author  : karinlee
# @FileName: merge_workbook.py
# @Software: PyCharm
# @Blog    ：https://blog.csdn.net/weixin_43972976

import os
import xlrd
import pandas as pd

def xlsx_filename_list(file_dir):
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


def merge_workbook(path,sheet_index=1,title_row=1):
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
    wks = xlsx_filename_list(path)  # 通过file_name函数获取path路径下所有xlsx文件的文件名
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

if __name__ == '__main__':
    path = "D:\我的坚果云\个人学习文档\python小工具\合并工作薄"  # 待输入合并工作簿总的路径
    path = path.replace("\\", "\\\\")
    merge_workbook(path)
    print('合并完成！')


