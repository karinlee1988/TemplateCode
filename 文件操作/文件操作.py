#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2023/4/15 22:24 
# @Author : karinlee
# @FileName : 文件操作.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976
# @github : https://github.com/karinlee1988/
# @Personal website : https://karinlee.cn/

import os
import csv


def get_allfolder_fullfilename(folder_path: str, filetype: list) -> list:
    """
    获取待处理文件夹下面所有文件，并返回文件夹里面指定类型的文件名列表（包含子文件夹里面的文件）

    20230429 test OK

    :param folder_path: 文件夹路径
    :type  folder_path: str

    :param filetype: 指定一种或多种类型文件的后缀列表（如['.xlsx']或['.xlsx','.xls']）
    :type  filetype: list

    :return: 文件夹里面所有全文件名列表（包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if os.path.splitext(file)[1] in filetype:
                filename_list.append(os.path.join(root, file))
    return filename_list


def get_singlefolder_fullfilename(folder_path: str, filetype: list) -> list:
    """
    获取待处理文件夹里指定后缀的文件名（单个文件夹，不包括子文件夹的文件）

    20230429 test OK

    :param folder_path: 文件夹路径
    :type  folder_path: str

    :param filetype: 指定一种或多种类型文件的后缀列表（如['.xlsx']或['.xlsx','.xls']）
    :type  filetype: list

    :return: 文件夹里面所有全文件名列表（不包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    files = os.listdir(folder_path)
    for file in files:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file)[1] in filetype:
            filename_list.append(folder_path + '\\' + file)
    return filename_list


def get_normal_filename(fullfilename: str) -> str:
    """
    从全文件名（包含绝对路径的文件名）转换为普通文件名
    如： "K:/Project/FilterDriver/DriverCodes/hello.txt" 通过转换变成"hello.txt"

    20230429 test OK

    :param fullfilename: 要写入的内容
    :type  fullfilename: list

    """
    return os.path.basename(fullfilename)


def record_csv(content_list: list, csv_filename: str) -> None:
    """
    将content_list列表的内容按行 追加 写入csv文件中(开始时默认保留第一行标题行，从第二行开始写入)
    一般为一维列表 ，如['a','b','c','d','e','f']，将上述abcdef字母填进csv中的ABCDEF列中

    如果用了二维列表，如[['a','b','c'],['d','e','f']]，则对应行A列单元格内容为字符串['a','b','c']，
    B列单元格内容为字符串['d','e','f']

    csv文件编码：utf-8

    20230429 test OK

    :param content_list: 要写入的内容列表（一般为一维列表）
    :type  content_list: list

    :param csv_filename: 要写入的csv文件名
    :type  csv_filename: str

    :return: None
    """
    # 创建文件对象
    with open(csv_filename, 'a', newline='', encoding='utf-8') as csvfile:
        # 基于文件对象构建 csv写入对象
        csv_writer = csv.writer(csvfile)
        # 写入文件
        csv_writer.writerow(content_list)


def record_txt(content: str, txt_filename: str) -> None:
    """
    将content字符串内容按行追加写入txt文件中
    txt文件编码：utf-8

    20230429 test OK

    :param content: 要写入txt文件的字符串
    :type content:  str

    :param txt_filename: txt文件全文件名
    :type txt_filename:  str

    :return: None
    :rtype: None
    """
    with open(txt_filename, "a", encoding='utf-8') as f:
        f.write(content + '\n')


if __name__ == '__main__':
    # 测试1 get_allfolder_fullfilename()
    # test1 = get_allfolder_fullfilename('tests\\', ['.xlsx', '.docx'])
    # print(test1)
    # 测试2 get_singlefolder_fullfilename()
    # test2 = get_singlefolder_fullfilename('tests\\folder1', ['.xlsx', '.docx', '.txt'])
    # print(test2)
    # 测试3 get_normal_filename()
    # test3 = get_normal_filename("K:/Project/FilterDriver/DriverCodes/hello.txt")
    # print(test3)
    # 测试4 record_csv()
    # list_1d = ['a','b','c','d','e','f']
    # list_2d = [['a','b','c'],['d','e','f']]
    # record_csv(list_1d,'tests\\rec.csv')
    #测试5 record_txt()
    # str_ = 'abcdefg'
    # record_txt(str_,'tests\\rec.txt')
    pass