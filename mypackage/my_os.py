#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/2/11 22:54
# @Author : karinlee
# @FileName : my_os.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

"""
本模块包含相关自建函数，用于处理文档相关操作。
"""

import os
import csv

def get_filename(folder_path):
    """
    获取文件夹下面所有文件，并返回文件夹里面所有文件名列表（包含子文件夹里面的文件）
    :param folder_path: 文件夹路径
    :type  folder_path: str

    :return: 文件夹里面所有全文件名列表（包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            filename_list.append(os.path.join(root, file))
    return filename_list

def get_design_filename(folder_path,filetype):
    """
    获取文件夹下面所有文件，并返回文件夹里面指定类型的文件名列表（包含子文件夹里面的文件）
    :param folder_path: 文件夹路径
    :type  folder_path: str

    :param filetype: 指定类型文件的后缀（如.xlsx）
    :type  filetype: str

    :return: 文件夹里面所有全文件名列表（包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if os.path.splitext(file)[1] == filetype:
                filename_list.append(os.path.join(root, file))
    return filename_list

def get_singlefolder_filename(folder_path,filetype):
    """
    获取待处理文件夹里指定后缀的文件名（单个文件夹，不包括子文件夹的文件）
    :param folder_path: 文件夹路径
    :type  folder_path: str

    :param filetype: 指定类型文件的后缀（如.xlsx）
    :type  filetype: str

    :return: 文件夹里面所有全文件名列表（包含子文件夹里面的文件）
    :rtype:  list
    """
    filename_list = []
    files = os.listdir(folder_path)
    for file in files:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(file)[1] == filetype:
            filename_list.append(folder_path+file)
    return filename_list

def record_csv(content_list,csv_filename):
    """
    将content_list列表的内容按行追加写入csv文件中
    csv文件编码：utf-8

    :param content_list: 要写入的内容
    :type  list

    :param csv_filename: 要写入的csv文件名
    :type  str

    :return: None
    :rtype: None
    """
    # 创建文件对象
    with open(csv_filename,'a',newline='',encoding='utf-8') as csvfile:
        # 基于文件对象构建 csv写入对象
        csv_writer = csv.writer(csvfile)
        # 写入文件
        csv_writer.writerow(content_list)

def record_txt(content,txt_filename):
    """
    将content字符串内容按行追加写入txt文件中
    txt文件编码：utf-8

    :param content: 要写入txt文件的字符串
    :type content:  str

    :param txt_filename: txt文件全文件名
    :type txt_filename:  str

    :return: None
    :rtype: None
    """
    with open (txt_filename,"a",encoding='utf-8') as f:
        f.write(content+'\n')
