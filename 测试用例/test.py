#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/2/12 11:11
# @Author : karinlee
# @FileName : test.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

"""
测试用例
"""

from mypackage.my_excel import *
from mypackage.my_os import *

def test_myos_getfilename():
    filename = get_filename("D:\\我的坚果云\\学习文档\\python小工具\\测试用例\\")
    return filename

def test_myos_get_design_filename():
    filename = get_design_filename("D:\\我的坚果云\\学习文档\\python小工具\\测试用例\\",'.bmp')
    return filename

def test_get_singlefolder_filename():
    filename = get_singlefolder_filename("D:\\我的坚果云\\学习文档\\python小工具\\测试用例\\测试文件夹1\\",'.docx')
    return filename


def test_record_csv():
    conlist = ["1",2,3,"eng","你好"]
    record_csv(conlist,"D:\\我的坚果云\\学习文档\\python小工具\\测试用例\\testcsv.csv")

def test_record_txt():
    strings = "hello，python"
    record_txt(strings,"D:\\我的坚果云\\学习文档\\python小工具\\测试用例\\testtxt.txt")


def test_get_xlsx_full_filename():
    filelist = get_xlsx_full_filename("D:\\MyNutstore\\PersonalStudy\\python_tools\\测试用例")
    return filelist

def test_get_xlsx_filename():
    filelist = get_xlsx_filename("D:\\MyNutstore\\PersonalStudy\\python_tools\\测试用例")
    return filelist


if __name__ == '__main__':
    print(test_get_xlsx_filename())