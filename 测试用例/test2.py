#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/4/25 12:22
# @Author : karinlee
# @FileName : test2.py.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976


# def getdict(workbook):
#     pass
#
#
# import os
# f_l = ['111111.xlsx','222222.xlsx']
# for f in f_l:
#     os.startfile(f,'print')


class TestObj1():

    def __init__(self):
        self.a = 1
        self.b = 2

    def Math1(self):
        return self.a + self.b


class TestObj2(TestObj1):

    def __init__(self):
        super(TestObj2, self).__init__()
        # self.a = 3
        # self.b = 5

    def Math2(self):
        return self.a - self.b

t1 = TestObj1()

t2 = TestObj2()
print(t1.Math1())
print(t2.Math2())