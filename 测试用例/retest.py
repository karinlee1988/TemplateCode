#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/8/16 23:27
# @Author : karinlee
# @FileName : retest.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976


import re
#
# pattern = re.compile(r'0?(13|14|15|17|18|19)[0-9]{9}')
res = r'0?(13|14|15|17|18|19)[0-9]{9}'
s = u"sadasdasd13376670002afdasfdrgfhdf"
#
# res2 = re.search(pattern,s)
# print(res2.group())

class RegularExpression(object):

    def __init__(self,content,re_str):
        self.content = content
        self.re_str = re_str
        self.pattern = re.compile(re_str)

    def search(self):

        result = re.search(self.pattern,self.content)
        try:
            return result.group()
        except AttributeError:
            return 0

    def match(self):
        result = re.match(self.pattern,self.content)
        try:
            return result.group()
        except AttributeError:
            return 0

if __name__ == '__main__':
    r = RegularExpression(s,res)
    r2 = r.search()
    print(r2)