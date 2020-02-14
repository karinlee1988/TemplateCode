#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/2/12 15:08
# @Author : karinlee
# @FileName : my_data_processing.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

"""
本模块用于处理各种类型的数据
"""

class IdCardNumber(object):

    def __init__(self,id_card_number):
        self.id_card_number = id_card_number

    def get_birth(self):

        number_length = len(self.id_card_number)
        if number_length == 18:
            birth = self.id_card_number[6:14]
            return birth
        elif number_length == 15:
            birth = '19' + self.id_card_number[6:12]  #目前的15位身份证当成19xx年出生
            return birth
        else:
            return False
