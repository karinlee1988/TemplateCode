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

import re

class IdCardNumber(object):
    """
    用于对身份证号码进行处理
    """

    def __init__(self,id_card_number):
        """
        构造函数，传入身份证号码实例化
        :param id_card_number: 身份证号码（15位或者18位）
        :type id_card_number: str
        """
        self.id_card_number = id_card_number

    def get_birth(self):
        """
        通过身份证号码获取出生日期(8位字符串）
        :return birth:出生日期
        :rtype birth: str
        """
        number_length = len(self.id_card_number)
        if number_length == 18:
            birth = self.id_card_number[6:14]
            return birth
        elif number_length == 15:
            birth = '19' + self.id_card_number[6:12]  #目前的15位身份证是19xx年出生
            return birth
        else:
            return "身份证号码格式不正确"

    @staticmethod
    def get_checkcode(digital_ontology_code):
        """
        静态方法，从身份证号码前17位数字本体码计算第18位校验码
        :param digital_ontology_code:   17位数字本体码
        :type digital_ontology_code: str
        :return str(vi[remainder]): 18位身份证的最后1位校验码
        :rtype str(vi[remainder]): str
        """
        ai = []  # 17位数字本体码按位分割的列表,先创建列表，后面再赋值
        wi = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]  # 加权因子列表
        vi = [1, 0, 'X', 9, 8, 7, 6, 5, 4, 3, 2]   # 校验码列表
        remainder = None # 用于在校验码列表中找校验码的模 （先赋值为None，后面再计算赋值）
        if len(digital_ontology_code) == 17:
            s = 0
            for i in digital_ontology_code:
                ai.append(int(i))
            for i in range(17):
                s = s + wi[i] * ai[i]  # 计算和
            remainder = s % 11  # 计算模
            return  str(vi[remainder])  # 通过模的值在校验码列表中找到对应的校验码并返回str
        else:
            return False

    def fifteen_to_eighteen(self):
        """
        15位身份证转18位身份证

        15变18，就是出生年份由2位变为4位，最后加了一位用于验证。验证位的规则如下：
        1、将前面的身份证号码17位数分别乘以不同的系数。从第一位到第十七位的系数分别为:7. 9 .10 .5. 8. 4. 2. 1. 6. 3. 7. 9. 10. 5. 8. 4. 2.
        2、将这17位数字分别和系数相乘的结果相加。
        3、用加出来和除以11，看余数是多少?
        4 、余数只可能有0 、1、 2、 3、 4、 5、 6、 7、 8、 9、 10这11个数字。其分别对应的最后一位身份证的号码为1 .0. X. 9. 8. 7. 6. 5. 4. 3. 2.。
        5、通过上面得知如果余数是2，就会在身份证的第18位数字上出现罗马数字的Ⅹ。如果余数是10，身份证的最后一位号码就是2。

        :return: 18位身份证号码
        :rtype: str
        """
        if len(self.id_card_number) == 15:
            digital_ontology_code = self.id_card_number[0:6] + '19' + self.id_card_number[6:15]
            return digital_ontology_code + self.get_checkcode(digital_ontology_code)
        else:
            return "身份证号码格式不正确"

    def eighteen_to_fifteen(self):
        """
        18位身份证转15位身份证
        :return: 15位身份证
        :rtype: str
        """
        if len(self.id_card_number) == 18:
            # 去掉第6，7位和第18位即可18位身份证转15位身份证
            return self.id_card_number[0:6] + self.id_card_number[8:17]


class RegularExpression(object):
    """
    使用正则表达式从混乱的数据（string）中匹配符合格式要求的数据

    * 常用正则表达式：

        r'0?(13|14|15|17|18|19)[0-9]{9}'   ->   手机号码
        r'\d{17}[\d|x|X]|\d{15}'             -> 15位或18位身份证号码（身份证尾号x或X均可通过校验）

    """

    def __init__(self,content,re_str):
        """

        :param content: 需处理的字符串
        :type content: str
        :param re_str: 正则表达式
        :type re_str: str
        """
        self.content = content
        self.re_str = re_str
        #
        self.pattern = re.compile(re_str)

    def search(self):
        """
        扫描整个字符串并返回第一个成功的匹配

        :return:
        :rtype:
        """
        # re.search 匹配整个字符串，直到找到一个匹配。
        result = re.search(self.pattern,self.content)
        try:
            return result.group()
        # 匹配失败的话返回0
        except AttributeError:
            return 0

    def match(self):
        """
        尝试从字符串的起始位置匹配一个模式，如果不是起始位置匹配成功的话，match()就返回none

        :return:
        :rtype:
        """
        # re.match 只匹配字符串的开始，如果字符串开始不符合正则表达式，则匹配失败，函数返回 None
        result = re.match(self.pattern,self.content)
        try:
            return result.group()
        # 匹配失败的话返回0
        except AttributeError:
            return 0


if __name__ == '__main__':
    # iaa = IdCardNumber("440228640524722")
    # print(iaa.fifteen_to_eighteen())
    r = RegularExpression(u"阿大撒大撒大苏打ada441881198808150214.。。，asdsad阿达阿三",r'\d{17}[\d|x|X]|\d{15}')
    r2 = r.search()
    print(r2)