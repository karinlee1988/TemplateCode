#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 20200117
# @Author  : karinlee
# @FileName: count_date.py
# @Software: pycharm
# @Blog    ：https://blog.csdn.net/weixin_43972976



class CountDate(object):
    """
    用于计算2个日期之间相关数据的类
    """

    def __init__(self,begin_date,end_date):
        """

        :param begin_date: 开始日期
        :type begin_date: str
        :param end_date:  结束日期
        :type end_date: str
        """
        self.begin_date = begin_date
        self.end_date = end_date

    def count_month(self):
        """
        用于计算2个日期之间的月数（用于多发待遇月数计算，包含头月与尾月）
        :return:  开始日期与结束日期的月数（包含头尾）
        :rtype:  int
        """
        full_month = (int(self.end_date[0:4]) - int(self.begin_date[0:4])) * 12 + (int(self.end_date[4:6]) - int(self.begin_date[4:6])) + 1
        return full_month


if __name__ == '__main__':
    print("======计算多发待遇月数计算器by李加林v1.0======")
    print("++++++日期示例格式：6位或者8位年月，如201912或20191201++++++")
    print("------计算的结果包含头尾月份数------")
    flag = 1
    while flag:
        begin_date = input("请输入开始日期：")
        end_date = input("请输入结束日期：")
        # CountDate类实例化
        count = CountDate(begin_date,end_date)
        month = count.count_month()
        print("\n\n")
        print(f" 日期 {begin_date} 与 日期 {end_date} 之间一共有 {month} 个月 ")
        print("\n\n")
        flag = input("***输入任意字符继续计算，按enter直接退出程序......***")









