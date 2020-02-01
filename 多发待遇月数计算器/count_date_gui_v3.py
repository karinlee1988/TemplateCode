#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 20200117
# @Author  : karinlee
# @FileName:
# @Software: pycharm
# @Blog    ：https://blog.csdn.net/weixin_43972976

import tkinter as tk

class CountDate(object):
    """
    社保参保月数计算器，根据起始日期和结束日期，计算某参保时段月数（包含头尾月份）
    日期格式为yymm或yymmdd   日期格式示例：201912 或 20191201
    """

    def __init__(self):
        self.master = tk.Tk()
        self.photo = tk.PhotoImage(file="img\\si2.gif")
        self.v1 = tk.StringVar()
        self.v2 = tk.StringVar()
        self.v3 = tk.StringVar()
        self.main()
        self.frame()
        self.master.mainloop()

    def main(self):
        self.master.geometry("800x666+600+100")  # 设置窗口大小和位置
        self.master.title("社保参保月数计算器by李加林 gui版v2.0")
        imgLable = tk.Label(self.master, image=self.photo)
        imgLable.pack()
        theLabel = tk.Label(self.master, text="* 计算某个参保时段的参保月数用 *\n\n* 输入需要计算的参保时段起始日期和结束日期，得到月数 *\n ", font=("黑体", 20))
        theLabel.pack()
        theLabel2 = tk.Label(self.master, text="* 日期格式为yymm或yymmdd *\n* 日期格式示例：201912 或 20191201 *\n\n", font=("黑体", 16))
        theLabel2.pack()

    def frame(self):
        frame = tk.Frame(self.master)
        frame2 = tk.Frame(self.master)
        frame.pack()
        frame2.pack()
        testCMD = self.master.register(self.test)
        tk.Label(frame, text="起始日期：").grid(row=0, column=0)
        self.e1 = tk.Entry(frame, textvariable=self.v1, validate="key", validatecommand=(testCMD, "%P")).grid(row=0, column=1)
        tk.Label(frame, text="结束日期：").grid(row=1, column=0)
        self.e2 = tk.Entry(frame, textvariable=self.v2, validate="key", validatecommand=(testCMD, "%P")).grid(row=1, column=1)
        tk.Label(frame, text="月数：").grid(row=0, column=2)
        self.e3 = tk.Entry(frame, textvariable=self.v3, state="readonly").grid(row=0, column=3)
        tk.Button(frame2, text="开始计算", command=self.calc, font=("微软雅黑", 16)).grid(pady=33)

    def test(self,content):
        return content.isdigit()

    def calc(self):
        """
        根据起始日期和结束日期，计算参保时段月数，set v3
        :return:  None
        :rtype: None
        """
        # 根据起始日期和结束日期，计算参保时段月数
        full_month = (int(self.v2.get()[0:4]) - int(self.v1.get()[0:4])) * 12 + (int(self.v2.get()[4:6]) - int(self.v1.get()[4:6])) + 1
        self.v3.set(str(full_month))

if __name__ == '__main__':
    app = CountDate()



















