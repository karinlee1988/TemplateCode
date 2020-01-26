#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 20200117
# @Author  : karinlee
# @FileName: count_date_gui.py
# @Software: pycharm
# @Blog    ：https://blog.csdn.net/weixin_43972976

# from tkinter import *
import tkinter as tk

master = tk.Tk()
master.geometry("800x400+600+300") # 设置窗口大小和位置
master.title("多发待遇月数计算器by李加林 gui版v1.0")
theLabel = tk.Label(master, text="* 多发待遇月数计算用 *\n\n* 输入起始日期和结束日期，得到月数 *\n (月数包括头尾月份）")
theLabel.pack()
theLabel2 = tk.Label(master, text="* 日期格式为yymm或yymmdd *\n* 日期格式示例：201912 或 20191201 *\n\n")
theLabel2.pack()

frame = tk.Frame(master)

frame.pack()

v1 = tk.StringVar()
v2 = tk.StringVar()
v3 = tk.StringVar()


def test(content):
    return content.isdigit()


testCMD = master.register(test)
tk.Label(frame, text="起始日期：").grid(row=0, column=0)
e1 = tk.Entry(frame, textvariable=v1, validate="key", validatecommand=(testCMD, "%P")).grid(row=0, column=1)
tk.Label(frame, text="结束日期：").grid(row=0, column=2)
e2 = tk.Entry(frame, textvariable=v2, validate="key", validatecommand=(testCMD, "%P")).grid(row=0, column=3)
tk.Label(frame, text="月数：").grid(row=0, column=4)
e3 = tk.Entry(frame, textvariable=v3, state="readonly").grid(row=0, column=5)



def calc():
#     result = int(v1.get()) + int(v2.get())
#     v3.set(str(result))
    full_month = (int(v2.get()[0:4]) - int(v1.get()[0:4])) * 12 + (int(v2.get()[4:6]) - int(v1.get()[4:6])) + 1
    v3.set(str(full_month))

tk.Button(frame, text="开始计算月数！", command=calc).grid(row=6, column=2, pady=20)

tk.mainloop()