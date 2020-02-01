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
master.geometry("800x666+600+100") # 设置窗口大小和位置
master.title("社保参保月数计算器by李加林 gui版v2.0")
photo = tk.PhotoImage(file="img\\si2.gif")
imgLable = tk.Label(master,image=photo)
imgLable.pack()
theLabel = tk.Label(master, text="* 计算某个参保时段的参保月数用 *\n\n* 输入需要计算的参保时段起始日期和结束日期，得到月数 *\n ",font=("黑体",20))
theLabel.pack()
theLabel2 = tk.Label(master, text="* 日期格式为yymm或yymmdd *\n* 日期格式示例：201912 或 20191201 *\n\n",font=("黑体",16))
theLabel2.pack()

frame = tk.Frame(master)
frame2 = tk.Frame(master)
frame.pack()
frame2.pack()

v1 = tk.StringVar()
v2 = tk.StringVar()
v3 = tk.StringVar()


def test(content):
    return content.isdigit()


testCMD = master.register(test)
tk.Label(frame, text="起始日期：").grid(row=0, column=0)
e1 = tk.Entry(frame, textvariable=v1, validate="key", validatecommand=(testCMD, "%P")).grid(row=0, column=1)
tk.Label(frame, text="结束日期：").grid(row=1, column=0)
e2 = tk.Entry(frame, textvariable=v2, validate="key", validatecommand=(testCMD, "%P")).grid(row=1, column=1)
tk.Label(frame, text="月数：").grid(row=0, column=2)
e3 = tk.Entry(frame, textvariable=v3, state="readonly").grid(row=0, column=3)



def calc():
    """
    根据起始日期和结束日期，计算参保时段月数，set v3
    :return:  None
    :rtype: None
    """
#     result = int(v1.get()) + int(v2.get())
#     v3.set(str(result))
    full_month = (int(v2.get()[0:4]) - int(v1.get()[0:4])) * 12 + (int(v2.get()[4:6]) - int(v1.get()[4:6])) + 1
    v3.set(str(full_month))

tk.Button(frame2, text="开始计算", command=calc,font=("微软雅黑",16)).grid(pady=33)


tk.mainloop()