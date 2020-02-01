#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/2/1 15:12
# @Author : karinlee
# @FileName : gui_temp_basic.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

import tkinter as tk

class GuiTempBasic(object):
    """
    tkinter GUI界面模板（不用选择文件操作版）
    """
    def __init__(self):
        """
        创建界面
        """
        # 新建窗口
        self.master = tk.Tk()
        # 在界面顶部添加横幅图片
        self.photo = tk.PhotoImage(file="si2.gif")
        # self.flag 当程序运行完成后给用户提示信息
        # self.v1 self.v2接收用户参数 待选用，不够可以加
        # 注意！这些都是tk.StringVar()对象，不是str，其他地方要用的话要用get()方法获取str
        self.flag = tk.StringVar()
        self.v1 = tk.StringVar()
        self.v2 = tk.StringVar()
        # 窗口图片横幅
        self.img_pack()
        # 主界面布局
        self.window()
        # 框架布局
        self.frame()
        # 保持运行
        self.master.mainloop()

    def img_pack(self):
        """
        窗口图片横幅
        :return:
        :rtype:
        """
        # 图片贴上去
        img_lable = tk.Label(self.master, image=self.photo)
        img_lable.pack()

    def window(self):
        """
        主界面布局设置
        :return:
        :rtype:
        """
        # 调整窗口默认大小及在屏幕上的位置
        self.master.geometry("800x600+600+100")
        # 窗口的标题栏，自己修改
        self.master.title("xxxxxxby李加林v1.0")
        # 把标题贴上去，自己修改
        tk.Label(self.master,text="------这个是大标题-------",font=("黑体",20)).pack()
        tk.Label(self.master,text="------这个是小标题-------\n\n",font=("黑体",16)).pack()

    def frame(self):
        """
        框架布局设置
        :return:
        :rtype:
        """
        # 框架贴上去，再在框架里添加Lable，Entry，Button等控件
        frame1 = tk.Frame(self.master)
        frame1.pack()
        # 输入框，标记，按键
        tk.Label(frame1, text="参数1", font=("黑体", 16)).grid(row=2, column=0)
        tk.Entry(frame1, textvariable=self.v1).grid(row=2, column=1)
        tk.Label(frame1, text="参数2:", font=("黑体", 16)).grid(row=3, column=0)
        tk.Entry(frame1, textvariable=self.v2).grid(row=3, column=1)
        # 按这个按钮执行主程序
        tk.Button(frame1, text="开始处理", command=self.main, font=("黑体", 16)).grid(row=4, column=1,pady=66)
        tk.Entry(frame1, textvariable=self.flag,state="readonly").grid(row=4, column=2)

    def main(self):
        """
        这个是主程序，最好将非gui版的程序封装成类，接收该GUI类的参数后实例化运行
        :return:
        :rtype:
        """
        #
        #  这里写主程序
        #
        # 标志设置为处理完成
        self.flag.set("处理完成！")

if __name__ == '__main__':
    # 实例化运行
    app = GuiTempBasic()