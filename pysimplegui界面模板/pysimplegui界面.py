#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2022/2/8 9:07
# @Author : karinlee
# @FileName : pysimplegui界面.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

import PySimpleGUI as sg


class BaseGUI(object):
    """
    基本的一个pysimplegui界面类

    20220208 test ok
    """

    def __init__(self):
        # 设置pysimplegui主题，不设置的话就用默认主题
        sg.ChangeLookAndFeel('Purple')
        # 定义2个常量，供下面的layout直接调用，就不用一个个元素来调字体了
        # 字体和字体大小
        self.FONT = ("微软雅黑", 16)
        # 可视化界面上元素的大小
        self.SIZE = (20, 1)
        # 界面布局
        self.layout = [
            # sg.Image()插入图片，支持gif和png
            [sg.Image(filename=r"peppa.png",
                      pad=(150, 0))],
            # sg.Text()显示文本
            [sg.Text('', font=self.FONT, size=self.SIZE)],
            # sg.Input()是输入框
            # 添加选择文件按钮，使用sg.FileBrowse()
            [sg.Text('请选择文件：', font=self.FONT, size=(30, 1))],
            [sg.Input('  ', key="_FILE_", readonly=True,  # readonly=True时不能在图形界面上直接修改该输入框内容
                      size=(36, 1), font=self.FONT),
             sg.FileBrowse(button_text='选择文件', size=(10, 1), font=self.FONT)],
            # 添加选择文件夹按钮，使用
            [sg.Text('请选择文件夹：', font=self.FONT, size=(30, 1))],
            [sg.Input('  ', key="_FOLDER_", readonly=True,
                      size=(36, 1), font=self.FONT),
             sg.FolderBrowse(button_text='选择文件夹', size=(10, 1), font=self.FONT)],
            [sg.Text(' 请输入数据:', font=self.FONT, size=self.SIZE),
             sg.Input(key='_DATA_', font=self.FONT, size=(10, 1))],
            [sg.Text(' 这里是返回的结果:', font=self.FONT, size=self.SIZE),
             sg.Input(key='_RESULT_', font=self.FONT, size=(10, 1), readonly=True)],
            [sg.Text('')],
            # sg.Btn()是按钮
            [sg.Btn('按我开始', key='_SUMMIT_', font=("微软雅黑", 16), size=(20, 1))],
            # sg.Output()可以在程序运行时，将原本在控制台上显示的内容输出到一个图形文本框里（如print命令的输出）
            [sg.Output(size=(72, 6), font=("微软雅黑", 10), background_color='light gray')]
        ]
        # 创建窗口，引入布局，并进行初始化
        # 创建时，必须要有一个名称，这个名称会显示在窗口上
        self.window = sg.Window('这是一个基本的pysimplegui窗口模板', layout=self.layout, finalize=True)

    # 窗口持久化
    def run(self):
        # 创建一个事件循环，否则窗口运行一次就会被关闭
        while True:
            # 监控窗口情况
            event, value = self.window.Read()
            # 当获取到事件时，处理逻辑（按钮绑定事件，点击按钮即触发事件）
            # sg.Input(),sg.Btn()都带有一个key，监控它们的状况，读取或写入信息
            if event == '_SUMMIT_':
                filepath = value['_FILE_']
                folderpath = value['_FOLDER_']
                data = value['_DATA_']
                # 获取到这些信息后，就可以进行处理了。

                # 例：计算输入数据的长度
                lenth = len(data)
                # 处理完毕，若要返回结果到图形界面上,使用self.window.Element().Updata()进行更新：
                self.window.Element("_RESULT_").Update(lenth)
                # 使用print()让sg.Output()捕获信息，直接print就可以了。哪里print都可以。
                print(f"选择的文件路径是:{filepath}")
                print(f"选择的文件夹路径是:{folderpath}")
                print(f"数据的长度是{lenth}")
            # 如果事件的值为 None，表示点击了右上角的关闭按钮，则会退出窗口循环
            if event is None:
                break
        self.window.close()


if __name__ == '__main__':
    # 实例化后运行
    tablegui = BaseGUI()
    tablegui.run()
