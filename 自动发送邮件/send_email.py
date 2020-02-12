#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/2/1 15:24
# @Author : karinlee
# @FileName : send_email.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976

"""
本模块用于自动发送邮件
"""

import smtplib
from email.mime.text import MIMEText

class SendEmail(object):
    """
    用于自动发送邮件
    可作为程序运行完毕的通知使用
    """

    def __init__(self):
        """
        填写收件邮箱，smtp server，读取account.txt的账号密码，标题，正文
        """
        self.to_list = ['263384909@qq.com']
        self.server_host = 'smtp.163.com'
        account_file = open('account.txt')
        self._username = account_file.readline().strip()
        self._password = account_file.readline().strip()
        # 邮件标题
        self.sub = "这个是由python发出的自动邮件测试"
        self.subtype = 'plain'    # 邮件正文的模式，通过_subtype设为html,默认是plain
        # 邮件正文
        self.content = """
        你好，这是一条python自动发出的信息1
        你好，这是一条python自动发出的信息2
        你好，这是一条python自动发出的信息3
        """

    def send(self):
        """
        发送邮件
        """
        me = "AUTO" + "<" + self._username + ">"
        # _subtype 可以设为html,默认是plain
        msg = MIMEText(self.content, _subtype=self.subtype)
        msg['Subject'] = self.sub
        msg['From'] = me
        msg['To'] = ';'.join(self.to_list)
        try:
            server = smtplib.SMTP()
            server.connect(self.server_host)
            server.login(self._username, self._password)
            server.sendmail(me, self.to_list, msg.as_string())
            server.close()
        except Exception as e:
            print(str(e))

if __name__ == '__main__':
    email = SendEmail()
    email.send()