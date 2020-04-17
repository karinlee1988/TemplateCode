#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time : 2020/3/17 20:40
# @Author : karinlee
# @FileName : wechat_groupchatreply.py
# @Software : PyCharm
# @Blog : https://blog.csdn.net/weixin_43972976


"""
本模块使用itchat库，实现微信群回复通知


"""

import itchat
# from itchat.content import *
import time

# 自动回复
# 封装好的装饰器，当接收到的消息是Text，即文字消息
# @itchat.msg_register('Text')
# 封装好的装饰器，当接收到的消息是[TEXT, PICTURE,SHARING,ATTACHMENT,VIDEO]
# 对于不同消息的类型，采取不同的处理方法
# from itchat.content import *
# ==================================itchat.content=======================================
# TEXT       = 'Text'
# MAP        = 'Map'
# CARD       = 'Card'
# NOTE       = 'Note'
# SHARING    = 'Sharing'
# PICTURE    = 'Picture'
# RECORDING  = VOICE = 'Recording'
# ATTACHMENT = 'Attachment'
# VIDEO      = 'Video'
# FRIENDS    = 'Friends'
# SYSTEM     = 'System'
#
# INCOME_MSG = [TEXT, MAP, CARD, NOTE, SHARING, PICTURE,
#     RECORDING, VOICE, ATTACHMENT, VIDEO, FRIENDS, SYSTEM]
# =======================================================================================
# 引入后可使用itchat.content里面的常量，但也可以不引入 直接用字符串注册
# @itchat.msg_register([TEXT, PICTURE,RECORDING,ATTACHMENT,SHARING,VIDEO])

def wechat_groupchatreply():
    # 在注册时增加isGroupChat=True将判定为群聊回复
    @itchat.msg_register('Text', isGroupChat=True)
    def groupchat_reply(msg):
        """
        用于微信群聊@我自动回复
        """
        if msg['Type'] == 'Text':
            reply_message = msg['Text']

        # 获取所有群聊名称
        group = itchat.get_chatrooms(update=True)
        # 通过群聊名称（NickName） 找到群聊的username  赋予to_group变量 （群聊的username是一串很长的字符串
        # 如@@b26aec2461d611ae73683f302635de72a19317943d79860d350efe5ec20c196d）

        for g in group:
            if g['NickName'] == '测试功能':
                print(g['UserName'])
                to_group = g['UserName']


        # 当消息不是由自己发出的时候
        if not msg['FromUserName'] == my_user_name:
            if msg['User']["NickName"] == '测试功能':
                return u'111111'

        while True:
            itchat.send_msg(u"12345678", toUserName=to_group)
            # 定时发送的时间
            time.sleep(10)

    #登录微信
    itchat.auto_login()
    # 获取自己的user_name
    my_user_name = itchat.get_friends(update=True)[0]["UserName"]

    #开始运行
    itchat.run()

if __name__ == '__main__':
    wechat_groupchatreply()

