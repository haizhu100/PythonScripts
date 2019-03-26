#!/usr/bin/python
#_*_ coding: utf-8 _*_

"""
通过smtp协议发送邮件，定义了一个smtp的类，封装一些方法
Python对SMTP支持有smtplib和email两个模块，email负责构造邮件，smtplib负责发送邮件。
"""
__author__ = 'haizhu_cheng@163.com'

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class controlsmtp:
    # 定义邮件服务器 smtp.163.com / smtp.qq.com / mail.shtel.com.cn
    mailhost = ""
    # 发件人
    fromaddrs = ""
    # 发件人登陆密码
    userpwd = ""
    # 收件人
    toaddrs = []
    # 抄送地址
    ccaddrs = []
    # 定义邮件主题
    subject = ""
    # 定义邮件正文
    content = ""

    def __init__(self, mailhost, fromaddrs ,userpwd ,toaddrs, ccaddrs, subject ,content):
        self.mailhost = mailhost
        self.fromaddrs = fromaddrs
        self.userpwd = userpwd
        self.toaddrs = toaddrs
        self.ccaddrs = ccaddrs
        self.subject = subject
        self.content = content

    def sendemail(self,email_dict,filelist):
        """
        :param email_dict: 发送邮件相关内容字典
        :param file: 附件列表
        :return: 不返回内容
        """
        # 第三方 SMTP 服务
        mailhost = email_dict['mailhost']  # SMTP服务器
        fromaddr = email_dict['fromaddr']  # 发件人
        password = email_dict['password']  # 邮箱密码
        toaddrs = email_dict['toaddrs'].split(",") # 收件人地址
        ccaddrs = email_dict['ccaddrs'].split(",") # 抄送人地址
        subject = email_dict['subject'] #
        content = email_dict['content']
        # 添加附件邮件实例
        message = MIMEMultipart()
        # 邮件发件人
        message['From'] = fromaddr
        # 邮件收件人
        message['To'] = ";".join(toaddrs)
        # 抄送邮件
        message['Cc'] = ";".join(ccaddrs)
        # 邮件主题
        message['Subject'] = subject
        # 添加邮件正文
        message.attach(MIMEText(content, 'plain', 'utf-8'))
        # 添加附件(根据附件列表来确定)
        for i in len(filelist):
            excelApart = MIMEApplication(open(filelist[i], 'rb').read())
            excelApart.add_header('Content-Disposition', 'attachment', filename=filelist[i])
            message.attach(excelApart)

        try:
            # 普通发送端口
            smtpObj = smtplib.SMTP(mailhost, 25)
            smtpObj.login(fromaddr, password)
            smtpObj.sendmail(fromaddr, toaddrs, message.as_string())
            print("邮件发送成功")
        except smtplib.SMTPException as e:
            print("邮件发送失败")
            print('error:', e)  # 打印错误
        finally:
            smtpObj.quit()


    # 通过ssl发送邮件
    def sendemailssl(self,email_dict,filelist):
        """
        :param email_dict: 发送邮件相关内容字典
        :param file: 附件内容
        :return: 不返回内容
        """
        # 第三方 SMTP 服务
        mailhost = email_dict['mailhost']  # SMTP服务器
        fromaddr = email_dict['fromaddr']  # 发件人
        password = email_dict['password']  # 邮箱密码
        toaddrs = email_dict['toaddrs'].split(",") # 收件人地址
        ccaddrs = email_dict['ccaddrs'].split(",") # 抄送人地址
        subject = email_dict['subject'] #
        content = email_dict['content']
        # 添加附件邮件实例
        message = MIMEMultipart()
        # 邮件发件人
        message['From'] = fromaddr
        # 邮件收件人
        message['To'] = ";".join(toaddrs)
        # 抄送邮件
        message['Cc'] = ";".join(ccaddrs)
        # 邮件主题
        message['Subject'] = subject
        # 添加邮件正文
        message.attach(MIMEText(content, 'plain', 'utf-8'))
        # 添加附件(根据附件列表来确定)
        for i in len(filelist):
            excelApart = MIMEApplication(open(filelist[i], 'rb').read())
            excelApart.add_header('Content-Disposition', 'attachment', filename=filelist[i])
            message.attach(excelApart)

        try:
            smtpObj = smtplib.SMTP_SSL(mailhost, 465)
            smtpObj.login(fromaddr, password)
            smtpObj.sendmail(fromaddr, toaddrs, message.as_string())
            print("邮件发送成功")
        except smtplib.SMTPException as e:
            print("邮件发送失败")
            print('error:', e)  # 打印错误
        finally:
            smtpObj.quit()


