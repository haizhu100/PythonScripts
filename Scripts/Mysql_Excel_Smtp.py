#!/usr/bin/python3
#_*_ coding: UTF-8 _*_
"""
该脚本功能包括从mysql数据库获取数据，然后将数据存到excel，最后作为附件发送邮件！数据库信息和邮件信息从配置库里读取
"""

import pymysql
import time
import xlwt
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from configparser import ConfigParser

# 从数据库获取数据
def get_data_from_mysql(mysqldb_dict):
    # 连接数据库
    connect = pymysql.Connect(
        host=mysqldb_dict["host"],
        port=mysqldb_dict["port"],
        user=mysqldb_dict["user"],
        passwd=mysqldb_dict["passwd"],
        db=mysqldb_dict["db"],
        charset=mysqldb_dict["charset"]
    )
    # 获取游标
    cursor = connect.cursor()
    # 获取当前时间
    nowtime = time.time()
    signtimenow = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(nowtime))
    # 获取一周前时间  604800=7*24*60*60
    signpassoneweek = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(nowtime - 604800))
    print(signtimenow)
    print(signpassoneweek)
    sql = "SELECT * FROM `equipment_sign` WHERE `sign_time` < %s AND `sign_time` > %s"
    # 执行SQL语句
    cursor.execute(sql, (signtimenow, signpassoneweek))
    # 获取所有记录列表
    results = cursor.fetchall()
    print('获取数据成功')
    # 关闭数据库
    connect.close()
    return results

def write_excel(outputfilepath, outputfilename ,results):
    f = xlwt.Workbook() # 创建工作簿
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok = True)  # 创建sheet
    row0 = [u'主键',u'应用编号',u'上传人',u'工号',u'设备编号',u'经度',u'纬度',u'网格ID',u'grid_name',u'sgin_des',u'picture_url',u'签到时间',u'签到类型',u'运营商']
    # 确定栏位宽度
    col_width = []
    for i in range(len(results)):
        for j in range(len(results[i])):
            if i == 0:
                col_width.append(len_byte(results[i][j]))
            else:
                if col_width[j] < len_byte(str(results[i][j])):
                    col_width[j] = len_byte(results[i][j])

    # 设置栏位宽度，栏位宽度小于10时候采用默认宽度
    for i in range(len(col_width)):
        if col_width[i] > 10:
            sheet1.col(i).width = 256 * (col_width[i] + 1)
    # 设置excel的风格
    style1 = set_style('Times New Roman', 220, True)
    # 还原时间格式
    style2 = set_style('Times New Roman', 220, True)
    # 生成第一行
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i], style1)
    # 装入results的数据,j=1从第二行开始写入数据
    j = 1
    for row in results:
        for k in range(0, len(row)):
            if(k == 11):
                # 这一列是打卡时间，需要还原时间格式
                style2.num_format_str = 'yyyy-mm-dd h:mm:ss'
                sheet1.write(j, k, row[k], style2)
                continue
            sheet1.write(j, k, row[k], style1)
        j = j+1
    output_path = outputfilepath + outputfilename
    f.save(output_path)
    print('数据保存到excel成功')

# excel设置样式
def set_style(font_name, font_height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = font_name  # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = font_height

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 设置居中
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    alignment.vert = xlwt.Alignment.VERT_TOP  # 垂直方向

    style.font = font
    style.borders = borders
    return style

# 发送带附件的邮件
def send_email(outputfilepath, outputfilename ,email_dict):
    # 添加附件邮件实例
    message = MIMEMultipart()
    # 第三方 SMTP 服务
    mailhost = email_dict['mailhost']  # SMTP服务器
    fromaddr = email_dict['fromaddr']
    message['From'] = fromaddr
    print("发件人：" + message['From'])
    password = email_dict['password']
    toaddrs = email_dict['toaddrs'].split(",")
    message['To'] = ",".join(toaddrs)
    print("收件人：" + message['To'])
    if 'ccaddrs' in email_dict:
        ccaddrs = email_dict['ccaddrs'].split(",")
        message['Cc'] = ";".join(ccaddrs)
        print("抄送人：" + message['Cc'])
    subject = email_dict['subject']
    content = email_dict['content']
    message['Subject'] = subject
    # 添加邮件正文
    message.attach(MIMEText(content, 'plain', 'utf-8'))
    file = outputfilepath + outputfilename
    # 添加附件
    excelApart = MIMEApplication(open(file, 'rb').read())
    # 邮件里呈现的文件名要去除路径
    excelApart.add_header('Content-Disposition', 'attachment', filename=outputfilename)
    message.attach(excelApart)

    try:
        # smtpObj = smtplib.SMTP_SSL(mailhost, 465)
        smtpObj = smtplib.SMTP(mailhost,25)
        smtpObj.login(fromaddr, password)
        smtpObj.sendmail(fromaddr, toaddrs, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException as e:
        print("邮件发送失败")
        print('error:', e)  # 打印错误
    finally:
        smtpObj.quit()


    # 获取字符串长度，一个中文的长度为2
def len_byte(value):
    valuestring = str(value)
    length = len(valuestring)
    utf8_length = len(valuestring.encode('utf-8'))
    length = (utf8_length - length) / 2 + length
    return int(length)

if __name__ == '__main__':

    # 初始化类
    cp = ConfigParser()
    # 如果配置里有中文的话，需要添加encoding="utf-8-sig"
    cp.read("../config/config_mysql_excel_smtp.cfg", encoding="utf-8-sig")

    # 得到所有的section，以列表的形式返回
    section = cp.sections()
    # print(section)
    for section in cp.sections():
        # print(section)
        if section == 'mysql_db':
            mysqldb_dict = {'host': cp.get(section, "host"),
                            'port': cp.getint(section, "port"),
                            'db': cp.get(section,"db"),
                            'user': cp.get(section, "user"),
                            'passwd': cp.get(section, "passwd"),
                            'charset': cp.get(section, "charset")}
            # print(mysqldb_dict)
        if section == 'excel':
            file_dict = {
                'outputfilepath': cp.get(section, "filepath"),
                'outputfilename': cp.get(section, "filename")
            }
        if section == 'email':
            if 'ccaddrs' in cp.options(section):
                email_dict = {'mailhost': cp.get(section, "mailhost"),
                              'fromaddr': cp.get(section, "fromaddr"),
                              'password': cp.get(section, "password"),
                              'toaddrs': cp.get(section, "toaddrs"),
                              'ccaddrs': cp.get(section, "ccaddrs"),
                              'subject': cp.get(section, "subject"),
                              'content': cp.get(section, "content")}
            else:
                email_dict = {'mailhost': cp.get(section, "mailhost"),
                              'fromaddr': cp.get(section, "fromaddr"),
                              'password': cp.get(section, "password"),
                              'toaddrs': cp.get(section, "toaddrs"),
                              'subject': cp.get(section, "subject"),
                              'content': cp.get(section, "content")}

   # 从数据库获取数据
    results = get_data_from_mysql(mysqldb_dict)
   # 获取当前时间
    nowtime = time.time()
    # 定义输出文件路径
    outputfilepath = file_dict['outputfilepath']
    # 定义输出文件名
    outputfilename = file_dict['outputfilename'] + str(nowtime) + '.xls'
   # 写数据的数据到excel文件
    write_excel(outputfilepath, outputfilename , results)
   # 将excel文件作为附件发送邮件
    send_email(outputfilepath, outputfilename ,email_dict)