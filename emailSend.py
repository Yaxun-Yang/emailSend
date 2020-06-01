#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : yyx

import smtplib
from email.mime.text import MIMEText
from email.header import Header
import xlrd

# 表格所在路径
path = 'D:\\python\\emailTest.xlsx'
data = xlrd.open_workbook(path)
sheets = data.sheets()
# 通过序号获取表格
sheet_1_by_function = data.sheets()[0]
sheet_1_by_index = data.sheet_by_index(0)
# 通过名称获取表格
sheet_by_name = data.sheet_by_name("小游戏提交作品通知")

# 按行获取邮箱地址
n_of_rows = sheet_by_name.nrows
# n_of_cols=sheet_1_by_name.ncols

# 控制台检测获取邮箱是否正确
for i in range(1, n_of_rows):
    print(sheet_by_name.cell(i, 4))

# 第三方 SMTP 服务
mail_host = "smtp.qq.com"  # 设置服务器
mail_user = "*********@qq.com"  # 用户名
mail_pass = "************"  # 授权码
sender = "发件人昵称"

for i in range(1, n_of_rows):
    receiversName = sheet_by_name.cell(i, 1).value # 1由收件人昵称在表格中的列数替代
    receivers = sheet_by_name.cell(i, 4).value # 4由收件人邮箱在表格中的列数替代
    print(receiversName, receivers)



    # 以称呼作为参数
    mail_msg = """
    <p> %s: </p>

​	<p> 你们好，我们是微信小程序华中赛区的赛事负责团队。</p>

​	<p>截止6月1日12:00赛事方仍未收到贵队提交的作品。微信小程序应用开发赛的作品提交截止时间为6月30日21:00。为减少在最后提交时间网络拥堵造成的不必要风险，希望贵队安排好时间提交高质量的作品。</p>

    <p align="right">微信小程序华中赛区</p>
    """ % receiversName

    # 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
    message = MIMEText(mail_msg, 'html', 'utf-8')
    message['From'] = Header(sender, 'utf-8') # 发送者昵称格式调整
    message['To'] = Header(receiversName, 'utf-8') #接收者昵称

    subject = '微信小程序应用开发大赛作品提交通知'
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(mail_user, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException:
        print("Error: 无法发送邮件")
