# -*- coding: utf-8 -*-
import smtplib
import string
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication # 用于添加附件

smtplib.SMTP()
host_server = 'smtp.163.com'
user = 'as256952950@163.com'
password = 'VGENJBEJIJCQBCHT'
sender = 'Email Auto Sent'
receiver = ['611118776@qq.com']
cc = ['xx@xxx.cn', 'xx@xxx.cn']
mail_title = 'Python自动发送FT数据的邮件'

mail_content = """
晚上好！
	该邮件为自动发送，请勿回复
"""
msg = MIMEMultipart()
msg["Subject"] = Header(mail_title,'utf-8')
msg["From"] = user
# msg["To"] = Header("测试邮箱","utf-8")
msg["To"] = ','.join(receiver)
msg['Cc'] = ','.join(cc)

msg.attach(MIMEText(mail_content))
att = MIMEApplication(open(r'C:\Users\Wufanzhengshu\Desktop\FT测试报告\2023-09-11\FTReport_20230911_095434.xlsx', 'rb').read())
# attachment = MIMEApplication(open(r'C:\Users\Wufanzhengshu\Desktop\FT\FTreport.xlsx','rb').read(),'base64','utf-8')
file_name = 'FTreport -' + '2023_7_18' + '.xlsx'
att["Content-Disposition"] = 'attachment;filename=' + file_name
att["Content-Type"] = 'application/octet-stream'
# 给附件重命名

att.add_header('Content-Dispositon','attachment',filename=('utf-8', '', file_name))
msg.attach(att)

# smtp = smtplib.SMTP_SSL(host_server, 465) # ssl登录连接到邮件服务器
# smtp.login(user,password)
# smtp.sendmail(user,receiver,msg.as_string())

try:
	smtp = smtplib.SMTP_SSL(host_server, 465) # ssl登录连接到邮件服务器
	# smtp.set_debuglevel(1) # 0是关闭，1是开启debug
	# smtp.ehlo(host_server) # 跟服务器打招呼，告诉它我们准备连接，最好加上这行代码
	smtp.login(user,password)
	smtp.sendmail(user,receiver + cc12,msg.as_string())
	# smtp.quit()
	print("邮件发送成功")
except smtplib.SMTPException as e:
    print("无法发送邮件", e)
