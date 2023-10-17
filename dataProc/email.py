import smtplib
# import string
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication # 用于添加附件
from email.utils import formataddr

def emailSender(mail_content, report_path):
	smtplib.SMTP()
	host_server = 'smtp.qiye.aliyun.com'
	port = 465
	user = 'as256952950@163.com'
	password = 'VGENJBEJIJCQBCHT'
	sender = 'Email Auto Sender'
	receiver = ['611118776@qq.com']
	cc = []
	timestamp = date.strftime('%Y.%m.%d')
	mail_title = f'[{timestamp}]生产数据汇总'

	# mail_content = f"""
	# 晚上好！
	# 	{a}
	# 	该邮件为自动发送，请勿回复
	# """
	msg = MIMEMultipart()
	msg["Subject"] = Header(mail_title,'utf-8')
	msg["From"] = user
	# msg["To"] = Header("测试邮箱","utf-8")
	msg["To"] = ','.join(receiver)
	msg['Cc'] = ','.join(cc)

	msg.attach(MIMEText(mail_content))
	att = MIMEApplication(open(report_path, 'rb').read())
	# attachment = MIMEApplication(open(r'C:\Users\Wufanzhengshu\Desktop\FT\FTreport.xlsx','rb').read(),'base64','utf-8')
	file_name = os.path.basename(report_path)
	att["Content-Disposition"] = 'attachment;filename=' + file_name
	att["Content-Type"] = 'application/octet-stream'
	# 给附件重命名

	att.add_header('Content-Dispositon','attachment',filename=('utf-8', '', file_name))
	msg.attach(att)

	# smtp = smtplib.SMTP_SSL(host_server, 465) # ssl登录连接到邮件服务器
	# smtp.login(user,password)
	# smtp.sendmail(user,receiver,msg.as_string())

	try:
		smtp = smtplib.SMTP_SSL(host_server, port) # ssl登录连接到邮件服务器
		# smtp.set_debuglevel(1) # 0是关闭，1是开启debug
		# smtp.ehlo(host_server) # 跟服务器打招呼，告诉它我们准备连接，最好加上这行代码
		smtp.login(user,password)
		smtp.sendmail(user,receiver + cc,msg.as_string())
		# smtp.quit()
		print("邮件发送成功")
	except smtplib.SMTPException as e:
		print("无法发送邮件", e)

def aliMailSender(mail_content, report_path):
	'''~~~smtp认证使用的邮箱帐号密码~~~'''
	username = 'wfzs@sichainsemi.com'
	password = 'Xc951223'#SXYkEspHp77hdySZ

	'''~~~定义发件地址~~~'''
	From = formataddr([Header('吴帆正树', 'utf-8').encode(), 'wfzs@sichainsemi.com'])  #昵称(邮箱没有设置外发指定自定义昵称时有效)+发信地址(或代发)
	# replyto = ''  #回信地址

	'''定义收件对象'''
	to = ','.join(['wfzs@sichainsemi.com'])  #收件人
	cc = ','.join(['swh@sichainsemi.com'])  #抄送
	bcc = ','.join(['', ''])  #密送
	rcptto = [to,cc,bcc]  #完整的收件对象

	'''定义主题'''
	timestamp = date.strftime('%Y.%m.%d')
	Subject = f'[{timestamp}]生产数据汇总'

	'''~~~开始构建message~~~'''
	msg = MIMEMultipart('alternative')
	'''1.1 收发件地址、回信地址、Message-id、发信时间、邮件主题'''
	msg['From'] = From
	# msg['Reply-to'] = replyto
	msg['TO'] = to
	msg['Cc'] = cc
	# msg['Bcc'] = bcc  #建议密送地址在邮件头中隐藏
	# msg['Message-id'] = email.utils.make_msgid()
	# msg['Date'] = email.utils.formatdate()
	msg['Subject'] = Subject
	''''1.2 正文text/plain部分'''
	textplain = MIMEText(mail_content, _subtype='plain', _charset='UTF-8')
	msg.attach(textplain)
	'''1.3 封装附件'''
	file = report_path   #指定本地文件，请换成自己实际需要的文件全路径。
	filename = os.path.basename(file)
	att = MIMEText(open(file, 'rb').read(), 'base64', 'utf-8')
	att["Content-Type"] = 'application/octet-stream'
	att.add_header("Content-Disposition", "attachment", filename=filename)
	msg.attach(att)

	'''~~~开始连接验证服务~~~'''
	try:
	    client = smtplib.SMTP_SSL('smtp.qiye.aliyun.com', 465)
	    print('smtp_ssl----连接服务器成功，现在开始检查帐号密码')
	except Exception as e1:
	    client = smtplib.SMTP('smtp.qiye.aliyun.com', 25, timeout=5) 
	    print('smtp----连接服务器成功，现在开始检查账号密码')
	except Exception as e2:
	    print('抱歉，连接服务超时')
	    exit(1)
	try:
	    client.login(username, password)
	    print('帐密验证成功')
	except:
	    print('抱歉，帐密验证失败')
	    exit(1)

	'''~~~发送邮件并结束任务~~~'''
	client.sendmail(username, (','.join(rcptto)).split(','), msg.as_string())
	client.quit()
	print('邮件发送成功')

def emailEdit(info_dict):
	mail_content = f"""大家好！\n<content>\n    该邮件为自动发送，请勿回复\n\n祝生活顺利\n清纯半导体"""
	if len(info_dict) == 0:
		content = "    今日未进行FT数据测试，或程序未能抓取有效数据文件，详情请咨询相关人员，谢谢！"
		mail_content = mail_content.replace('<content>', content)
	else:
		content='    今日生产测试数据摘要如下：\n'
		for device, d in info_dict.items():
			content = content + f'        {device}共测试{len(d.keys())}个批次：\n'
			for lot, info in d.items():
				count = info['COUNT']
				good = round(info['GOOD']*100, 2)
				content = content + f'            {lot}测试器件{count}颗, 良率为{good}%\n'
		content += "    详细报告已随附件发送，如有任何问题请联系相关人员，谢谢！"
		mail_content = mail_content.replace('<content>', content)
	return mail_content