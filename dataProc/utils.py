<<<<<<< HEAD
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
=======
import Configuration, re
from datetime import datetime

#TODO There is still a BUG need to be fixed(the value of TIME always included the filename suffix)
def scanFilename(filename):
	mod = Configuration.FILENAME_MOD
	fn_list  = filename.split('_')
	mod_list = []
	res_dict = {}
	s = ''
	flag = True

	for m in mod:
		if m == '_' and flag:
			mod_list.append(s)
			s = ''
			continue
		elif m == '<':
			flag = False
		elif m == '>':
			flag = True
		else:
			s += m

	mod_list.append(s.split('.')[0])

	if len(fn_list) != len(mod_list):
		return False

	for mod,fn in zip(mod_list, fn_list):
		res_dict[mod] = fn

	return res_dict
	# print(com_name, test_site, device, test_program_rev, lot, retest_num, operation_site, date, time)

def searchFiles(path, 
				lots = None, 
				test_site = 'DC', 
				start_date = None, 
				end_date = None
				):
	
	file_list = []

	# update_time = int(time.mktime(time.strptime(self.last_time, "%Y-%m-%d %H:%M:%S")))
	# update_time = int(datetime.timestamp(datetime.strptime(Configuration.UPDATE_TIME, "%Y-%m-%d %H:%M:%S")))

	for item in os.scandir(path):
		if item.is_dir():
			# print(f"Find a dir({item})")
			l = getFiles(item.path)
			file_list.extend(l)
		else:
			# print(item.path)
			info_dict = scanFilename(os.path.basename(item))
			if not info_dict or info_dict[Configuration.TEST_FLAG] != test_site:
				continue

			#TODO: get file_time from info_dict
			file_time = os.stat(item.path).st_mtime
			lot = info_dict[Configuration.LOT_FLAG]	

			start_time_mark = file_time > start_date if start_date != None else True
			end_time_mark = file_time < end_date if end_date != None else True
			lot_mark = lot in lots if lots != None else True

			if start_time_mark & end_time_mark & lot_mark:
				print(f'[Find File]: Find a new file({item.path})')
			# time = datetime.fromtimestamp(os.stat(item.path).st_mtime)
				info_dict[Configuration.PATH_FLAG] = item.path
				file_list.append(info_dict)

	return file_list

def updateConfTime(time = None):
	if time == None:
		time = date.strftime('%Y-%m-%d %H:%M:%S')

	Configuration.conf.set('setting', 'last_time', time)

def fileVerify(path):

	devices = re.findall(r'(S\w+[M|D]\d+[A|H]?D?)',path)
	lots = re.findall(r'([B|BP]\d*)',path)
	
	if len(devices) != 2 or devices[0] != devices[1]:
		print(f'[Filename Error]: Device verify ERROR, please check the device of the file({path})')
		return False

	if len(lots) != 2 or lots[0] != lots[1]:
		print(f'[Filename Error]: LOT verify ERROR, please check the lot of the file({path})')
		return False

	return devices[0], lots[0]

a = '1234'
b = ['1234', '123']

print(a in b)
>>>>>>> 9aa4e6bac2d28b00d9b03addb14a1f3b9b57c056
