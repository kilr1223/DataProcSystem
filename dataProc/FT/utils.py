import Configuration, re, os
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
				devices = None,
				lots = None, 
				test_site = 'DC', 
				start_date = None, 
				end_date = None
				):
	
	file_list = []
	dir_list = [path]

	# update_time = int(time.mktime(time.strptime(self.last_time, "%Y-%m-%d %H:%M:%S")))
	# update_time = int(datetime.timestamp(datetime.strptime(Configuration.UPDATE_TIME, "%Y-%m-%d %H:%M:%S")))

	for dir_path in dir_list:
		for item in os.scandir(dir_path):
			if item.is_dir():
				# print(f"Find a dir({item})")
				dir_list.append(item)
			else:
				# print(item.path)
				info_dict = scanFilename(os.path.basename(item))
				if not info_dict or info_dict[Configuration.TEST_FLAG] != test_site:
					continue

				#TODO: get file_time from info_dict
				file_time = os.stat(item.path).st_mtime
				lot = info_dict[Configuration.LOT_FLAG]	
				device = info_dict[Configuration.DEVICE_FLAG]

				start_time_mark = file_time > start_date if start_date != None else True
				end_time_mark = file_time < end_date if end_date != None else True
				lot_mark = lot in lots if lots != None else True
				device_mark = device in devices if devices != None else True

				if start_time_mark & end_time_mark & lot_mark & device_mark:
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

a = searchFiles(r'C:\Users\Wufanzhengshu\Desktop\生产数据')
# print(a)