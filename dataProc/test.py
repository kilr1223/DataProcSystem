import re, csv, configparser, json, re, xlrd, xlwt, xlutils, os, datetime, openpyxl, configparser
import pandas as pd
import numpy as np

from tkinter import filedialog
from dataProc import DataObject
def traversal_files(path):

	file_list = []

	for item in os.scandir(path):
		if item.is_dir():
			l = traversal_files(item.path)
			file_list.extend(l)
		else:
			# os.path.join()
			file_list.append(item.path)

	return file_list

def delAllFiles(path):

	for item in os.scandir(path):
		if item.is_dir():
			delAllFiles(item.path)
		else:
			# print(os.path.isfile(item))
			os.remove(item)

def infoDel(file_path):
	
	old_csv = csv.reader(open(file_path, 'r', encoding = 'gbk'))
	old_filepath = file_path.replace('S1M075120B1', 'S1M075120B1 - DEL')
	new_csv = csv.writer(open(old_filepath, 'w', encoding = 'gbk', newline=''))

	for row in old_csv:
		if row[0] in ['SPEC_LOW', 'SPEC_HIGH']:
			continue

		if  len(row) < 8:
			new_csv.writerow(row)
			continue
		# print(row)
		new_row = row[:7] + row[17:]
		if row[0] == 'PARAMETER_ID':
			flag = False
			num = 1
			for i in range(len(new_row)):
				# print(row[i])
				if row[i] == '1':
					# print("change")
					flag = True

				if flag:
					new_row[i] = num
					# print(new_row[i])
					num += 1
		new_csv.writerow(new_row)

def mergeData(path):

	file_list = traversal_files(path)
	df = pd.DataFrame()

	for file in file_list:
		file_name = os.path.basename(file)
		if file_name[-3:] not in ['csv', 'CSV']:
			continue

		lot = re.findall(r'([B|BP]\d*)',file_name)[0]

		df_csv = pd.read_csv(file, header = 25)

		df_csv.insert(0, 'LOT', lot)
		df = pd.concat([df, df_csv])

	df.to_excel(path + "/all data.xlsx", sheet_name='DATA')

def concatFile(file1, file2):
	
	# for row in csv.reader(open(file1, 'r', encoding = 'gbk' )):
	# 	file_len += 1
	file_name = os.path.basename(file1)
	path = r'C:\Users\Wufanzhengshu\Desktop\12'

	new_file = csv.writer(open(os.path.join(path, file_name), 'a', encoding = 'gbk', newline=''))

	first_file = csv.reader(open(file1, 'r', encoding = 'gbk', newline=''))
	second_file = csv.reader(open(file2, 'r', encoding = 'gbk' ))

	file_len = -getHeader(file1)

	for row in first_file:
		new_file.writerow(row)
		file_len += 1

	header = getHeader(file2)
	while header != 0:
		second_row = next(second_file)
		header -= 1

	for new_row in second_file:
		new_row[0] = str(int(new_row[0]) + file_len)
		new_file.writerow(new_row)

def getHeader(file):

	data_iter = csv.reader(open(file, 'r', encoding = 'gbk'))
	data = next(data_iter)
	header_len = 0
	flag = 2

	while flag > 0:
		if len(data) > 0 and '---' in data[0]:
			flag -= 1
		data = next(data_iter)
		header_len += 1

	while True:
		if len(data) != 0:
			try:
				int(data[0])

				break
			except ValueError:
				pass
		data = next(data_iter)
		header_len += 1

	return header_len

def linkTwoFile():

	path = filedialog.askdirectory()
	# concatFile(file1, file2)
	file1_list = []
	file2_list = []

	for file in os.listdir(path):
		if not file.endswith(('.csv', '.CSV')):
			continue

		if '@' not in file:
			file1_list.append(file)
		elif '@3' in file:
			file2_list.append(file)

	for file1 in file1_list:
		target_end = re.findall(r'\s(.*)$', file1)[0]
		print(target_end)
		for file2 in file2_list:
			end = re.findall(r'\s(.*)$', file2)[0]

			if end == target_end:
				concatFile(os.path.join(path, file1), os.path.join(path, file2))

def clearData(file):
	path = os.path.dirname(file)
	def checkRow(row):
		index, bv, vth, idss, igss, rdson = row
		try:
			if all(-20 < value < 20 for value in [bv, vth, rdson]) and all(value < 5 for value in [idss, igss]):
				return True
			else:
				return False
		except:
			return False
	wb = openpyxl.load_workbook(filename = file)
	ws_list = wb.sheetnames
	ws_list.remove('DATA')

	info = ''
	for sheet in ws_list:
		ws = wb[sheet].values
		next(ws)

		for row in ws:
			
			if not checkRow(row):
				info += f'Delete Row {row[0]} from Sheet {sheet} \n'
				# print(row)
	# print(info)
	with open(os.path.join(path, 'Info2.txt'),"w") as f:
		f.write(info)
# path = filedialog.askdirectory()
# df = pd.read_excel(r'C:\Users\Wufanzhengshu\Desktop\3193\Data.xlsx', sheet_name='14H-BI90')
# print(int(df['INDEX'][0]))

def matchFilename(filename):

	mod = 'SHC_<TEST_SITE>_<DEVICE>_<PROGRAM_REV>_<LOT>_<RETEST_NUM>_<OPERATION_SITE>_<DATE>_<TIME>'
	mod_list = []
	fn_list = filename.split('_')
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
	mod_list.append(s)

	if len(fn_list) != len(mod_list):
		return False

	for mod,fn in zip(mod_list, fn_list):
		res_dict[mod] = fn

	return res_dict

def getHeader(file):

		data_iter = csv.reader(open(file, 'r', encoding = 'gbk'))
		data = next(data_iter)
		header_len = 0
		flag = 2

		while flag > 0:
			if len(data) > 0 and '---' in data[0]:
				flag -= 1
			data = next(data_iter)
			header_len += 1

		while True:
			if len(data) != 0:
				try:
					int(data[0])

					break
				except ValueError:
					pass
			data = next(data_iter)
			header_len += 1

		return header_len

def test(path):

	file_list = []
	df = pd.DataFrame()
	for file in os.listdir(path):
		if 'BIN' in file or file[-3:] not in ['csv', 'CSV']:
			continue
		file_list.append(file)

	for file in file_list:
		num = int(re.findall(r'-(\d*)-RAWDATA', file)[0])
		title = ['NUM', 'CHIP', 'X', 'Y', 'BIN', 'MRBPIN', 'CONT', 'VTH0', 'IGSS0', 'IGSS1', 'VTH1', 'ABSDEL', 'VTH2', 'IGSS2', 'VTH3', 'ABSDEL', 'IDSS1', 'IDSS2', 'DELAY', 'IGSS3', 'DELAY', 'VTH4', 'VTH6', 'RDON1', 'RDON2', 'DELAY', 'VFSD1', 'DELAY', 'IGSS4', 'IGSS5']
# c = csv.reader(open(file, 'r', encoding = 'gbk'))
		print(file)
		# h = getHeader(os.path.join(path, file))
		df1 = pd.read_csv(os.path.join(path, file), header = 26, encoding=u'gbk')
		if len(df1.columns) > len(title):
			last_col = df1.shape[1] - 1
			df1 = df1.drop(df1.columns[last_col], axis=1)
		df1.columns = title
		# df1.insert(0, 'NUM', num)
		# print(df1)
		# break
		df = pd.concat([df, df1])
	return df

clearData(r'D:\可靠性实验\试验报告\S1M040120H\DATA\Data.xlsx')