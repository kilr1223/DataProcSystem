import re, csv, os
import pandas as pd

def getWLBIcolname(file):
	rows_iter = csv.reader(open(file, 'r', encoding = 'gbk'))
	row = next(rows_iter)
	header = 0

	while True:
		if len(row) == 0:
			continue
		elif row[0] == 'Test Name':
			end_list = row
		elif row[0] == 'Site#':
			head_list = row
		elif row[0] == '1':
			break
		row = next(rows_iter)
		header += 1

	while '' in head_list:
		head_list.remove('')

	for item in end_list[1:]:
		if len(item) > 0:
			head_list.append(item)

	return header, head_list

def getCPData(file):
	rows_iter = csv.reader(open(file, 'r', encoding = 'gbk'))
	row = next(rows_iter)
	header = 0
	flag = 2

	while flag > 0:
		if len(row) == 0:
			continue
		elif '-----' in row[0]:
			flag -= 1

		row = next(rows_iter)
		header += 1

	return header

def getDataFromFolder():
	df_CP = pd.DataFrame()
	df_WLBI = pd.DataFrame()

	path = r'C:\Users\Wufanzhengshu\Desktop\3193\WLBI'
	for file in os.listdir(path):

		if file[-3:] not in ['csv', 'CSV']:
			continue

		file_path = os.path.join(path, file)

		header, col_list = getWLBIcolname(file_path)
		df = pd.read_csv(file_path, header = header, names = col_list)
		df = df.loc[(df['Site#'] == '1') | (df['Site#'] == '2'),:]
		wafer = re.search(r'-(\d*)[A-Z]\d', file, re.I).group(1) 
		df.insert(0, 'WAFER', str(wafer))
		df_WLBI = pd.concat([df_WLBI, df])

	path = r'C:\Users\Wufanzhengshu\Desktop\3193\\CP'		
	for file in os.listdir(path):

		if file[-3:] not in ['csv', 'CSV']:
			continue

		file_path = os.path.join(path, file)

		header = getCPData(file_path)
		df = pd.read_csv(file_path, header = header)
		# cols = list(df.columns)

		# for i in range(3):
		# 	cols[i] = ['X', 'Y', 'BIN'][i]
		df.columns = getColName(file_path)
		df = df.loc[df['BIN'] == 1, :]
		wafer = re.search(r'(\d*),?.csv$', file, re.I).group(1)
		df.insert(0, 'WAFER', str(wafer))

		if '' in df.columns:
			df = df.drop(np.NaN)

		# df_CP = df_CP.append(df)
		if len(df_CP) == 0:
			df_CP = df
		else:
			df_CP = pd.concat([df_CP, df], ignore_index = True)

	df_CP['X'] = df_CP['X'] + 4	
	df_CP['Y'] = df_CP['Y'] + 14	
	print(df_CP)
	return df_CP, df_WLBI

def getColName(file):	

	header = getCPData(file)
	data_iter = csv.reader(open(file, 'r', encoding = 'gbk'))
	data = next(data_iter)
	index_dict = {	'No.' : [], 
					'Conditon1' : [], 
					'Conditon2' : [], 
					'Conditon3' : []}

	for i in range(0, header):
		if len(data) > 0 and data[0] in index_dict.keys():
			index_dict[data[0]] = data
		data = next(data_iter)

	index_dict['No.'][0] = 'X'
	index_dict['No.'][1] = 'Y'

	for i in range(4, len(index_dict['No.'])):

		if len(index_dict['Conditon1'][i]) > 0 and index_dict['Conditon1'][i] != '-':
			index_dict['No.'][i] = index_dict['No.'][i] + '@' + index_dict['Conditon1'][i]
		else:
			continue

		if index_dict['Conditon2'][i] != '-':
			index_dict['No.'][i] = index_dict['No.'][i] + '&' + index_dict['Conditon2'][i]

		if len(index_dict['Conditon3']) > 0 and index_dict['Conditon3'][i] != '-':
			index_dict['No.'][i] = index_dict['No.'][i] + '&' + index_dict['Conditon3'][i]	

	return index_dict['No.']

def checkError(df_CP, df_WLBI):

	for index, row in df_CP.iterrows():
		wafer, X, Y = row
		if wafer not in df_WLBI['WAFER'].unique():
			continue
		df = df_WLBI.loc[(df_WLBI['XCoord'] == X) & (df_WLBI['YCoord'] == Y) & (df_WLBI['WAFER'] == wafer), :]
		# softbin = df_WLBI.loc[(df_WLBI['XCoord'] == X) & (df_WLBI['YCoord'] == Y) & (df_WLBI['WAFER'] == wafer), 'SoftBin'].tolist()
		if len(df) == 0:
			print(f"[ERROR] Wafer{wafer}|Die ({X},{Y}): Do NOT found data in the WLBI device")
			continue

		softbin = df['SoftBin'].tolist()[0]
		delta_igss = df['IGSS_25V_POST'].tolist()[0]/df['IGSS_25V_PRE'].tolist()[0]
		delta_idss = df['IDSS_1200V_POST'].tolist()[0]/df['IDSS_1200V_PRE'].tolist()[0]

		if softbin != 1:
			print(f"[ERROR] Wafer{wafer}|Die({X},{Y}): FAIL in the WLBI device")
		elif delta_igss > 5:
			print(f"[ERROR] Wafer{wafer}|Die({X},{Y}): IGSS exception during WLBI test (IGSS increase from {round(df['IGSS_25V_PRE'].tolist()[0], 4)} to {round(df['IGSS_25V_POST'].tolist()[0], 4)})")
		elif delta_idss > 5:
			print(f"[ERROR] Wafer{wafer}|Die({X},{Y}): IDSS exception during WLBI test (IDSS increase from {round(df['IDSS_1200V_PRE'].tolist()[0], 4)} to {round(df['IDSS_1200V_POST'].tolist()[0], 4)})")
df1, df2 = getDataFromFolder()
print(df1)
# checkError(df1, df2)

# a=getColName(r'C:\Users\Wufanzhengshu\Desktop\3193\\CP\GTA-S1M013120B1-A0,GTA-S1M014120B-B0016382,04.csv')
# print(a)