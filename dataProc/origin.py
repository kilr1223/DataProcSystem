import originpro as op
import os, re, sys
import pandas as pd
from tkinter import filedialog

BASELINE_SHEETNAME = 'Baseline'
FAILURE_SHEETNAME = 'Failure'
TEMPLATE_NAME = 'autoGraph'
BASELINE_5_INDEX = 1
BASELINE_10_INDEX = 2
BASELINE_P20_INDEX = 3
BASELINE_N20_INDEX = 4

def origin_shutdown_exception_hook(self, exctype, value, traceback):
    '''Ensures Origin gets shut down if an uncaught exception'''
    op.exit()
    sys.__excepthook__(exctype, value, traceback)

class BaselineStyle():
	'''
	mode == 1: The default style is prepared for IDSS, IGSS or IR, 
		the baseline will be set to 5x with a x-axis from -0.5 to 5.5.
		The fluctuations regarding leakage current will be displayed in magnification style

	mode == 3: The parameter will be changed to match humidity test.
		For those tests with humidity, the baseline should be set to 10x with a y-axis from -1 to 11

	mode == 2 or 4: The parameter will be set as percentage.
		The others data will be showed as percentage	
	'''
	def __init__(self, mode):
		self.blueline_coly = BASELINE_N20_INDEX
		if mode == 1:
			self.toMultiple5()
		elif mode == 3:
			self.toMultiple10()
		else:
			self.toPercentage()

	def toMultiple5(self):
		self.redline_coly = BASELINE_5_INDEX
		self.y_begin = -0.5
		self.y_end = 5.5
		self.y_step = 1

	def toMultiple10(self):
		self.redline_coly = BASELINE_10_INDEX
		self.y_begin = -1
		self.y_end = 11
		self.y_step = 2

	def toPercentage(self):
		self.redline_coly = BASELINE_P20_INDEX
		self.y_begin = -25
		self.y_end = 25
		self.y_step = 10

class OriginObject():

	def __init__(self, path):

		if op and op.oext:
	  		sys.excepthook = origin_shutdown_exception_hook

	    # Set Origin instance visibility.
		# Important for only external Python.
		# Should not be used with embedded Python. 
		# if op.oext:
		#     op.set_show(True)

		self.path = path
		self.xls_file = os.path.join(path, 'Data.xlsx')
		self.wb = op.new_book(lname = 'Data', hidden = True)
		self.ws_baseline = self.wb.add_sheet(BASELINE_SHEETNAME)
		self.ws_failure = self.wb.add_sheet(FAILURE_SHEETNAME)

		self.getImgPath()
		self.getSheetList()
		self.createBaseLine()

	def getImgPath(self):

		img_file = os.path.join(self.path, 'graph')

		if not os.path.exists(img_file):
			os.mkdir(img_file)

		self.img_file = img_file
		return img_file

	def getSheetList(self):
		sheet_list = list(pd.read_excel(self.xls_file, sheet_name=None).keys())

		if 'DATA' in sheet_list:
			# del sheet_list['DATA']
			sheet_list.remove('DATA')

		self.sheet_list = sheet_list

		return sheet_list

	def createBaseLine(self):
		ws_baseline = self.ws_baseline
		ws_baseline.from_list(BASELINE_5_INDEX, [5,5], lname = 'MULTI5')
		ws_baseline.from_list(BASELINE_10_INDEX, [10,10], lname = 'MULTI10')
		ws_baseline.from_list(BASELINE_P20_INDEX, [20,20], lname = 'PERCENT-P')
		ws_baseline.from_list(BASELINE_N20_INDEX, [-20,-20], lname = 'PERCENT-N')

	def markErrorPoints(self, df, mode):
		ws_failure = self.ws_failure
		ws_failure.clear()
		
		if mode == 1:
			value_range = (0,5)
		elif mode == 3:
			value_range = (0,10)
		else:
			value_range = (-20,20)

		for index, value in df.items():
			if not value_range[0] < value < value_range[1]:
				ws_failure.from_list(ws_failure.cols, [index, index], axis = 'X')
				ws_failure.from_list(ws_failure.cols, [-30, 30], axis = 'Y')

	def drawImg(self):
		print('----------------Drawing Images----------------')   
		# print(self.sheet_list)
		# for sheet_name in ['14H-RB45']:
		for sheet_name in self.sheet_list:
			ws = self.wb.add_sheet(sheet_name)

			df = pd.read_excel(self.xls_file, sheet_name=sheet_name)
			
			test_num = re.findall(r'-(.*[A-Za-z])\d', sheet_name)[0]

			# 
			colname_list = list(df.columns)
			ws.from_df(df)
			start_index = int(df['INDEX'][0])
			data_count = ws.rows
			end_index = start_index + data_count - 1

			for col_name in colname_list[1:]:
				
				if all(key not in col_name for key in ['IDSS', 'IGSS', 'IR']):
					ws.set_label(col_name, val='Δ' + col_name + '  (%)', type = 'L')
				else:
					ws.set_label(col_name, val='Δ' + col_name, type = 'L')

			# df = df.set_index('INDEX')
			
			self.ws_baseline.from_list(0, [-50, end_index*2], lname = 'X')

			for column_name in colname_list[1:]:
				
				colname = re.findall('^(.*)@', column_name)[0]
				graph = op.new_graph(lname = test_num + '-' + colname, template = TEMPLATE_NAME)				
				graph[0].add_plot(ws, colx = 0, coly = colname_list.index(column_name), type = 's')

				mode = 0
				if test_num == 'F':
					mode = mode | 0b10
				if colname in ['IDSS', 'IGSS', 'IR']:
					mode = mode | 0b01
				bls = BaselineStyle(mode)

				ay = graph[0].axis('y')
				ay.set_limits(begin = bls.y_begin, end = bls.y_end, step = bls.y_step)
				graph[0].add_plot(self.ws_baseline, colx = 0, coly = bls.redline_coly, type = 'l')
				graph[0].add_plot(self.ws_baseline, colx = 0, coly = bls.blueline_coly, type = 'l')
				
				self.markErrorPoints(df[column_name], mode)

				for col_fail in range(0, self.ws_failure.cols, 2):
					fail_line = graph[0].add_plot(self.ws_failure, colx = col_fail, coly = col_fail + 1, type = 'l')
					fail_line.set_cmd('-d 0')
					fail_line.color = (255,0,0)
					# print(fail_line.shapelist)

				#Depending on the count of DUTs, the x-axis range will be changed
				ax = graph[0].axis('x')
				# ax.set_limits(begin = start_index - 1)
				if data_count > 80:
					ax.set_limits(begin = start_index - 5, end = end_index + 5, step = 30)
				elif data_count < 70:
					ax.set_limits(begin = start_index - 1, end = end_index + 1, step = 5)
				graph.save_fig(self.img_file, type = 'jpg')
				print(f'Graph exported: {os.path.join(self.img_file, graph.lname)}')
			# 	op.save(r'C:\Users\Wufanzhengshu\Desktop\new workbook.opju')

	
# print(graph[0].label('Line1').get_int('Width'))

# path = filedialog.askdirectory()
if op and op.oext:
    sys.excepthook = origin_shutdown_exception_hook

a = OriginObject(r'D:\可靠性实验\试验报告\S1M040120H\DATA')
print("Initing")

a.drawImg()

# op.save(r'C:\Users\Wufanzhengshu\Desktop\new workbook.opju')
if op.oext:
	op.exit()
