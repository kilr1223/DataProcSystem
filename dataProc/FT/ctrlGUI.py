
title_page = r'''
 —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— 
|															   |
|    ______ ______                                       __    |
|       / ____//_  __/  _____ ___   ____   ____   _____ / /_   |
|      / /_     / /    / ___// _ \ / __ \ / __ \ / ___// __/   |
|     / __/    / /    / /   /  __// /_/ // /_/ // /   / /_     |
|    /_/      /_/    /_/    \___// .___/ \____//_/    \__/     |
|       _____  _        __      /_/   _                        |
|      / ___/ (_)_____ / /_   ____ _ (_)____                   |
|      \__ \ / // ___// __ \ / __ `// // __ \                  |
|     ___/ // // /__ / / / // /_/ // // / / /                  |
|    /____//_/ \___//_/ /_/ \__,_//_//_/ /_/                   |
|															   |
|						version: 1.0.1						   |
|															   |
|						programer: Kilr						   |
|															   |
 —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— —— 
'''

menu_page = r'''
Welcome to the Sichain data system, please select what function you what:
1.Get data report by lot number
2.Get data report by date
3.Get data summary by lot number
4.Get data summary by date
5.Configuration

Input ESC to quit...
'''

def getInput(word=None):
	lot = input('Please input the lot number you what to check:')
	print('Get data report by lot number')

def getReportByLot():
	lot = input('Please input the lot number you what to check:')
	print('Get data report by lot number')

def getReportByDate():
	date = input('Please input the start date you what to check:')
	print('Get data report by date')

def getSummaryByLot():
	lot = input('Please input the lot number you what to check:')
	print('Get data summary by lot number')

def getSummaryByDate():
	date = input('Please input the start date you what to check:')
	print('Get data summary by lot number')

def quit():
	exit()

func_dict = {
	'1':getReportByLot,
	'2':getReportByDate,
	'3':getSummaryByLot,
	'4':getSummaryByDate,
	'5':quit,
	'ESC':quit
}

print(title_page)
menu_input = input(menu_page)

while menu_input.upper() not in func_dict.keys():
	menu_input = input('Your input is invalid, please check and input agian:\n')

func_dict[menu_input]()