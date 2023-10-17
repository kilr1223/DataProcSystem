# -*- coding: utf-8 -*-
import configparser, json

CONF_FILE = './Configuration.ini'
DEVICE_FLAG = 'DEVICE'
LOT_FLAG = 'LOT'
COL_FLAG = 'SITE_NUM'
TEST_FLAG = 'TEST_SITE'
PATH_FLAG = 'PATH'

conf = configparser.ConfigParser()
conf.read(CONF_FILE, encoding="utf-8-sig")
FILENAME_MOD = conf.get('setting', 'filename_mod')
DATA_DIR = conf.get('setting', 'data_dir')
SAVE_DIR = conf.get('setting', 'save_dir')
UPDATE_TIME = conf.get('setting', 'last_time')
TEST_KEY = json.loads(conf.get('setting', 'test_key'))
SUMMARY_LIST = json.loads(conf.get('report', 'summary_list'))
BIN_DICT = json.loads(conf.get('report', 'BIN'))
PARA_KEY = json.loads(conf.get('report', 'para_key'))
USERNAME = conf.get('email', 'username')
PASSWORD = conf.get('email', 'password')
RECCIVER = json.loads(conf.get('email', 'recciver'))
CC = json.loads(conf.get('email', 'cc'))

