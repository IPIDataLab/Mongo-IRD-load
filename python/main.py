#!/usr/bin/python

from xlrd import open_workbook
import sys
from sys import argv, exit
import json

from excell_parse import parse
#############################
### KNOWN ISSUES:
### UTF ENCODING NOT WORKING utils.py 
### GEOCODING API LIMIT utils.py
### 
### 
### 
#############################

def main(argv):

	data = []

	####ADD ERROR HANDLING TO CHECK FOR EXCEL
	wb = open_workbook(argv[0]+".xlsx")
	# if not <something to check for validity>:
	# 	print 'Please specify valid excel workbook'
	# 	exit()
	for sheet in wb.sheets():
		data = parse(sheet, data)


	# update = get_new_ppp_data()
	# norm_data = normalize(update)
	# mongo_load(norm_data, username, password)


	# username = raw_input('Enter your MongoDB username:')
	# password = raw_input('Enter your MongoDB password:')
	# if len(username) < 1 or len(password) < 1:
	# 	print "You must enter a username and password"
	# 	sys.exit()
	# else:
	# 	update = get_new_ppp_data()
	# 	norm_data = normalize(update)
	# 	mongo_load(norm_data, username, password)

	# print json.dumps(data[0], indent=4, separators=(',', ':'))

	print_out = open(argv[0]+".json", "w")
	print_out.write(json.dumps(data, indent=4, separators=(',', ':')))

	print_out.close()


if __name__ == '__main__':
	main(argv[1:])