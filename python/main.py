#!/usr/bin/python

from xlrd import open_workbook
# import sys
from sys import argv, exit
import json

from excell_parse import parse
from data_load import mongo_load
#############################
### This file takes source and 
### destination file arguments
### and asks for MongoDB admin
### validation username and 
### password to load directly
### mongodb instance
### 
### 
### KNOWN ISSUES:
### UTF ENCODING NOT WORKING FOR FOREIGN LANG utils.py 
### GEOCODING API LIMIT utils.py
#############################

def main(argv):

	arguments = len(argv)

	username = raw_input('Enter your MongoDB username [empty for no db]: ')
	if username != '':
		password = raw_input('Enter your MongoDB password: ')
		db = raw_input('Enter MongoDB you want to insert into: ')
		collection = raw_input('Enter ' + db + '.collection you want to insert into: ')

	data = []
	####ADD ERROR HANDLING TO CHECK FOR EXCEL
	wb = open_workbook(argv[0]+".xlsx")
	# if not <something to check for validity>:
	# 	print 'Please specify valid excel workbook'
	# 	sys.exit()
	for sheet in wb.sheets():
		data = parse(sheet, data)
	##Write out raw JSON file
	if arguments == 1:
		print_out = open(argv[0]+".json", "w")
		print_out.write(json.dumps(data, indent=4, separators=(',', ':')))
		print_out.close()
	elif arguments == 2:
		print_out = open(argv[1]+".json", "w")
		print_out.write(json.dumps(data, indent=4, separators=(',', ':')))
		print_out.close()

	# load into mongo i
	if username != '':
		mongo_load(data,db,collection,username,password)
		print "output stored in %s and MongoDB" %  argv[0]+".json"
	else:
		print "output stored in %s" % argv[0]+".json"

if __name__ == '__main__':
	main(argv[1:])