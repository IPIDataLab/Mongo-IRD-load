#!/usr/bin/python

from xlrd import cellname
from geopy.geocoders import GoogleV3
import re, json

###UTILITY FUNCTIONS FOR MONGO INIT PACKAGE
#checks for na and blank values

def na_check(sheet,row, column):
	if sheet.cell(row,column).value.encode('utf8') == "n.a.":
		pass
	elif sheet.cell(row,column).value.encode('utf8') == "": 
		pass
	else:
		value = sheet.cell(row,column).value.encode('utf8')
		return value.strip()

#splits string into array
def split_str_array(string_in, delimiter):
	if not string_in:
		pass
	else:
		string_array = string_in.split(delimiter)
		return string_array

# # gets value from row
def get_cell(sheet, field, row_index, lkey, str_split = False):
	a = na_check(sheet, row_index, lkey[field.lower()])
	if not a:
		pass
	elif a == 'n.i.':
		a = 'No information'
	elif str_split:
		a = split_str_array(a, ';')
		# remove trailing ';'
		a = [ i.strip(';').strip(' ') for i in a if i.strip(';').strip(' ') ]
	return a

def geocode(address):
	## load geocoding cache
	geolocator = GoogleV3()
	try:
		geocache = json.load(open('geocache.json'))
	except Exception as e:
		geocache = {}

	address = address.decode("utf8")

	try:
		return geocache[address]
	except KeyError:
		pass

	try:
		address1 = address
		address1 = re.sub(r'PO Box [0-9]+, ', '', address1)
		address1 = re.sub(r'The German Colony, ', '', address1)
		address1 = re.sub(r'The University of Cambodia, ', '', address1)
		address1 = re.sub(r'New Taipei City 228, China', 'New Taipei City 228, Taiwan', address1)
		address1 = re.sub(r'DB3 9BS', 'CB3 9BS', address1)
		print address1
		location = geolocator.geocode(address1)
		res = { 'lat': location[1][0], 'lon': location[1][1] }
		geocache[address] = res
		with open('geocache.json', 'w') as outfile:
			json.dump(geocache, outfile, sort_keys = True, indent = 4, ensure_ascii=True)
			#import ipdb; ipdb.set_trace()#
		return res

	except Exception as e:
		print e
		print u"Couldn't geocode %s" % address
		#import ipdb; ipdb.set_trace()
		return {}

