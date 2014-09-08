#!/usr/bin/python

from xlrd import cellname
from geopy.geocoders import GoogleV3


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

def geocode():
	# define geolatcaor scheme. GoogleV3 is open streetmaps but can also choose a number of other schemes. See docs.
	geolocator = GoogleV3()


		# ### ADD ADDRESS AND GEOLOCATION
		# ###
		# # main address, skip if NA
		# # if IOsheet.cell(row_index,5).value.encode('utf8') != 'n.a.':
		# # 	address = IOsheet.cell(row_index,5).value.encode('utf8')
		# # 	data[-1]["address"] = address

		# # 	# error handling for particularly problematic addresses
		# # 	if address =='':
		# # 		pass
		# # 	elif address == '43a Emek Refaim St., PO Box 8771, The German Colony, Jerusalem 91086, Israel ':
		# # 		location = geolocator.geocode('Emek Refaim St 43 Jerusalem, 93141')
		# # 	elif address == 'The University of Cambodia, No. 143 Preah Norodom Boulevard, PO Box 166, Phnom Penh 12000, Cambodia':
		# # 		location = geolocator.geocode('University of Cambodia, Northbridge Road, Sangkat Toek Thla, Khan Sen Sok, Phnom Penh, Kingdom of Cambodia')
		# # 	else:
		# # 		location = geolocator.geocode(address)
		# # 	if not location:
		# # 		print IOsheet.cell(row_index,0).value.encode('utf8')
		# # 		print type(location)

		# # 	data[-1]['loc'] = {'x':location.latitude, 'y':location.longitude}
		# # print data[-1]['name_en']

		# ######################
		# ### NEED TO HANDLE ARRAY
		# ### OF ALTERNATE ADDRESSES
		# ### SEPARATED BY COLUMN
		# ######################