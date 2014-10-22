#!/usr/bin/python

from xlrd import cellname
from geopy.geocoders import GoogleV3
import re, json

#############################
#############################
# This file contains a number
# of utility functions used by
# excell_parse.py for data
# validation and error catching
#############################
#############################


###UTILITY FUNCTIONS FOR MONGO INIT PACKAGE
#checks for na and blank values

def na_check(sheet,row, column):
	val = sheet.cell(row,column).value

	# cast e.g. "12.0" as string
	if type(val) is float:
		val = "%.1f" % val

	if val == "n.a.":
		pass
	elif val == "": 
		pass
	else:
		return val.encode('utf8').strip()

#splits string into array
def split_str_array(string_in, delimiter):
	if not string_in:
		pass
	else:
		string_array = re.split(delimiter,string_in)
		return string_array

# # gets value from row
def get_cell(sheet, field, row_index, lkey, str_split = False):
	a = na_check(sheet, row_index, lkey[field.lower()])

	if type(str_split) is bool:
		if str_split:
			str_split=';'

	if not a:
		pass
	elif a == 'n.i.':
		a = 'No information'
	elif str_split:
		a = split_str_array(a, str_split)
		# remove trailing ';'
		a = [ i.strip(';').strip(' ') for i in a if i.strip(str_split).strip(' ') ]
	return a

def geocode(address):
	## load geocoding cache
	geolocator = GoogleV3()
	try:
		geocache = json.load(open('geocache.json'))
	except Exception as e:
		geocache = {}

	address = address.decode("utf8")
	if address == 'No information':
		return {}

	try:
		res = geocache[address]
		return res
	except KeyError:
		pass

	try:
		address1 = address
		address1 = re.compile('PO Box [0-9]+, ', re.IGNORECASE).sub('', address1)
		address1 = re.sub(r'The German Colony, ', '', address1)
		address1 = re.sub(r'The Middle East Council of Churches', '', address1)
		address1 = re.sub(r'The University of Cambodia, ', '', address1)
		address1 = re.sub(r'New Taipei City 228, China', 'New Taipei City 228, Taiwan', address1)
		address1 = re.sub(r'DB3 9BS', 'CB3 9BS', address1)
		address1 = re.sub(r'Regionalprogramm Politischer Dialog Westafrika', 'Les Cocotiers, Cotonou, Benin', address1)
		address1 = re.sub(r'P. O. Box: 811633 Amman 11181 Jordan', 'luzmila hospital, Amman 11181 Jordan', address1)
		if re.search(r'(150, route de Ferney|Route de Ferney 150)', address1):
			address1 = 'Route de Ferney 150, 1202 Geneve, Suisse'
		if re.search(r'(8 Montreal Road, Glen Anil)', address1):
			address1 = '8 Montreal Road, Glen Anil, Durban, South Africa'
		if re.search(r'(1004 Apartments.*Victoria Island, Lagos)', address1):
			address1 = '1004 Apartments, Victoria Island, Lagos'
		if re.search(r'(De Nobili College, Ramwadi, Pune)', address1):
			address1 = 'De Nobili College, Ramwadi, Pune, India'
		if re.search(r'(21st Km, Kanakapura Road, Udayapura, Bangalore)', address1):
			address1 = '21st Km, Kanakapura Road, Udayapura, Bangalore, India'
		if re.search(r'(Parmarth Niketan)', address1) and re.search(r'(Rishikesh)', address1):
			address1 = 'Parmarth Niketan Swargashram Rishikesh, india'
		if re.search(r'Fumon Hall.*Suginami', address1):
			address1 = '2 Chome-6-1 Wada, Suginami-ku, Tokyo-to, Japan'
		if re.search(r'(Sera Jey Monastic University, PO Bylakuppe - 571104)', address1):
			address1 = 'Sera Jey Monastic University, Bylakuppe - 571104, Mysore District, India'
		if re.search(r'(Dharmaram Vidya Kshetram, Pontifical Athenaeum of Philosophy, Theology and Canon Law, Bangalore-560 029, India)', address1):
			address1 = 'Dharmaram Vidya Kshetram, DVK Rd, Bhavani Nagar, Sadduguntepalya, Bengaluru, Karnataka 560029, India'
		print "=> %s" % address1
		location = geolocator.geocode(address1)
		print "== %s" % location.address
		res = { 'lat': location.latitude, 'lon': location.longitude, 'normalized': location.address }
		geocache[address] = res
		with open('geocache.json', 'w') as outfile:
			json.dump(geocache, outfile, sort_keys = True, indent = 4, ensure_ascii=True)
			#import ipdb; ipdb.set_trace()#
		return res

	except Exception as e:
		print e
		print u"!!ERROR: Couldn't geocode %s" % address
		#import ipdb; ipdb.set_trace()
		return {}

