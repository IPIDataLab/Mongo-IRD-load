#!/usr/bin/python

import pymongo
from pymongo import MongoClient

#############################
#############################
# This file loads normalized
# IRD data into the mongodb
#############################
#############################

def mongo_load(data,db_name,collection_name,username,password):

	# Make db connection
	client = MongoClient('localhost',27017)
	db = client[db_name]
	collection = db[collection_name]

	try:
		db.authenticate(username,password)
		print "Authenticated Mongo connection."

		collection.insert(data)

		print str(len(data)) + " documents inserted into " + collection_name + " collection in the " + db_name + " database."
	
	except:
		print "Wrong username or password!"


if __name__ == '__main__':
	mongo_load(data,db_name,collection_name,username,password)