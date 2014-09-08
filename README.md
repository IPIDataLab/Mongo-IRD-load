MongoDB IRD Data Initializer
==========
This is a package of utility functions that parse incoming KAICIID Excel data into JSON and load directly into a Mongo instance. This will be a one shot use when the IRD is migrated from shared Excel editing to a server-based Mongo instance with backend API access for new and updated records.

## Dependencies
-	[XLRD](http://www.python-excel.org/) - Python to Excel parsing library
-	[GeoPy](https://github.com/geopy/geopy) - Python client for geocoding web-services
-	json, datetime and re modules
 
## To Run from Command-line
1.	Requires separation of IRD data from IRD support worksheets and some amount of manual typo editing.	
2.	Open terminal and navigate to python subfolder in application folder.
3	Run ```python main.py path/to/file/<input filename> path/to/file/<target filename>(optional)``` (**Note: target file will be a json-ized version of input file if the second argument is left blank**)
4.	Terminal will prompt for MongoDB username, password, target database and target collection
5.	The package will write documents to target collection, raw JSON file and will verify with the number of documents inserted.

## TODO

