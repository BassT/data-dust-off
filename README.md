data-dust-off
=============
Clean data from messy, inconsistent excel sheet with this python library.

Introduction
------------
This library emerged out of a *Data for GOOD* DataThon in Toronto in March 2014. 
Most of the time was spent converting and cleaning messy excel sheets from NPOs and NGOs.
In order to reduce this time in the future, a simple library which speeds up this process is desirable.  
Due to the heterogenity of the data sources, we assume that the conversion and cleaning process *can't*
be done 100% automatically. Therefore, data scientist doing GOOD will be asked to support the automatic
process when necessary.

Dependencies
------------
We're using [openpyxl](http://pythonhosted.org/openpyxl/) to read/write Excel 2007 xlsx/xlsm files.
	
	pip install openpyxl

Functions
---------
1. [Convert string values to number values](#convert-string-values-to-number-values)
2. [Clean category data](#clean-category-data)

### Convert string values to number values
This method converts numbers of an excel-sheet which are typed as strings to numbers that are actually 
typed as numbers. Reasons for numbers typed as strings might be unnecessary spaces or currency signs.

### Clean category data
This methods cleans a column with messy categories.  
*Example:*  
  Say you have an array like ['class 1', 'class   1', 'clas 2', 'class 2'],
  in this case you can use this method in order to get a clean array like
  ['class 1', 'class 1', 'class 2', 'class 2'] which can be used for data
  analysis.

Example usage
-------------
First, we need to import the module:  

	import dustoff as do

Second, we want to import the selection of an excel sheet which we want to clean:

	session = do.import_col('/your/path/to/file.xlsx')

In doing so, you'll be asked to specify the desired sheet, column, start and end row of your selection. Example:

	Available sheets: [Sheet 1, Sheet 2]
	Enter the name of the sheet you want to select (case-sensitive): Sheet 1
	Enter the name of the column you want to select: B
	Enter the first row you want to include in the selection: 4
	Enter the last row you want to include in the selection: 8
	Data extracted from your selection:
	Data[0]: 100.00
	Data[1]: Less than 500
	Data[2]: 300$
	Data[3]: 200
	Data[4]: 150.00

The selected data is typed as strings, but in this case we would like to convert them to numerical values:

	session['data'] = do.convert_str_to_num(session['data'])

	Converting data[0]: 100.00
	Converted data[0] to: 100.00 (type: float)

	Converting data[1]: Less than 500
	Could not convert 'Less than 500' to a number
	...

In case the conversion can't be done automatically, you'll be asked to help the program out.

	...
	Please enter a number or leave blank for a missing value: 
	Converted data[1] to: missing value
	...

Here, we decide to interpret this vague data point as a missing value. The automatic conversion continues:

	Converting data[2]: 300$
	Converted data[2] to: 300.00 (type float)

	Converting data[3]: 200
	Converted data[3] to: 200.00 (type float)

	Converting data[4]: 150.00
	Converted data[4] to: 150.00 (type float)

Finally, the conversion is done and we can save the converted excel sheet as a new revision:

	do.save_rev(session)
	Saved new revision to '/your/path/to/file - rev from 2014-03-18 12-26-00.xlsx' with converted values for column 'B'.