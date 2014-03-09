import openpyxl as px
import re

def import_col(file):
	""" Opens an excel file and guides the user using prompts to select
	the desired cells.

	:param file: the path to the excel file
	"""

	data = [];

	wb = px.load_workbook(file)
	sheet_names = wb.get_sheet_names()
	print "Available sheets: [" + ", ".join(sheet_names) + "]"
	sheet = raw_input("Enter the name of the sheet you want to select (case-sensitive): ")
	while(sheet not in sheet_names):
		sheet = raw_input("Enter the name of the sheet you want to select (case-sensitive): ")
	ws = wb[sheet]
	col = raw_input("Enter the name of the column you want to select: ").upper()
	start = int(raw_input("Enter the first row you want to include in the selection: "))
	end = int(raw_input("Enter the last row you want to include in the selection: "))
	for i in range(start, end):
		print ws[col + str(i)].value
		data.append(ws[col + str(i)].value)

	return data


def convert_string_typed_number_to_number_typed_number(data_in):
	""" This method converts numbers of an excel-sheet which are typed as strings
	to numbers that are actually typed as numbers.
	Reasons for numbers typed as strings might be unnecessary spaces or currency
	signs.
	
	:param data_in: a string array
	"""

	data_out = []

	for i in range(0,len(data_in)):
		data_out.append(float(re.sub(r'\s*\$*\s*([0-9]{0,3}),*([0-9]{0,3}),*([0-9]{1,3})(\.[0-9]+)\s*', r'\1\2\3\4', data(i))))

	return data_out

data = import_col("/home/sebastian/Documents/Datathon TO/MASTER - Program Tracking SheetATMarch2014.xlsx")
print "Success"