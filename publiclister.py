#!/usr/bin/python
# -*- coding: utf-8 -*-

# on debian "apt-get install python-xlrd"
# everywhere "pip install xlrd"
import xlrd

# opening the list document and loading the first sheet
document = xlrd.open_workbook('pub.xlsx')
print("pub.xlsx loaded.")
sheet = document.sheet_by_index(0)
rows = sheet.nrows
print("Table with " + str(rows) + " rows loaded.")

current_row = 1

while(current_row < rows):
	# print sheet.row(current_row)
	if (sheet.cell_type(current_row, 7) != 0) or (sheet.cell_type(current_row, 6) != 0):
		print ("Year: " + str(sheet.cell_value(current_row, 5))),
		print ("Type: " + sheet.cell_value(current_row, 6).encode('utf8')),
		print ("Author: " + sheet.cell_value(current_row, 7).encode('utf8')),
		print ("Title: " + sheet.cell_value(current_row, 12).encode('utf8'))
	else:
		print "NEW YEAR"
	current_row += 1
