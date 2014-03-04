#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
on debian "apt-get install python-xlrd"
everywhere "pip install xlrd"
'''
import xlrd
import xlsxwriter


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
    worksheet.write(rowcounter, 0, repr(sheet.cell_value(current_row, 5)).split(".")[0])
    print ("Year: " + str(sheet.cell_value(current_row, 5)).split(".")[0]),
    worksheet.write(rowcounter, 1, repr(sheet.cell_value(current_row, 6).encode('utf8')))
    print ("Type: " + sheet.cell_value(current_row, 6).encode('utf8')),
    worksheet.write(rowcounter, 2, repr(sheet.cell_value(current_row, 7).encode('utf8')))
    print ("Author: " + sheet.cell_value(current_row, 7).encode('utf8')),
    worksheet.write(rowcounter, 3, repr(sheet.cell_value(current_row, 12).encode('utf8').replace('\n', '').replace('  ', ' ').replace('\"', '')))
    print ("Title: " + sheet.cell_value(current_row, 12).encode('utf8').replace('\n', '').replace('  ', ' ').replace('\"', ''))
    rowcounter = rowcounter + 1
  else:
    print "NEW YEAR"
    rowcounter = 1
    newTable = xlsxwriter.Workbook(str(sheet.cell_value(current_row, 5))[:4] + ".xlsx")
    worksheet = newTable.add_worksheet()
    worksheet.write(0, 0, str(sheet.cell_value(current_row, 5)).split(".")[0])

  current_row += 1

