#!/usr/bin/python
# -*- coding: utf-8 -*-

import xlrd

# opening the list document and loading the first sheet
document = xlrd.open_workbook('pub.xlsx')
print("pub.xlsx loaded.")
sheet = document.sheet_by_index(0)
rows = sheet.nrows
print("Table with " + str(rows) + " rows loaded.")
