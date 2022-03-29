#!/usr/bin/env python3
# source: https://stackoverflow.com/questions/17684610/python-convert-csv-to-xlsx

import csv
from xlsxwriter.workbook import Workbook
from sys import argv

if len(argv)!=3:
  sys.exit("Usage: "+argv[0]+" input.csv output.xlsx")
input_filename = argv[1]
output_filename = argv[2] 

workbook = Workbook(output_filename)
worksheet = workbook.add_worksheet()
for r, row in enumerate(csv.reader(open(input_filename, 'rt', encoding='utf8'), delimiter=";")):
    for c, col in enumerate(row):
        worksheet.write(r, c, col)
workbook.close()