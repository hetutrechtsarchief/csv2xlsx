#!/usr/bin/env python3

# source: https://stackoverflow.com/questions/17684610/python-convert-csv-to-xlsx

import csv,sys,json
from xlsxwriter.workbook import Workbook
from sys import argv
from tqdm import tqdm

delimiter=";"
encoding="utf-8"

if len(argv)<3:
  sys.exit("Usage: "+argv[0]+" input.csv output.xlsx [settings.json]")
input_filename = argv[1]
output_filename = argv[2] 
settings_filename = argv[3] if len(argv)>3 else None

workbook = Workbook(output_filename,
    {'strings_to_numbers': True}
)

worksheet = workbook.add_worksheet()
rows = [ row for row in csv.reader(open(input_filename, 'rt', encoding=encoding), delimiter=delimiter) ]
for r, row in enumerate(tqdm(rows)):
    for c, col in enumerate(row):
        worksheet.write(r, c, col)

# load settings if supplied
settings = json.load(open(settings_filename)) if settings_filename else None


#############
# onderstaande settings nu uit JSON lezen!!

# workbook.set_size(2000, 700)

# worksheet.freeze_panes(1, 0)  # Freeze the first row.
# cell_format = workbook.add_format({'bold': True})
# worksheet.set_row(0, 20, cell_format) # for row 0, height=20, format=bold

# # col: text
# cell_format = workbook.add_format()
# cell_format.set_text_wrap()
# worksheet.set_column(2, 2, 100, cell_format) # text
# worksheet.set_column(3, 3, 20, cell_format) # Datering
# worksheet.set_column(4, 4, 5, cell_format) # Aantal
# worksheet.set_column(5, 5, 20, cell_format) # UiterlijkeVorm
# worksheet.set_column(6, 6, 30, cell_format) # Notabene

# #col: code
# cell_format = workbook.add_format({'num_format': '@'})
# for r, row in enumerate(rows):
#     if r==0:
#         continue
#     worksheet.write_string(r, 1, row[1], cell_format)

# worksheet.autofilter(0,0,len(rows),len(row)-1)


#############

workbook.close()