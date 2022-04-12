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

if "width" in settings and "height" in settings:
    workbook.set_size(settings["width"], settings["height"])

if "first_row_freeze" in settings and settings["first_row_freeze"]:
    worksheet.freeze_panes(1, 0)  # Freeze the first row.

if "first_row_bold" in settings and settings["first_row_bold"]:
    cell_format = workbook.add_format({'bold': True})
    worksheet.set_row(0, 20, cell_format) # for row 0, height=20, format=bold

if "first_row_autofilter" in settings and settings["first_row_autofilter"]:
    worksheet.autofilter(0,0,len(rows),len(row)-1)

if "column_widths" in settings:
    for c, col_width in enumerate(settings["column_widths"]):
        cell_format = workbook.add_format()
        
        if "text_wrap" in settings and settings["text_wrap"]:
            cell_format.set_text_wrap()

        worksheet.set_column(c, c, col_width, cell_format)

#############

workbook.close()