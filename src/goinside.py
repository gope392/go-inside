from openpyxl import load_workbook
import sys

#import xlsx file
#xls_file = input('Enter Excel file path')
xls_file = sys.argv[1]
workbook = load_workbook(xls_file)
worksheets = workbook.worksheets

for worksheet in worksheets:
    workbook.active = worksheet
    print(f'Worksheet: {workbook.active.title}')
    columns = list(workbook.active.iter_cols(max_row=1))
    for column in columns:
        cell, *other_cells = column
        print(f'Cell {cell.column_letter}{cell.col_idx} = {cell.value}')


#Identify worksheets
#wsheets = wbook.worksheets

#for wsheet in wsheets:
#    print 
#    print('max row: ' + whseet.max_row)

#print('Worksheets:')
#for name in wbook.sheetnames:
#    print(name)

