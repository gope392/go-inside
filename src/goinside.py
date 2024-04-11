from openpyxl import load_workbook
import sys

#import xlsx file
#xls_file = input('Enter Excel file path')
xls_file = sys.argv[1]

wbook = load_workbook(xls_file)

print('Please Select Worksheet:')
for name in wbook.sheetnames:
    print(name)

def menu():
    print('select:')
    
