from openpyxl import load_workbook
import sys

#xls_file = input('Enter Excel file path')
xls_file = sys.argv[1]

wbook = load_workbook(filename=xls_file)
