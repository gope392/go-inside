from openpyxl import load_workbook
from collections import Counter
import sys
from mapping import *
from classes import Ticket

#import xlsx file
#xls_file = input('Enter Excel file path')
xls_file = sys.argv[1]
workbook = load_workbook(xls_file)

tickets = []

for worksheet in workbook.worksheets:
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        ticket = Ticket(Ticket_Id=row[Ticket_Id],
                        Tipo=row[Tipo],
                        Resumen=row[Resumen],
                        Prioridad=row[Prioridad],
                        Estado=row[Estado],
                        Fecha_de_creacion=row[Fecha_de_creacion],
                        Fecha_de_cambio=row[Fecha_de_cambio],
                        Fecha_de_resolucion=row[Fecha_de_resolucion],
                        Resolucion=row[Resolucion],
                        Solicitante=row[Solicitante],
                        Asignado=row[Asignado],
                        Descripcion=row[Descripcion],
                        Fecha_estimada=row[Fecha_estimada],
                        Worksheet=worksheet.title)
        tickets.append(ticket)

def attr_count(attr, obj_list):
    attr_list = [obj.attr for obj in obj_list]
    attr_counts={}
    for value in attr_list:
        attr_counts[value] = attr_counts.get(value, 0) + 1
    print(f'---{attr} total---')
    for value, count in attr_counts.items():
        print(f'{value} : {count}')

attr_count('Tipo', tickets)
#types = [obj.Tipo for obj in tickets] 
#statuses = [obj.Estado for obj in tickets]
#list = types + statuses

#value_counts={}

#for value in list:
#    value_counts[value] = value_counts.get(value, 0) + 1

#for value, count in value_counts.items():
#    print(f'{value} : {count}')
#def cell_value(cell):
#    return cell.value

#for worksheet in workbook.worksheets:
#    print(f'Worksheet: {worksheet.title} - Dimensions: {worksheet.dimensions}')
#    column_titles = next(worksheet.rows)
#    for title in column_titles:
#        if (title.value.lower() == 'estado' or title.value.lower() == 'tipo'):
#            status_col = list(worksheet[title.column_letter])
#            status_col.pop(0)
#            status = map(cell_value, status_col)
#            counts = Counter(status)
#            for value, count in counts.items():
#              print(f'{title.value} {value}: {count}')
