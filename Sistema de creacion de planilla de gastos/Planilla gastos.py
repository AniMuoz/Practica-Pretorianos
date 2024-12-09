import os
import openpyxl
import datetime
import os.path as path
from openpyxl.styles import Font

fecha = datetime.date.today()
dia = str(fecha.year) + str(fecha.month) + str(fecha.day)
print("Codigo de dia: ", dia)

#def gastos(dia):
#Excel ordenada de menos asistentes a mas asistentes
if path.exists(f"Detalle_gastos_pretorianos_seguridad_{dia}.xlsx"):
    guardias = openpyxl.load_workbook(f"Detalle_gastos_pretorianos_seguridad_{dia}.xlsx")
else:
    guardias = openpyxl.Workbook()

hoja = guardias.active

mes = input("Escriba el mes en el que quiere hacer la nomina: ")

hoja['B3'] = 'EMPRESA: PRETORIANOS SEGURIDAD'
hoja['B4'] = 'DIRECCIÓN: MANUEL BULNES Nº 920, OFICINA 208, QUILPUÉ'
hoja['B6'] = f'DETALLE GENERAL DE GASTOS MES {mes} DEL AÑO 2024'
hoja['B6'].font = Font(bold=True, size=12)

hoja['B8'] = 'CLASIFICACION DEL GASTO'
hoja['B8'].font = Font(bold=True, size=12)
hoja['C8'] = 'MONTO'
hoja['C8'].font = Font(bold=True, size=12)

hoja['B9'] = 'CONSUMOS BASICOS'		
hoja['B10'] = 'TELEFONO E INTERNET'		
hoja['B11'] = 'GASTOS COMUNES'		
hoja['B12'] = 'ARRIENDO DE OFICINA'		
hoja['B13'] = 'COMBUSTIBLE'		
hoja['B14'] = 'ESCRITORIO Y OFICINA'		
hoja['B15'] = 'ESTACIONAMIENTO'
hoja['B16'] = 'ARTICULOS DE ASEO'		
hoja['B17'] = 'GASTOS DE REPRESENTACIÓN'		
hoja['B18'] = 'VESTUARIO Y CALZADO'		
hoja['B19'] = 'PASAJES, PEAJES Y CORREOS'		
hoja['B20'] = 'MANTENIMIENTO, REPARACIÓN Y SEGURIDAD'		
hoja['B21'] = 'EQUIPAMIENTO'		
hoja['B22'] = 'ALIMENTACIÓN'	
hoja['B23'] = 'TOTAL GASTOS DEL MES'	

topicos = ['CONSUMOS BASICOS', 'TELEFONO E INTERNET', 'GASTOS COMUNES', 'ARRIENDO DE OFICINA', 'COMBUSTIBLE', 'ESCRITORIO Y OFICINA', 
           'ESTACIONAMIENTO', 'ARTICULOS DE ASEO', 'GASTOS DE REPRESENTACIÓN', 'VESTUARIO Y CALZADO', 'PASAJES, PEAJES Y CORREOS',
           'MANTENIMIENTO, REPARACIÓN Y SEGURIDAD', 'EQUIPAMIENTO', 'ALIMENTACIÓN']

for i in range (9, 23):
    valor = int(input(f"Monto de {topicos[i - 9]}: "))
    hoja.cell(row = i, column = 3, value = valor)

hoja['C23'] = '=SUM(C9:C22)'

hoja['D26'] = 'FREDDY ANDRES MUÑOZ OLIVARES'
hoja['D26'].font = Font(bold=True, size=12)
hoja['D27'] = 'GERENTE GENERAL'
hoja['D27'].font = Font(bold=True, size=12)

#i = 22
#hoja.cell(row = i + 2, column = 2, value = mes)
#hoja.cell(row = i + 2, column = 3, value = "$" + mes)

guardias.save(f"Detalle_gastos_pretorianos_seguridad_{dia}.xlsx")