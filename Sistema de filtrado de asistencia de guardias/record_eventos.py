import openpyxl
import datetime
import os.path as path
import os

#variables
nombres =[]
ruts = []
record = []
data = []

#ruta de archivos
folder = r"C:\Users\acer\Desktop\practica pretorianos\eventos"

#ciclo para leer todos los archivos de la ruta
for filename in os.listdir(folder):
    if filename.endswith(".xlsx"):  # Procesar solo archivos .xlsx
        filepath = os.path.join(folder, filename)
        
        excel = openpyxl.load_workbook(filepath) #abre excel
        
        hoja = excel.active #abre la hoja de excel
        
        evento = [hoja.cell(row=2,column=1).value] #se rescata el titulo del evento
        print(evento[0])

        nombres = [hoja.cell(row=i,column=3).value for i in range(5,hoja.max_row+1)] #se rescata el nombre de los gaurdias
        print("Nombres: ",nombres)

        ruts = [hoja.cell(row=i,column=4).value for i in range(5,hoja.max_row+1)] #se rescata ruts de los guardias
        print("Ruts: ",ruts)

        asistencia = [hoja.cell(row=i,column=2).value for i in range(5,hoja.max_row+1)] #se rescata asistencia al evento de los guardias
        print("Asistencia (1 = si | 2 = no): ", asistencia)
            
        print("-"*50)
            
        if len(data) == 0:
            data = [nombres, ruts, record]
        else:
            if len(data[0]) != len(nombres):  # Asegúrate de que data[0] y nombres tienen longitudes diferentes
                for i in range(len(nombres)):
                    if nombres[i] not in data[0]:  # Solo añade si no está ya en data[0]
                        print("añadiendo:", nombres[i])
                        data[0].append(nombres[i])  # Añade el nombre
                        data[1].append(ruts[i])     # Añade el rut correspondiente
                        if record[i] == None:
                            record.insert(i, 0)
                        data[2].append(record[i])
        excel.close()
        excel.save(filename)
    for i in range(0, len(data[0])):
        if len(record) == 0 :
            record.append(0)
        if asistencia[i] == 1:
            if record[i] == 0:
                record.insert(i, 1)
            else:
                record[i] = record[i] + 1
if len(record) > len(data[0]):
    record.pop(-1)

print(data[0])
#print(data[2])
data[2].pop()
print(data[2])

#Ordenar de mayor a menor
indices_ordenados = sorted(range(len(data[2])), key=lambda i: data[2][i], reverse=True)

# Aplicamos el orden a cada sublista
data_ordenada= [[sublista[i] for i in indices_ordenados] for sublista in data]

#Ordenar de menor a mayor
indices_ordenadosinv = sorted(range(len(data[2])), key=lambda i: data[2][i], reverse=False)

# Aplicamos el orden a cada sublista
data_ordenadainv= [[sublista[i] for i in indices_ordenadosinv] for sublista in data]

fecha = datetime.date.today()
dia = str(fecha.year) + str(fecha.month) + str(fecha.day)
print("Codigo de dia: ", dia)

#Excel ordenado de mas asistentes a menos
if path.exists(f"Record_guardias{dia}.xlsx"):
    guardias = openpyxl.load_workbook(f"Record_guardias{dia}.xlsx")
else:
    guardias = openpyxl.Workbook()

hoja = guardias.active

hoja['A1'] = 'Nombre'
hoja['B1'] = 'Rut'
hoja['C1'] = 'Cantidad de eventos asistidos'

for i in range(0, len(data[0])):
    hoja.cell(row = i + 2, column = 1, value = data_ordenada[0][i])
    hoja.cell(row = i + 2, column = 2, value = data_ordenada[1][i])
    hoja.cell(row = i + 2, column = 3, value = data_ordenada[2][i])

guardias.save(f"Record_guardias{dia}.xlsx")

#Excel ordenada de menos asistentes a mas asistentes
if path.exists(f"Menos_asistencia_guardias{dia}.xlsx"):
    guardias = openpyxl.load_workbook(f"Menos_asistencia_guardias{dia}.xlsx")
else:
    guardias = openpyxl.Workbook()

hoja = guardias.active

hoja['A1'] = 'Nombre'
hoja['B1'] = 'Rut'
hoja['C1'] = 'Cantidad de eventos asistidos'

for i in range(0, len(data[0])):
    hoja.cell(row = i + 2, column = 1, value = data_ordenadainv[0][i])
    hoja.cell(row = i + 2, column = 2, value = data_ordenadainv[1][i])
    hoja.cell(row = i + 2, column = 3, value = data_ordenadainv[2][i])

guardias.save(f"Menos_asistencia_guardias{dia}.xlsx")