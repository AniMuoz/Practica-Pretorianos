import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import datetime
import os

# Aquí va el código de tu lógica existente (el que proporcionaste)

nombres =[]
ruts = []
record = []
event = []
inasitencia = []
control_insasitencia = [[], [], [], []]  # Cuatro listas vacías
data = []

# Para simplificar, aquí solo se incluyen las funciones necesarias
# Asegúrate de incluir todas las funciones de tu código original

def busca_archivos(folder, data):
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
            
            # Procesar asistencia
            for i in range(len(asistencia)):
                if asistencia[i] == None:  # Si la asistencia es 0
                    inasitencia.append("Inasistencia")  # Añadir a la lista de inasistencia
                    control_insasitencia[0].append(nombres[i])
                    control_insasitencia[1].append(ruts[i])
                    control_insasitencia[2].append("Inasistencia")
                    control_insasitencia[3].append(evento[0])  # Usar el título del evento
                    print(f"Inasistencia al evento: {nombres[i]}")  # Imprimir encontrado
            print("-"*50)
            # Añadir datos a data
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

    data[2].pop()

    #Print de debuging
    #print("Guardias con inasistencia:", inasitencia)
    #print(data[0])
    #print(data[2])
    #print(data[2])
    #print(data[3])
    #print("\nDatos insacistencia: ", control_insasitencia)

    #Ordenar de mayor a menor
    indices_ordenados = sorted(range(len(data[2])), key=lambda i: data[2][i], reverse=True)

    # Aplicamos el orden a cada sublista
    data_ordenada= [[sublista[i] for i in indices_ordenados] for sublista in data]

    #Ordenar de menor a mayor
    indices_ordenadosinv = sorted(range(len(data[2])), key=lambda i: data[2][i], reverse=False)

    # Aplicamos el orden a cada sublista
    data_ordenadainv= [[sublista[i] for i in indices_ordenadosinv] for sublista in data]

    return data_ordenada,data_ordenadainv

def mas_asistencia(data_ordenada, dia):
    #Excel ordenado de mas asistentes a menos
    if os.path.exists(f"Record_guardias{dia}.xlsx"):
        guardias = openpyxl.load_workbook(f"Record_guardias{dia}.xlsx")
    else:
        guardias = openpyxl.Workbook()

    hoja = guardias.active

    hoja['A1'] = 'Nombre'   
    hoja['B1'] = 'Rut'
    hoja['C1'] = 'Cantidad de eventos asistidos'

    for i in range(0, len(data_ordenada[0])):
        hoja.cell(row = i + 2, column = 1, value = data_ordenada[0][i])
        hoja.cell(row = i + 2, column = 2, value = data_ordenada[1][i])
        hoja.cell(row = i + 2, column = 3, value = data_ordenada[2][i])

    guardias.save(f"Record_guardias{dia}.xlsx")

def menos_asistencia(data_ordenadainv, dia):
    #Excel ordenada de menos asistentes a mas asistentes
    if os.path.exists(f"Menos_asistencia_guardias_{dia}.xlsx"):
        guardias = openpyxl.load_workbook(f"Menos_asistencia_guardias_{dia}.xlsx")
    else:
        guardias = openpyxl.Workbook()

    hoja = guardias.active

    hoja['A1'] = 'Nombre'
    hoja['B1'] = 'Rut'
    hoja['C1'] = 'Cantidad de eventos asistidos'

    for i in range(0, len(data_ordenadainv[0])):
        hoja.cell(row = i + 2, column = 1, value = data_ordenadainv[0][i])
        hoja.cell(row = i + 2, column = 2, value = data_ordenadainv[1][i])
        hoja.cell(row = i + 2, column = 3, value = data_ordenadainv[2][i])

    guardias.save(f"Menos_asistencia_guardias_{dia}.xlsx")

def inasistencia(dia):
    #Excel ordenada de menos asistentes a mas asistentes
    if os.path.exists(f"Inasistencia_guardias_{dia}.xlsx"):
        guardias = openpyxl.load_workbook(f"Inasistencia_guardias_{dia}.xlsx")
    else:
        guardias = openpyxl.Workbook()

    hoja = guardias.active

    hoja['A1'] = 'Nombre'
    hoja['B1'] = 'Rut'
    hoja['C1'] = 'Evento'

    for i in range(0, len(control_insasitencia[0])):
        hoja.cell(row = i + 2, column = 1, value = control_insasitencia[0][i])
        hoja.cell(row = i + 2, column = 2, value = control_insasitencia[1][i])
        hoja.cell(row = i + 2, column = 3, value = control_insasitencia[3][i])

    guardias.save(f"Inasistencia_guardias_{dia}.xlsx")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Filtrado de Asistencia de Guardias")

        self.label = tk.Label(root, text="Indique la ruta donde están los eventos:")
        self.label.pack()

        self.folder_path = tk.StringVar()
        self.entry = tk.Entry(root, textvariable=self.folder_path, width=50)
        self.entry.pack()

        self.browse_button = tk.Button(root, text="Buscar Carpeta", command=self.browse_folder)
        self.browse_button.pack()

        self.process_button = tk.Button(root, text="Procesar Eventos", command=self.process_events)
        self.process_button.pack()

        self.result_label = tk.Label(root, text="")
        self.result_label.pack()

        self.option_var = tk.IntVar()
        self.option_label = tk.Label(root, text="Seleccione una opción:")
        self.option_label.pack()
        
        self.option1 = tk.Radiobutton(root, text="1. Obtener record de guardias", variable=self.option_var, value=1)
        self.option1.pack()

        self.option2 = tk.Radiobutton(root, text="2. Obtener guardias que menos asistencia tiene", variable=self.option_var, value=2)
        self.option2.pack()

        self.option3 = tk.Radiobutton(root, text="3. Que evento faltó cada guardia", variable=self.option_var, value=3)
        self.option3.pack()

        self.execute_button = tk.Button(root, text="Ejecutar Opción", command=self.execute_option)
        self.execute_button.pack()

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def process_events(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Error", "Por favor, seleccione una carpeta.")
            return
        # Aquí llamas a la función que procesa los archivos
        self.orden, self.ordeninv = busca_archivos(folder, data)
        messagebox.showinfo("Éxito", "Eventos procesados correctamente.")

    def execute_option(self):
        #dia = str(datetime.date.today().year) + str(datetime.date.today().month) + str(datetime.date.today().day)
        fecha = datetime.date.today()
        dia = str(fecha.year) + str(fecha.month) + str(fecha.day)
        print("Codigo de dia: ", dia)
        opcion = self.option_var.get()
        if opcion == 1:
            mas_asistencia(self.orden, dia)
            messagebox.showinfo("Éxito", "Se generó el archivo Excel con los datos de asistencia.")
        elif opcion == 2:
            menos_asistencia(self.ordeninv, dia)
            messagebox.showinfo("Éxito", "Se generó el archivo Excel con los datos de menos asistencia.")
        elif opcion == 3:
            inasistencia(dia)
            messagebox.showinfo("Éxito", "Se generó el archivo Excel con los datos de inasistencia.")
        else:
            messagebox.showerror("Error", "Por favor, seleccione una opción válida.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()

