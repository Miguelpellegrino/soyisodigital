import openpyxl
import pandas as pd
import os

filename = 'datos.xlsx'

def add():
    os.system('cls')
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook['Sheet1']
    name = input("Nombre del producto: ")
    amount = input("Precio del producto: ")
    inventory = input("Cantidad del producto: ")
    sheet.append([name,amount,inventory])
    workbook.save(filename)
    print("")
    print("Se agrego el producto correctamente")
    operation()

def update():
    os.system('cls')
    df = pd.read_excel(filename)
    print("")
    print(df)
    line = input("Ingrese el numero de la linea que desea editar: ")
    try:
        if int(line):
            df = pd.read_excel(filename, sheet_name='Sheet1')
            df.loc[int(line), 'Productos'] = input("Nombre del producto: ")
            df.loc[int(line), 'Precios'] = input("Precio del producto: ")   
            df.loc[int(line), 'Disponible'] = input("Cantidad del producto: ")
            df.to_excel(filename, sheet_name='Sheet1', index=False)
    except:
        print("Debe ingresar un numero")
        update()
    operation()
    






def record():
    os.system('cls')
    df = pd.read_excel(filename)
    print("")
    print(df)
    operation()

def exit():
    return True

def delete():
    df = pd.read_excel(filename)
    print("")
    print(df)
    df = pd.read_excel(filename)
    dels = input("Ingrese el numero de la linea que desea eliminar: ")
    try:
        if int(dels):
            df = df.drop(int(dels))
            df.to_excel(filename, index=False)
    except:
        print("Debe ingresar un numero")
        delete()
    os.system('cls')
    print("Se ha eliminado el registro seleccionado")
    operation()



def option(option):
    options = {
        '1':'add()', 
        '2':'delete()', 
        '3':'update()', 
        '4':'record()', 
        '5':'exit()'
    }
    try:
        if int(option):
            eval(options.get(option))
    except:
        print("")
        print("Debe ingresar un numero")
        operation()



def operation():
    number = input("""
    Ingrese una de las siguientes opciones (numericas)
    [1] Agregar
    [2] Eliminar
    [3] Actualizar
    [4] Ver Registros
    [5] Salir
    """)
    option(number)
    


def create_xls():
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    worksheet.title = 'Sheet1'

    worksheet['A1'] = 'Productos'
    worksheet['B1'] = 'Precios'
    worksheet['C1'] = 'Disponible'

    worksheet['A2'] = 'Pantalones'
    worksheet['A3'] = 'Camisas'
    worksheet['A4'] = 'Corbatas'
    worksheet['A5'] = 'Casacas'

    worksheet['B2'] = '200.00'
    worksheet['B3'] = '120.00'
    worksheet['B4'] = '50.00'
    worksheet['B5'] = '350.00'

    worksheet['C2'] = '50'
    worksheet['C3'] = '45'
    worksheet['C4'] = '30'
    worksheet['C5'] = '15'

    workbook.save(filename)

if os.path.isfile(filename):
    print(f"El archivo {filename} existe en el directorio actual.")
    operation()
else:
    create_xls()
    operation()


