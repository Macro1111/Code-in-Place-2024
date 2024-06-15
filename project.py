# Importando las bibliotecas necesarias: pandas, openpyxl, docxtpl, num2words, calendar, datetime, locale
# Importing necessary libraries: pandas, openpyxl, docxtpl, num2words, calendar, datetime, locale
import calendar
import datetime
import locale
import pandas as pd
from docxtpl import DocxTemplate
from num2words import num2words

# Ruta al archivo Excel
# Path to the Excel file
excel_file_path = "Prueba base de datos.xlsx"

# Leyendo el archivo Excel en un DataFrame usando pandas
# Reading the Excel file into a DataFrame using pandas
df = pd.read_excel(excel_file_path)

def main():
    # Función principal del programa
    # Main function of the program
    variables_documento()
    fecha()
    consecutivo2 = consecutivo(linea_excel)

    # Solicitando al usuario que ingrese el nombre del documento final
    # Asking the user to input the final document name
    nombre_documento_final = nombre_nuevo_archivo(input("Ingresa el nombre del documento final: "))
    
    # Cargar la plantilla del documento Word
    # Load the Word document template
    doc = DocxTemplate("Prueba Template.docx")
    
    # Definir los datos para reemplazar en la plantilla
    # Define data to replace in the template
    data = {"Nombre": nombre, 
            "Cedula": cedula, 
            "Correo": correo, 
            "Valor": valor_numeros, 
            "Numero": celular, 
            "Direccion": direccion, 
            "Valorletras": valor_en_letras, 
            "Dia": fecha_dia, 
            "Mes": nombre_mes, 
            "Año": año, 
            "Consecutivo": consecutivo2}
    
    # Renderizar la plantilla con los datos
    # Render the template with the data
    doc.render(data)
    
    # Guardar el documento final
    # Save the final document
    doc.save(f"Documentos Finales\\{nombre_documento_final}")

    # imprime Documento creado exitosamente
    # print Document created successfully
    print("Documento creado exitosamente")
    

def fecha():
    # Configurar la configuración regional para español
    # Set the locale to Spanish
    locale.setlocale(locale.LC_TIME, 'es_ES.utf8')  # Puedes ajustar 'es_ES.utf8' según tu sistema operativo
    
    # Obtener el nombre del mes actual, día y año
    # Get the name of the current month, day, and year
    global nombre_mes; nombre_mes = (calendar.month_name[datetime.datetime.now().month]).capitalize()
    global fecha_dia; fecha_dia = datetime.datetime.now().day
    global año; año = datetime.datetime.now().year

def variables_documento():
    while True:
        while True:
            try:
                global linea_excel; linea_excel = int(input("Ingresa el número de la línea que contiene la información: "))
                linea_df = linea_excel - 2
                if linea_df < 0:
                    if linea_df == -1:
                        print("Esta línea contiene los nombres de las columnas, ingresa otro valor")
                        continue
                    else:
                        print("Línea no válida")
                        continue
                break
            except ValueError:
                print("La información ingresada no es un número")
                continue

        try:
            global nombre; nombre = (df.loc[linea_df, "Nombre"]).upper()
            global cedula; cedula = "{:,}".format(df.loc[linea_df, "Cedula"]).replace(",", ".")
            global correo; correo = df.loc[linea_df, "Correo"]
            global valor_numeros; valor_numeros = "{:,}".format(df.loc[linea_df, "Valor"]).replace(",", ".")
            global celular; celular = df.loc[linea_df, "Celular"]
            global direccion; direccion = df.loc[linea_df, "Direccion"]
            global valor_en_letras; valor_en_letras = numero_letras((df.loc[linea_df, "Valor"]))
            break

        except KeyError:
            print("No hay información en las casillas")
            # No information in the cells

def nombre_nuevo_archivo(s):
    # Agregar la extensión .docx al nombre ingresado por el usuario
    # Add the .docx extension to the user-entered name
    nombre_usuario = s + ".docx"
    return nombre_usuario

def consecutivo(s):
    # Calcular un número consecutivo sumando 300 al número de línea
    # Calculate a consecutive number by adding 300 to the line number
    consecutivo = s + 300
    return consecutivo

def numero_letras(s):
    # Convertir un valor numérico a su representación en palabras en español
    # Convert a numeric value to its representation in words in Spanish
    letras = num2words(s, lang='es').upper()
    return letras

if __name__ == "__main__":
    main()