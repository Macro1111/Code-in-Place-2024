# Version en Español 
## Automatizador de cuentas de cobro

# Introducción:
El Programa esta diseñado para la generación automatizada de cuentas de cobro utilizando una plantilla en formato Word y completandola con datos extraidos de un documento en Excel. Esta tarea se lleva a cabo mediante el uso de diversas bibliotecas, cada una especializada en una tarea específica. Entre las principales bibliotecas utilizadas se encuentran pandas para la manipulación de datos en formato Excel, docxtpl para la manipulación de plantillas de Word, y num2words para la conversión de valores numéricos a palabras en español.


## Función main:
1. Inicialización:
El script inicia ejecutando la función main, que a su vez inicia llamando a otras funciones esenciales. variables_documento() se encarga de configurar las variables necesarias para el documento, mientras que fecha() establece la configuración regional y obtiene la fecha actual. Posteriormente, se solicita al usuario el nombre del documento final.

2. Procesamiento de la Plantilla:
El script carga una plantilla de documento Word utilizando la biblioteca DocxTemplate. Para completar esta plantilla, se definen los datos en un diccionario llamado data. Este diccionario contiene información crucial que se utilizará para llenar la plantilla, como el nombre, la cédula, el correo, entre otros.

3. Renderización y Guardado:
La plantilla, ahora completada con los datos proporcionados, se renderiza y el documento final se guarda en una ruta específica en el escritorio del usuario. El nombre del archivo se compone del nombre ingresado por el usuario y lleva la extensión .docx. Finalmente, se imprime un mensaje indicando que el documento ha sido creado exitosamente.

## Función fecha:
Esta función desencadena la configuración regional para español mediante la función locale. Posteriormente, utiliza las bibliotecas calendar y datetime para obtener el nombre del mes actual, el día y el año. Esta información es relevante para incluirla en el documento final.

## Función variables_documento:

1. Manejo de la Entrada del Usuario:
Esta función utiliza dos bucles while para asegurar una entrada correcta del usuario. En primer lugar, solicita al usuario el número de línea que contiene la información en el archivo de Excel.

2. Obtención de Información del DataFrame:
Posteriormente, la función accede al DataFrame utilizando la biblioteca pandas. Se manejan posibles excepciones, como el ingreso de una línea que contiene nombres de columnas o la falta de información en las casillas correspondientes. Los datos obtenidos se almacenan en variables globales para su uso posterior en la generación del documento.

## Función nombre_nuevo_archivo:
Esta función recibe un nombre ingresado por el usuario y le añade la extensión .docx. El propósito es generar un nombre de archivo válido para el documento final.

## Función consecutivo:
Esta función calcula un número consecutivo sumando 300 a la línea de Excel proporcionada. El valor resultante se utilizará en la generación del documento como un identificador único. 

## Función numero_letras:
Utilizando la biblioteca num2words, esta función convierte un valor numérico a su representación en palabras en español, lo cual es útil para incluir cantidades en letras en el documento final.

## Ejecución Principal:
El código incluye una condición if __name__ == "__main__": que asegura que el script se esté ejecutando directamente y no siendo importado como un módulo. Si es el caso, se llama a la función main() para iniciar la ejecución del script.




# English Version
## Collection Account Automator

# Introduction:
The program is designed for the automated generation of collection accounts using a Word template and filling it with data extracted from an Excel document. This task is accomplished through the use of various libraries, each specialized in a specific task. Key libraries include pandas for Excel data manipulation, docxtpl for Word template manipulation, and num2words for converting numeric values to words in Spanish.

## Main Function:
1. Initialization:
The script starts by executing the main function, which in turn calls other essential functions. variables_documento() configures the necessary variables for the document, while fecha() sets the regional configuration and retrieves the current date. Subsequently, the user is prompted to enter the name of the final document.

2. Template Processing:
The script loads a Word document template using the DocxTemplate library. To complete this template, data is defined in a dictionary called data. This dictionary contains crucial information that will be used to fill the template, such as name, ID, email, among others.

3. Rendering and Saving:
The template, now filled with the provided data, is rendered, and the final document is saved to a specific path on the user's desktop. The filename is composed of the user-entered name and carries the .docx extension. Finally, a message is printed indicating that the document has been successfully created.

## Function fecha:
This function triggers the regional configuration for Spanish using the locale function. Later, it uses the calendar and datetime libraries to obtain the current month's name, day, and year. This information is relevant for inclusion in the final document.

## Function variables_documento:

1. User Input Handling:
This function uses two while loops to ensure correct user input. Firstly, it asks the user for the line number containing the information in the Excel file.

2. Retrieving DataFrame Information:
Subsequently, the function accesses the DataFrame using the pandas library. It handles possible exceptions, such as entering a line containing column names or missing information in the corresponding cells. The obtained data is stored in global variables for later use in document generation.

## Function nombre_nuevo_archivo:
This function takes a name entered by the user and adds the .docx extension. The purpose is to generate a valid file name for the final document.

## Function consecutivo:
This function calculates a consecutive number by adding 300 to the provided Excel line. The resulting value will be used in document generation as a unique identifier.

## Function numero_letras:
Using the num2words library, this function converts a numeric value into its representation in words in Spanish, useful for including quantities in letters in the final document.

## Main Execution:
The code includes an if __name__ == "__main__": condition to ensure the script is being run directly and not imported as a module. If so, the main() function is called to initiate script execution.