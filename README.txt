Para poder ejecutar nuestro script es necesario tener las librerias "os", "docx", "openpyxl", "pypdf2" las cuales nos sirven para lo siguiente:

os: os es una biblioteca estándar de Python que proporciona funciones para interactuar con el sistema operativo subyacente en el que se ejecuta Python. Proporciona métodos para trabajar con archivos, directorios, rutas de archivos, variables de entorno y más. En este código, utilizamos la función os.listdir() para obtener una lista de archivos en un directorio especificado, y os.path.join() para construir rutas de archivos cruz-plataforma.

docx (python-docx): python-docx es una biblioteca de Python que permite leer, escribir y manipular archivos de Microsoft Word en formato .docx. Con esta biblioteca, podemos extraer metadatos de archivos .docx, como el título, autor, asunto, palabras clave y comentarios. En el código, utilizamos la clase Document para abrir un archivo .docx y acceder a sus propiedades principales.

openpyxl: openpyxl es una biblioteca de Python para leer y escribir archivos de Excel en formato .xlsx. Con esta biblioteca, podemos acceder a las propiedades de un libro de Excel, como el título, autor, asunto, palabras clave y descripción. En el código, utilizamos la función load_workbook() para cargar un archivo .xlsx y acceder a sus propiedades.

PyPDF2: PyPDF2 es una biblioteca de Python para trabajar con archivos PDF. Con esta biblioteca, podemos extraer texto, metadatos y manipular archivos PDF de varias formas. En el código, utilizamos la clase PdfReader para abrir un archivo PDF y acceder a sus metadatos, como el título, autor, asunto y palabras clave.

datetime: datetime es una biblioteca estándar de Python que proporciona clases para manejar fechas y horas. En el código, utilizamos la clase datetime para formatear y manipular fechas y horas, especialmente para obtener la fecha de creación y modificación de un archivo PDF.
_______________________________________________________________________________________________________________________________________________________________________

El script esta diseñado para que lea los metadatos de la carpeta "Files" que esta en el mismo directorio que el archivo "get_meta.py", por lo tanto basta con ejecutar el script desde Visual Studio Code, o desde una linea de comandos de Windows o Linux, de esta forma se mostraran los datos extraidos de cada tipo de archivo.