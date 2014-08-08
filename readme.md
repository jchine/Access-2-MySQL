# MS Access to MySQL 

Partiendo de la información de la Base de datos de MS Access, se genera un 
fichero SQL imitando el resultado de la instrucción mysqldump de MySQL.
 
## USAGE

1. Descarga el archivo **exportSQLFile.bas**.
2. Importarlo dentro de los módulos de tú Base de Datos de MS Access. _Recuerda 
que has de entrar en Visual Basic del MS Access (Alt + f11) para poder 
importarlo correctamente_.
3. Utilizando cualquier formulario, o sino creando uno para la ocasión, crear un
botón e incluir la instrucción: 
> generateSQLFile

### generateSQLFile.bas

Es la función principal del archivo, el resto son utilidades, y se puede 
utilizar en cualquier acción de elemento o de formulario.
Por defecto utiliza:

- la Base de Datos del archivo de MS Access (currentDB)
- export.sql como nombre del archivo donde se generaran todas las intruccciones
 
Si se desa incluir una base de datos diferente o indicar un nombre especifico
para el fichero utilizar:
> generateSQLFile otherDB, "otherfile.sql"


Copyright (c) 2014 Jonathan Chiné

MIT Licensed