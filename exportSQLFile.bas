Option Compare Database

' This file is part of the Access 2 MySQL.
'
' (c) 2014. Jonathan Chiné
'
' For the full copyright and license information, please view the LICENSE
' file that was distributed with this source code.

' *****************************************************************************
' USAGE
'
'   En cualquier acción de elemento o de formulario incluir la llamada a la
'   instrucción para generar el fichero SQL:
'       generateSQLFile
'
'   Si se desa incluir una base de datos diferente o indicar un nombre especifico
'   para el fichero utilizar:
'       generateSQLFile otherDB, "otherfile.sql"

' *****************************************************************************
' Definition default params
'
' -------------- change these constants before use!
Const CONFIG_COMMENTS_SQL = True                            ' Include comments before SQL statements
Const CONFIG_DROP_TABLE = True                              ' Include Drop SQL statements
Const CONFIG_USED_RELATIONS = True                          ' Includes SQL statements to insert values into tables
Const CONFIG_LOCK_TABLE = True                              ' Lock of table to insert values
Const CONFIG_INSERT_VALUES = True                           ' Includes SQL statements to insert values into tables

Const CONFIG_OUTPUT_FILENAME = "export"                     ' Default File name
Const CONFIG_OUTPUT_FILE_PATH = ""                          ' Current DB path
Const CONFIG_OUTPUT_FILE_EXTENSION = "sql"                  ' Default extension for dump SQL

' -------------- utilizamos para poder exportar los archivos incluidos en la BD!
'               (imagenes, audios, videos, words...)
Const CONFIG_EXPORT_FILES = False                           ' Export Ole Objetcs as file
Const CONFIG_BLOCK_SIZE = 32768                             ' NOT CHANGE
Const CONFIG_EXPORT_FILES_EXTENSION = "jpg"                 ' Extension of File
Const CONFIG_FILES_DIRECTORY = CONFIG_OUTPUT_FILE_PATH _
                            & "/files"                      ' Directory to save files from OLE Objects
                             

' Summanry
'   Generamos y guardamos las instrucciones SQL
'
' Params
'   Database db        OPTIONAL Dababase for generate dump sql
'   String strFileName OPTIONAL Output filename with ".sql"
Public Sub generateSQLFile(Optional db As Database, Optional strFileName As String)
    Dim numFile As Integer
    Dim strPath As String
    Dim line As Variant
    Dim listLines As Collection, listSQL As Collection
    Dim table As TableDef
        
    If (IsMissing(db) Or (db Is Nothing)) Then
        Set db = CurrentDb
    End If
    
    If IsMissing(strFileName) Or ("" = strFileName) Then
        If "" = Trim(CONFIG_OUTPUT_FILENAME & "") Then
            ' Utilizamos el nombre de la Base de datos
            strFileName = GetNameDB(db.Name)
        Else
            strFileName = CONFIG_OUTPUT_FILENAME
        End If
        
        strFileName = strFileName & "." & CONFIG_OUTPUT_FILE_EXTENSION
    End If
    
    ' Establecemos el indicador del archivo
    numFile = FreeFile
    
    ' Contruimos el nombre y la ruta del fichero SQL
    If 0 = InStr(strFileName, "." & CONFIG_OUTPUT_FILE_EXTENSION) Then
        strFileName = strFileName & "." & CONFIG_OUTPUT_FILE_EXTENSION
    End If
    If ("" = Trim(CONFIG_OUTPUT_FILE_PATH & "")) Then
        strPath = CurrentProject.Path
    Else
        strPath = Trim(CONFIG_OUTPUT_FILE_PATH)
    End If
    strPath = strPath & "\" & UCase(strFileName)
    
    ' Eliminamos el fichero si existe
    DeleteFile strPath
    
    Set listLines = New Collection
    
        If CONFIG_COMMENTS_SQL Then
            listLines.Add "--"
            listLines.Add "-- Base de datos: `" & GetNameDB(db.Name) & "`"
            listLines.Add "--"
            listLines.Add ""
        End If
        
        listLines.Add "CREATE DATABASE IF NOT EXISTS `" & GetNameDB(db.Name) & "`"
        listLines.Add "CHARACTER SET `utf8`;"
        listLines.Add ""
        listLines.Add "USE `" & GetNameDB(db.Name) & "`;"

        If CONFIG_COMMENTS_SQL Then
            listLines.Add "-- --------------------------------------------------------"
        End If

        For Each table In db.TableDefs
            ' Descartamos las tablas del sistema de MS Access "MSys..."
            If ("MSys" <> Left(table.Name, 4)) Then
                Set listSQL = printTable(db, table)
                For Each line In listSQL
                    listLines.Add "" & line
                Next
            End If
        Next table

        Set listSQL = PrintRelations(db)
        For Each line In listSQL
            listLines.Add "" & line
        Next
        
        writeOut listLines, strPath
        
    MsgBox ("Fichero '" & strFileName & "' creado en la carpeta " & CONFIG_OUTPUT_FILE_PATH _
            & " o en su defecto en la carpeta de la Base de Datos (" & CurrentProject.Path & ")")
End Sub

' Summanry
'   Generamos la lista con las instrucciones SQL necesarias para poder crear la tabla y
'   sus valores.
'
' Params
'   Database db     La Base de Datos
'   TableDef Table  La tabla de la base de datos de la cual crearemos las instrucciones SQL
'
' Return
'   Collection Lista con las instrucciones SQL
Public Function printTable(db As Database, table As TableDef) As Collection
    Dim strSQL As String, strTableName As String, strFieldName As String, strValues As String, strFields As String
    Dim i, pos As Integer
    Dim boolIndex As Boolean
    Dim strVariant As Variant
    Dim Field As Field
    Dim rec As Recordset
    Dim indexSQL As Collection, txtSQL As Collection
    
    Set txtSQL = New Collection
    strTableName = "`" & HelperClearCharacters(table.Name) & "`"
    
    If CONFIG_COMMENTS_SQL Then
        txtSQL.Add "--"
        txtSQL.Add "-- Estructura de tabla para la tabla " & strTableName
        txtSQL.Add "--"
    End If
    
    If CONFIG_DROP_TABLE Then
        txtSQL.Add ""
        txtSQL.Add "DROP TABLE IF EXISTS " & strTableName & ";"
        txtSQL.Add ""
    End If
        
        strSQL = "CREATE TABLE IF NOT EXISTS " & strTableName & " (" & vbCr
        pos = 0
        For Each Field In table.Fields
            strFieldName = "`" & HelperClearCharacters(Field.Name) & "`"
            If (pos > 0) Then
                strFields = strFields & ", "
            End If
            strFields = strFields & strFieldName
            
            strSQL = strSQL & strFieldName & " " & FieldTypeName(Field)
            If (IsNull(Field.DefaultValue) Or (Trim(Field.DefaultValue) <> "")) Then
                strSQL = strSQL & " DEFAULT '" & Field.DefaultValue & "'"
            End If
            
            If (Field.Required) Or (Not (Field.Attributes And dbAutoIncrField) = 0&) Then
                strSQL = strSQL & " NOT NULL"
            Else
                strSQL = strSQL & " NULL"
            End If
            
            If (pos < table.Fields.Count - 1) Or (table.Indexes.Count > 0) Then
                strSQL = strSQL & ","
            End If
            strSQL = strSQL & vbCr
            pos = pos + 1
        Next Field
        
        
        Set indexSQL = PrintIndex(table)
        If (indexSQL.Count > 0) Then
            For pos = 1 To indexSQL.Count
                strSQL = strSQL & indexSQL.Item(pos)
                If (pos < indexSQL.Count) Then
                    strSQL = strSQL & ","
                End If
                strSQL = strSQL & vbCr
            Next pos
        End If
        
        strSQL = strSQL & ") ENGINE=InnoDB DEFAULT CHARSET=utf8;" & vbCr
        txtSQL.Add strSQL
        
    If CONFIG_COMMENTS_SQL And CONFIG_INSERT_VALUES Then
        txtSQL.Add "--"
        txtSQL.Add "-- Dumping data for table " & strTableName
        txtSQL.Add "-- Volcado de datos para la tabla " & strTableName
        txtSQL.Add "--"
    End If
    
    If CONFIG_INSERT_VALUES Then
        Set rec = db.OpenRecordset(table.Name, dbOpenTable)
        If rec.EOF And rec.BOF Then
            ' Empty values in table
            txtSQL.Add "-- Empty values"
        Else
            strSQL = "INSERT INTO " & strTableName & " (" & strFields & ") VALUES" & vbCr
            
            rec.MoveFirst
            pos = 0
            Do While Not rec.EOF
                strValues = ""
                For i = 0 To rec.Fields.Count - 1
                    'If ((I > 0) And (I < rec.Fields.Count - 1)) Then
                    If (i > 0) Then
                        strValues = strValues & ", "
                    End If
                    
                    If (Not IsNull(rec.Fields(i).Value)) Then
                        If (isText(rec.Fields(i))) Then
                            strValues = strValues & "'" & FixQuote(rec.Fields(i).Value) & "'"
                        ElseIf (isTime(rec.Fields(i))) Then
                            strValues = strValues & "'" & FormatDateTime(rec.Fields(i).Value) & "'"
                        ElseIf (isFile(rec.Fields(i))) Then
                            ' Exportamos los datos (OLE) como datos Hexadecimales
                            If ("" = Trim(rec.Fields(i))) Then
                                strValues = strValues & "NULL"
                            Else
                                If CONFIG_EXPORT_FILES Then
                                    strValues = strValues & ExportOleObject(HelperClearCharacters(table.Name), rec.Fields(i), pos)
                                Else
                                    strValues = strValues & OleToHex(rec.Fields(i).Value)
                                End If
                            End If
                        Else
                            strValues = strValues & rec.Fields(i).Value
                        End If
                    Else
                        strValues = strValues & "NULL"
                    End If
                Next i
                
                strSQL = strSQL & "(" & strValues & ")"
                If (pos < (table.RecordCount - 1)) Then
                    strSQL = strSQL & "," & vbCr
                End If
                rec.MoveNext
                pos = pos + 1
            Loop
            
            strSQL = strSQL & ";" & vbCr
            If CONFIG_LOCK_TABLE Then
                ' Lock the table for insert values
                txtSQL.Add "LOCK TABLES " & strTableName & " WRITE;"
                txtSQL.Add strSQL
                txtSQL.Add "UNLOCK TABLES;"
            Else
                txtSQL.Add strSQL
            End If
        End If
    End If
    
    Set printTable = txtSQL
End Function

' Summanry
'   Generamos la lista con los indices existentes en la tabla con instrucciones SQL.
'   Revisamos los indices para no encontrar duplicados
'
' Params
'   TableDef table La table con la lista de indices
'
' Return
'   Collection Lista con los indices con instrucciones SQL
Private Function PrintIndex(table As TableDef) As Collection
    Dim returnCollection As Collection
    Dim primaryList As Collection, uniqueList As Collection
    Dim indElement As Index
    Dim intKey As Integer
    
    Set returnCollection = New Collection
    
    If 0 < table.Indexes.Count Then
        ' Obtenemos todos los indices del tipo Primary
        ' Obtenemos todos los indices del tipo Unique
        Set primaryList = New Collection
        Set uniqueList = New Collection
        
        For Each indElement In table.Indexes
            If indElement.Primary Then
                primaryList.Add indElement.Name
                returnCollection.Add "PRIMARY KEY (" & getKeys(indElement) & ")"
            ElseIf indElement.Unique Then
                uniqueList.Add indElement.Name
                returnCollection.Add "UNIQUE KEY (" & getKeys(indElement) & ")"
            End If
        Next indElement
        
        For Each indElement In table.Indexes
            If Not isContains(primaryList, indElement.Name) _
                And Not isContains(uniqueList, indElement.Name) _
                And (Not indElement.Foreign) Then
                returnCollection.Add "KEY `" & indElement.Name & "` (" & getKeys(indElement) & ")"
            End If
'            If indElement.Primary Then
'                returnCollection.Add "PRIMARY KEY (" & getKeys(indElement) & ")"
'            ElseIf indElement.Unique Then
'                returnCollection.Add "UNIQUE KEY (" & getKeys(indElement) & ")"
'            ElseIf (Not indElement.Foreign) Then
'                returnCollection.Add "KEY `" & Index.Name & "` (" & getKeys(indElement) & ")"
'            End If
        Next indElement
    End If
    Set PrintIndex = returnCollection
End Function

' Summanry
'   Generamos la lista con las instrucciones SQL necesarias para poder crear las
'   relaciones de las tablas de la Base de datos
'
'   Inspirado en la página web de Ivan Cachicatari's(1)
'
' Ref:
'   (1) http://en.latindevelopers.com/ivancp/2012/ms-access-to-mysql-with-relationships/
'
' Params
'   Database db La Base de Datos
'
' Return
'   Collection Lista con las instrucciones SQL
Private Function PrintRelations(db As Database) As Collection
    Dim strSQL As String, fk As String
    Dim i As Integer, J As Integer
    Dim table As TableDef
    Dim txtSQL As Collection
    
    ' grab a reference to this once, otherwise when we retrieve a table below,
    ' we will get an 'Object Invalid or No Longer Set' error.
    Set txtSQL = New Collection
    
    For i = 0 To db.Relations.Count - 1
        Set table = db.TableDefs.Item(db.Relations(i).table)
        If ((table.Attributes And TableDefAttributeEnum.dbSystemObject) = 0) Then
        
            If CONFIG_COMMENTS_SQL Then
                txtSQL.Add "--"
                txtSQL.Add "-- Filtros para la tabla `" & HelperClearCharacters(table.Name) & "`"
                txtSQL.Add "--"
            End If
 
           strSQL = "ALTER TABLE `" & HelperClearCharacters(db.Relations(i).ForeignTable) & _
               "` ADD CONSTRAINT `" & HelperClearCharacters(db.Relations(i).Name) & "` FOREIGN KEY ("
           fk = "("
           For J = 0 To db.Relations(i).Fields.Count - 1
               strSQL = strSQL & "`" & HelperClearCharacters(db.Relations(i).Fields(J).ForeignName) & "` ,"
               fk = fk & "`" & HelperClearCharacters(db.Relations(i).Fields(J).Name) & "` ,"
           Next J
 
           strSQL = Left(strSQL, Len(strSQL) - 1)
           fk = Left(fk, Len(fk) - 1)
           fk = fk & ")"
           strSQL = strSQL & ") REFERENCES `" & HelperClearCharacters(db.Relations(i).table) & "`" & fk
 
           If (db.Relations(i).Attributes And RelationAttributeEnum.dbRelationUpdateCascade) Then
               strSQL = strSQL & " ON UPDATE CASCADE"
           End If
 
           If (db.Relations(i).Attributes And RelationAttributeEnum.dbRelationDeleteCascade) Then
               strSQL = strSQL & " ON DELETE CASCADE"
           End If
 
           strSQL = strSQL & ";"
 
           txtSQL.Add strSQL
        End If
    Next i
    
    Set PrintRelations = txtSQL
End Function

' Summanry
'   Elimina Acentos, la ñ y la ç del texto indicado
'
' Params
'   String strText Text for replace accents characters
' Return
'   String El texto sin caracteres acentuados
Public Function HelperRemoveAccents(strText As String) As String
    Dim lngText As Long, lngPos As Long, pos As Long
    Dim strCharacter As String, strNormalized As String, strSearchAccents As String, strReplace As String

    strSearchAccents = "áàäâãåçéèêëíìîïñóòôöõðšúùûüýÿž" & "ÁÀÄÂÃÅÇÉÈÊËÍÌÎÏÑÓÒÔÖÕÐŠÚÙÛÜÝŸŽ"
    strReplace = "aaaaaaceeeeiiiinoooooosuuuuyyz" & "AAAAAACEEEEIIIINOOOOOOSUUUUYYZ"
    
    lngText = Len(strText)
    If lngText = 0 Then
        HelperRemoveAccents = ""
        Exit Function
    End If
    
    For pos = 1 To lngText
        strCharacter = Mid(strText, pos, 1)
        'comparamos el caracter con la cadena con acentos
        lngPos = InStr(1, strSearchAccents, strCharacter, vbBinaryCompare)
        'si se ha encontrado coincidencia ...
        If lngPos <> 0 Then
            'sustituímos el caracter con el que tiene la misma
            'posición en la cadena sin acentos (o sea la letra sin acentos)
            strCharacter = Mid(strReplace, lngPos, 1)
        End If
        '... y si no, pues seguimos como si nada
        strNormalized = strNormalized & strCharacter
    Next pos
    
    HelperRemoveAccents = strNormalized
End Function

' Summanry
'   Elimina caracteres no deseados, solo acepta los caracteres "a..z", "A..Z",
'   "0..9", "-"  y "_".
'
' Params
'   String strText Text for clearing characters
' Return
'   String El texto con los caracteres aceptados (validos)
Public Function HelperClearCharacters(strText As String) As String
    Dim strReturn As String, strNumber As String
    Dim pos As Integer
    
    strReturn = ""
    ' Eliminamos acentos
    strText = HelperRemoveAccents(strText)
    strText = UCase(strText)
    
    ' Eliminamos caracteres no deseados.
    For pos = 1 To (Len(strText))
        ' Coge sólo un carácter
        strNumber = Mid(strText, pos, 1)
        Select Case Asc(strNumber)
            ' strNumbers del 0 al 9
            Case 48 To 57
                strReturn = strReturn & strNumber
            ' letras de la "A" a la "Z"
            Case 65 To 90
                strReturn = strReturn & strNumber
            ' letras de la "a" a la "z"
            Case 97 To 122
                strReturn = strReturn & strNumber
            ' caracter "-"
            Case 45
                strReturn = strReturn & strNumber
            ' caracter "_"
            Case 95:
                strReturn = strReturn & strNumber
            ' no necesitamos el caracter
            Case Else
        End Select
    Next
    
    HelperClearCharacters = Trim(strReturn)
End Function

' Summary
'   Comprobamos si existe el fichero en el directorio indicado
'
' Params
'   String strFilePath Nombre del fichero más la ruta
' Return
'   Boolean
Private Function isFileExist(strFilePath As String) As Boolean
    If Len(Dir(strFilePath)) = 0 Then
        isFileExist = False
    Else
        isFileExist = True
    End If
End Function

' Summary
'   Elimina el fichero indicado
'
' Params
'   String strFile Nombre del fichero más la ruta
' Return
'   Boolean
Private Function DeleteFile(strFile As String) As Boolean
    If isFileExist(strFile) Then
        SetAttr strFile, vbNormal
        Kill strFile
        DeleteFile = True
    Else
        DeleteFile = False
    End If
End Function

' Summanry
'   Creamos el fichero con el juego de caracteres UTF-8 donde se guardan todas
'   las intrucciones SQL, que permiten recrear la Base de Datos.
'
'   Inspirado en el Código de JoBrad(1)
'
' Ref:
'   (1) https://gist.github.com/JoBrad/1023484
' Params
'   Collection  lines   Lista con las lineas que se han de escribir en el archivo
'   String      strFile Nombre del Archivo donde se ha de escribir las lineas de Texto
'
' Return
'   String Nombre + Ruta del nuevo fichero creado
Private Function writeOut(lines As Collection, strFile As String) As Boolean
On Error GoTo errHandler
    Dim line, fsT As Variant
    
    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")
     
    'Specify stream type - we want To save text/string data.
    fsT.Type = 2
     
    'Specify charset For the source text data.
    fsT.Charset = "utf-8"
     
    'Open the stream And write binary data To the object
    fsT.Open
    
    For Each line In lines
        fsT.writetext line & vbCrLf
    Next line
    
    'Save binary data To disk
    fsT.SaveToFile strFile, 2
     
    GoTo finish
     
errHandler:
    MsgBox (Err.Description)
    writeOut = False
    Exit Function
    
finish:
    writeOut = True
End Function


' Summanry
'   Obtener el nombre de la Base de datos de la ruta indicada
'
' Params
'   String strPath Ruta completa al fichero de la Base de Datos
'
' Return
'   String El nombre de la Base de datos
Private Function GetNameDB(strPath As String) As String
    Dim strName As String
   
    strName = Mid(Mid(strPath, InStrRev(strPath, "/") + 1), InStrRev(strPath, "\") + 1)
    strName = Mid(strName, 1, InStr(strName, ".") - 1)
    
    GetNameDB = HelperClearCharacters(strName)
End Function

' Summanry
'   Identificamos los tipos de cada campo y su equivalente en MySQL.
'   Inspirado en el código de Allen Browne(1).
'
' Ref
'   (1) http://allenbrowne.com/func-06.html
' Params
'   DAO.Field fld Campo de la tabla
'
' Return
'   String El tipo de MySQL equivalente al tipo de la tabla
Private Function FieldTypeName(fld As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim pos As Integer
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean:     strReturn = "tinyint(1)"                    ' 1 "Yes/No"
        Case dbByte:        strReturn = "tinyint(8) unsigned"           ' 2 "Byte"
        Case dbInteger:     strReturn = "int(6)"                        ' 3 "Integer"
        Case dbLong                                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "int(11) unsigned"                          '  "Long Integer"
            Else
                strReturn = "int(11) unsigned AUTO_INCREMENT"           '  "AutoNumber"
            End If
        Case dbCurrency:    strReturn = "decimal(19,4)"                 ' 5 "Currency"
        Case dbSingle:      strReturn = "float"                         ' 6 "Single"
        Case dbDouble:      strReturn = "double"                        ' 7 "Double"
        Case dbDate:        strReturn = "datetime"                      ' 8 "Date/Time"
        Case dbBinary:      strReturn = "bit(1)"                        ' 9 (no interface) "Binary"
        Case dbText
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "varchar"                                   '10 "Text"
            Else
                strReturn = "text"                                      '10 (no interface) "Text (fixed width)"
            End If
        
        Case dbLongBinary:
            If CONFIG_EXPORT_FILES Then
                strReturn = "text"                                      '11 "Change for Text"
            Else
                strReturn = "longblob"                                  '11 "OLE Object"
            End If
        Case dbMemo
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "text"                                      '12 "Memo"
            Else
                strReturn = "text"                                      '12 "Hyperlink"
            End If
        Case dbGUID:        strReturn = "char(36)"                      '15 "GUID" | binary(16)

        'Attached tables only: cannot create these in JET.
        Case dbBigInt:      strReturn = "bigint(20)"                    '16 "Big Integer"
        Case dbVarBinary:   strReturn = "varbinary"                     '17 "VarBinary"
        Case dbChar:        strReturn = "char(1)"                       '18 "Char"
        Case dbNumeric:     strReturn = "int(6)"                        '19 "Numeric"
        Case dbDecimal:     strReturn = "decimal(10,0)"                 '20 "Decimal"
        Case dbFloat:       strReturn = "float"                         '21 "Float"
        Case dbTime:        strReturn = "time"                          '22 "Time"
        Case dbTimeStamp:   strReturn = "timestamp DEFAULT " _
                                    & "CURRENT_TIMESTAMP ON UPDATE " _
                                    & "CURRENT_TIMESTAMP"               '23 "Time Stamp"

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&:
            If CONFIG_EXPORT_FILES Then
                            strReturn = "text"                          'dbAttachment       "Change for Text"
            Else
                            strReturn = "longblob"                      'dbAttachment       "Attachment"
            End If
        Case 102&:          strReturn = "tinyint(8)"                    'dbComplexByte      "Complex Byte"
        Case 103&:          strReturn = "int(6)"                        'dbComplexInteger   "Complex Integer"
        Case 104&:          strReturn = "int(11)"                       'dbComplexLong      "Complex Long"
        Case 105&:          strReturn = "float"                         'dbComplexSingle    "Complex Single"
        Case 106&:          strReturn = "double"                        'dbComplexDouble    "Complex Double"
        Case 107&:          strReturn = "char(36)"                      'dbComplexGUID      "Complex GUID"
        Case 108&:          strReturn = "decimal(65,9)"                 'dbComplexDecimal   "Complex Decimal"
        Case 109&:          strReturn = "longtext"                      'dbComplexText      "Complex Text"
        Case Else
            Exit Function
    End Select

    If (fld.Size > 0 And ((fld.Attributes And dbAutoIncrField) = 0&) And (Not isTime(fld))) Then
        pos = InStr(1, strReturn, "(")
        ' remove "()"
        If (Not IsNull(pos) And pos > 0) Then
            strReturn = Left(strReturn, pos - 1)
        End If
        
        strReturn = strReturn & "(" & fld.Size & ")"
    End If
    
    FieldTypeName = strReturn
End Function

' Summanry
'   Comprobamos si el campo de la tabla de algún tipo de texto.
'
' Params
'   Field fld Campo de la Tabla
'
' Return
'   Boolena
Private Function isText(fld As DAO.Field) As String
    Dim pos As Integer
    Dim boolReturn As Boolean

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbText: boolReturn = True
        Case dbMemo: boolReturn = True
        
        'Attached tables only: cannot create these in JET.
        Case dbChar: boolReturn = True

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 109&: boolReturn = True
        Case Else: boolReturn = False
    End Select
    
    isText = boolReturn
End Function

' Summanry
'   Comprobamos si el campo de la tabla de algún tipo de tiempo.
'
' Params
'   Field fld Campo de la Tabla
'
' Return
'   Boolena
Private Function isTime(fld As DAO.Field) As Boolean
    Dim boolReturn As Boolean
    
    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbDate: boolReturn = True
        Case dbTime: boolReturn = True
        Case dbTimeStamp: boolReturn = True
        Case Else: boolReturn = False
    End Select

    isTime = boolReturn
End Function

' Summanry
'   Comprobamos si el campo de la tabla es de algún tipo de archivo.
'
' Params
'   Field fld Campo de la Tabla
'
' Return
'   Boolena
Private Function isFile(fld As DAO.Field) As Boolean
    Dim boolReturn As Boolean
    
    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbLongBinary: boolReturn = True
        Case 101&: boolReturn = True
        
        Case Else: boolReturn = False
    End Select

    isFile = boolReturn
End Function

' Summanry
'   Indica si en la lista pasada existe un elemento con el indice indicado
'
' Params
'   Collection list     Lista donde se buscara el indice
'   String     strIndex Indice que se esta comprobando
'
' Return
'   Boolean
Private Function isContains(listSearch As Collection, strIndex As String) As Boolean
    Dim boolReturn As Boolean
    Dim elementList As Variant
    
    For Each elementList In listSearch
        If (elementList = strIndex) Then
            boolReturn = True
            Exit For
        End If
    Next elementList
    
    isContains = boolReturn
End Function

' Summanry
'   Modificamos el texto cmabiando las comillas simples(') y comillas dobles (")
'   para poder utilizarlas dentro del código
'
' Params
'   String FQText Texto donde aparece la comilla simple o la comilla doble
'
' Return
'   String
Private Function FixQuote(FQText As String) As String
    FixQuote = Replace(FQText, "'", "\'") ' Chr(39)
    'FixQuote = Replace(FQText, """", "\""") ' Chr(34)
End Function

' Summanry
'   Aplicamos el formato de tiempo que necesitamos (yyyy-mm-dd hh:nn:ss)
'
' Params
'   String strTime El valor del tiempo
'
' Return
'   String Con el tiempo en el formato adecuado
Private Function FormatDateTime(strTime As String) As String
    Dim strReturn As String
    
    strReturn = Format(strTime, "yyyy-mm-dd") & " " & Format(strTime, "hh:nn:ss")
    FormatDateTime = strReturn
End Function

' Summanry
'   Obtenemos los campos clave de la tabla
'
' Params
'   Index ind Lista de indices de la Tabla
'
' Return
'   String Con los indices separados por comas(,)
Private Function getKeys(ind As Index) As String
    Dim pos As Integer
    Dim strReturn As String, strType As String
    Dim fld As Field
    
    'Select Case intType
    '    Case 1: strType = "PRIMARY KEY"
    '    Case 2: strType = "UNIQUE"
    '    Case Else: strType = "KEYS"
    'End Select
    pos = 0
    For Each fld In ind.Fields
        If (pos > 0) Then
            strReturn = strReturn & ","
        End If
        strReturn = strReturn & "`" & HelperClearCharacters(fld.Name) & "`"
        pos = pos + 1
    Next fld
    
    'If (strReturn <> "") Then
    '    strReturn = strType & " (" & strReturn & ")" & vbCr
    'End If
    
    getKeys = strReturn
End Function

' Summanry
'   Convertimos el código guardado del Ole Object en un código Hexadecimal para MySQL
'
'   Inspirado en el Blog de happycodings.com(1)
'
' Ref:
'   (1) http://visualbasic.happycodings.com/applications-vba/code4.html
' Params
'   String BinNum Código almacenado del Ole Object
'
' Return
'   String Con el código de la imagen en formato Hexadecimal
Private Function OleToHex(BinNum As String) As String
    Dim intByte As Integer, intChr As Integer, NumBlocks As Integer, LeftOver As Integer, i As Integer
    Dim strAux As String
    Dim hexAux, hexReturn As Variant
   
    intByte = 8
    NumBlocks = Len(BinNum) \ intByte     ' 8 = byte
    LeftOver = Len(BinNum) Mod intByte
    
    If 8 < Len(BinNum) Then
        For i = 1 To NumBlocks + 1
            If (i = NumBlocks + 1) And (0 < LeftOver) Then
                strAux = Mid(BinNum, ((i - 1) * intByte) + 1, LeftOver)
            Else
                strAux = Mid(BinNum, ((i - 1) * intByte) + 1, intByte)
            End If
            If ("" <> strAux) Then
                intChr = Asc(strAux)
                hexAux = Hex(intChr)
                If Len(hexAux) < 2 Then
                    hexAux = "0" & hexAux
                End If
            Else
                hexAux = "00"
            End If
            hexReturn = hexReturn & hexAux
        Next i
    Else
        intChr = Asc(BinNum)
        hexAux = Hex(intChr)
        If Len(hexAux) < 2 Then
          hexAux = "0" & hexAux
        End If
        hexReturn = hexAux
    End If

    OleToHex = "0x" & hexReturn
End Function

' Summanry
'   Creamos el fichero de la imagen con el código Ole Object guadado en la
'   Base de datos.
'   IMPORTANTE: No es una practica recomendable, lo mejor es guardar la referencia
'               al archivo, no guardar el archivo dentro de la Base de datos.
'
'   Inspirado en los comentarios de soporte de Windows(1)
'
' Ref:
'   (1) http://support.microsoft.com/kb/210486
' Params
'   String BinNum Código almacenado del Ole Object
'
' Return
'   String Nombre + Ruta del nuevo fichero creado
Private Function ExportOleObject(strTable As String, fld As DAO.Field, intCount As Integer) As String
    Dim NumBlocks As Integer, DestFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim strFileName As String, FileData As String
    Dim RetVal As Variant

    On Error GoTo Err_WriteBLOB
    
    ' Get Name to file
    strFileName = strTable & "_" & fld.Name & "_" & intCount & "." & CONFIG_EXPORT_FILES_EXTENSION
    
    ' Get the size of the field.
    FileLength = fld.FieldSize()
    If FileLength = 0 Then
        ExportOleObject = "NULL"
        Exit Function
    End If
    
    ' Calculate number of blocks to write and leftover bytes.
    NumBlocks = FileLength \ CONFIG_BLOCK_SIZE
    LeftOver = FileLength Mod CONFIG_BLOCK_SIZE

    ' Remove any existing destination file.
    DestFile = FreeFile
    Open strFileName For Output As DestFile
    Close DestFile

    ' Open the destination file.
    Open strFileName For Binary As DestFile

    ' Write the leftover data to the output file.
    FileData = fld.GetChunk(0, LeftOver)
    Put DestFile, , FileData

    ' Write the remaining blocks of data to the output file.
    For i = 1 To NumBlocks
        ' Reads a chunk and writes it to output file.
        FileData = fld.GetChunk((i - 1) * CONFIG_BLOCK_SIZE + LeftOver, CONFIG_BLOCK_SIZE)
        Put DestFile, , FileData
    Next i

    Close DestFile
    ExportOleObject = CONFIG_FILES_DIRECTORY & "/" & strFileName
    Exit Function

Err_WriteBLOB:
    WriteBLOB = -Err
    Exit Function
    
End Function
