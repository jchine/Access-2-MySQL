Attribute VB_Name = "exportSQLFile"
Option Compare Database

' This file is part of the access2mysql.
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
Const CONFIG_COMMENTS_SQL = True                        ' Include comments before SQL statements
Const CONFIG_DROP_TABLE = True                          ' Include Drop SQL statements
Const CONFIG_USED_RELATIONS = True                      ' Includes SQL statements to insert values into tables
Const CONFIG_LOCK_TABLE = True                          ' Lock of table to insert values
Const CONFIG_INSERT_VALUES = True                       ' Includes SQL statements to insert values into tables

Const CONFIG_OUTPUT_FILENAME = "export"                 ' Default File name
Const CONFIG_OUTPUT_FILE_PATH = ""                      ' Current DB path
Const CONFIG_OUTPUT_FILE_EXTENSION = "sql"              ' Default extension for dump SQL

' Summanry
'   Elimina Acentos, la ñ y la ç del texto indicado
'
' Params
'   String strText Text for replace accents characters
Public Function HelperRemoveAccents(strText As String) As String
    Dim lngText, lngPos, pos As Long
    Dim strCharacter, strNormalized, strSearchAccents, strReplace As String

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
Public Function HelperClearCharacters(strText As String) As String
    Dim strReturn, strNumber As String
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

' Summanry
'   Generamos y guardamos las instrucciones SQL
'
' Params
'   Database db        OPTIONAL Dababase for generate dump sql
'   String strFileName OPTIONAL Output filename with ".sql"
Public Sub generateSQLFile(Optional db As Database, Optional strFileName As String)
    Dim numFile As Integer
    Dim line As Variant
    Dim listSQL As Collection
    Dim Table As TableDef
        
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
        
    OpenFile strFileName, numFile
    
        If CONFIG_COMMENTS_SQL Then
            Print #numFile, "--"
            Print #numFile, "-- Base de datos: `" & GetNameDB(db.Name) & "`"
            Print #numFile, "--"
            Print #numFile, ""
        End If
        
        Print #numFile, "CREATE DATABASE IF NOT EXISTS `" & GetNameDB(db.Name) & "`"
        Print #numFile, "CHARACTER SET `utf8`;"
        Print #numFile, ""
        Print #numFile, "USE `" & GetNameDB(db.Name) & "`;"
        
        If CONFIG_COMMENTS_SQL Then
            Print #numFile, "-- --------------------------------------------------------"
        End If

        For Each Table In db.TableDefs
            ' Descartamos las tablas del sistema de MS Access "MSys..."
            If ("MSys" <> Left(Table.Name, 4)) Then
                Set listSQL = printTable(db, Table)
                For Each line In listSQL
                    Print #numFile, line
                Next
            End If
        Next Table
        
        Set listSQL = printRelations(db)
        For Each line In listSQL
            Print #numFile, line
        Next
        
    CloseFile numFile
    
    MsgBox ("Fichero '" & strFileName & "' creado en la carpeta " & CONFIG_OUTPUT_FILE_PATH _
            & " o en su defecto en la carpeta de la Base de Datos (" & CurrentProject.Path & ")")
End Sub

' Summanry
'   Creamos y abrirmos el fichero donde se guardará las instrucciones SQL
'
' Params
'   String strFileName Output filename with ".sql"
'   Integer FILENumber Identifier of filename
Private Sub OpenFile(strFileName As String, FILENumber As Integer)
    Dim strPath As String
    
    If 0 = InStr(strFileName, "." & CONFIG_OUTPUT_FILE_EXTENSION) Then
        strFileName = strFileName & "." & CONFIG_OUTPUT_FILE_EXTENSION
    End If
    
    If ("" = Trim(CONFIG_OUTPUT_FILE_PATH & "")) Then
        strPath = CurrentProject.Path
    Else
        strPath = Trim(CONFIG_OUTPUT_FILE_PATH)
    End If
    
    Open strPath & "\" & UCase(strFileName) For Output As FILENumber
End Sub

' Summanry
'   Cerramos el fichero donde se guardará las instrucciones SQL
'
' Params
'   Integer FILENumber Identifier of filename
Private Sub CloseFile(FILENumber As Integer)
    Dim strPath As String
    
    If ("" = Trim(CONFIG_OUTPUT_FILE_PATH)) Then
        strPath = CurrentProject.Path
    Else
        strPath = Trim(CONFIG_OUTPUT_FILE_PATH)
    End If
    
    Close (FILENumber)
End Sub

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
                strReturn = "text"                                      '10 "Text"
            Else
                strReturn = "text"                                      '10 (no interface) "Text (fixed width)"
            End If
        
        Case dbLongBinary:  strReturn = "longblob"                      '11 "OLE Object"
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
        Case 101&:          strReturn = "longblob"                      'dbAttachment       "Attachment"
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
'   Modificamos el texto cmabiando las comillas simples(') y comillas dobles (")
'   para poder utilizarlas dentro del código
'
' Params
'   String FQText Texto donde aparece la comilla simple o la comilla doble
'
' Return
'   String
Private Function FixQuote(FQText As String) As String
    FQText = Replace(FQText, "'", "''") ' Chr(39)
    FixQuote = Replace(FQText, """", """""") ' Chr(34)
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
    Dim strReturn, strType As String
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
'   Generamos la lista con las instrucciones SQL necesarias para poder crear la tabla y
'   sus valores.
'
' Params
'   Database db     La Base de Datos
'   TableDef Table  La tabla de la base de datos de la cual crearemos las instrucciones SQL
'
' Return
'   Collection Lista con las instrucciones SQL
Public Function printTable(db As Database, Table As TableDef) As Collection
    Dim strSQL, strTableName, strFieldName, strValues, strFields As String
    Dim i, pos As Integer
    Dim boolIndex As Boolean
    Dim strVariant As Variant
    Dim Field As Field
    Dim Index As Index
    Dim rec As Recordset
    Dim indexSQL, txtSQL As Collection
    
    Set txtSQL = New Collection
    strTableName = "`" & HelperClearCharacters(Table.Name) & "`"
    
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
        For Each Field In Table.Fields
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
            
            If (pos < Table.Fields.Count - 1) Or (Table.Indexes.Count > 0) Then
                strSQL = strSQL & ","
            End If
            strSQL = strSQL & vbCr
            pos = pos + 1
        Next Field
        
        Set indexSQL = New Collection
        For Each Index In Table.Indexes
            If Index.Primary Then
                indexSQL.Add "PRIMARY KEY (" & getKeys(Index) & ")"
            ElseIf Index.Unique Then
                indexSQL.Add "UNIQUE KEY (" & getKeys(Index) & ")"
            ElseIf (Not Index.Foreign) Then
                indexSQL.Add "KEY `" & Index.Name & "` (" & getKeys(Index) & ")"
            End If
        Next Index
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
        Set rec = db.OpenRecordset(Table.Name, dbOpenTable)
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
                        Else
                            strValues = strValues & rec.Fields(i).Value
                        End If
                    Else
                        strValues = strValues & "NULL"
                    End If
                Next i
                
                strSQL = strSQL & "(" & strValues & ")"
                If (pos < (Table.RecordCount - 1)) Then
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
Function printRelations(db As Database) As Collection
    Dim strSQL, fk, TableName As String
    Dim i, J As Integer
    Dim Table As TableDef
    Dim txtSQL As Collection
    
    ' grab a reference to this once, otherwise when we retrieve a table below,
    ' we will get an 'Object Invalid or No Longer Set' error.
    Set txtSQL = New Collection
    
    For i = 0 To db.Relations.Count - 1
        Set Table = db.TableDefs.Item(db.Relations(i).Table)
        If ((Table.Attributes And TableDefAttributeEnum.dbSystemObject) = 0) Then
            If CONFIG_COMMENTS_SQL Then
                txtSQL.Add "--"
                txtSQL.Add "-- Filtros para la tabla `" & Table.Name & "`"
                txtSQL.Add "--"
            End If
 
           strSQL = "ALTER TABLE `" & db.Relations(i).ForeignTable & _
               "` ADD CONSTRAINT `" & db.Relations(i).Name & "` FOREIGN KEY ("
           fk = "("
           For J = 0 To db.Relations(i).Fields.Count - 1
               strSQL = strSQL & "`" & db.Relations(i).Fields(J).ForeignName & "` ,"
               fk = fk & "`" & db.Relations(i).Fields(J).Name & "` ,"
           Next J
 
           strSQL = Left(strSQL, Len(strSQL) - 1)
           fk = Left(fk, Len(fk) - 1)
           fk = fk & ")"
           strSQL = strSQL & ") REFERENCES `" & db.Relations(i).Table & "`" & fk
 
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
    
    Set printRelations = txtSQL
End Function
