Option Compare Database
Option Explicit

' Create temp table if it doesn't exist.
' schemaSql example: "[Id] LONG|AUTOINCREMENT NOT NULL, [CtName] TEXT(100) NOT NULL, [ReturnDate1] DATETIME"
' pkField example: "Id"
Public Function EnsureTempTableAdo( _
    ByVal tableName As String, _
    ByVal schemaSql As String, _
    Optional ByVal pkField As String = "" _
) As Boolean

    Dim db As DAO.Database
    Set db = CurrentDb

    If TableExists(db, tableName) Then
        EnsureTempTableAdo = False
        Exit Function
    End If

    db.Execute "CREATE TABLE [" & tableName & "] (" & schemaSql & ");", dbFailOnError

    If Len(pkField) > 0 Then
        Dim tdf As DAO.TableDef
        Dim idx As DAO.Index

        Set tdf = db.TableDefs(tableName)
        Set idx = tdf.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Unique = True
        idx.Fields.Append idx.CreateField(pkField)
        tdf.Indexes.Append idx
        tdf.Indexes.Refresh
    End If

    EnsureTempTableAdo = True
End Function

' Clear all rows from the temp table.
Public Sub ClearTempTableDao(ByVal tableName As String)
    CurrentDb.Execute "DELETE FROM [" & tableName & "];", dbFailOnError
End Sub

' Insert one row using parameter names = "p" + field name.
' fieldsCsv example: "Id,CtName,ReturnDate1"
' typesCsv example: "LONG,TEXT(100),DATETIME"
' values: dictionary-like object
Public Sub InsertTempRowDao(ByVal tableName As String, ByVal fieldsCsv As String, ByVal typesCsv As String, ByVal values As Object)
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim xe As XError

    On Error GoTo TCError

    Set db = CurrentDb
    Set qd = BuildInsertQd(db, tableName, fieldsCsv, typesCsv)

    BindParamsFromValues qd, values
    qd.Execute dbFailOnError

    CloseObj qd
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Get single row by PK as Dictionary (param name = "p" + pkField).
Public Function GetTempByIdDao(ByVal tableName As String, ByVal pkField As String, ByVal pkType As String, ByVal idValue As Variant) As Object
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim result As Object
    Dim xe As XError

    Dim paramName As String
    paramName = "p" & pkField

    On Error GoTo TCError

    Set db = CurrentDb
    Set qd = db.CreateQueryDef("", _
        "PARAMETERS [" & paramName & "] " & pkType & ";" & vbCrLf & _
        "SELECT * FROM [" & tableName & "] WHERE [" & pkField & "] = [" & paramName & "];")

    qd.Parameters(paramName).Value = idValue
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    Set result = RecordsetToDictionaryDao(rs)

    CloseObj rs
    CloseObj qd

    Set GetTempByIdDao = result
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Convert the current record in a DAO.Recordset into a Dictionary.
Public Function RecordsetToDictionaryDao(ByVal rs As DAO.Recordset) As Object
    If rs Is Nothing Then
        Set RecordsetToDictionaryDao = Nothing
        Exit Function
    End If

    If (rs.BOF And rs.EOF) Then
        Set RecordsetToDictionaryDao = Nothing
        Exit Function
    End If

    Dim row As Object
    Set row = NewDictionary()

    Dim f As DAO.Field
    For Each f In rs.Fields
        row(f.Name) = f.Value
    Next

    Set RecordsetToDictionaryDao = row
End Function

' Check whether a table exists in the database.
Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next
    TableExists = False
End Function

' Build a parameterized INSERT QueryDef (param names = "p" + field name).
Private Function BuildInsertQd(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldsCsv As String, ByVal typesCsv As String) As DAO.QueryDef
    Dim fields() As String
    Dim types() As String

    fields = Split(fieldsCsv, ",")
    types = Split(typesCsv, ",")

    If UBound(fields) <> UBound(types) Then
        XRaise "XDaoUtil.BuildInsertQd", "fieldsCsv and typesCsv must have the same number of items."
    End If

    Dim paramsSql As String
    Dim colsSql As String
    Dim valsSql As String

    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim f As String
        Dim t As String
        Dim p As String

        f = Trim$(fields(i))
        t = Trim$(types(i))
        p = "p" & f

        If Len(paramsSql) > 0 Then paramsSql = paramsSql & ", "
        paramsSql = paramsSql & "[" & p & "] " & t

        If Len(colsSql) > 0 Then colsSql = colsSql & ", "
        colsSql = colsSql & "[" & f & "]"

        If Len(valsSql) > 0 Then valsSql = valsSql & ", "
        valsSql = valsSql & "[" & p & "]"
    Next

    Dim sqlText As String
    sqlText = "PARAMETERS " & paramsSql & ";" & vbCrLf & _
              "INSERT INTO [" & tableName & "] (" & colsSql & ")" & vbCrLf & _
              "VALUES (" & valsSql & ");"

    Set BuildInsertQd = db.CreateQueryDef("", sqlText)
End Function

' Bind QueryDef parameters from the values object (missing -> Null).
Private Sub BindParamsFromValues(ByVal qd As DAO.QueryDef, ByVal values As Object)
    Dim p As DAO.Parameter

    For Each p In qd.Parameters
        Dim fieldName As String
        fieldName = NormalizeParamName(p.Name)

        ' expected parameter naming: p + FieldName
        If LCase$(Left$(fieldName, 1)) = "p" And Len(fieldName) > 1 Then
            fieldName = Mid$(fieldName, 2)
        End If

        If HasField(values, fieldName) Then
            p.Value = values(fieldName)
        Else
            p.Value = Null
        End If
    Next
End Sub

' Normalize a parameter name by stripping surrounding brackets.
Private Function NormalizeParamName(ByVal paramName As String) As String
    If Left$(paramName, 1) = "[" And Right$(paramName, 1) = "]" Then
        NormalizeParamName = Mid$(paramName, 2, Len(paramName) - 2)
        Exit Function
    End If
    NormalizeParamName = paramName
End Function

' Test whether the values object contains the given field key.
Private Function HasField(ByVal values As Object, ByVal fieldName As String) As Boolean
    On Error GoTo TCError

    Dim v As Variant
    v = values(fieldName)

    HasField = True
    Exit Function

TCError:
    HasField = False
End Function

' Create a local Access table from an ADO Recordset and load its data.
' Create temp table from ADO RS if missing; otherwise clear it; then load data.
Public Sub EnsureTempTableFromAdoRs( _
    ByVal rs As ADODB.Recordset, _
    ByVal tableName As String, _
    Optional ByVal pkFieldName As String = "" _
)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim i As Long
    Dim xe As XError

    On Error GoTo TCError

    If rs Is Nothing Then XRaise "XDaoUtil.EnsureTempTableFromAdoRs", "Recordset is Nothing."
    If rs.State = 0 Then XRaise "XDaoUtil.EnsureTempTableFromAdoRs", "Recordset is closed."

    Set db = CurrentDb()

    If TableExists(db, tableName) Then
        ' Table exists -> clear
        db.Execute "DELETE FROM [" & tableName & "];", dbFailOnError
    Else
        ' Table missing -> create (offline TableDef then append)
        Set tdf = db.CreateTableDef(tableName)

        For i = 0 To rs.Fields.Count - 1
            tdf.Fields.Append MapAdoFieldToDaoField(tdf, rs.Fields(i))
        Next i

        If Len(pkFieldName) > 0 Then
            If FieldExistsInTdf(tdf, pkFieldName) Then
                AddPrimaryKey tdf, pkFieldName
            Else
                XRaise "XDaoUtil.EnsureTempTableFromAdoRs", "PK field not found: " & pkFieldName
            End If
        End If

        db.TableDefs.Append tdf
        db.TableDefs.Refresh
    End If

    ' Always load data after ensuring table exists and is empty
    InsertAdoRecordsetRows rs, tableName
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Create a DAO Field based on an ADO Field definition.
Private Function MapAdoFieldToDaoField( _
    ByVal tdf As DAO.TableDef, _
    ByVal adoFld As ADODB.Field _
) As DAO.Field

    Dim daoType As DAO.DataTypeEnum
    daoType = MapAdoTypeToDaoType(adoFld.Type)

    Select Case daoType

        Case dbText
            If adoFld.DefinedSize > 0 And adoFld.DefinedSize <= 255 Then
                Set MapAdoFieldToDaoField = tdf.CreateField(adoFld.Name, dbText, adoFld.DefinedSize)
            Else
                Set MapAdoFieldToDaoField = tdf.CreateField(adoFld.Name, dbMemo)
            End If

        Case dbBinary
            If adoFld.DefinedSize > 0 And adoFld.DefinedSize <= 255 Then
                Set MapAdoFieldToDaoField = tdf.CreateField(adoFld.Name, dbBinary, adoFld.DefinedSize)
            Else
                Set MapAdoFieldToDaoField = tdf.CreateField(adoFld.Name, dbLongBinary)
            End If

        Case Else
            Set MapAdoFieldToDaoField = tdf.CreateField(adoFld.Name, daoType)

    End Select
End Function

' Map an ADO data type to the closest DAO data type.
Private Function MapAdoTypeToDaoType(ByVal adoType As ADODB.DataTypeEnum) As DAO.DataTypeEnum
    Select Case adoType
        Case adSmallInt:  MapAdoTypeToDaoType = dbInteger
        Case adInteger:   MapAdoTypeToDaoType = dbLong

        Case adBigInt
            If SupportsDaoBigInt() Then
                MapAdoTypeToDaoType = 20 ' dbBigInt
            Else
                MapAdoTypeToDaoType = dbDouble
            End If

        Case adUnsignedTinyInt, adTinyInt
            MapAdoTypeToDaoType = dbByte

        Case adBoolean
            MapAdoTypeToDaoType = dbBoolean

        Case adSingle
            MapAdoTypeToDaoType = dbSingle

        Case adDouble
            MapAdoTypeToDaoType = dbDouble

        Case adCurrency
            MapAdoTypeToDaoType = dbCurrency

        Case adDecimal, adNumeric
            MapAdoTypeToDaoType = dbDouble

        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            MapAdoTypeToDaoType = dbDate

        Case adVarChar, adWChar, adVarWChar, adChar, adBSTR
            MapAdoTypeToDaoType = dbText

        Case adLongVarChar, adLongVarWChar
            MapAdoTypeToDaoType = dbMemo

        Case adBinary, adVarBinary
            MapAdoTypeToDaoType = dbBinary

        Case adLongVarBinary
            MapAdoTypeToDaoType = dbLongBinary

        Case Else
            MapAdoTypeToDaoType = dbText
    End Select
End Function

' Detect whether the current Access version supports DAO BigInt.
Private Function SupportsDaoBigInt() As Boolean
    On Error GoTo Nope
    Dim tdf As DAO.TableDef
    Dim f As DAO.Field

    Set tdf = CurrentDb.CreateTableDef("")
    Set f = tdf.CreateField("x", 20) ' dbBigInt
    SupportsDaoBigInt = True
    Exit Function
Nope:
    SupportsDaoBigInt = False
End Function

' Add a single-field primary key to a TableDef.
Private Sub AddPrimaryKey(ByRef tdf As DAO.TableDef, ByVal fieldName As String)
    Dim idx As DAO.Index

    Set idx = tdf.CreateIndex("PK_" & tdf.Name)
    With idx
        .Primary = True
        .Unique = True
        .Fields.Append .CreateField(fieldName)
    End With

    tdf.Indexes.Append idx
End Sub

' Check whether a field exists in a TableDef.
Private Function FieldExistsInTdf(ByVal tdf As DAO.TableDef, ByVal fieldName As String) As Boolean
    On Error GoTo Nope
    Dim f As DAO.Field
    Set f = tdf.Fields(fieldName)
    FieldExistsInTdf = True
    Exit Function
Nope:
    FieldExistsInTdf = False
End Function

' Insert all rows from an ADO Recordset into a local Access table.
Private Sub InsertAdoRecordsetRows(ByVal rs As ADODB.Recordset, ByVal tableName As String)
    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim sql As String
    Dim i As Long
    Dim rowNum As Long
    Dim xe As XError

    On Error GoTo TCError

    Set db = CurrentDb()

    sql = "PARAMETERS " & BuildDaoParameters(rs) & vbCrLf & _
          "INSERT INTO [" & tableName & "] (" & BuildColumnList(rs) & ") " & vbCrLf & _
          "VALUES (" & BuildParameterPlaceholders(rs) & ");"

    Set qd = db.CreateQueryDef("", sql)
    qd.ReturnsRecords = False

    db.BeginTrans

    rowNum = 0
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            rowNum = rowNum + 1
            For i = 0 To rs.Fields.Count - 1
                qd.Parameters("p" & i).Value = rs.Fields(i).Value
            Next i
            qd.Execute dbFailOnError
            rs.MoveNext
        Loop
    End If

    db.CommitTrans
    CloseObj qd
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    ' BUG FIX: Wrapped Rollback in On Error Resume Next to prevent secondary crash if transaction never started
    On Error Resume Next
    If Not db Is Nothing Then db.Rollback
    On Error GoTo 0
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, "Transfer failed at row " & rowNum & ": " & xe.ErrDesc
End Sub

' Build a comma-separated list of bracketed column names from a Recordset.
Private Function BuildColumnList(ByVal rs As ADODB.Recordset) As String
    Dim i As Long, s As String
    For i = 0 To rs.Fields.Count - 1
        s = s & IIf(i > 0, ",", "") & "[" & rs.Fields(i).Name & "]"
    Next
    BuildColumnList = s
End Function

' Build a comma-separated list of parameter placeholders (p0, p1, ...).
Private Function BuildParameterPlaceholders(ByVal rs As ADODB.Recordset) As String
    Dim i As Long, s As String
    For i = 0 To rs.Fields.Count - 1
        s = s & IIf(i > 0, ",", "") & "[p" & i & "]"
    Next
    BuildParameterPlaceholders = s
End Function

' Build the PARAMETERS clause for the insert QueryDef.
Private Function BuildDaoParameters(ByVal rs As ADODB.Recordset) As String
    Dim i As Long, s As String
    For i = 0 To rs.Fields.Count - 1
        s = s & IIf(i > 0, ", ", "") & "p" & i & " Variant"
    Next
    BuildDaoParameters = s & ";"
End Function
