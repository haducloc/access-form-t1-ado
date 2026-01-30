Option Compare Database
Option Explicit

' Create temp table if it doesn't exist. 
' schemaSql example: "[Id] LONG NOT NULL, [CtName] TEXT(100) NOT NULL, [ReturnDate1] DATETIME" 
' pkField example: "Id"
Public Sub EnsureTempTableDao(ByVal tableName As String, ByVal schemaSql As String, Optional ByVal pkField As String = "Id")
  Dim db As DAO.Database
  Set db = CurrentDb

  If TableExists(db, tableName) Then
    Exit Sub
  End If

  db.Execute "CREATE TABLE [" & tableName & "] (" & schemaSql & ");", dbFailOnError

  If Len(pkField) > 0 Then
    db.Execute "CREATE UNIQUE INDEX [PK_" & tableName & "] ON [" & tableName & "] ([" & pkField & "]);", dbFailOnError
  End If
End Sub

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
  Set qd = Nothing
  Set db = Nothing
  Exit Sub

TCError:
  Set xe = ToXError(Err)

  CloseObj qd
  Set qd = Nothing
  Set db = Nothing

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
  Set rs = Nothing
  Set qd = Nothing
  Set db = Nothing

  If result Is Nothing Then
    Set result = NewDictionary()
  End If

  Set GetTempByIdDao = result
  Exit Function

TCError:
  Set xe = ToXError(Err)

  CloseObj rs
  CloseObj qd
  Set rs = Nothing
  Set qd = Nothing
  Set db = Nothing

  Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Convert the current record in a DAO.Recordset into a Dictionary.
Public Function RecordsetToDictionaryDao(ByVal rs As DAO.Recordset) As Object
  Dim row As Object
  Set row = NewDictionary()

  If rs Is Nothing Then
    Set RecordsetToDictionaryDao = row
    Exit Function
  End If

  If (rs.BOF And rs.EOF) Then
    Set RecordsetToDictionaryDao = row
    Exit Function
  End If

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
