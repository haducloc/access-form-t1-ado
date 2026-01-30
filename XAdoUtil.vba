Option Compare Database
Option Explicit

' ADO constants (late-binding)
Public Const adUseClient As Long = 3
Public Const adOpenStatic As Long = 3
Public Const adLockReadOnly As Long = 1
Public Const adCmdText As Long = 1
Public Const adExecuteNoRecords As Long = &H80
Public Const adParamInput As Long = 1

' DataTypeEnum
Public Const adSmallInt As Long = 2
Public Const adInteger As Long = 3
Public Const adBigInt As Long = 20
Public Const adUnsignedTinyInt As Long = 17
Public Const adSingle As Long = 4
Public Const adDouble As Long = 5
Public Const adCurrency As Long = 6
Public Const adNumeric As Long = 131
Public Const adBoolean As Long = 11
Public Const adGUID As Long = 72
Public Const adDBDate As Long = 133
Public Const adDBTime As Long = 134
Public Const adDBTimeStamp As Long = 135
Public Const adChar As Long = 129
Public Const adWChar As Long = 130
Public Const adVarChar As Long = 200
Public Const adLongVarChar As Long = 201
Public Const adVarWChar As Long = 202
Public Const adLongVarWChar As Long = 203
Public Const adBinary As Long = 128
Public Const adVarBinary As Long = 204
Public Const adLongVarBinary As Long = 205

' Begin a transaction.
Public Sub BeginTransAdo(ByVal cn As Object)
    If Not cn Is Nothing Then cn.BeginTrans
End Sub

' Commit a transaction.
Public Sub CommitTransAdo(ByVal cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then cn.CommitTrans
    On Error GoTo 0
End Sub

' Roll back a transaction.
Public Sub RollbackTransAdo(ByVal cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then cn.RollbackTrans
    On Error GoTo 0
End Sub

' Create ADO command from connection + SQL.
Public Function CreateCommandAdo(ByVal cn As Object, ByVal sql As String) As Object
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    Set CreateCommandAdo = cmd
End Function

' Execute command (no records) and return records affected.
Public Function ExecuteUpdateAdo(ByVal cmd As Object) As Long
    Dim ra As Long
    cmd.Execute ra, , adExecuteNoRecords
    ExecuteUpdateAdo = ra
End Function

' Execute command and return first column of first row (or Null).
Public Function ExecuteScalarAdo(ByVal cmd As Object) As Variant
    Dim rs As Object
    Dim xe As XError
    On Error GoTo TCError

    Set rs = cmd.Execute

    If (rs Is Nothing) Or (rs.EOF And rs.BOF) Then
        ExecuteScalarAdo = Null
    Else
        ExecuteScalarAdo = rs.Fields(0).Value
    End If

    CloseObj rs
    Exit Function

TCError:
    ' Preserve error, cleanup, then rethrow
    Set xe = ToXError(Err)

    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Open disconnected recordset (client-side static, read-only).
Public Function ExecuteQueryAdo(ByVal cmd As Object, Optional ByVal disconnect As Boolean = True) As Object
    Dim rs As Object
    Dim xe As XError
    On Error GoTo TCError

    Set rs = CreateObject("ADODB.Recordset")

    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly

    If disconnect Then Set rs.ActiveConnection = Nothing
    Set ExecuteQueryAdo = rs
    Exit Function

TCError:
    ' Preserve error, cleanup, then rethrow
    Set xe = ToXError(Err)

    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Add SMALLINT parameter.
Public Sub ParamInt2Ado(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adSmallInt, name, value
End Sub

' Add INT parameter.
Public Sub ParamInt4Ado(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adInteger, name, value
End Sub

' Add BIGINT parameter.
Public Sub ParamInt8Ado(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adBigInt, name, value
End Sub

' Add BIT/Boolean parameter.
Public Sub ParamBoolAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adBoolean, name, value
End Sub

' Add Single parameter.
Public Sub ParamFloatAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adSingle, name, value
End Sub

' Add Double parameter.
Public Sub ParamDoubleAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adDouble, name, value
End Sub

' Add Currency parameter.
Public Sub ParamCurrencyAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adCurrency, name, value
End Sub

' Add DECIMAL/NUMERIC parameter.
Public Sub ParamDecimalAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adNumeric, name, value
End Sub

' Add DECIMAL/NUMERIC parameter with precision & scale.
Public Sub ParamDecimalPSAdo(ByVal cmd As Object, ByVal name As String, _
                            ByVal value As Variant, ByVal precision As Byte, ByVal numScale As Byte)
    Dim p As Object
    Set p = cmd.CreateParameter(name, adNumeric, adParamInput)

    p.Precision = precision
    p.NumericScale = numScale

    If IsNull(value) Or IsEmpty(value) Then
        p.Value = Null
    Else
        p.Value = value
    End If

    cmd.Parameters.Append p
End Sub

' Add GUID parameter.
Public Sub ParamGuidAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adGUID, name, value
End Sub

' Add CHAR(n) parameter.
Public Sub ParamCharAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant, ByVal size As Long)
    AddParam cmd, adChar, name, value, size
End Sub

' Add NCHAR(n) parameter.
Public Sub ParamNCharAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant, ByVal size As Long)
    AddParam cmd, adWChar, name, value, size
End Sub

' Add VARCHAR(n) parameter.
Public Sub ParamVarcharAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant, ByVal size As Long)
    AddParam cmd, adVarChar, name, value, size
End Sub

' Add NVARCHAR(n) parameter.
Public Sub ParamNVarcharAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant, ByVal size As Long)
    AddParam cmd, adVarWChar, name, value, size
End Sub

' Add VARCHAR(MAX) parameter.
Public Sub ParamVarcharMaxAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adLongVarChar, name, value, -1
End Sub

' Add NVARCHAR(MAX) parameter.
Public Sub ParamNVarcharMaxAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adLongVarWChar, name, value, -1
End Sub

' Adds a VARCHAR LIKE parameter for non-Unicode string searches
Public Sub ParamLikeAdo( _
    ByVal cmd As Object, _
    ByVal name As String, _
    ByVal value As Variant, _
    Optional ByVal maxLikeSize As Long = 255, _
    Optional ByVal dbType As XDbType = XDbType.Db_SQLServer _
)
    DoParamLikeAdo cmd, adVarChar, name, value, maxLikeSize, dbType
End Sub

' Adds a NVARCHAR LIKE parameter for Unicode string searches
Public Sub ParamNLikeAdo( _
    ByVal cmd As Object, _
    ByVal name As String, _
    ByVal value As Variant, _
    Optional ByVal maxLikeSize As Long = 255, _
    Optional ByVal dbType As XDbType = XDbType.Db_SQLServer _
)
    DoParamLikeAdo cmd, adVarWChar, name, value, maxLikeSize, dbType
End Sub

' Core helper that normalizes, truncates, and applies LIKE formatting before adding the parameter
Private Sub DoParamLikeAdo( _
    ByVal cmd As Object, _
    ByVal adDataType As Long, _
    ByVal name As String, _
    ByVal value As Variant, _
    ByVal maxLikeSize As Long, _
    ByVal dbType As XDbType _
)
    If IsNull(value) Or IsEmpty(value) Then
        AddParam cmd, adDataType, name, Null, 1
        Exit Sub
    End If

    Dim str As String: str = CStr(value)

    If maxLikeSize > 0 And Len(str) > maxLikeSize Then
        str = Left$(str, maxLikeSize)
    End If

    Dim likeParamValue As Variant
    likeParamValue = ToLikeParamValue(str, dbType)

    AddParam cmd, adDataType, name, likeParamValue, Len(CStr(likeParamValue))
End Sub

' Add DATETIME/TIMESTAMP parameter.
Public Sub ParamDateTimeAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adDBTimeStamp, name, value
End Sub

' Add DATE parameter.
Public Sub ParamDateAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adDBDate, name, value
End Sub

' Add TIME parameter.
Public Sub ParamTimeAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adDBTime, name, value
End Sub

' Add VARBINARY parameter.
Public Sub ParamVarBinaryAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant, ByVal size As Long)
    AddParam cmd, adVarBinary, name, value, size
End Sub

' Add VARBINARY(MAX) parameter.
Public Sub ParamVarBinaryMaxAdo(ByVal cmd As Object, ByVal name As String, ByVal value As Variant)
    AddParam cmd, adLongVarBinary, name, value, -1
End Sub

' Create and append a parameter safely.
Private Sub AddParam(ByVal cmd As Object, ByVal adDataType As Long, ByVal name As String, ByVal value As Variant, Optional ByVal size As Long = 0)
    Dim p As Object

    If size > 0 Then
        Set p = cmd.CreateParameter(name, adDataType, adParamInput, size)
    Else
        Set p = cmd.CreateParameter(name, adDataType, adParamInput)
    End If

    If IsNull(value) Or IsEmpty(value) Then
        p.Value = Null
    Else
        p.Value = value
    End If

    cmd.Parameters.Append p
End Sub

' Convert current record in recordset into a Scripting.Dictionary.
Public Function RecordToDictAdo(ByVal rs As Object) As Object
    Dim d As Object: Set d = NewDictionary()
    Dim i As Long

    For i = 0 To rs.Fields.Count - 1
        d(rs.Fields(i).Name) = rs.Fields(i).Value
    Next

    Set RecordToDictAdo = d
End Function

' Convert a 2-column ADO Recordset (Value, DisplayName) into XDropdownOptions
Private Function ToDropdownOptionsAdo(ByVal rs As Object) As XDropdownOptions
    Dim opts As XDropdownOptions
    Set opts = New XDropdownOptions

    If (rs Is Nothing) Then
        Set ToDropdownOptionsAdo = opts
        Exit Function
    End If

    If rs.Fields.Count <> 2 Then
        XRaise "XAdoUtil.ToDropdownOptionsAdo", "Recordset must have only 2 columns (Value, DisplayName)."
    End If

    ' Read rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            opts.Add rs.Fields(0).Value, Nz(rs.Fields(1).Value, vbEmptyString)
            rs.MoveNext
        Loop
    End If

    Set ToDropdownOptionsAdo = opts
End Function

' Execute an ADO Command and convert the 2-column result into XDropdownOptions
Public Function ExecuteDropdownOptionsAdo(ByVal cmd As Object) As XDropdownOptions
    Dim rs As Object
    Dim xe As XError
    On Error GoTo TCError

    Set rs = cmd.Execute
    Set ExecuteDropdownOptionsAdo = ToDropdownOptionsAdo(rs)

    CloseObj rs
    Exit Function

TCError:
    ' Preserve error, cleanup, then rethrow
    Set xe = ToXError(Err)

    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Execute a SQL query and convert the 2-column result into XDropdownOptions
Public Function ExecuteDropdownOptionsSqlAdo(ByVal connAdo As Object, ByVal sql As String) As XDropdownOptions
    Dim cmd As Object
    Dim xe As XError
    On Error GoTo TCError

    Set cmd = CreateCommandAdo(connAdo, sql)
    Set ExecuteDropdownOptionsSqlAdo = ExecuteDropdownOptionsAdo(cmd)

    CloseObj cmd
    Exit Function

TCError:
    ' Preserve error, cleanup, then rethrow
    Set xe = ToXError(Err)

    CloseObj cmd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function
