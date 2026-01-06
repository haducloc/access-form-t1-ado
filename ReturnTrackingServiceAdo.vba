Option Compare Database
Option Explicit

'Insert a ReturnTracking row using a parameterized ADO command (positional ? params)
Public Sub InsertReturnTrackingAdo(ByVal connAdo As Object, _
    ByVal applicantId As Long, _
    ByVal soundex As Variant, _
    ByVal ctName As Variant, _
    ByVal internetSource As Variant, _
    ByVal returnDate1 As Variant, _
    ByVal returnDate2 As Variant, _
    ByVal returnDate3 As Variant, _
    ByVal isCompleted As Variant, _
    ByVal comments As Variant _
)
    Dim cmd As Object
    Dim xe As XError

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Parameterized Command
    Set cmd = CreateCommandAdo(connAdo, _
        "INSERT INTO ReturnTracking " & _
        "(ApplicantID, Soundex, CtName, InternetSource, ReturnDate1, ReturnDate2, ReturnDate3, IsCompleted, Comments) " & _
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")

    ' Register parameters in the correct order for ? placeholders (type/Null handling matters)
    ParamInt4Ado cmd, "@applicantId", applicantId
    ParamVarcharAdo cmd, "@soundex", soundex, 50
    ParamVarcharAdo cmd, "@ctName", ctName, 50
    ParamBoolAdo cmd, "@internetSource", internetSource
    ParamDateAdo cmd, "@returnDate1", returnDate1
    ParamDateAdo cmd, "@returnDate2", returnDate2
    ParamDateAdo cmd, "@returnDate3", returnDate3
    ParamBoolAdo cmd, "@isCompleted", isCompleted
    ParamVarcharAdo cmd, "@comments", comments, 255

    ' Execute
    ExecuteUpdateAdo cmd

    GoTo TCFinally

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)

    GoTo TCFinallyRethrow

TCFinally:
    CloseObj cmd
    Exit Sub

TCFinallyRethrow:
    CloseObj cmd
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub


'Update ReturnTracking by ApplicantID (ApplicantID parameter must be last to match WHERE ?)
Public Sub UpdateReturnTrackingAdo(ByVal connAdo As Object, _
    ByVal applicantId As Long, _
    ByVal soundex As Variant, _
    ByVal ctName As Variant, _
    ByVal internetSource As Variant, _
    ByVal returnDate1 As Variant, _
    ByVal returnDate2 As Variant, _
    ByVal returnDate3 As Variant, _
    ByVal isCompleted As Variant, _
    ByVal comments As Variant _
)
    Dim cmd As Object
    Dim xe As XError

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Parameterized Command
    Set cmd = CreateCommandAdo(connAdo, _
        "UPDATE ReturnTracking " & _
        "SET Soundex = ?, CtName = ?, InternetSource = ?" & _
        ", ReturnDate1 = ?, ReturnDate2 = ?, ReturnDate3 = ?, IsCompleted = ?, Comments = ? " & _
        "WHERE ApplicantID = ?")

    ' Register parameters in the correct order for ? placeholders (type/Null handling matters)
    ParamVarcharAdo cmd, "@soundex", soundex, 50
    ParamVarcharAdo cmd, "@ctName", ctName, 50
    ParamBoolAdo cmd, "@internetSource", internetSource
    ParamDateAdo cmd, "@returnDate1", returnDate1
    ParamDateAdo cmd, "@returnDate2", returnDate2
    ParamDateAdo cmd, "@returnDate3", returnDate3
    ParamBoolAdo cmd, "@isCompleted", isCompleted
    ParamVarcharAdo cmd, "@comments", comments, 255
    ParamInt4Ado cmd, "@applicantId", applicantId

    ' Execute
    ExecuteUpdateAdo cmd

    GoTo TCFinally

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)

    GoTo TCFinallyRethrow

TCFinally:
    CloseObj cmd
    Exit Sub

TCFinallyRethrow:
    CloseObj cmd
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub


'Delete ReturnTracking by ApplicantID using a parameterized ADO command
Public Sub DeleteReturnTrackingAdo(ByVal connAdo As Object, ByVal applicantId As Long)
    Dim cmd As Object
    Dim xe As XError

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Parameterized Command
    Set cmd = CreateCommandAdo(connAdo, "DELETE FROM ReturnTracking WHERE ApplicantID = ?")

    ' Register parameter to match the single ? placeholder
    ParamInt4Ado cmd, "@applicantId", applicantId

    ' Execute
    ExecuteUpdateAdo cmd

    GoTo TCFinally

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)

    GoTo TCFinallyRethrow

TCFinally:
    CloseObj cmd
    Exit Sub

TCFinallyRethrow:
    CloseObj cmd
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub


'Fetch a single ReturnTracking row by ApplicantID and return it as a Dictionary (or Nothing)
Public Function GetReturnTrackingByIdAdo(ByVal connAdo As Object, ByVal applicantId As Long) As Object
    Dim cmd As Object
    Dim rs As Object
    Dim xe As XError

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Parameterized Command
    Set cmd = CreateCommandAdo(connAdo, _
        "SELECT * FROM ReturnTracking WHERE ApplicantID = ?")

    ' Register parameter to match the single ? placeholder
    ParamInt4Ado cmd, "@applicantId", applicantId

    ' Execute Query
    Set rs = ExecuteQueryAdo(cmd, True)

    If (rs Is Nothing) Or rs.EOF Then
        Set GetReturnTrackingByIdAdo = Nothing
    Else
        Set GetReturnTrackingByIdAdo = RecordToDictAdo(rs)
    End If

    GoTo TCFinally

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)

    GoTo TCFinallyRethrow

TCFinally:
    CloseObj rs
    CloseObj cmd
    Exit Function

TCFinallyRethrow:
    CloseObj rs
    CloseObj cmd
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function


'Search ReturnTracking with optional ApplicantID and CtName filters using NULL-aware predicates
Public Sub SearchReturnTrackingAdo( _
    ByVal connAdo As Object, ByVal allValid As Boolean, ByVal applicantId As Variant, ByVal ctName As Variant, _
    ByRef cmd As Object, ByRef rs As Object)

    Dim xe As XError

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Ensure outputs start clean
    Set rs = Nothing
    Set cmd = Nothing

    If Not allValid Then
        ' No valid filters: return no rows
        Set cmd = CreateCommandAdo(connAdo, "SELECT * FROM ReturnTracking WHERE 1 <> 1")
    Else
        ' Parameterized Command
        Dim sql As String
        sql = "SELECT * FROM ReturnTracking " & _
              "WHERE (? IS NULL OR ApplicantID = ?) " & _
              "AND (? IS NULL OR CtName LIKE ?)"

        Set cmd = CreateCommandAdo(connAdo, sql)

        ' (? IS NULL OR ApplicantID = ?)
        ParamInt4Ado cmd, "@p1", applicantId
        ParamInt4Ado cmd, "@p2", applicantId

        ' (? IS NULL OR CtName LIKE ?)

        ' ctName column length is 50.
        ' For the ctName LIKE parameter, use a size >= 50; 255 is sufficient.
        ParamLikeAdo cmd, "@p3", ctName, 255, Db_SQLServer
        ParamLikeAdo cmd, "@p4", ctName, 255, Db_SQLServer
    End If

    ' Execute Query
    Set rs = ExecuteQueryAdo(cmd, True)
    Exit Sub

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)

    CloseObj rs
    CloseObj cmd

    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub
