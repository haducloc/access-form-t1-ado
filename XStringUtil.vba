Option Compare Database
Option Explicit

Public Const vbEmptyString As String = ""

'Trim a value to String; return Null if value is Null or trims to empty.
Public Function TrimToNull(ByVal val As Variant) As Variant
    If IsNull(val) Then
        TrimToNull = Null
        Exit Function
    End If

    Dim s As String
    s = Trim$(CStr(val))

    If Len(s) = 0 Then
        TrimToNull = Null
    Else
        TrimToNull = s
    End If
End Function

'Trim a value to String; return "" if value is Null or trims to empty.
Public Function TrimToEmpty(ByVal val As Variant) As String
    If IsNull(val) Then
        TrimToEmpty = vbEmptyString
    Else
        TrimToEmpty = Trim$(CStr(val))
    End If
End Function
