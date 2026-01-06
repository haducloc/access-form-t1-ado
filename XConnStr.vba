' ===== Class Module: XConnStr =====
Option Compare Database
Option Explicit

Private parts As Collection

'Initialize class members
Private Sub Class_Initialize()
    Set parts = New Collection
End Sub

'Add a key/value pair to the connection string (Null/Empty becomes "Key=;")
Public Sub Add(ByVal key As String, ByVal value As Variant)
    If IsNull(value) Or IsEmpty(value) Or (VarType(value) = vbString And Len(value) = 0) Then
        parts.Add Trim$(key) & "=;"
    Else
        parts.Add Trim$(key) & "=" & CStr(value) & ";"
    End If
End Sub


'Concatenate all stored parts into the final connection string
Public Function Build() As String
    Dim s As String
    Dim item As Variant

    For Each item In parts
        s = s & CStr(item)
    Next

    Build = s
End Function
