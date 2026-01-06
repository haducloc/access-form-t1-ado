Option Compare Database
Option Explicit

Public Const X_ERR_NUMBER As Long = vbObjectError + 1000

'Raise a standardized X error
Public Sub XRaise(ByVal source As String, ByVal message As String)
    Err.Raise X_ERR_NUMBER, source, message
End Sub

' Close and release late-bound object safely.
Public Sub CloseObj(ByRef obj As Object, Optional ByVal closeMethod As String = "Close")
    On Error Resume Next
    If Not obj Is Nothing Then
        CallByName obj, closeMethod, VbMethod
        Set obj = Nothing
    End If
    On Error GoTo 0
End Sub