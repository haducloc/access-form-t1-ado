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

Public Function ToXError(ByVal err As VBA.ErrObject) As XError
    Dim xe As XError
    Set xe = New XError

    xe.ErrNum = err.Number
    xe.ErrDesc = err.Description
    xe.ErrSrc = err.source

    Set ToXError = xe
End Function

Public Function NewDictionary() As Object
  Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function