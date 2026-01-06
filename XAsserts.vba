Option Compare Database
Option Explicit

' Returns True if the given Variant contains a meaningful value
Private Function HasValue(ByVal v As Variant) As Boolean
    ' Null check
    If IsNull(v) Then Exit Function

    ' uninitialized Variant
    If IsEmpty(v) Then Exit Function

    ' String
    If VarType(v) = vbString Then
        HasValue = Len(Trim$(v)) > 0
        Exit Function
    End If
    
    HasValue = True
End Function

' Raises an error if the given Variant does not contain a value; otherwise returns it
Public Function AssertHasValue(ByVal v As Variant, ByVal source As String, ByVal message As String) As Variant
    If Not HasValue(v) Then
        XRaise source, message
    End If
    AssertHasValue = v
End Function

' Raises an error if the given object is nothing; otherwise returns it
Public Function AssertNotNothing(ByVal v As Object, ByVal source As String, ByVal message As String) As Object
    If v Is Nothing Then
        XRaise source, message
    End If
    Set AssertNotNothing = v
End Function

' Raises an error if the given Boolean expression evaluates to False
Public Sub AssertTrue(ByVal booleanExpression As Boolean, ByVal source As String, ByVal message As String)
    If Not booleanExpression Then
        XRaise source, message
    End If
End Sub
