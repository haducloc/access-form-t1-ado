' ===== Class Module: XDropdownOptions =====
Option Compare Database
Option Explicit

Private options As Collection

'Initialize class members
Private Sub Class_Initialize()
    Set options = New Collection
End Sub

'Add an empty (Null/"") option as the first entry
Public Sub AddEmptyFirst()
    Add Null, vbEmptyString
End Sub

'Add a (value, displayName) option pair
Public Sub Add(ByVal Value As Variant, ByVal displayName As String)
    Dim v As Variant
    Dim d As String

    If IsNull(Value) Then
        v = Null
    ElseIf VarType(Value) = vbString Then
        v = Replace(Value, ";", ".")
    Else
        v = Value
    End If

    d = Replace(displayName, ";", ".")
    options.Add ValueToken(v) & ";" & TextToken(d)
End Sub


' Add options to the given combobox
Public Sub ToValueList(ByVal cbo As ComboBox)
    Dim opt As Variant
    For Each opt In options
        cbo.AddItem opt
    Next opt
End Sub


'Convert a value to its Value List token representation
Private Function ValueToken(ByVal v As Variant) As String
    If IsNull(v) Then
        ValueToken = """" & """"          ' -> ""
    ElseIf IsNumeric(v) Then
        ValueToken = CStr(v)
    Else
        ValueToken = TextToken(CStr(v))
    End If
End Function

'Convert a string to a quoted Value List token with escaped quotes
Private Function TextToken(ByVal s As String) As String
    s = Replace(s, """", """""")
    TextToken = """" & s & """"
End Function
