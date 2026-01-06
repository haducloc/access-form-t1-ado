Option Compare Database
Option Explicit

' Returns True if a form object exists in the project (regardless of whether it's open).
Public Function FormExists(ByVal formName As String) As Boolean
    On Error Resume Next
    Dim ao As AccessObject: Set ao = CurrentProject.AllForms(formName)
    FormExists = (Err.Number = 0)
    Err.Clear: On Error GoTo 0
End Function

' Returns the Form object if the form is loaded/open; otherwise returns Nothing.
Public Function FormLoaded(ByVal formName As String) As Form
    On Error Resume Next
    Set FormLoaded = Forms(formName)
    On Error GoTo 0
End Function

' Returns True if a control exists on an OPEN form.
Public Function ControlExists(ByVal formName As String, ByVal ControlName As String) As Boolean
    On Error Resume Next
    Dim ctl As Control
    Set ctl = Forms(formName).Controls(ControlName)
    ControlExists = (Err.Number = 0)
    Err.Clear: On Error GoTo 0
End Function

' Returns the Control object if it exists on the given (open) form; otherwise returns Nothing.
Public Function GetControl(ByVal frm As Form, ByVal ControlName As String) As Control
    On Error Resume Next
    Set GetControl = frm.Controls(ControlName)
    On Error GoTo 0
End Function

' Initializes a datasheet form and enforces Datasheet DefaultView
Public Sub InitCustomDatasheetForm(ByVal frm As Form, Optional ByVal isModal As Boolean = False)
    AssertTrue frm.DefaultView = acDefViewDatasheet, _
           "XFormUtil.InitCustomAccessForm", "The DefaultView of " & frm.name & " must be Datasheet Form."
    frm.Modal = isModal
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = False
    frm.AllowFilters = False
    
    frm.NavigationButtons = True
    frm.RecordSelectors = False
End Sub

' Initializes an custom form and enforces Single Form DefaultView
Public Sub InitCustomAccessForm(ByVal frm As Form, Optional ByVal isModal As Boolean = False)
    AssertTrue frm.DefaultView = acDefViewSingle, _
           "XFormUtil.InitCustomAccessForm", "The DefaultView of " & frm.name & " must be Single Form."
    frm.Modal = isModal
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = False
    frm.AllowFilters = False
    
    frm.NavigationButtons = False
    frm.RecordSelectors = False
End Sub

' Invokes a Public method on a loaded form by name.
Public Sub InvokeFormMethod(ByVal formName As String, ByVal methodName As String)
    Dim frm As Form
    Set frm = FormLoaded(formName)
    If frm Is Nothing Then Exit Sub

    ' Invoke the method on the form.
    CallByName frm, methodName, VbMethod
End Sub

' Invokes a Public method with one argument on a loaded form by name.
Public Sub InvokeFormMethod1(ByVal formName As String, ByVal methodName As String, ByVal argument1 As Variant)
    Dim frm As Form
    Set frm = FormLoaded(formName)
    If frm Is Nothing Then Exit Sub

    ' Invoke the method on the form.
    CallByName frm, methodName, VbMethod, argument1
End Sub

' Invokes a Public method with two arguments on a loaded form by name.
Public Sub InvokeFormMethod2(ByVal formName As String, ByVal methodName As String, ByVal argument1 As Variant, ByVal argument2 As Variant)
    Dim frm As Form
    Set frm = FormLoaded(formName)
    If frm Is Nothing Then Exit Sub

    ' Invoke the method on the form.
    CallByName frm, methodName, VbMethod, argument1, argument2
End Sub

' Invokes a Public method with three arguments on a loaded form by name.
Public Sub InvokeFormMethod3(ByVal formName As String, ByVal methodName As String, ByVal argument1 As Variant, ByVal argument2 As Variant, ByVal argument3 As Variant)
    Dim frm As Form
    Set frm = FormLoaded(formName)
    If frm Is Nothing Then Exit Sub

    ' Invoke the method on the form.
    CallByName frm, methodName, VbMethod, argument1, argument2, argument3
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ToRsOrderByAdo
' Purpose   : Converts Access OrderBy string to ADO Recordset.Sort format.
'---------------------------------------------------------------------------------------
Public Function ToRsOrderByAdo(ByVal datasheetFormOrderBy As String) As String
    Static reQual As Object, reWs As Object
    Dim parts() As String
    Dim i As Long
    Dim s As String, out As String
    Dim fieldTok As String, dirTok As String, inner As String

    ' Create regex engines only once to save memory
    If reQual Is Nothing Then
        Set reQual = CreateObject("VBScript.RegExp")
        reQual.IgnoreCase = True
        reQual.Global = False
        ' Matches Table or [Table] followed by a dot
        reQual.Pattern = "^\s*(\[[^\]]+\]|\w+)\."

        Set reWs = CreateObject("VBScript.RegExp")
        reWs.Global = True
        ' Matches multiple spaces
        reWs.Pattern = "\s+"
    End If

    ' Return empty if input is null or blank
    datasheetFormOrderBy = Trim$(datasheetFormOrderBy & "")
    If Len(datasheetFormOrderBy) = 0 Then Exit Function

    ' Split into individual sort columns
    parts = Split(datasheetFormOrderBy, ",")

    For i = LBound(parts) To UBound(parts)
        s = Trim$(parts(i))
        If Len(s) = 0 Then GoTo NextPart

        ' Remove table or form prefixes
        Do While reQual.Test(s)
            s = Trim$(reQual.Replace(s, ""))
        Loop

        ' Clean up extra spaces
        s = Trim$(reWs.Replace(s, " "))
        If Len(s) = 0 Then GoTo NextPart

        ' Extract Direction from the end of the string
        dirTok = vbEmptyString
        If Len(s) >= 4 Then
            If UCase$(Right$(s, 4)) = " ASC" Then
                dirTok = "ASC"
                fieldTok = Trim$(Left$(s, Len(s) - 4))
            ElseIf Len(s) >= 5 And UCase$(Right$(s, 5)) = " DESC" Then
                dirTok = "DESC"
                fieldTok = Trim$(Left$(s, Len(s) - 5))
            Else
                fieldTok = s
            End If
        Else
            fieldTok = s
        End If

        ' Check for field names wrapped in brackets
        If Left$(fieldTok, 1) = "[" And Right$(fieldTok, 1) = "]" Then
            inner = Mid$(fieldTok, 2, Len(fieldTok) - 2)
            ' Strip brackets if field name is a single word
            If InStr(1, inner, " ") = 0 Then
                fieldTok = inner
            Else
                ' Keep brackets if field name contains spaces
                fieldTok = "[" & inner & "]"
            End If
        Else
            ' Force brackets if name has spaces but is missing them
            If InStr(1, fieldTok, " ") > 0 Then
                fieldTok = "[" & fieldTok & "]"
            End If
        End If

        ' Combine field name and sort direction
        s = fieldTok
        If Len(dirTok) > 0 Then s = s & " " & dirTok

        ' Build the final comma separated result
        If Len(out) > 0 Then out = out & ", "
        out = out & s

NextPart:
    Next i

    ToRsOrderByAdo = out
End Function
