' ===== Class Module: XStateCollection =====
Option Compare Database
Option Explicit

Private states As Collection

'Initialize the internal collection
Private Sub Class_Initialize()
    Set states = New Collection
End Sub

'Add a single XInputState to the collection
Public Sub Add(ByVal state As XInputState)
    states.Add state
End Sub

'Add multiple XInputState objects with type checks
Public Sub AddStates(ParamArray items() As Variant)
    Dim i As Long
    Dim s As XInputState

    For i = LBound(items) To UBound(items)

        If Not IsObject(items(i)) Then
            XRaise "XStateCollection.AddStates", "Item " & (i + 1) & " is not an object."
        End If

        If Not TypeOf items(i) Is XInputState Then
            XRaise "XStateCollection.AddStates", "Item " & (i + 1) & " is not an XInputState."
        End If

        Set s = items(i)
        states.Add s
    Next i
End Sub

'Return True only if every state is valid
Public Function AllValid() As Boolean
    Dim s As XInputState
    
    For Each s In states
        If Not s.IsValid Then
            AllValid = False
            Exit Function
        End If
    Next s

    AllValid = True
End Function

Public Function AllNull() As Boolean
    If states.count = 0 Then
        AllNull = False
        Exit Function
    End If
    
    AllNull = (CountNull() = states.count)
End Function

Public Function CountNull() As Long
    Dim count As Long
    Dim s As Variant

    For Each s In states
        If IsNull(s.value) Then
            count = count + 1
        End If
    Next s

    CountNull = count
End Function

'Build a numbered list of validation error messages
Public Function ToErrorString() As String
    Dim i As Long
    Dim msg As String
    Dim s As XInputState
    
    For i = 1 To states.Count
        Set s = states(i)
        If Not s.IsValid Then
            msg = msg & "* " & s.ErrorMessage & vbCrLf
        End If
    Next i
    
    ToErrorString = RTrim$(msg)
End Function
