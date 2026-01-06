Option Compare Database
Option Explicit

' Create a Yes/No XDropdownOptions list, with configurable ordering.
Public Function YesNoDropdownOptions(Optional ByVal yesThenNo As Boolean = True) As XDropdownOptions

    Dim opts As XDropdownOptions
    Set opts = New XDropdownOptions

    If yesThenNo Then
        opts.Add True, "Yes"
        opts.Add False, "No"
    Else
        opts.Add False, "No"
        opts.Add True, "Yes"
    End If

    Set YesNoDropdownOptions = opts
End Function

' Configure a ComboBox for a 2-column Value List with hidden value column.
Public Sub InitDropdownVL2(ByVal cbo As Access.ComboBox, Optional ByVal firstEmpty As Boolean = True, Optional ByVal displayWidthInches As String = "2in")
    With cbo
        .RowSourceType = "Value List"
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0in;" & displayWidthInches
        .LimitToList = True
        .AllowValueListEdits = False
    End With
    
    ' "";"" at index 0
    If firstEmpty Then
        cbo.AddItem """" & """" & ";" & """" & """", 0
    End If
End Sub
