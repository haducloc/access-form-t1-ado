Option Compare Database
Option Explicit

' Declare datasheet delegate object
Private datasheetDelegate As XDatasheetDelegate

' Initializes the datasheet delegate when the form loads
Private Sub Form_Load()
    ' Init Custom Datasheet Form
    InitCustomDatasheetForm Me
    
    ' Init Datasheet Form Delegate
    Set datasheetDelegate = New XDatasheetDelegate

    datasheetDelegate.Init _
        datasheetForm:=Me, _
        parentForm:="ReturnTracking_SearchForm", _
        reloadMethod:="DoSearch", _
        editForm:="ReturnTracking_EditForm", _
        idField:="ApplicantID", _
        timerIntervalMs:=100
End Sub

' Forwards the timer event to the datasheet delegate
Private Sub Form_Timer()
    datasheetDelegate.Form_Timer
End Sub

' Forwards the double-click event to the datasheet delegate
Private Sub Form_DblClick(Cancel As Integer)
    datasheetDelegate.Form_DblClick Cancel
End Sub

' Forwards the ApplyFilter event to the datasheet delegate
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    datasheetDelegate.Form_ApplyFilter Cancel, ApplyType
End Sub

' Forwards the form error event to the datasheet delegate
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    datasheetDelegate.Form_Error DataErr, Response
End Sub
