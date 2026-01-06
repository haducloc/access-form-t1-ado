Option Compare Database
Option Explicit

' ADO Objects
Private connAdo As Object

' Store applicantId
Private applicantId As Variant

' Get ADO connection
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

'Open connection and decide whether this form is Add or Edit mode
Private Sub Form_Load()
    ' Init Custom Form Properties plus modal
    InitCustomAccessForm Me, True

    ' Init dropdowns
    InitDropdowns
        
    ' Parse applicantId
    Dim hasApplicantId As Boolean
    hasApplicantId = ParseInt4(Me.OpenArgs, applicantId)
    
    Me.btnDelete.Enabled = Not IsNull(applicantId)
    Me.txtApplicantID.Enabled = IsNull(applicantId)
    
    ' Load existing record
    If Not IsNull(applicantId) Then
        LoadReturnTracking applicantId
        Exit Sub
    End If
End Sub

' Init dropdowns
Private Sub InitDropdowns()
    ' cboCompleted
    Dim opts As XDropdownOptions
    Set opts = ExecuteDropdownOptionsSqlAdo(GetConn(), "SELECT Value, DisplayName FROM YesNoDS")
    
    InitDropdownVL2 Me.cboCompleted
    opts.ToValueList Me.cboCompleted

    ' Other Dropdowns
End Sub

'Load a ReturnTracking record by ApplicantID and bind values to controls
Public Sub LoadReturnTracking(ByVal applicantId As Long)
    Dim dict As Object

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Get record as Dictionary (Nothing if not found)
    Set dict = GetReturnTrackingByIdAdo(GetConn(), applicantId)

    If dict Is Nothing Then
        MsgBox "Record not found for ApplicantID " & applicantId, vbInformation
        DoCmd.Close acForm, Me.name, acSaveNo
    Else
        ' Bind fields to controls
        Me.txtApplicantID = dict("ApplicantID")
        Me.txtSoundex = dict("Soundex")
        Me.txtCtName = dict("CtName")
        Me.chkInternetSource = dict("InternetSource")
        Me.dtReturnDate1 = dict("ReturnDate1")
        Me.dtReturnDate2 = dict("ReturnDate2")
        Me.dtReturnDate3 = dict("ReturnDate3")
        Me.cboCompleted = dict("IsCompleted")
        Me.txtComments = dict("Comments")
    End If

    Exit Sub

TCError:
    MsgBox "Failed to load record: " & err.Description, vbCritical
End Sub

'Validate inputs and save the record via INSERT (add) or UPDATE (edit)
Private Sub btnSave_Click()
    ' Input States
    Dim stApplicantId As XInputState: Set stApplicantId = GetInt4(Me.txtApplicantID, True)
    Dim stSoundex As XInputState: Set stSoundex = GetString(Me.txtSoundex)
    Dim stCtName As XInputState: Set stCtName = GetString(Me.txtCtName, True)
    Dim stInternetSource As XInputState: Set stInternetSource = GetBool(Me.chkInternetSource, True)

    Dim stReturnDate1 As XInputState: Set stReturnDate1 = GetDate(Me.dtReturnDate1, True)
    Dim stReturnDate2 As XInputState: Set stReturnDate2 = GetDate(Me.dtReturnDate2)
    Dim stReturnDate3 As XInputState: Set stReturnDate3 = GetDate(Me.dtReturnDate3)

    Dim stIsCompleted As XInputState: Set stIsCompleted = GetBool(Me.cboCompleted)
    Dim stComments As XInputState: Set stComments = GetString(Me.txtComments)

    ' State Collection
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stApplicantId, stSoundex, stCtName, stInternetSource, _
                     stReturnDate1, stReturnDate2, stReturnDate3, stIsCompleted, stComments

    ' If Invalid
    If Not states.allValid Then
        MsgBox "Please fix errors:" & vbCrLf & vbCrLf & states.ToErrorString, vbExclamation
        Exit Sub
    End If

    ' Try/Catch/Finally
    On Error GoTo TCError
    
    ' Save (insert or update)
    If Not IsNull(applicantId) Then
        UpdateReturnTrackingAdo GetConn(), stApplicantId.value, stSoundex.value, stCtName.value, _
                             stInternetSource.value, stReturnDate1.value, stReturnDate2.value, _
                             stReturnDate3.value, stIsCompleted.value, stComments.value
    Else
        InsertReturnTrackingAdo GetConn(), stApplicantId.value, stSoundex.value, stCtName.value, _
                             stInternetSource.value, stReturnDate1.value, stReturnDate2.value, _
                             stReturnDate3.value, stIsCompleted.value, stComments.value
    End If

    MsgBox "Record saved successfully.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub
    
TCError:
    MsgBox "Failed to save Record: " & err.Description, vbCritical
    
End Sub

'Confirm and delete the current record by ApplicantID
Private Sub btnDelete_Click()
    ' Input States
    Dim stApplicantId As XInputState: Set stApplicantId = GetInt4(Me.txtApplicantID, True)

    ' Confirm
    If MsgBox("Are you sure you want to delete this record: " & stApplicantId.value & "?", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If

    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Delete the record
    DeleteReturnTrackingAdo GetConn(), stApplicantId.value

    MsgBox "Record deleted.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to delete record: " & err.Description, vbCritical
    
End Sub

' Cleanup the connection and refresh the main form search results
Private Sub Form_Close()
    CloseObj connAdo

    InvokeFormMethod "ReturnTracking_SearchForm", "DoSearch"
End Sub
