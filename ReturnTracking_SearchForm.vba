Option Compare Database
Option Explicit

' ADODB Objects
Private connAdo As Object
Private rsAdo As Object
Private cmdAdo As Object

' Get ADO connection
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

'Initialize connection and run initial search
Private Sub Form_Load()
    ' Init Custom Form Properties
    InitCustomAccessForm Me
    
    ' Call DoSearch to load records
    DoSearch
End Sub

'Validate inputs, execute search, and bind results to the subform
Public Sub DoSearch(Optional ByVal orderByAdo As String = "")
    ' Close prior ADODB objects
    CloseObj rsAdo
    CloseObj cmdAdo

    ' Input States
    Dim stApplicantId As XInputState: Set stApplicantId = GetInt4(Me.txtApplicantID)
    Dim stCtName As XInputState: Set stCtName = GetString(Me.txtCtName)

    ' State Collection
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stApplicantId, stCtName
    
    ' Try/Catch/Finally
    On Error GoTo TCError

    ' Execute Search
    SearchReturnTrackingAdo GetConn(), states.allValid, stApplicantId.value, stCtName.value, cmdAdo, rsAdo

    ' Sorting
    rsAdo.Sort = IIf(orderByAdo <> "", orderByAdo, "ReturnDate1 DESC")

    ' Set the Recordset to the SubForm (Datasheet)
    Set Me.ReturnTracking_Datasheet.Form.Recordset = rsAdo

    Exit Sub

TCError:
    MsgBox "Search failed: " & err.Description, vbCritical
End Sub

'Open the edit form in add mode
Private Sub btnAddNew_Click()
    DoCmd.OpenForm "ReturnTracking_EditForm", , , , acFormAdd
End Sub

'Run search using current criteria
Private Sub btnSearch_Click()
    DoSearch
End Sub

'Release ADO objects when the form closes
Private Sub Form_Close()
    ' Clean up
    CloseObj rsAdo
    CloseObj cmdAdo
    CloseObj connAdo
End Sub
