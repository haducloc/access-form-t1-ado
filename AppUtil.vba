Option Compare Database
Option Explicit

'Build and return the SQL Server connection string used by this app
Public Function GetDbConString() As String
    Dim cs As XConnStr
    Set cs = New XConnStr

    cs.Add "Provider", "MSOLEDBSQL"
    cs.Add "Data Source", "localhost"
    cs.Add "Initial Catalog", "AccessDB"
    cs.Add "Integrated Security", "SSPI"
    cs.Add "Encrypt", "False"
    cs.Add "TrustServerCertificate", "True"

    GetDbConString = cs.Build()
End Function


'Return an open ADODB.Connection, creating/opening it if needed
Public Function GetConnection(ByRef cn As Object) As Object
    On Error GoTo TCError

    If cn Is Nothing Then
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionString = GetDbConString()
    End If

    If cn.state = 0 Then cn.Open

    Set GetConnection = cn
    Exit Function

TCError:
    MsgBox "Database Connection Error: " & Err.Description, vbCritical
    Set GetConnection = Nothing
End Function
