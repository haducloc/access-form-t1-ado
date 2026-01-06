Option Compare Database
Option Explicit

'Build and return the SQL Server connection string used by this app
Public Function GetDbConString(Optional ByVal timeoutSeconds As Long = 300) As String
    Dim cs As XConnStr
    Set cs = New XConnStr

    cs.Add "Provider", "MSOLEDBSQL"
    cs.Add "Data Source", "localhost"
    cs.Add "Initial Catalog", "AccessDB"
    cs.Add "Integrated Security", "SSPI"
    cs.Add "Encrypt", "False"
    cs.Add "TrustServerCertificate", "True"
    cs.Add "Connect Timeout", CStr(timeoutSeconds)

    ' SQL Server Authentication
    ' cs.Add "User ID", "dbuser"
    ' cs.Add "Password", "dbpassword"
    ' Delete the line: Integrated Security

    GetDbConString = cs.Build()
End Function

'Return an open ADODB.Connection, creating/opening it if needed
Public Function GetConnection(ByRef cn As Object, Optional ByVal timeoutSeconds As Long = 300) As Object
    Dim xe As XError
    On Error GoTo TCError

    If cn Is Nothing Then
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionString = GetDbConString(timeoutSeconds)
        ' cn.ConnectionTimeout = timeoutSeconds
    End If

    If cn.State = 0 Then cn.Open

    Set GetConnection = cn
    Exit Function

TCError:
    Set xe = ToXError(Err)

    CloseObj cn
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function
