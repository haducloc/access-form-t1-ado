Option Compare Database
Option Explicit

'Parse a Boolean from common text tokens; Null/empty -> Null (True), invalid -> False.
Public Function ParseBool(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant
    v = TrimToNull(val)

    If IsNull(v) Then
        ParseBool = True
        Exit Function
    End If

    Dim s As String: s = UCase$(CStr(v))
    Select Case s
        Case "TRUE", "YES", "T", "Y", "1", "-1"
            result = True
            ParseBool = True

        Case "FALSE", "NO", "F", "N", "0"
            result = False
            ParseBool = True

        Case Else
            ParseBool = False
    End Select
End Function

'Parse a SmallInt-sized whole number; Null/empty -> Null (True), invalid/out-of-range -> False.
Public Function ParseInt2(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseInt2 = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseInt2 = False: Exit Function

    Dim d As Double: d = CDbl(s)
    If d <> Fix(d) Then ParseInt2 = False: Exit Function

    If d < -32768# Or d > 32767# Then ParseInt2 = False: Exit Function

    result = CInt(Fix(d))
    ParseInt2 = True
End Function

'Parse an Int-sized whole number; Null/empty -> Null (True), invalid/out-of-range -> False.
Public Function ParseInt4(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseInt4 = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseInt4 = False: Exit Function

    Dim d As Double: d = CDbl(s)
    If d <> Fix(d) Then ParseInt4 = False: Exit Function

    If d < -2147483648# Or d > 2147483647# Then ParseInt4 = False: Exit Function

    result = CLng(Fix(d))
    ParseInt4 = True
End Function

'Parse a BigInt-sized whole number (64-bit only); Null/empty -> Null (True), invalid/out-of-range/platform -> False.
Public Function ParseInt8(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseInt8 = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseInt8 = False: Exit Function

    Dim d As Double: d = CDbl(s)
    If d <> Fix(d) Then ParseInt8 = False: Exit Function

#If VBA7 And Win64 Then
    On Error GoTo TCError
    result = CLngLng(Fix(d))
    ParseInt8 = True
    Exit Function
TCError:
    result = Null
    ParseInt8 = False
#Else
    ParseInt8 = False
#End If
End Function

'Parse a Single-precision number; Null/empty -> Null (True), invalid -> False.
Public Function ParseFloat(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseFloat = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseFloat = False: Exit Function

    On Error GoTo TCError
    result = CSng(s)
    ParseFloat = True
    Exit Function
TCError:
    result = Null
    ParseFloat = False
End Function

'Parse a Double-precision number; Null/empty -> Null (True), invalid -> False.
Public Function ParseDouble(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseDouble = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseDouble = False: Exit Function

    On Error GoTo TCError
    result = CDbl(s)
    ParseDouble = True
    Exit Function
TCError:
    result = Null
    ParseDouble = False
End Function

'Parse a Decimal (stored in a Variant); Null/empty -> Null (True), invalid -> False.
Public Function ParseDecimal(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseDecimal = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsNumeric(s) Then ParseDecimal = False: Exit Function

    On Error GoTo TCError
    result = CDec(s)
    ParseDecimal = True
    Exit Function
TCError:
    result = Null
    ParseDecimal = False
End Function

'Parse a Date-only value (DateValue); Null/empty -> Null (True), invalid -> False.
Public Function ParseDate(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseDate = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsDate(s) Then ParseDate = False: Exit Function

    On Error GoTo TCError
    result = DateValue(s)
    ParseDate = True
    Exit Function
TCError:
    result = Null
    ParseDate = False
End Function

'Parse a Time-only value (TimeValue); Null/empty -> Null (True), invalid -> False.
Public Function ParseTime(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseTime = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsDate(s) Then ParseTime = False: Exit Function

    On Error GoTo TCError
    result = TimeValue(s)
    ParseTime = True
    Exit Function
TCError:
    result = Null
    ParseTime = False
End Function

'Parse a Date+Time value (CDate); Null/empty -> Null (True), invalid -> False.
Public Function ParseDateTime(ByVal val As Variant, ByRef result As Variant) As Boolean
    result = Null

    Dim v As Variant: v = TrimToNull(val)
    If IsNull(v) Then ParseDateTime = True: Exit Function

    Dim s As String: s = CStr(v)
    If Not IsDate(s) Then ParseDateTime = False: Exit Function

    On Error GoTo TCError
    result = CDate(s)
    ParseDateTime = True
    Exit Function
TCError:
    result = Null
    ParseDateTime = False
End Function
